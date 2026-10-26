"""
xlwings Lite scripts pulling data from the tigzig yfin-h API for tickers
listed in the Tickers table (MASTER sheet):
  refresh_calendar_xlw  -> CALENDAR sheet: earnings dates, dividends
  refresh_prices_xlw    -> PRICES sheet:   adjusted closing prices over a date range,
                            one row per ticker, one column per date (wide/pivoted)

Shared plumbing (sheet setup, spinner, status, table header/formalize,
summary) lives in standalone helpers below -- both scripts call them
directly rather than duplicating the Excel choreography. Both scripts are
the SAME loop shape (one row per ticker, written incrementally batch by
batch via _run_row_per_item_pull) so both stay live the whole way through:
  - Calendar: data_columns = Ticker/ExDividendDate/DividendDate/etc.
  - Prices: data_columns = Ticker + one column per calendar day in the
    Settings StartDate..EndDate range + Status. The date columns are
    built directly from the Settings dates (every calendar day, in
    order) -- NOT discovered from what the API actually returns -- so
    the table header is known before the first fetch and every batch can
    write immediately. A date with no trading data (weekend, holiday)
    just ends up as a blank column across every ticker; that's
    deliberate, not a bug.

Source tables (looked up by TABLE NAME across all sheets -- doesn't care
which sheet they live on or what row they start at, so renaming a sheet or
inserting rows above a table will not break it):
  Settings  -> Name/Value  (BatchSize; StartDate/EndDate for prices only)
  Tickers   -> Ticker

Sheet layout (both output sheets, same shape): a "top zone" (info line,
spinner, status, summary) sits above the data table, sized to how many
spinner rows THIS run uses (see _layout_rows) -- so the gap between the
spinner block and the table stays tight regardless of spinner row count.

Live updating: @script writes and print() are queued and only appear once
the script finishes. `await book.flush()` applies them straight away, which
is why this script calls it after every write, and why the table fills in
batch by batch instead of arriving all at once at the end. It needs an
async script. See https://lite.xlwings.org/async-api.

Date columns: a date-like string such as "2026-08-10" is parsed into a real
date by Excel as it is written, which is what shows as "####" in a narrow
column. Setting the number format to text afterwards does not undo it, it
just displays the already-parsed serial number as digits. Set the format on
the empty cell first, then write, which is why the date columns here are
pre-formatted before the header goes down. It applies to every column
holding date-like text: CALENDAR's ExDividendDate, DividendDate and
NextEarningsDate all need it, the same as RefreshedAt.

Paste this into main.py in the xlwings Lite task pane. Two runnable
scripts: refresh_calendar_xlw, refresh_prices_xlw.
No requirements.txt entries needed (pandas is Pyodide-provided; pyodide.http
is built in).
"""

import asyncio
import datetime as dt
import time
from typing import Any, Optional, Tuple

import pandas as pd
import xlwings as xw
from pyodide.http import pyfetch
from xlwings import script

CALENDAR_SHEET_NAME = "CALENDAR"
PRICES_SHEET_NAME = "PRICES"
CALENDAR_BASE_URL = "https://yfin-h.tigzig.com/v1/get-calendar/?tickers="
PRICES_BASE_URL = "https://yfin-h.tigzig.com/v1/get-adj-close/"
API_MAX_TICKERS_PER_CALL = 25  # the API accepts at most 25 tickers per call, every endpoint

# --- Top-zone layout -------------------------------------------------------
INFO_ROW = 1
SPINNER_HEADER_ROW = 2
SPINNER_START_ROW = 3
DEFAULT_SPINNER_ROWS = 5

SPINNER_CHARS = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
SPINNER_WIDTH = 5    # glyphs concatenated per row
SPINNER_INTERVAL = 0.25  # seconds between ticks -- own clock, independent of batch progress

PREFORMAT_ROWS = 2000  # how many rows ahead to pre-format RefreshedAt as text

# Mirrors the final summary onto MASTER so the result is visible next to the
# Settings table. The output sheet's own copy is the primary one.
MASTER_STATUS_COL = "F"
MASTER_STATUS_ROW = 5


def _layout_rows(spinner_rows: int) -> Tuple[int, int]:
    """
    (status_row, table_header_row) sized to how many spinner rows this run
    uses, so the gap after a short spinner block doesn't leave a lot of
    dead blank space. Sizing the layout up front is what keeps it tight;
    the row delete at the end is only a cosmetic trim on top of it.
    """
    status_row = SPINNER_START_ROW + spinner_rows + 1  # 1 blank row after the spinner
    table_header_row = status_row + 5  # 4 summary rows + 1 blank row
    return status_row, table_header_row


def _col_letter(n: int) -> str:
    """1-indexed column number -> Excel column letter(s)."""
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def find_table_in_workbook(book: xw.Book, table_name: str) -> Tuple[Optional[xw.Sheet], Any]:
    """
    Searches all sheets for a table and returns both the sheet and table objects.
    Returns: (xw.Sheet, xw.Table) or (None, None) if not found.
    """
    for sheet in book.sheets:
        if table_name in sheet.tables:
            return sheet, sheet.tables[table_name]
    return None, None


def _chunk(items, size):
    """Split a list into chunks of `size`."""
    for i in range(0, len(items), size):
        yield items[i:i + size]


def _clean_tickers(values) -> list:
    """Filter a raw Ticker column list down to non-blank, non-NaN strings."""
    tickers = []
    for v in values:
        if v is None:
            continue
        s = str(v).strip()
        if not s or s.lower() == "nan":
            continue
        tickers.append(s)
    return tickers


def _fmt_elapsed(seconds: float) -> str:
    """Seconds under a minute; minutes (1 decimal) above."""
    if seconds < 60:
        return f"{seconds:.1f}s"
    return f"{seconds / 60:.1f} min"


def _parse_date_setting(value) -> str:
    """Settings dates may arrive as a real datetime (normal case, Excel date
    cell) or as text -- handle both, always return 'YYYY-MM-DD'."""
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    for fmt in ("%d %B %Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise Exception(f"Could not parse date setting: {value!r}")


def _spinner_row_text(tick: int, row_offset: int) -> str:
    """SPINNER_WIDTH braille glyphs concatenated, phase-shifted by `row_offset`."""
    n = len(SPINNER_CHARS)
    return "".join(SPINNER_CHARS[(tick + row_offset + k) % n] for k in range(SPINNER_WIDTH))


async def _retry_excel_op(func_, description: str, max_retries: int = 3, retry_delay: float = 1.5):
    """
    Run a zero-arg callable, retrying on failure. Covers transient Excel
    errors like 'InvalidOperationInCellEditMode' (user mid cell-edit
    elsewhere) which typically clear themselves within a couple seconds.
    Re-raises the last error if every attempt fails.
    """
    last_err = None
    for attempt in range(1, max_retries + 1):
        try:
            return func_()
        except Exception as e:
            last_err = e
            print(f"    {description}: attempt {attempt}/{max_retries} failed ({e})")
            if attempt < max_retries:
                await asyncio.sleep(retry_delay)
    raise last_err or RuntimeError(f"{description}: all {max_retries} attempts failed")


def _read_settings(book: xw.Book) -> Tuple[xw.Sheet, dict]:
    settings_sheet, settings_table = find_table_in_workbook(book, "Settings")
    if settings_table is None or settings_sheet is None:
        raise Exception("Settings table not found in workbook.")
    settings_df = settings_table.range.options(pd.DataFrame, index=False).value
    return settings_sheet, dict(zip(settings_df["Name"], settings_df["Value"]))


def _read_tickers(book: xw.Book) -> list:
    tickers_sheet, tickers_table = find_table_in_workbook(book, "Tickers")
    if tickers_table is None:
        raise Exception("Tickers table not found in workbook.")
    tickers_df = tickers_table.range.options(pd.DataFrame, index=False).value
    return sorted(_clean_tickers(tickers_df["Ticker"].tolist()))


def _cleanup_old_status_sheets(book: xw.Book):
    for leftover_name in ("STATUS", "STATUS_TEST"):
        for s in book.sheets:
            if s.name == leftover_name:
                s.delete()
                break


async def _setup_output_sheet(book: xw.Book, name: str) -> xw.Sheet:
    def _setup():
        _cleanup_old_status_sheets(book)
        for s in book.sheets:
            if s.name == name:
                s.delete()
        ns = book.sheets.add(name=name)
        ns.activate()
        return ns
    try:
        ns = await _retry_excel_op(_setup, "sheet setup", max_retries=5, retry_delay=2.0)
        await book.flush()
        return ns
    except Exception as e:
        raise Exception(f"Could not set up '{name}' after retries -- Excel may be busy: {e}")


async def _write_info_line(book: xw.Book, sheet: xw.Sheet, text: str):
    def _write():
        sheet[f"A{INFO_ROW}"].value = text
    await _retry_excel_op(_write, "info line write", max_retries=3, retry_delay=1.0)
    await book.flush()


async def _set_status(book: xw.Book, sheet: xw.Sheet, status_row: int, text: str):
    def _write():
        sheet[f"A{status_row}"].value = "Status"
        sheet[f"B{status_row}"].value = text
    try:
        await _retry_excel_op(_write, "status write", max_retries=2, retry_delay=1.0)
        await book.flush()
    except Exception as e:
        print(f"  (status write failed, non-fatal: {e})")


async def _set_spinner(book: xw.Book, sheet: xw.Sheet, tick: int, rows: int):
    def _write():
        for r in range(rows):
            row = SPINNER_START_ROW + r
            sheet[f"A{row}"].value = _spinner_row_text(tick, r * SPINNER_WIDTH)
    try:
        await _retry_excel_op(_write, "spinner write", max_retries=1, retry_delay=0.5)
        await book.flush()
    except Exception:
        pass  # cosmetic only


async def _spin_loop(book: xw.Book, sheet: xw.Sheet, rows: int):
    tick = 0
    while True:
        await _set_spinner(book, sheet, tick, rows)
        tick += 1
        await asyncio.sleep(SPINNER_INTERVAL)


async def _start_spinner(book: xw.Book, sheet: xw.Sheet, rows: int):
    def _write_header():
        sheet[f"A{SPINNER_HEADER_ROW}"].value = "Working..."
    await _retry_excel_op(_write_header, "spinner header write", max_retries=2, retry_delay=1.0)
    await book.flush()
    return asyncio.create_task(_spin_loop(book, sheet, rows))


async def _stop_spinner(spin_task, book: xw.Book, sheet: xw.Sheet, rows: int):
    spin_task.cancel()
    try:
        await spin_task
    except asyncio.CancelledError:
        pass

    def _clear():
        sheet[f"A{SPINNER_HEADER_ROW}"].value = None
        for r in range(rows):
            sheet[f"A{SPINNER_START_ROW + r}"].value = None
    try:
        await _retry_excel_op(_clear, "spinner clear", max_retries=1, retry_delay=0.5)
        await book.flush()
    except Exception:
        pass


async def _remove_spinner_gap(book: xw.Book, sheet: xw.Sheet, rows: int):
    """
    COSMETIC ONLY, call LAST -- after the table and summary are already
    finalized (and after any other post-processing, e.g. number formatting,
    since a successful delete here physically shifts every row below it up,
    which would invalidate any row number computed earlier in the caller).
    Pass shift="up" explicitly rather than relying on the default. A failure
    here is non-fatal: the layout is already sized tight by _layout_rows, so
    this only trims the last blank rows.
    """
    first_row = SPINNER_HEADER_ROW
    last_row = SPINNER_START_ROW + rows - 1

    def _delete():
        sheet.range(f"{first_row}:{last_row}").delete(shift="up")
    try:
        await _retry_excel_op(_delete, "spinner row delete (shift=up)", max_retries=2, retry_delay=1.0)
        await book.flush()
    except Exception as e:
        print(f"  (spinner row delete failed again, non-fatal -- blank rows remain: {e})")


async def _write_table_header(book: xw.Book, sheet: xw.Sheet, table_header_row: int, columns: list):
    def _write():
        empty_df = pd.DataFrame(columns=columns)
        sheet[f"A{table_header_row}"].options(index=False).value = empty_df
        header_range = sheet[f"A{table_header_row}"].resize(1, len(columns))
        header_range.color = "#4472C4"
        header_range.font.color = "#FFFFFF"
        header_range.font.bold = True
    await _retry_excel_op(_write, "table header write", max_retries=5, retry_delay=2.0)
    await book.flush()


async def _preformat_text_column(book: xw.Book, sheet: xw.Sheet, start_row: int, col_letter: str, n_rows: int):
    """
    Force a column range to TEXT format before any value lands in it, so a
    date-looking string never gets auto-parsed into a real date cell by
    Excel on entry (see the Date columns note in the module docstring).
    Must run BEFORE the range's first write. Two call shapes, both used below:
      - header cell only (n_rows=1, start_row=table_header_row): for a
        column whose HEADER LABEL is date-like text (e.g. PRICES' date
        columns are literally named "2026-07-01") but whose DATA must
        stay numeric (so later number formatting, e.g. "0.0", still works).
      - header + all data rows (start_row=table_header_row, n_rows=
        PREFORMAT_ROWS + 1): for a column that's genuinely text end to
        end (e.g. RefreshedAt, or CALENDAR's ExDividendDate/etc.).
    """
    def _fmt():
        rng = sheet[f"{col_letter}{start_row}"].resize(n_rows, 1)
        rng.number_format = "@"
    try:
        await _retry_excel_op(_fmt, "pre-format text column", max_retries=2, retry_delay=1.0)
        await book.flush()
    except Exception as e:
        print(f"  (pre-format text column failed, non-fatal: {e})")


async def _format_number_range(book: xw.Book, sheet: xw.Sheet, start_row: int, start_col: int,
                                n_rows: int, n_cols: int, number_format: str, description: str):
    if n_rows <= 0 or n_cols <= 0:
        return
    def _fmt():
        rng = sheet[f"{_col_letter(start_col)}{start_row}"].resize(n_rows, n_cols)
        rng.number_format = number_format
    try:
        await _retry_excel_op(_fmt, description, max_retries=2, retry_delay=1.0)
        await book.flush()
    except Exception as e:
        print(f"  ({description} failed, non-fatal: {e})")


async def _formalize_table(book: xw.Book, sheet: xw.Sheet, table_header_row: int, n_rows: int, n_cols: int):
    if n_rows <= 0:
        return
    table_range = sheet[f"A{table_header_row}"].resize(n_rows + 1, n_cols)

    def _add_table():
        sheet.tables.add(source=table_range)
    try:
        await _retry_excel_op(_add_table, "table creation", max_retries=2, retry_delay=1.5)
        await book.flush()
    except Exception as e:
        print(f"Table formatting failed (data is still on the sheet, just not a formal Table): {e}")


async def _write_final_summary(book: xw.Book, sheet: xw.Sheet, status_row: int,
                                settings_sheet: xw.Sheet, summary_rows: list):
    summary_df = pd.DataFrame(summary_rows)

    def _write():
        sheet[f"A{status_row}"].options(index=False, header=False).value = summary_df
    try:
        await _retry_excel_op(_write, "summary write", max_retries=2, retry_delay=1.0)
        await book.flush()
    except Exception as e:
        print(f"  (summary write failed, non-fatal: {e})")

    def _write_master():
        settings_sheet[f"{MASTER_STATUS_COL}{MASTER_STATUS_ROW}"].options(index=False, header=False).value = summary_df
    try:
        await _retry_excel_op(_write_master, "MASTER status mirror", max_retries=2, retry_delay=1.0)
        await book.flush()
    except Exception as e:
        print(f"  (MASTER status mirror failed, non-fatal -- the output sheet's copy is the reliable one: {e})")


# =============================================================================
# CALENDAR -- one row per ticker, written incrementally batch by batch
# =============================================================================

def _build_calendar_rows(batch_json: dict, tickers_in_batch: list, fetch_ok: bool = True) -> list:
    rows = []
    for ticker in tickers_in_batch:
        if not fetch_ok:
            rows.append({
                "Ticker": ticker, "ExDividendDate": None, "DividendDate": None,
                "NextEarningsDate": None, "Status": "Fetch failed",
            })
            continue

        data = batch_json.get(ticker)
        if data is None:
            rows.append({
                "Ticker": ticker, "ExDividendDate": None, "DividendDate": None,
                "NextEarningsDate": None, "Status": "No calendar data",
            })
            continue

        error = data.get("error")
        earnings_dates = data.get("earnings_dates") or []
        next_earnings = ", ".join(str(d) for d in earnings_dates) if earnings_dates else None

        rows.append({
            "Ticker": ticker,
            "ExDividendDate": data.get("ex_dividend_date"),
            "DividendDate": data.get("dividend_date"),
            "NextEarningsDate": next_earnings,
            "Status": "No calendar data" if error else "OK",
        })
    return rows


async def _run_row_per_item_pull(book: xw.Book, settings_sheet: xw.Sheet, output_sheet_name: str,
                                  data_columns: list, info_text: str, batches: list,
                                  fetch_batch_rows, refresh_time_str: str,
                                  spinner_rows: int = DEFAULT_SPINNER_ROWS,
                                  text_safe_columns: Optional[list] = None,
                                  header_only_text_columns: Optional[list] = None) -> Tuple[xw.Sheet, int, int]:
    """
    Shared engine for both scripts: sets up the output sheet, writes the
    header, then fetches and writes one batch at a time.

    fetch_batch_rows(batch) is an async callable that must NEVER raise. On a
    fetch error it returns placeholder rows instead, so failures show up in
    the table itself and not only in the summary.

    The two date-column arguments are easy to swap, and doing so is the one
    way to get this wrong:
      text_safe_columns         header AND data are date-like text, e.g.
                                CALENDAR's ExDividendDate. Formatted as text
                                before anything is written. RefreshedAt is
                                included automatically.
      header_only_text_columns  header is date-like text but the data is
                                numeric, e.g. PRICES columns named
                                "2026-07-01" holding prices. Only the header
                                cell is formatted, so the number format still
                                applies to the values below it.

    Returns (sheet, total_written, table_header_row) for caller-side
    post-processing such as number formatting or the row trim.
    """
    run_start = time.time()
    status_row, table_header_row = _layout_rows(spinner_rows)

    new_sheet = await _setup_output_sheet(book, output_sheet_name)
    await _write_info_line(book, new_sheet, info_text)

    spin_task = await _start_spinner(book, new_sheet, spinner_rows)
    await _set_status(book, new_sheet, status_row, f"Batch is running... {len(batches)} batches")

    all_columns = data_columns + ["RefreshedAt"]

    # Pre-format BEFORE writing the header, so a date-like header LABEL
    # (PRICES) or a date-like column end to end (CALENDAR/RefreshedAt)
    # never gets auto-parsed into a real date cell in the first place.
    for col_name in (header_only_text_columns or []):
        col_idx = all_columns.index(col_name) + 1
        await _preformat_text_column(book, new_sheet, table_header_row, _col_letter(col_idx), 1)

    for col_name in list(text_safe_columns or []) + ["RefreshedAt"]:
        col_idx = all_columns.index(col_name) + 1
        await _preformat_text_column(book, new_sheet, table_header_row, _col_letter(col_idx), PREFORMAT_ROWS + 1)

    await _write_table_header(book, new_sheet, table_header_row, all_columns)

    print(f"'{output_sheet_name}' ready. Writing results batch by batch as they arrive...")

    next_row = table_header_row + 1
    total_written = 0
    write_failed_batches = 0

    for i, batch in enumerate(batches, start=1):
        rows = await fetch_batch_rows(batch)
        batch_df = pd.DataFrame(rows, columns=data_columns)
        batch_df["RefreshedAt"] = refresh_time_str

        write_row = next_row

        def _do_write():
            new_sheet[f"A{write_row}"].options(index=False, header=False).value = batch_df

        try:
            await _retry_excel_op(_do_write, f"batch {i} write", max_retries=3, retry_delay=1.5)
            await book.flush()
            next_row += len(batch_df)
            total_written += len(batch_df)
            print(f"Batch {i}/{len(batches)} OK: {len(rows)} rows written "
                  f"| elapsed {_fmt_elapsed(time.time() - run_start)} | total rows {total_written}")
            await _set_status(book, new_sheet, status_row,
                               f"Batch {i}/{len(batches)} done - {total_written} rows so far "
                               f"({_fmt_elapsed(time.time() - run_start)} elapsed)")
        except Exception:
            write_failed_batches += 1
            print(f"Batch {i}/{len(batches)} FAILED TO WRITE after retries -- rows lost for this batch "
                  f"| elapsed {_fmt_elapsed(time.time() - run_start)}")
            await _set_status(book, new_sheet, status_row,
                               f"Batch {i}/{len(batches)} FAILED TO WRITE "
                               f"({_fmt_elapsed(time.time() - run_start)} elapsed)")

        if i < len(batches):
            await asyncio.sleep(0.3)  # light pacing between calls to the API

    await _stop_spinner(spin_task, book, new_sheet, spinner_rows)
    await _formalize_table(book, new_sheet, table_header_row, total_written, len(all_columns))

    total_elapsed = time.time() - run_start
    batches_ok = len(batches) - write_failed_batches
    print(f"Done. Total time {_fmt_elapsed(total_elapsed)} | Batches {batches_ok}/{len(batches)} done "
          f"| Rows written {total_written}")
    await book.flush()

    status_text = f"Done - {total_written} rows written ({batches_ok}/{len(batches)} batches ok)"
    summary_rows = [
        ("Status", status_text),
        ("Total Time", _fmt_elapsed(total_elapsed)),
        ("Batches", f"{batches_ok}/{len(batches)} done"),
        ("Rows Written", total_written),
    ]
    await _write_final_summary(book, new_sheet, status_row, settings_sheet, summary_rows)

    if write_failed_batches:
        print(f"{write_failed_batches} batch(es) never wrote to Excel -- re-run to fill in.")

    return new_sheet, total_written, table_header_row


@script(name="1. Pull Calendar")
async def refresh_calendar_xlw(book: xw.Book):
    refresh_time_str = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print("Reading Settings + Tickers...")
    settings_sheet, settings = _read_settings(book)
    batch_size = int(settings.get("BatchSize", 10) or 10)
    if batch_size > API_MAX_TICKERS_PER_CALL:
        print(f"  BatchSize {batch_size} exceeds the API's {API_MAX_TICKERS_PER_CALL}-ticker cap -- "
              f"capping this run to {API_MAX_TICKERS_PER_CALL}.")
        batch_size = API_MAX_TICKERS_PER_CALL
    tickers = _read_tickers(book)

    batches = list(_chunk(tickers, batch_size))
    info_text = f"Calendar pull: {len(tickers)} tickers, {len(batches)} batches"
    print(info_text)

    async def _fetch(batch):
        url = CALENDAR_BASE_URL + ",".join(batch)
        try:
            response = await pyfetch(url, method="GET", headers={"Accept": "application/json"})
            if not response.ok:
                print(f"  HTTP error: {response.status}")
                return _build_calendar_rows({}, batch, fetch_ok=False)
            batch_json = await response.json()
            return _build_calendar_rows(batch_json, batch, fetch_ok=True)
        except Exception as e:
            print(f"  request failed: {e}")
            return _build_calendar_rows({}, batch, fetch_ok=False)

    data_columns = ["Ticker", "ExDividendDate", "DividendDate", "NextEarningsDate", "Status"]
    new_sheet, total_written, table_header_row = await _run_row_per_item_pull(
        book, settings_sheet, CALENDAR_SHEET_NAME, data_columns,
        info_text, batches, _fetch, refresh_time_str, spinner_rows=5,
        text_safe_columns=["ExDividendDate", "DividendDate", "NextEarningsDate"])

    await _remove_spinner_gap(book, new_sheet, rows=5)


# =============================================================================
# PRICES -- one row per ticker, one column per date (wide/pivoted), written
# LIVE batch by batch. Date columns come from the Settings dates, not from
# the API response, so the header is known before the first fetch.
# =============================================================================

PRICES_SPINNER_ROWS = 3
PRICE_NUMBER_FORMAT = "0.0"  # one decimal place, everywhere


def _build_price_rows(batch_json: dict, tickers_in_batch: list, date_cols: list, fetch_ok: bool = True) -> list:
    rows = []
    for ticker in tickers_in_batch:
        row = {"Ticker": ticker}
        if not fetch_ok:
            for d in date_cols:
                row[d] = None
            row["Status"] = "Fetch failed"
            rows.append(row)
            continue

        any_price = False
        for d in date_cols:
            day = batch_json.get(d)
            v = day.get(ticker) if day else None
            # The API returns the string "Data not available" per date/ticker
            # for a symbol it cannot resolve, mixed in with real numbers from
            # other tickers in the same response. Only accept a numeric value.
            if not isinstance(v, (int, float)):
                v = None
            row[d] = v
            if v is not None:
                any_price = True
        row["Status"] = "OK" if any_price else "No data"
        rows.append(row)
    return rows


@script(name="2. Pull Prices")
async def refresh_prices_xlw(book: xw.Book):
    refresh_time_str = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print("Reading Settings + Tickers...")
    settings_sheet, settings = _read_settings(book)

    batch_size = int(settings.get("BatchSize", 10) or 10)
    if batch_size > API_MAX_TICKERS_PER_CALL:
        print(f"  BatchSize {batch_size} exceeds the API's {API_MAX_TICKERS_PER_CALL}-ticker cap -- "
              f"capping this run to {API_MAX_TICKERS_PER_CALL}.")
        batch_size = API_MAX_TICKERS_PER_CALL

    start_date_raw = settings.get("StartDate (For Prices Only)")
    end_date_raw = settings.get("EndDate (For Prices Only)")
    if start_date_raw is None or end_date_raw is None:
        raise Exception("Settings needs 'StartDate (For Prices Only)' and 'EndDate (For Prices Only)' rows.")
    start_date = _parse_date_setting(start_date_raw)
    end_date = _parse_date_setting(end_date_raw)
    # The API's end_date is EXCLUSIVE (confirmed by direct probing 2026-08-07:
    # end_date=2026-08-06 returns data only through 2026-08-05). Settings/the
    # displayed table columns treat EndDate as INCLUSIVE (what the user typed
    # is what they expect to see), so request one extra day from the API only
    # -- date_cols (the actual sheet columns) are unaffected.
    api_end_date = (dt.datetime.strptime(end_date, "%Y-%m-%d") + dt.timedelta(days=1)).strftime("%Y-%m-%d")

    date_cols = [d.strftime("%Y-%m-%d") for d in pd.date_range(start=start_date, end=end_date, freq="D")]

    tickers = _read_tickers(book)
    batches = list(_chunk(tickers, batch_size))
    info_text = (f"Prices pull: {len(tickers)} tickers, {start_date} to {end_date} "
                 f"({len(date_cols)} calendar days), {len(batches)} batches")
    print(info_text)

    async def _fetch(batch):
        url = f"{PRICES_BASE_URL}?tickers={','.join(batch)}&start_date={start_date}&end_date={api_end_date}"
        try:
            response = await pyfetch(url, method="GET", headers={"Accept": "application/json"})
            if not response.ok:
                print(f"  HTTP error: {response.status}")
                return _build_price_rows({}, batch, date_cols, fetch_ok=False)
            data = await response.json()
            return _build_price_rows(data, batch, date_cols, fetch_ok=True)
        except Exception as e:
            print(f"  request failed: {e}")
            return _build_price_rows({}, batch, date_cols, fetch_ok=False)

    data_columns = ["Ticker"] + date_cols + ["Status"]
    new_sheet, total_written, table_header_row = await _run_row_per_item_pull(
        book, settings_sheet, PRICES_SHEET_NAME, data_columns,
        info_text, batches, _fetch, refresh_time_str, spinner_rows=PRICES_SPINNER_ROWS,
        header_only_text_columns=date_cols)

    # Price columns start right after Ticker (column B = index 2), one column per date.
    await _format_number_range(book, new_sheet, table_header_row + 1, 2,
                                total_written, len(date_cols), PRICE_NUMBER_FORMAT,
                                "price decimal format")

    await _remove_spinner_gap(book, new_sheet, rows=PRICES_SPINNER_ROWS)
