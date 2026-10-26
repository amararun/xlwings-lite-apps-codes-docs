"""INTELISCAPE-X: AI web scraper in Excel with xlwings Lite.

You list URLs (URL_LIST) and describe the columns you want in plain English (COLUMN_INPUTS).
For each URL, Jina Reader turns the page into Markdown text, and Gemini extracts the rows as
structured JSON that follows a schema built from your column list. Rows go to DATA.

Keys come from environment variables (xlwings Lite > Settings > Environment Variables, scope
Local): GEMINI_API_KEY (required), JINA_API_KEY (optional; Jina Reader works without one at a
lower rate limit). Nothing is saved in the workbook.
"""
import asyncio
import json
import os
import time
from datetime import datetime
from typing import Any, Optional, Tuple

import pandas as pd
import xlwings as xw
from pyodide.http import pyfetch
from xlwings import script

CONTROL = "MASTER"
SETTINGS = {"model": "C16", "delay": "C17", "timeout": "C18", "max_tokens": "C19", "temperature": "C20"}
STATUS_CELL = "A26"
URL_SHEET, COLUMN_SHEET, DATA_SHEET = "URL_LIST", "COLUMN_INPUTS", "DATA"
DONE, SKIP, ERROR, RUNNING = "✅ DONE", "⏭️ SKIP", "❌ ERROR", "⏳ RUNNING"
RETRY_STATUS = {429, 500, 502, 503, 504}
BLUE, RED, HEADER_FILL = "#1F4E79", "#C00000", "#DDEBF7"


# -------------------------
# Small helpers
# -------------------------
def set_status(book: xw.Book, text: str, ok: bool):
    cell = book.sheets[CONTROL][STATUS_CELL]
    cell.value = text
    cell.font.bold = True
    cell.font.size = 12 if ok else 14
    cell.font.color = BLUE if ok else RED


def number(value: Any, default: float) -> float:
    try:
        return float(value) if value not in (None, "") else default
    except (TypeError, ValueError):
        return default


def read_settings(book: xw.Book) -> dict:
    c = book.sheets[CONTROL]
    return {
        "model": str(c[SETTINGS["model"]].value or "").strip() or "gemini-flash-latest",
        "delay": max(0.0, number(c[SETTINGS["delay"]].value, 1)),
        "timeout": max(10.0, number(c[SETTINGS["timeout"]].value, 120)),
        "max_tokens": int(number(c[SETTINGS["max_tokens"]].value, 50000)),
        "temperature": min(2.0, max(0.0, number(c[SETTINGS["temperature"]].value, 0.2))),
    }


def read_columns(book: xw.Book) -> Tuple[list, list, str]:
    """Column names (A3 down), descriptions (B) and the optional instructions in D2."""
    sh = book.sheets[COLUMN_SHEET]
    names = sh["A3"].expand("down").value if sh["A3"].value else []
    names = names if isinstance(names, list) else [names]
    descs = sh[f"B3:B{2 + max(len(names), 1)}"].value
    descs = descs if isinstance(descs, list) else [descs]
    fields, notes = [], []
    for i, n in enumerate(names):
        if n is not None and str(n).strip():
            fields.append(str(n).strip())
            notes.append(str(descs[i]).strip() if i < len(descs) and descs[i] else "")
    return fields, notes, str(sh["D2"].value or "").strip()


def read_urls(book: xw.Book) -> list:
    """(row, url, status) for every URL in URL_LIST from A2 down."""
    sh = book.sheets[URL_SHEET]
    if not sh["A2"].value:
        return []
    urls = sh["A2"].expand("down").value
    urls = urls if isinstance(urls, list) else [urls]
    status = sh[f"B2:B{1 + len(urls)}"].value
    status = status if isinstance(status, list) else [status]
    return [(i + 2, str(u).strip(), status[i] if i < len(status) else None)
            for i, u in enumerate(urls) if u and str(u).strip()]


async def fetch_text(url: str, tries: int = 3, **kwargs) -> Tuple[int, str]:
    """pyfetch with retries on busy / temporary errors. Returns (status, body text)."""
    status, text = 0, ""
    for attempt in range(1, tries + 1):
        r = await pyfetch(url, **kwargs)
        status, text = r.status, await r.text()
        if r.ok or status not in RETRY_STATUS or attempt == tries:
            break
        print(f"      ⏳ HTTP {status}, retrying ({attempt}/{tries - 1}) ...")
        await asyncio.sleep(5 * attempt)
    return status, text


def error_message(text: str) -> str:
    try:
        err = json.loads(text)
        err = err.get("error", err) if isinstance(err, dict) else err
        return str(err.get("message", err) if isinstance(err, dict) else err)[:300]
    except ValueError:
        return text[:300]


# -------------------------
# The two API calls
# -------------------------
async def read_page(url: str, jina_key: Optional[str], timeout: float) -> str:
    """Jina Reader: the page as Markdown text."""
    headers = {"X-Return-Format": "markdown", "X-Timeout": str(int(timeout))}
    if jina_key:
        headers["Authorization"] = f"Bearer {jina_key}"
    status, text = await fetch_text(f"https://r.jina.ai/{url}", headers=headers)
    if status != 200:
        raise RuntimeError(f"Jina Reader HTTP {status}: {error_message(text)}")
    if not text.strip():
        raise RuntimeError("Jina Reader returned an empty page")
    return text


def field_keys(fields: list) -> dict:
    """Column name -> JSON key (lowercase, underscores). Keeps keys unique."""
    keys: dict = {}
    for f in fields:
        k = "".join(ch if ch.isalnum() else "_" for ch in f.lower()).strip("_") or "field"
        base, n = k, 2
        while k in keys.values():
            k, n = f"{base}_{n}", n + 1
        keys[f] = k
    return keys


def build_prompt(page: str, fields: list, notes: list, instructions: str) -> str:
    lines = "\n".join(f"{i + 1}. {f}: {d}" for i, (f, d) in enumerate(zip(fields, notes)))
    rules = f"\nINSTRUCTIONS (follow these strictly, return only items that match):\n{instructions}\n" if instructions else ""
    return f"""You extract structured data from a web page that has been converted to Markdown.
{rules}
For EACH qualifying item on the page, return these fields:
{lines}

Return a JSON array with one object per item. Use an empty string when a field is not on the page.
Do not skip qualifying items.

PAGE:
{page}"""


async def extract_rows(page: str, fields: list, notes: list, instructions: str, key: str, cfg: dict) -> Tuple[list, dict]:
    """Gemini with a response schema: one object per item, every field a string."""
    keys = field_keys(fields)
    schema = {"type": "ARRAY", "items": {"type": "OBJECT",
                                         "properties": {k: {"type": "STRING"} for k in keys.values()},
                                         "propertyOrdering": list(keys.values()),
                                         "required": list(keys.values())}}
    payload = {"contents": [{"parts": [{"text": build_prompt(page, fields, notes, instructions)}]}],
               "generationConfig": {"temperature": cfg["temperature"], "maxOutputTokens": cfg["max_tokens"],
                                    "response_mime_type": "application/json", "response_schema": schema}}
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{cfg['model']}:generateContent"
    status, text = await fetch_text(url, method="POST", headers={"Content-Type": "application/json",
                                                                 "x-goog-api-key": key}, body=json.dumps(payload))
    if status != 200:
        raise RuntimeError(f"Gemini HTTP {status}: {error_message(text)}")
    data = json.loads(text)
    usage = data.get("usageMetadata", {})
    cand = (data.get("candidates") or [{}])[0]
    reason = cand.get("finishReason", "")
    if reason == "MAX_TOKENS":
        raise RuntimeError(f"the answer was cut off at {cfg['max_tokens']:,} output tokens. Raise MASTER "
                           f"{SETTINGS['max_tokens']} or narrow the instructions in COLUMN_INPUTS D2.")
    parts = cand.get("content", {}).get("parts", [])
    raw = "".join(p.get("text", "") for p in parts if not p.get("thought"))
    if not raw.strip():
        raise RuntimeError(f"Gemini returned no content (finish reason {reason or 'unknown'})")
    items = json.loads(raw)
    back = {v: k for k, v in keys.items()}
    rows = [{back.get(k, k): v for k, v in it.items()} for it in items if isinstance(it, dict)]
    return rows, usage


# -------------------------
# Writing results
# -------------------------
def existing_data(book: xw.Book, header: list, reprocess: set) -> Tuple[pd.DataFrame, str]:
    """Rows already in DATA for URLs not processed again in this run (kept across runs)."""
    names = [s.name for s in book.sheets]
    if DATA_SHEET not in names or not book.sheets[DATA_SHEET]["A1"].value:
        return pd.DataFrame(columns=header), ""
    old = book.sheets[DATA_SHEET]["A1"].expand().options(pd.DataFrame, index=False).value
    if list(old.columns) != header:
        return pd.DataFrame(columns=header), " Columns changed, so DATA was started fresh."
    return old[~old["URL"].isin(reprocess)].reset_index(drop=True), ""


def new_sheet(book: xw.Book, name: str, after: xw.Sheet) -> xw.Sheet:
    for sh in book.sheets:
        if sh.name == name:
            sh.delete()
    return book.sheets.add(name=name, after=after)


def write_table(sh: xw.Sheet, cell: str, df: pd.DataFrame, name: Optional[str] = None):
    sh[cell].options(index=False).value = df
    sh[cell].resize(1, df.shape[1]).color = HEADER_FILL
    if len(df):
        sh.tables.add(source=sh[cell].resize(len(df) + 1, df.shape[1]), name=name)


@script(name="1. Scrape URLs", button="[btn_scrape]MASTER!C23")
async def scrape_urls_from_list(book: xw.Book):
    """Scrape every URL in URL_LIST not marked DONE or SKIP, write rows to DATA as you go."""
    start = time.time()
    cfg = read_settings(book)
    gemini_key, jina_key = os.getenv("GEMINI_API_KEY"), os.getenv("JINA_API_KEY")
    if not gemini_key:
        set_status(book, "❌ GEMINI_API_KEY is not set. In xlwings Lite open Settings > Environment Variables, add "
                         "GEMINI_API_KEY with scope Local, then restart xlwings Lite (see Step 1).", ok=False)
        return
    fields, notes, instructions = read_columns(book)
    if not fields:
        set_status(book, f"❌ No columns in {COLUMN_SHEET}. Put column names in A3 down and descriptions in B.", ok=False)
        return
    urls = read_urls(book)
    todo = [(row, u) for row, u, s in urls if s not in (DONE, SKIP)]
    if not todo:
        set_status(book, f"❌ Nothing to scrape: every URL in {URL_SHEET} is marked DONE or SKIP. Clear column B "
                         f"for the ones to run again.", ok=False)
        return
    header = ["URL"] + fields
    kept, note = existing_data(book, header, {u for _, u in todo})
    print(f"▶ {len(todo)} URL(s) to scrape, {len(fields)} columns, model {cfg['model']}, "
          f"Jina key: {'yes' if jina_key else 'no (free rate limit)'}; keeping {len(kept)} earlier rows")

    url_sheet = book.sheets[URL_SHEET]
    data = new_sheet(book, DATA_SHEET, book.sheets[COLUMN_SHEET])
    data["A1"].value = header
    data["A1"].resize(1, len(header)).color = HEADER_FILL
    if len(kept):
        data["A2"].options(index=False, header=False).value = kept
    next_row = len(kept) + 2
    errors, per_url, tokens = [], [], {"in": 0, "out": 0}
    stopped = ""

    for i, (row, url) in enumerate(todo, start=1):
        url_sheet[f"B{row}:D{row}"].value = [RUNNING, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), None]
        set_status(book, f"⏳ URL {i} of {len(todo)}: {url}", ok=True)
        await book.flush()
        t0 = time.time()
        print(f"\n[{i}/{len(todo)}] {url}")
        try:
            page = await read_page(url, jina_key, cfg["timeout"])
            print(f"    page: {len(page):,} characters")
            rows, usage = await extract_rows(page, fields, notes, instructions, gemini_key, cfg)
            tokens["in"] += usage.get("promptTokenCount", 0)
            tokens["out"] += usage.get("candidatesTokenCount", 0) + usage.get("thoughtsTokenCount", 0)
            if not rows:
                raise RuntimeError("no matching items on the page")
            values = [[url] + [r.get(f, "") for f in fields] for r in rows]
            data[f"A{next_row}"].value = values
            next_row += len(values)
            url_sheet[f"B{row}:D{row}"].value = [DONE, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), len(values)]
            print(f"    ✅ {len(values)} rows in {time.time() - t0:.1f}s")
        except Exception as e:
            msg = str(e)
            print(f"    ❌ {msg}")
            errors.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), url, msg])
            if "HTTP 429" in msg:
                # Rate limit or quota still hit after the retries: every later URL would fail too.
                # Stop here and leave this URL and the rest pending, so the next run picks them up.
                url_sheet[f"B{row}:D{row}"].value = [None, None, None]
                stopped = (f" Stopped at URL {i} of {len(todo)}: the API rate limit or quota was reached. "
                           f"Run again later; finished URLs are kept and the rest are still pending.")
                await book.flush()
                break
            url_sheet[f"B{row}:D{row}"].value = [ERROR, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0]
            values = []
        per_url.append({"URL": url, "Rows": len(values), "Seconds": round(time.time() - t0, 1),
                        "Result": "OK" if values else "Error"})
        await book.flush()
        if i < len(todo) and cfg["delay"]:
            await asyncio.sleep(cfg["delay"])

    # Final tables: DATA as a table, VALIDATION (rows per URL, all runs), ERROR_LOG, DASHBOARD
    total_rows = next_row - 2
    if total_rows:
        data.tables.add(source=data["A1"].resize(total_rows + 1, len(header)), name="ExportedData")
    all_rows = pd.concat([kept, pd.DataFrame([{"URL": p["URL"]} for p in per_url for _ in range(p["Rows"])])],
                         ignore_index=True) if total_rows else pd.DataFrame(columns=["URL"])
    status_by_url = {u: s for _, u, s in urls}
    status_by_url.update({p["URL"]: DONE if p["Rows"] else ERROR for p in per_url})
    counts = all_rows.groupby("URL").size() if len(all_rows) else pd.Series(dtype=int)
    validation = pd.DataFrame([{"URL": u, "Status": status_by_url.get(u) or "", "Rows in DATA": int(counts.get(u, 0))}
                               for _, u, _ in urls])
    val = new_sheet(book, "VALIDATION", data)
    val["A1"].value = "Rows in DATA per URL (all runs)"
    val["A1"].font.bold = True
    write_table(val, "A3", validation)
    err = new_sheet(book, "ERROR_LOG", val)
    err["A1"].value = "Errors from the last run"
    err["A1"].font.bold = True
    write_table(err, "A3", pd.DataFrame(errors, columns=["Time", "URL", "Error"]))
    ok = [p for p in per_url if p["Rows"]]
    secs = [p["Seconds"] for p in per_url]
    summary = pd.DataFrame([
        {"Metric": "Run finished", "Value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
        {"Metric": "Total run time (seconds)", "Value": round(time.time() - start, 1)},
        {"Metric": "URLs processed", "Value": len(per_url)},
        {"Metric": "URLs OK", "Value": len(ok)},
        {"Metric": "URLs with errors", "Value": len(per_url) - len(ok)},
        {"Metric": "Rows extracted this run", "Value": sum(p["Rows"] for p in per_url)},
        {"Metric": "Rows in DATA (all runs)", "Value": total_rows},
        {"Metric": "Average seconds per URL", "Value": round(sum(secs) / len(secs), 1) if secs else 0},
        {"Metric": "Gemini input tokens", "Value": tokens["in"]},
        {"Metric": "Gemini output tokens (incl. thinking)", "Value": tokens["out"]},
        {"Metric": "Model", "Value": cfg["model"]},
    ])
    dash = new_sheet(book, "DASHBOARD", err)
    dash["A1"].value = "Last run"
    dash["A1"].font.bold = True
    write_table(dash, "A3", summary)
    write_table(dash, "D3", pd.DataFrame(per_url, columns=["URL", "Rows", "Seconds", "Result"]))

    all_ok = not stopped and len(ok) == len(per_url)
    msg = (f"{'✅' if all_ok else '❌'} {len(ok)} of {len(per_url)} URL(s) OK, "
           f"{sum(p['Rows'] for p in per_url)} rows this run, {total_rows} rows in DATA.{note}"
           + (f" {len(per_url) - len(ok)} error(s): see ERROR_LOG." if len(ok) < len(per_url) else "") + stopped)
    print("\n" + msg)
    set_status(book, msg, ok=all_ok)
    data.activate()
