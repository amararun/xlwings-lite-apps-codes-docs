"""Technical Analysis Reports from Excel with xlwings Lite.

1. Pull Prices & Charts: daily prices from the TIGZIG Yahoo Finance API, weekly bars built from
   them, indicators computed in pandas, and charts drawn with matplotlib onto the CHARTS sheet.
2. Get AI Report: sends the symbol and dates to the TIGZIG Technical Analysis API, where an LLM
   reads the charts and indicators and writes the report. Returns PDF and HTML links.
Both APIs are open (no key). They allow requests from xlwings Lite (CORS origin addin.xlwings.org).
"""
import datetime as dt
import json
from typing import Any, Optional, Tuple
from urllib.parse import quote

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlwings as xw
from pyodide.http import pyfetch
from xlwings import script

CONTROL = "MASTER"
SYMBOL_CELL = "C7"
DATE_CELLS = {"daily_start": "C8", "daily_end": "C9", "weekly_start": "C10", "weekly_end": "C11"}
MODEL_CELL = "C12"         # report model; empty = the API default (Claude Sonnet 5.5)
STATUS_CELL = "A22"
LINK_ROW = 24              # A24 PDF link, A25 HTML link, A26:A27 the raw URLs
PRICES_API = "https://yfin-h.tigzig.com/get-all-prices/"
REPORT_API = "https://ta.tigzig.com/v1/technical-analysis"

BLUE, ORANGE, GREY = "#1F4E79", "#E67E22", "#7F7F7F"
RED = "#C00000"


# -------------------------
# Small helpers
# -------------------------
def set_status(book: xw.Book, text: str, ok: bool):
    cell = book.sheets[CONTROL][STATUS_CELL]
    cell.value = text
    cell.font.bold = True
    cell.font.size = 12 if ok else 14
    cell.font.color = BLUE if ok else RED


def to_date(value: Any) -> Optional[dt.date]:
    """Excel date, serial number or 'YYYY-MM-DD' text -> date. Empty -> None."""
    if value is None or value == "":
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)):
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(value))
    return dt.date.fromisoformat(str(value).strip()[:10])


def read_inputs(book: xw.Book) -> Tuple[str, dict]:
    """Symbol and the four dates. Empty dates: daily = last 6 months, weekly = last 3 years, ending today."""
    c = book.sheets[CONTROL]
    symbol = str(c[SYMBOL_CELL].value or "").strip().upper()
    raw = {k: to_date(c[a].value) for k, a in DATE_CELLS.items()}
    today = dt.date.today()
    daily_end = raw["daily_end"] or today
    weekly_end = raw["weekly_end"] or today
    daily_start = raw["daily_start"] or daily_end - dt.timedelta(days=182)
    weekly_start = raw["weekly_start"] or weekly_end - dt.timedelta(days=3 * 365)
    if daily_start >= daily_end or weekly_start >= weekly_end:
        raise ValueError("each start date must be before its end date")
    return symbol, {"daily_start": daily_start, "daily_end": daily_end,
                    "weekly_start": weekly_start, "weekly_end": weekly_end}


def replace_sheet(book: xw.Book, name: str, after: xw.Sheet) -> xw.Sheet:
    for sheet in book.sheets:
        if sheet.name == name:
            sheet.delete()
    return book.sheets.add(name=name, after=after)


# -------------------------
# Prices
# -------------------------
async def get_prices(symbol: str, start: dt.date, end: dt.date) -> pd.DataFrame:
    url = f"{PRICES_API}?tickers={quote(symbol)}&start_date={start}&end_date={end}"
    r = await pyfetch(url)
    text = await r.text()
    if not r.ok:
        try:
            err = json.loads(text).get("error", {})
            text = err.get("message", text) if isinstance(err, dict) else str(err)
        except ValueError:
            pass
        raise RuntimeError(f"price API HTTP {r.status}: {text[:300]}")
    data = json.loads(text)
    if isinstance(data, dict) and data.get("error"):
        raise RuntimeError(f"price API: {data['error']}")
    rows = [dict(v[symbol], Date=day) for day, v in data.items() if isinstance(v, dict) and symbol in v]
    if not rows:
        raise RuntimeError(f"no prices for {symbol} between {start} and {end}. Check the Yahoo Finance symbol.")
    df = pd.DataFrame(rows)
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.sort_values("Date").reset_index(drop=True)
    return df[["Date", "Open", "High", "Low", "Close", "Volume"]]


def weekly_bars(daily: pd.DataFrame) -> pd.DataFrame:
    """Weeks ending Friday: first open, highest high, lowest low, last close, total volume."""
    w = daily.resample("W-FRI", on="Date").agg(
        {"Open": "first", "High": "max", "Low": "min", "Close": "last", "Volume": "sum"})
    return w.dropna().reset_index()


# -------------------------
# Indicators (plain pandas, standard definitions)
# -------------------------
def add_indicators(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    close, high, low = out["Close"], out["High"], out["Low"]
    out["EMA_12"] = close.ewm(span=12).mean()
    out["EMA_26"] = close.ewm(span=26).mean()
    delta = close.diff()
    gain = delta.clip(lower=0).ewm(alpha=1 / 14).mean()
    loss = (-delta.clip(upper=0)).ewm(alpha=1 / 14).mean()
    out["RSI_14"] = 100 - 100 / (1 + gain / loss)
    out["ROC_14"] = close.pct_change(14) * 100
    out["MACD_12_26"] = out["EMA_12"] - out["EMA_26"]
    out["MACD_SIGNAL_9"] = out["MACD_12_26"].ewm(span=9).mean()
    mid, sd = close.rolling(20).mean(), close.rolling(20).std()
    out["BB_UPPER_20_2"], out["BB_MIDDLE_20"], out["BB_LOWER_20_2"] = mid + 2 * sd, mid, mid - 2 * sd
    hh, ll = high.rolling(14).max(), low.rolling(14).min()
    out["STOCH_K_14"] = (close - ll) / (hh - ll) * 100
    out["STOCH_D_3"] = out["STOCH_K_14"].rolling(3).mean()
    out["WILLIAMS_R_14"] = (hh - close) / (hh - ll) * -100
    return out


# -------------------------
# Charts
# -------------------------
def make_chart(df: pd.DataFrame, symbol: str, frequency: str):
    """Price with EMAs and Bollinger Bands, MACD, RSI and ROC. Blue / orange, no red-green pairs."""
    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(8, 8), height_ratios=[2, 1, 1], sharex=True)
    d = df["Date"]
    width = max((d.iloc[-1] - d.iloc[0]).days / len(df) * 0.8, 0.5)
    ax1.plot(d, df["Close"], color="black", lw=1.4, label="Close")
    ax1.plot(d, df["BB_UPPER_20_2"], color=GREY, ls="--", lw=0.9, label="BB upper / lower")
    ax1.plot(d, df["BB_LOWER_20_2"], color=GREY, ls="--", lw=0.9)
    ax1.plot(d, df["BB_MIDDLE_20"], color=GREY, ls=":", lw=0.9, label="BB middle")
    ax1.plot(d, df["EMA_12"], color=BLUE, lw=1.8, label="EMA 12")
    ax1.plot(d, df["EMA_26"], color=ORANGE, lw=1.8, label="EMA 26")
    vol = ax1.twinx()
    up = df["Close"].diff().fillna(0) >= 0
    vol.bar(d, df["Volume"], width=width, color=np.where(up, "#9DC3E6", "#F5CBA7"), alpha=0.6)
    vol.set_ylim(0, df["Volume"].max() * 4 if df["Volume"].max() > 0 else 1)
    vol.set_yticks([])
    ax1.set_title(f"{symbol} - price, EMAs and Bollinger Bands ({frequency})", fontsize=12, weight="bold")
    ax1.legend(loc="upper left", fontsize=8)
    hist = df["MACD_12_26"] - df["MACD_SIGNAL_9"]
    ax2.bar(d, hist, width=width, color=np.where(hist >= 0, "#9DC3E6", "#F5CBA7"))
    ax2.plot(d, df["MACD_12_26"], color=BLUE, lw=1.3, label="MACD")
    ax2.plot(d, df["MACD_SIGNAL_9"], color=ORANGE, lw=1.3, label="Signal")
    ax2.axhline(0, color=GREY, lw=0.7)
    ax2.set_title(f"MACD (12, 26, 9) - {frequency}", fontsize=10, weight="bold")
    ax2.legend(loc="upper left", fontsize=8)
    ax3.plot(d, df["RSI_14"], color=BLUE, lw=1.3, label="RSI 14")
    for level in (30, 70):
        ax3.axhline(level, color=GREY, ls="--", lw=0.7)
    ax3.set_ylim(0, 100)
    roc = ax3.twinx()
    roc.plot(d, df["ROC_14"], color=ORANGE, lw=1.1, label="ROC 14")
    roc.axhline(0, color=GREY, lw=0.5)
    ax3.set_title(f"RSI and ROC - {frequency}", fontsize=10, weight="bold")
    h1, l1 = ax3.get_legend_handles_labels()
    h2, l2 = roc.get_legend_handles_labels()
    ax3.legend(h1 + h2, l1 + l2, loc="upper left", fontsize=8)
    for ax in (ax1, ax2, ax3):
        ax.grid(True, alpha=0.2)
    fig.autofmt_xdate()
    fig.subplots_adjust(left=0.1, right=0.9, top=0.95, bottom=0.1, hspace=0.35)
    return fig


def write_prices(book: xw.Book, name: str, after: xw.Sheet, df: pd.DataFrame, title: str) -> xw.Sheet:
    sh = replace_sheet(book, name, after)
    sh["A1"].value = title
    sh["A1"].font.bold = True
    out = df.copy()
    out.columns = [c.upper() for c in out.columns]
    sh["A2"].options(index=False).value = out
    sh.tables.add(source=sh["A2"].resize(len(out) + 1, out.shape[1]))
    sh["A3"].resize(len(out), 1).number_format = "dd-mmm-yyyy"
    return sh


@script(name="1. Pull Prices & Charts", button="[btn_tech]MASTER!C14")
async def create_technicals(book: xw.Book):
    """Daily and weekly prices, indicators and charts, all computed inside Excel."""
    try:
        symbol, d = read_inputs(book)
    except ValueError as e:
        set_status(book, f"❌ {e}. Dates are in C8:C11 (empty = defaults).", ok=False)
        return
    if not symbol:
        set_status(book, f"❌ Type a Yahoo Finance symbol in {SYMBOL_CELL}, for example ^NSEI or AAPL.", ok=False)
        return
    print(f"▶ {symbol}: daily {d['daily_start']} to {d['daily_end']}, weekly {d['weekly_start']} to {d['weekly_end']}")
    set_status(book, f"⏳ Pulling prices for {symbol} ...", ok=True)
    await book.flush()
    try:
        daily = add_indicators(await get_prices(symbol, d["daily_start"], d["daily_end"]))
        weekly = add_indicators(weekly_bars(await get_prices(symbol, d["weekly_start"], d["weekly_end"])))
    except Exception as e:
        print(f"❌ {e}")
        set_status(book, f"❌ {e}", ok=False)
        return
    control = book.sheets[CONTROL]
    fmt = "%d %b %Y"
    ds = write_prices(book, "PRICES_DAILY", control, daily,
                      f"Daily prices and indicators: {symbol}, {d['daily_start']:{fmt}} to {d['daily_end']:{fmt}}")
    ws = write_prices(book, "PRICES_WEEKLY", ds, weekly,
                      f"Weekly prices and indicators (weeks ending Friday): {symbol}, "
                      f"{d['weekly_start']:{fmt}} to {d['weekly_end']:{fmt}}")
    charts = replace_sheet(book, "CHARTS", ws)
    charts["A1"].value = f"{symbol}: daily (left) and weekly (right)"
    charts["A1"].font.bold = True
    for df, freq, anchor in ((daily, "Daily", "A3"), (weekly, "Weekly", "L3")):
        fig = make_chart(df, symbol, freq)
        charts.pictures.add(fig, name=f"{freq}Chart", update=True, anchor=charts[anchor])
        plt.close(fig)
    msg = (f"✅ {symbol}: {len(daily)} daily and {len(weekly)} weekly rows, indicators and 2 charts. "
           f"See PRICES_DAILY, PRICES_WEEKLY and CHARTS.")
    print(msg)
    set_status(book, msg, ok=True)
    charts.activate()


@script(name="2. Get AI Report (PDF + HTML)", button="[btn_ai]MASTER!C19")
async def get_ai_report(book: xw.Book):
    """Ask the TIGZIG Technical Analysis API for an AI-written report and show the links."""
    try:
        symbol, d = read_inputs(book)
    except ValueError as e:
        set_status(book, f"❌ {e}. Dates are in C8:C11 (empty = defaults).", ok=False)
        return
    if not symbol:
        set_status(book, f"❌ Type a Yahoo Finance symbol in {SYMBOL_CELL}, for example ^NSEI or AAPL.", ok=False)
        return
    c = book.sheets[CONTROL]
    c[f"A{LINK_ROW}"].resize(4, 1).clear_contents()
    set_status(book, f"⏳ Writing the report for {symbol}. This takes about a minute ...", ok=True)
    await book.flush()
    payload = {"ticker": symbol,
               "daily_start_date": str(d["daily_start"]), "daily_end_date": str(d["daily_end"]),
               "weekly_start_date": str(d["weekly_start"]), "weekly_end_date": str(d["weekly_end"])}
    model = str(c[MODEL_CELL].value or "").strip()
    if model:
        payload["model"] = model
    print(f"▶ POST {REPORT_API} {payload}")
    r = await pyfetch(REPORT_API, method="POST", headers={"Content-Type": "application/json"},
                      body=json.dumps(payload))
    text = await r.text()
    if not r.ok:
        try:
            body = json.loads(text)
            detail = body.get("detail") or body.get("error") or body
            if isinstance(detail, dict):
                detail = detail.get("message") or json.dumps(detail)
        except ValueError:
            detail = text
        hint = " The API allows 20 reports a day per IP address." if r.status == 429 else ""
        set_status(book, f"❌ Report API HTTP {r.status}: {str(detail)[:300]}{hint}", ok=False)
        return
    links = json.loads(text)
    pdf, html = links.get("pdf_url", ""), links.get("html_url", "")
    c[f"A{LINK_ROW}"].value = [[f'=HYPERLINK("{pdf}", "Open PDF report")'],
                               [f'=HYPERLINK("{html}", "Open HTML report")'], [pdf], [html]]
    msg = f"✅ Report ready for {symbol} ({model or 'default model'}). Click the links below (PDF and HTML)."
    print(msg, pdf, html)
    set_status(book, msg, ok=True)
