"""YLENS: Yahoo Finance Analyzer in Excel with xlwings Lite.

1. Company Profile: company details, valuation, dividends, share statistics and key executives.
2. Financial Statements: annual P&L, balance sheet and cash flow, and quarterly P&L.
3. Prices & Charts: daily prices, weekly bars, indicators computed in pandas, charts with matplotlib.
4. AI Report: the TIGZIG Technical Analysis API writes a report from the charts and indicators (PDF + HTML).
Data comes from the TIGZIG Yahoo Finance API (yfin-h.tigzig.com). Both APIs are open (no key) and allow
requests from xlwings Lite (CORS origin addin.xlwings.org).
"""
import asyncio
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
STATUS_CELL = "A23"
LINK_ROW = 25              # A25 PDF link, A26 HTML link, A27:A28 the raw URLs
YFIN_API = "https://yfin-h.tigzig.com"
PRICES_API = f"{YFIN_API}/get-all-prices/"
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
# Company profile and financial statements
# -------------------------
PROFILE_SECTIONS = {
    "1. Company Information": [
        "longBusinessSummary", "fullTimeEmployees", "address1", "address2", "city", "zip", "country", "phone", "fax", "website",
        "industry", "industryKey", "industryDisp", "sector", "sectorKey", "sectorDisp",
        "symbol", "longName", "shortName", "displayName", "typeDisp", "quoteType",
        "language", "region", "market", "exchange", "exchangeTimezoneName",
        "exchangeTimezoneShortName", "fullExchangeName", "messageBoardId", "irWebsite",
        "corporateActions", "executiveTeam", "maxAge"
    ],
    "2. Stock Price & Market Activity": [
        "currentPrice", "previousClose", "open", "dayLow", "dayHigh",
        "regularMarketPrice", "regularMarketPreviousClose", "regularMarketOpen",
        "regularMarketDayLow", "regularMarketDayHigh", "regularMarketChange",
        "regularMarketChangePercent", "regularMarketVolume", "volume",
        "averageVolume", "averageVolume3Month", "averageVolume10days",
        "averageDailyVolume10Day", "regularMarketDayRange", "bid", "ask",
        "bidSize", "askSize", "priceHint", "preMarketPrice", "preMarketTime",
        "preMarketChange", "preMarketChangePercent", "marketState",
        "hasPrePostMarketData", "quoteSourceName", "triggerable",
        "customPriceAlertConfidence", "gmtOffSetMilliseconds", "exchangeDataDelayedBy"
    ],
    "3. Valuation & Ratios": [
        "marketCap", "enterpriseValue", "priceToSalesTrailing12Months",
        "priceToBook", "trailingPE", "forwardPE", "trailingPegRatio", "beta",
        "enterpriseToRevenue", "enterpriseToEbitda", "payoutRatio", "debtToEquity",
        "recommendationMean", "recommendationKey", "numberOfAnalystOpinions",
        "averageAnalystRating", "priceEpsCurrentYear"
    ],
    "4. Dividend Information": [
        "dividendRate", "dividendYield", "trailingAnnualDividendRate",
        "trailingAnnualDividendYield", "fiveYearAvgDividendYield",
        "lastDividendValue", "lastDividendDate", "dividendDate", "exDividendDate"
    ],
    "5. Earnings & Financials": [
        "trailingEps", "forwardEps", "epsTrailingTwelveMonths", "epsForward",
        "epsCurrentYear", "netIncomeToCommon", "totalRevenue", "revenuePerShare",
        "grossProfits", "ebitda", "operatingCashflow", "freeCashflow",
        "totalCash", "totalCashPerShare", "totalDebt", "bookValue",
        "earningsGrowth", "revenueGrowth", "earningsQuarterlyGrowth",
        "returnOnAssets", "returnOnEquity", "grossMargins", "operatingMargins",
        "ebitdaMargins", "quickRatio", "currentRatio", "financialCurrency"
    ],
    "6. Share Statistics": [
        "sharesOutstanding", "impliedSharesOutstanding", "floatShares",
        "sharesShort", "sharesShortPriorMonth", "sharesShortPreviousMonthDate",
        "dateShortInterest", "sharesPercentSharesOut", "shortPercentOfFloat",
        "shortRatio", "heldPercentInsiders", "heldPercentInstitutions"
    ],
    "7. Analyst Price Targets": [
        "targetHighPrice", "targetLowPrice", "targetMeanPrice", "targetMedianPrice"
    ],
    "8. Technical Indicators": [
        "fiftyDayAverage", "fiftyDayAverageChange", "fiftyDayAverageChangePercent",
        "twoHundredDayAverage", "twoHundredDayAverageChange",
        "twoHundredDayAverageChangePercent", "fiftyTwoWeekLow", "fiftyTwoWeekHigh",
        "fiftyTwoWeekLowChange", "fiftyTwoWeekLowChangePercent",
        "fiftyTwoWeekHighChange", "fiftyTwoWeekHighChangePercent",
        "fiftyTwoWeekRange", "fiftyTwoWeekChange", "fiftyTwoWeekChangePercent",
        "SandP52WeekChange", "52WeekChange"
    ],
    "9. Earnings Call / Fiscal Calendar": [
        "lastFiscalYearEnd", "nextFiscalYearEnd", "mostRecentQuarter",
        "earningsTimestamp", "earningsTimestampStart", "earningsTimestampEnd",
        "earningsCallTimestampStart", "earningsCallTimestampEnd",
        "isEarningsDateEstimate"
    ],
    "10. Other": [
        "lastSplitFactor", "lastSplitDate", "firstTradeDateMilliseconds",
        "esgPopulated", "cryptoTradeable", "tradeable", "sourceInterval"
    ]
}


STATEMENTS = [  # sheet, endpoint, title
    ("PL", "excel/get-income-statement", "Income statement (annual)"),
    ("BS", "excel/get-balance-sheet", "Balance sheet (annual)"),
    ("CF", "excel/get-cash-flow", "Cash flow (annual)"),
    ("QTLY", "excel/get-quarterly-income-statement", "Income statement (quarterly)"),
]


def last_sheet(book: xw.Book, default: xw.Sheet) -> xw.Sheet:
    """Price sheets go after the profile and statement sheets when they exist, else after MASTER."""
    names = [s.name for s in book.sheets]
    for name in ("QTLY", "CF", "BS", "PL", "PROFILE"):
        if name in names:
            return book.sheets[name]
    return default


async def get_json(path: str, symbol: str) -> dict:
    """GET a yfin endpoint for one symbol; returns that symbol's block or raises with a readable message."""
    r = await pyfetch(f"{YFIN_API}/{path}/?tickers={quote(symbol)}")
    text = await r.text()
    if not r.ok:
        raise RuntimeError(f"{path} HTTP {r.status}: {text[:200]}")
    data = json.loads(text)
    block = data.get(symbol) if isinstance(data, dict) else None
    if not isinstance(block, dict):
        raise RuntimeError(f"no data for {symbol}. Check the Yahoo Finance symbol.")
    if block.get("error"):
        raise RuntimeError(f"{symbol}: {block['error']}")
    return block


def cell_value(v: Any) -> Any:
    """Numbers, text and booleans as they are; lists and dicts as short JSON text."""
    if isinstance(v, (list, dict)):
        return json.dumps(v)[:500]
    return v


def read_symbol(book: xw.Book) -> str:
    symbol = str(book.sheets[CONTROL][SYMBOL_CELL].value or "").strip().upper()
    if not symbol:
        raise ValueError(f"Type a Yahoo Finance symbol in {SYMBOL_CELL}, for example GOOG or RELIANCE.NS.")
    return symbol


@script(name="1. Company Profile", button="[btn_profile]MASTER!C14")
async def get_profile(book: xw.Book):
    """Company details, market data, valuation, dividends, share statistics and key executives."""
    try:
        symbol = read_symbol(book)
    except ValueError as e:
        set_status(book, f"❌ {e}", ok=False)
        return
    set_status(book, f"⏳ Pulling the profile for {symbol} ...", ok=True)
    await book.flush()
    try:
        block = await get_json("get-detailed-info", symbol)
    except Exception as e:
        print(f"❌ {e}")
        set_status(book, f"❌ {e}", ok=False)
        return
    info, officers = block.get("main_info") or {}, block.get("officers") or []
    sh = replace_sheet(book, "PROFILE", book.sheets[CONTROL])
    sh["A1"].value = f"Company profile: {symbol} ({info.get('longName') or info.get('shortName') or ''})"
    sh["A1"].font.bold = True
    row, written = 3, 0
    for section, fields in PROFILE_SECTIONS.items():
        rows = [[f, cell_value(info[f])] for f in fields if info.get(f) not in (None, "", [], {})]
        if not rows:
            continue
        sh[f"A{row}"].value = section
        sh[f"A{row}"].font.bold = True
        sh[f"A{row}"].font.color = BLUE
        for k, (_, v) in enumerate(rows):
            if isinstance(v, str):  # keep text as text: "20:1" would become a time, "94043" a number
                sh[f"B{row + 2 + k}"].number_format = "@"
        sh[f"A{row + 1}"].value = [["Item", "Value"]] + rows
        sh[f"A{row + 1}:B{row + 1}"].font.bold = True
        written += len(rows)
        row += len(rows) + 3
    if officers:
        sh[f"A{row}"].value = "11. Key Executives"
        sh[f"A{row}"].font.bold = True
        sh[f"A{row}"].font.color = BLUE
        table = [["Name", "Title", "Age", "Total Pay"]] + [
            [o.get("name"), o.get("title"), o.get("age"), o.get("totalPay")] for o in officers]
        sh[f"A{row + 1}"].value = table
        sh[f"A{row + 1}:D{row + 1}"].font.bold = True
    msg = f"✅ {symbol}: {written} profile items and {len(officers)} executives on PROFILE."
    print(msg)
    set_status(book, msg, ok=True)
    sh.activate()


@script(name="2. Financial Statements", button="[btn_fin]MASTER!C16")
async def get_financials(book: xw.Book):
    """Annual P&L, balance sheet, cash flow and quarterly P&L, one sheet each."""
    try:
        symbol = read_symbol(book)
    except ValueError as e:
        set_status(book, f"❌ {e}", ok=False)
        return
    set_status(book, f"⏳ Pulling financial statements for {symbol} ...", ok=True)
    await book.flush()
    results = await asyncio.gather(*(get_json(ep, symbol) for _, ep, _ in STATEMENTS), return_exceptions=True)
    after = book.sheets["PROFILE"] if "PROFILE" in [s.name for s in book.sheets] else book.sheets[CONTROL]
    done, failed = [], []
    for (name, _, title), res in zip(STATEMENTS, results):
        if isinstance(res, BaseException):
            failed.append(f"{name}: {res}")
            continue
        dates, data = res.get("dates") or [], res.get("data") or []
        sh = replace_sheet(book, name, after)
        after = sh
        sh["A1"].value = f"{title}: {symbol}. Values as reported by Yahoo Finance, in the reporting currency."
        sh["A1"].font.bold = True
        table = [["Metric"] + dates] + [[d.get("metric")] + [d.get(x) for x in dates] for d in data]
        sh["A3"].value = table
        if data:
            sh.tables.add(source=sh["A3"].resize(len(table), len(table[0])))
            sh["B4"].resize(len(data), len(dates)).number_format = "#,##0"
        done.append(f"{name} {len(data)} rows")
    if failed:
        msg = f"❌ {symbol}: " + "; ".join(failed) + (f". Done: {', '.join(done)}" if done else "")
        set_status(book, msg[:400], ok=False)
    else:
        set_status(book, f"✅ {symbol}: " + ", ".join(done) + ".", ok=True)
    print(done, failed)


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


@script(name="3. Prices & Charts", button="[btn_tech]MASTER!C19")
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
    ds = write_prices(book, "PRICES_DAILY", last_sheet(book, control), daily,
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


@script(name="4. AI Report (PDF + HTML)", button="[btn_ai]MASTER!C21")
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
