"""AI Workflows: LLM API calls from Excel with xlwings Lite.

1a / 1b send the first rows of an Excel table to an LLM (Gemini, or any model on OpenRouter) and get back
structured JSON: which columns are categorical or numeric, and a PostgreSQL type and
description for every column. 2 runs a Python EDA on the columns the LLM picked.

API keys are read from environment variables (xlwings Lite > Settings > Environment
Variables, scope Local), so they stay in your browser and are never saved in the workbook.
"""
import asyncio
import json
import os
from typing import Any, Optional, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlwings as xw
from pyodide.http import pyfetch
from xlwings import script

CONTROL = "MASTER"
TABLE_CELL = "C12"         # table to analyze
GEMINI_MODEL_CELL = "C13"
OPENROUTER_MODEL_CELL = "C14"
STATUS_CELL = "A21"        # red (error) / dark blue (success) message
SCHEMA_SHEET = "SCHEMA"    # results of 1a / 1b; 2. Run EDA reads B3:B5 from here
SAMPLE_ROWS = 10

BLUE = "#1F4E79"
RED = "#C00000"
HEADER_FILL = "#DDEBF7"


# -------------------------
# Small helpers
# -------------------------
def set_status(book: xw.Book, text: str, ok: bool):
    cell = book.sheets[CONTROL][STATUS_CELL]
    cell.value = text
    cell.font.bold = True
    cell.font.size = 12 if ok else 14
    cell.font.color = BLUE if ok else RED


def find_table(book: xw.Book, name: str) -> Tuple[Optional[xw.Sheet], Any]:
    """Find an Excel table by name in any sheet (case-insensitive)."""
    for sheet in book.sheets:
        for table in sheet.tables:
            if table.name.lower() == name.lower():
                return sheet, table
    return None, None


def delete_sheet(book: xw.Book, name: str):
    for sheet in book.sheets:
        if sheet.name == name:
            sheet.delete()


def read_settings(book: xw.Book) -> Tuple[str, str, str]:
    c = book.sheets[CONTROL]
    table = str(c[TABLE_CELL].value or "").strip()
    gemini_model = str(c[GEMINI_MODEL_CELL].value or "").strip() or "gemini-flash-latest"
    openrouter_model = str(c[OPENROUTER_MODEL_CELL].value or "").strip() or "openai/gpt-6-luna"
    return table, gemini_model, openrouter_model


def load_table(book: xw.Book, table_name: str) -> Optional[pd.DataFrame]:
    _, table = find_table(book, table_name)
    if table is None:
        return None
    return table.range.options(pd.DataFrame, index=False).value


# -------------------------
# Prompts and JSON schemas (the same two tasks for both providers)
# -------------------------
def classify_prompt(sample: pd.DataFrame) -> str:
    return f"""Analyze these rows from a dataset (including headers):

{sample.to_string(index=False)}

Task: identify categorical and numeric variables for exploratory data analysis.
Rules:
1. Exclude ID columns, meaningless numbers and dates.
2. Categorical: text columns and numeric codes that represent categories.
3. Numeric: continuous or discrete numbers suitable for statistical analysis.
Use the column names exactly as they appear in the header."""


def schema_prompt(sample: pd.DataFrame) -> str:
    return f"""You are a PostgreSQL schema generator. Analyze this sample data (first rows, with headers):

{sample.to_string(index=False)}

Requirements:
1. Use standard PostgreSQL types: TEXT, NUMERIC, DATE, TIMESTAMP (do not use INTEGER).
2. Column names must be SQL-safe (lowercase letters, digits and underscores).
3. Use NUMERIC for any value that can have decimals, DATE or TIMESTAMP for dates, TEXT when unsure.
4. Give each column a description of up to 50 words: what it holds, whether it looks continuous,
   discrete, categorical, temporal or an ID, and example values from the data."""


CLASSIFY_SCHEMA = {
    "type": "object",
    "properties": {
        "categorical_variables": {"type": "array", "items": {"type": "string"}},
        "numeric_variables": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["categorical_variables", "numeric_variables"],
    "additionalProperties": False,
}
COLUMNS_SCHEMA = {
    "type": "object",
    "properties": {
        "columns": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string"},
                    "type": {"type": "string"},
                    "description": {"type": "string"},
                },
                "required": ["name", "type", "description"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["columns"],
    "additionalProperties": False,
}


def gemini_schema(schema: dict) -> dict:
    """Gemini's response_schema uses the same shape without additionalProperties."""
    out = {k: v for k, v in schema.items() if k != "additionalProperties"}
    if "properties" in out:
        out["properties"] = {k: gemini_schema(v) for k, v in out["properties"].items()}
    if "items" in out:
        out["items"] = gemini_schema(out["items"])
    return out


# -------------------------
# API calls (pyfetch is built into Pyodide, no package to install)
# -------------------------
RETRY_STATUS = {429, 500, 502, 503, 504}   # busy or temporary errors: wait and try again


async def post_json(url: str, headers: dict, payload: dict, tries: int = 3) -> dict:
    for attempt in range(1, tries + 1):
        r = await pyfetch(url, method="POST", headers={"Content-Type": "application/json", **headers},
                          body=json.dumps(payload))
        text = await r.text()
        if r.ok or r.status not in RETRY_STATUS or attempt == tries:
            break
        print(f"⏳ HTTP {r.status}, retrying ({attempt}/{tries - 1}) ...")
        await asyncio.sleep(3 * attempt)
    if not r.ok:
        try:
            message = json.loads(text).get("error", {}).get("message", text)
        except ValueError:
            message = text
        raise RuntimeError(f"HTTP {r.status}: {str(message)[:300]}")
    return json.loads(text)


async def ask_gemini(model: str, key: str, prompt: str, schema: dict) -> dict:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0, "response_mime_type": "application/json",
                             "response_schema": gemini_schema(schema)},
    }
    data = await post_json(url, {"x-goog-api-key": key}, payload)
    parts = data["candidates"][0]["content"]["parts"]
    return json.loads("".join(p.get("text", "") for p in parts if not p.get("thought")))


async def ask_openrouter(model: str, key: str, prompt: str, schema: dict, name: str) -> dict:
    """OpenRouter: one key and the OpenAI-style API for models from many providers."""
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "response_format": {"type": "json_schema",
                            "json_schema": {"name": name, "schema": schema, "strict": True}},
        "provider": {"require_parameters": True},  # only route to providers that support the JSON schema
    }
    data = await post_json("https://openrouter.ai/api/v1/chat/completions",
                           {"Authorization": f"Bearer {key}", "X-Title": "xlwings Lite AI Workflows"}, payload)
    return json.loads(data["choices"][0]["message"]["content"])


# -------------------------
# Writing the SCHEMA sheet
# -------------------------
def keep_real_columns(names: list, df: pd.DataFrame) -> Tuple[list, list]:
    """Match the LLM's column names to the table's real columns (case-insensitive)."""
    lookup = {str(c).lower(): str(c) for c in df.columns}
    kept, dropped = [], []
    for n in names:
        real = lookup.get(str(n).strip().lower())
        if real and real not in kept:
            kept.append(real)
        elif not real:
            dropped.append(str(n))
    return kept, dropped


def write_schema_sheet(book: xw.Book, table_name: str, source: str, classes: dict,
                       columns: list, df: pd.DataFrame) -> Tuple[int, int, list]:
    cats, d1 = keep_real_columns(classes["categorical_variables"], df)
    nums, d2 = keep_real_columns(classes["numeric_variables"], df)
    nums = [c for c in nums if c not in cats]
    delete_sheet(book, SCHEMA_SHEET)
    sh = book.sheets.add(name=SCHEMA_SHEET, after=book.sheets[CONTROL])
    sh["A1"].value = f"Schema detected by {source}"
    sh["A1"].font.bold = True
    sh["A1"].font.size = 14
    sh["A3"].value = [["Table", table_name],
                      ["Categorical columns", ", ".join(cats)],
                      ["Numeric columns", ", ".join(nums)]]
    sh["A3:A5"].font.bold = True
    sh["B4:B5"].color = "#FFFFCC"
    sh["A6"].value = "Edit B4 / B5 (comma-separated column names) if you disagree with the LLM, then run 2. Run EDA."
    sh["A6"].font.italic = True
    schema_df = pd.DataFrame(columns)[["name", "type", "description"]]
    schema_df.columns = ["Column Name", "PostgreSQL Type", "Description"]
    sh["A8"].options(index=False).value = schema_df
    sh.tables.add(source=sh["A8"].resize(len(schema_df) + 1, 3), name="T_SCHEMA")
    return len(cats), len(nums), d1 + d2


async def detect_schema(book: xw.Book, provider: str):
    table_name, gemini_model, openrouter_model = read_settings(book)
    key_name = "GEMINI_API_KEY" if provider == "Gemini" else "OPENROUTER_API_KEY"
    key = os.getenv(key_name)
    if not key:
        set_status(book, f"❌ {key_name} is not set. In xlwings Lite open Settings > Environment Variables, "
                         f"add {key_name} with scope Local, then restart xlwings Lite (see Step 1).", ok=False)
        return
    df = load_table(book, table_name)
    if df is None:
        set_status(book, f"❌ Table '{table_name}' not found. Type the name of an Excel table in {TABLE_CELL} "
                         f"(the sample tables are RBICC and RESP).", ok=False)
        return
    sample = df.head(SAMPLE_ROWS)
    model = gemini_model if provider == "Gemini" else openrouter_model
    print(f"🤖 {provider} {model}: sending {len(sample)} rows of {table_name} ({df.shape[1]} columns)")
    set_status(book, f"⏳ Waiting for {provider} ({model}). This can take up to a minute ...", ok=True)
    await book.flush()
    try:
        if provider == "Gemini":
            classes, cols = await asyncio.gather(
                ask_gemini(model, key, classify_prompt(sample), CLASSIFY_SCHEMA),
                ask_gemini(model, key, schema_prompt(sample), COLUMNS_SCHEMA))
        else:
            classes, cols = await asyncio.gather(
                ask_openrouter(model, key, classify_prompt(sample), CLASSIFY_SCHEMA, "variable_classification"),
                ask_openrouter(model, key, schema_prompt(sample), COLUMNS_SCHEMA, "postgresql_schema"))
    except Exception as e:
        print(f"❌ {e}")
        set_status(book, f"❌ {provider} call failed ({model}): {e}", ok=False)
        return
    print("📄 Variable classification:", json.dumps(classes))
    print("📄 PostgreSQL schema:", json.dumps(cols)[:2000])
    n_cat, n_num, dropped = write_schema_sheet(book, table_name, f"{provider} ({model}) for table {table_name}",
                                               classes, cols["columns"], df)
    msg = (f"✅ {provider} ({model}): {n_cat} categorical, {n_num} numeric, {len(cols['columns'])} columns "
           f"described. See the SCHEMA sheet, then run 2. Run EDA.")
    if dropped:
        msg += f" Ignored names not in the table: {', '.join(dropped)}."
    print(msg)
    set_status(book, msg, ok=True)
    book.sheets[SCHEMA_SHEET].activate()


@script(name="1a. Detect Schema (Gemini)", button="[run_gemini_btn]MASTER!C17")
async def analyze_table_schema_gemini(book: xw.Book):
    """Ask Gemini for the column classification and a PostgreSQL schema."""
    await detect_schema(book, "Gemini")


@script(name="1b. Detect Schema (OpenRouter)", button="[run_openrouter_btn]MASTER!C19")
async def analyze_table_schema_openrouter(book: xw.Book):
    """Ask a model on OpenRouter for the column classification and a PostgreSQL schema."""
    await detect_schema(book, "OpenRouter")


# -------------------------
# EDA (pure pandas: takes and returns DataFrames)
# -------------------------
def parse_columns(text: Any) -> list:
    text = str(text or "").strip().strip("[]")
    return [c.strip().strip("'\"") for c in text.split(",") if c.strip().strip("'\"")]


def numeric_stats(num: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    def add(label: str, values: pd.Series):
        rows.append({"Statistic": label, **{c: values[c] for c in num.columns}})
    add("count", num.count())
    add("missing", num.isna().sum())
    add("missing %", (num.isna().mean() * 100).round(2))
    add("mean", num.mean())
    add("std", num.std())
    add("variance", num.var())
    add("skewness", num.skew())
    add("kurtosis", num.kurtosis())
    for p in [0, 1, 5, 10, 25, 50, 75, 90, 95, 99, 100]:
        add({0: "min", 50: "median", 100: "max"}.get(p, f"p{p}"), num.quantile(p / 100))
    return pd.DataFrame(rows)


def as_labels(s: pd.Series) -> pd.Series:
    """Category values as text; whole-number codes read from Excel as 1.0 show as 1."""
    num = pd.to_numeric(s, errors="coerce")
    if s.notna().any() and num.notna().sum() == s.notna().sum() and (num.dropna() % 1 == 0).all():
        return num.astype("Int64").astype(str).replace("<NA>", "(blank)")
    return s.fillna("(blank)").astype(str)


def category_counts(s: pd.Series, top: int = 50) -> pd.DataFrame:
    counts = as_labels(s).value_counts()
    if len(counts) > top:
        counts = pd.concat([counts.iloc[:top], pd.Series({"(all other values)": counts.iloc[top:].sum()})])
    return pd.DataFrame({"Value": counts.index, "Count": counts.values,
                         "Percent": (counts.values / len(s) * 100).round(2)})


def density(s: pd.Series, x: np.ndarray) -> np.ndarray:
    """Gaussian kernel density (Scott's rule), the smooth line drawn over each histogram."""
    v = s.to_numpy(dtype=float)
    h = v.std(ddof=1) * len(v) ** (-0.2)
    z = (x[:, None] - v[None, :]) / h
    return np.exp(-0.5 * z ** 2).sum(axis=1) / (len(v) * h * np.sqrt(2 * np.pi))


def write_block(sheet: xw.Sheet, row: int, title: str, df: pd.DataFrame) -> int:
    """Title row + a table. Returns the next free row."""
    sheet[f"A{row}"].value = title
    sheet[f"A{row}"].font.bold = True
    sheet[f"A{row}"].resize(1, max(3, df.shape[1])).color = HEADER_FILL
    sheet[f"A{row + 1}"].options(index=False).value = df
    sheet.tables.add(source=sheet[f"A{row + 1}"].resize(len(df) + 1, df.shape[1]))
    return row + len(df) + 4


@script(name="2. Run EDA", button="[run_eda_btn]MASTER!C24")
def perform_eda(book: xw.Book):
    """Statistics, distributions, correlations and charts for the columns on the SCHEMA sheet."""
    names = [s.name for s in book.sheets]
    if SCHEMA_SHEET not in names:
        set_status(book, "❌ No SCHEMA sheet yet. Run 1a or 1b first.", ok=False)
        return
    sch = book.sheets[SCHEMA_SHEET]
    table_name = str(sch["B3"].value or "").strip()
    df = load_table(book, table_name)
    if df is None:
        set_status(book, f"❌ Table '{table_name}' (SCHEMA!B3) not found in this workbook.", ok=False)
        return
    cats, bad1 = keep_real_columns(parse_columns(sch["B4"].value), df)
    nums, bad2 = keep_real_columns(parse_columns(sch["B5"].value), df)
    if bad1 or bad2:
        set_status(book, f"❌ Not columns of {table_name}: {', '.join(bad1 + bad2)}. Fix SCHEMA!B4:B5.", ok=False)
        return
    if not cats and not nums:
        set_status(book, "❌ SCHEMA!B4 and B5 are both empty.", ok=False)
        return
    print(f"📊 EDA on {table_name}: {len(df):,} rows, {len(nums)} numeric, {len(cats)} categorical")
    num = df[nums].apply(pd.to_numeric, errors="coerce") if nums else pd.DataFrame()

    # Calculate everything first, then write.
    blocks = []
    if nums:
        blocks.append(("Numeric variables", numeric_stats(num)))
    for c in cats:
        blocks.append((f"{c}: distribution", category_counts(df[c])))
    corr = num.corr().round(2) if len(nums) > 1 else None
    if corr is not None:
        blocks.append(("Correlation matrix", corr.reset_index().rename(columns={"index": "Variable"})))

    distros_name, plot_name = f"DISTROS_{table_name}"[:31], f"PLOT_{table_name}"[:31]
    for n in (distros_name, plot_name):
        delete_sheet(book, n)
    distros = book.sheets.add(name=distros_name, after=book.sheets[SCHEMA_SHEET])
    plots = book.sheets.add(name=plot_name, after=distros)
    row = 1
    for title, block in blocks:
        row = write_block(distros, row, title, block)

    plt.rcParams.update({"figure.facecolor": "white", "axes.grid": True, "grid.alpha": 0.3,
                         "axes.titlesize": 9, "axes.labelsize": 8, "xtick.labelsize": 7, "ytick.labelsize": 7})
    plots["A1"].value = f"Charts for {table_name}"
    plots["A1"].font.bold = True
    slots = [f"{col}{r}" for r in range(20, 2000, 16) for col in ("A", "F", "K")]

    def place(fig, anchor: str, name: str):
        fig.tight_layout()
        plots.pictures.add(fig, name=name, update=True, anchor=plots[anchor])
        plt.close(fig)

    if corr is not None:
        fig, ax = plt.subplots(figsize=(6, 4))
        im = ax.imshow(corr.values, cmap="coolwarm", vmin=-1, vmax=1, aspect="auto")
        for i in range(len(corr)):
            for j in range(len(corr)):
                ax.text(j, i, f"{corr.iloc[i, j]:.2f}", ha="center", va="center", fontsize=6)
        ax.set_xticks(range(len(corr)), corr.columns, rotation=45, ha="right")
        ax.set_yticks(range(len(corr)), corr.columns)
        ax.set_title("Correlation matrix", weight="bold")
        fig.colorbar(im, ax=ax, shrink=0.8)
        place(fig, "A3", "correlation_heatmap")
    i = 0
    for c in nums:
        s = num[c].dropna()
        if s.empty:
            continue
        fig, ax = plt.subplots(figsize=(3.2, 2.6))
        ax.hist(s, bins=30, density=True, color="#9DC3E6", edgecolor="#5B7A99")
        if s.nunique() > 1:
            x = np.linspace(s.min(), s.max(), 200)
            ax.plot(x, density(s, x), color=BLUE, lw=1.5)
        ax.set_title(c, weight="bold")
        place(fig, slots[i], f"dist_{c}")
        i += 1
        fig, ax = plt.subplots(figsize=(3.2, 2.6))
        ax.boxplot(s, widths=0.6, patch_artist=True, boxprops=dict(facecolor="#9DC3E6"),
                   medianprops=dict(color=BLUE, lw=2))
        ax.set_title(c, weight="bold")
        place(fig, slots[i], f"box_{c}")
        i += 1
    for c in cats:
        counts = as_labels(df[c]).value_counts().head(15)
        fig, ax = plt.subplots(figsize=(3.2, 2.6))
        ax.bar(range(len(counts)), counts.values, color="#9DC3E6", edgecolor="#5B7A99")
        ax.set_xticks(range(len(counts)), counts.index, rotation=45, ha="right")
        ax.set_title(c + (" (top 15)" if df[c].nunique() > 15 else ""), weight="bold")
        place(fig, slots[i], f"bar_{c}")
        i += 1

    msg = (f"✅ EDA done for {table_name}: {len(nums)} numeric, {len(cats)} categorical, {i} charts. "
           f"See {distros_name} and {plot_name}.")
    print(msg)
    set_status(book, msg, ok=True)
    distros.activate()
