"""Database & ML in Excel with xlwings Lite.

Database (steps 1-3): query a Neon Postgres database straight from Excel over Neon's HTTPS SQL
endpoint. No server in between: the connection string comes from the environment variable
NEON_DATABASE_URL (xlwings Lite > Settings > Environment Variables, scope Local) and is not saved
in the workbook.
Analysis (steps 4-5): EDA and an XGBoost response model on any Excel table, for example the sample
table DEF that ships with the workbook, or a table pulled from the database.
"""
import json
import os
from typing import Any, Optional, Tuple
from urllib.parse import urlparse

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlwings as xw
from pyodide.http import pyfetch
from xlwings import script

CONTROL = "MASTER"
CELLS = {"table": "C14", "rows": "C15", "mode": "C16", "sql": "A20",
         "analysis_table": "C23", "id_col": "C24", "target_col": "C25"}
STATUS_CELL = "A30"
MAX_ROWS = 50_000          # rows written to Excel from one pull or query
EXCLUDE_FEATURES = {"score_probability", "score_decile"}   # outputs of an earlier model run, never inputs
BLUE, ORANGE, GREY, RED = "#1F4E79", "#E67E22", "#7F7F7F", "#C00000"
HEADER_FILL = "#DDEBF7"


# -------------------------
# Helpers
# -------------------------
def set_status(book: xw.Book, text: str, ok: bool):
    cell = book.sheets[CONTROL][STATUS_CELL]
    cell.value = text
    cell.font.bold = True
    cell.font.size = 12 if ok else 14
    cell.font.color = BLUE if ok else RED


def setting(book: xw.Book, key: str) -> Any:
    return book.sheets[CONTROL][CELLS[key]].value


def new_sheet(book: xw.Book, name: str) -> xw.Sheet:
    for sh in book.sheets:
        if sh.name == name:
            sh.delete()
    return book.sheets.add(name=name, after=book.sheets[CONTROL])


def write_block(sheet: xw.Sheet, row: int, title: str, df: pd.DataFrame, col: str = "A",
                name: Optional[str] = None) -> int:
    """Title row + a table. Returns the next free row."""
    sheet[f"{col}{row}"].value = title
    sheet[f"{col}{row}"].font.bold = True
    sheet[f"{col}{row}"].resize(1, max(3, df.shape[1])).color = HEADER_FILL
    sheet[f"{col}{row + 1}"].options(index=False).value = df
    if len(df):
        sheet.tables.add(source=sheet[f"{col}{row + 1}"].resize(len(df) + 1, df.shape[1]), name=name)
    return row + len(df) + 4


def find_table(book: xw.Book, name: str) -> Tuple[Optional[xw.Sheet], Any]:
    for sheet in book.sheets:
        for table in sheet.tables:
            if table.name.lower() == str(name).strip().lower():
                return sheet, table
    return None, None


def quote_ident(name: str) -> str:
    """'public.my table' -> "public"."my table" (safe identifier quoting)."""
    return ".".join('"' + part.strip().replace('"', '""') + '"' for part in str(name).split("."))


# -------------------------
# Neon over HTTPS
# -------------------------
NUMERIC_TYPES = {20, 21, 23, 26, 700, 701, 1700}
DATE_TYPES = {1082, 1114, 1184}


async def neon_query(sql: str, params: Optional[list] = None) -> pd.DataFrame:
    conn = os.getenv("NEON_DATABASE_URL", "").strip()
    if not conn:
        raise RuntimeError("NEON_DATABASE_URL is not set. In xlwings Lite open Settings > Environment Variables, add "
                           "NEON_DATABASE_URL (your Neon connection string) with scope Local, then restart xlwings Lite.")
    host = urlparse(conn).hostname
    if not host or "neon" not in host:
        raise RuntimeError("NEON_DATABASE_URL does not look like a Neon connection string (postgresql://...neon.tech/...).")
    # No Content-Type header: Neon's CORS rules do not list it, and the browser would block the call.
    r = await pyfetch(f"https://{host}/sql", method="POST",
                      headers={"Neon-Connection-String": conn, "Neon-Array-Mode": "true"},
                      body=json.dumps({"query": sql, "params": params or []}))
    text = await r.text()
    if not r.ok:
        try:
            message = json.loads(text).get("message", text)
        except ValueError:
            message = text
        raise RuntimeError(f"Neon HTTP {r.status}: {str(message)[:300]}")
    data = json.loads(text)
    fields = data.get("fields", [])
    df = pd.DataFrame(data.get("rows", []), columns=[f["name"] for f in fields])
    for f in fields:
        if f["dataTypeID"] in NUMERIC_TYPES:
            df[f["name"]] = pd.to_numeric(df[f["name"]], errors="coerce")
        elif f["dataTypeID"] in DATE_TYPES:
            df[f["name"]] = pd.to_datetime(df[f["name"]], errors="coerce")
    return df


def write_result(book: xw.Book, sheet_name: str, title: str, df: pd.DataFrame) -> str:
    sh = new_sheet(book, sheet_name)
    shown = df.head(MAX_ROWS)
    sh["A1"].value = title + (f" (first {MAX_ROWS:,} of {len(df):,} rows)" if len(df) > MAX_ROWS else "")
    sh["A1"].font.bold = True
    if shown.empty:
        sh["A3"].value = "No rows returned."
        return sh.name
    obj_cols = [c for c in shown.columns if shown[c].dtype == object]
    for c in obj_cols:   # keep codes like 00123 and date-like text as text
        idx = shown.columns.get_loc(c)
        sh["A3"].offset(0, idx).resize(len(shown) + 1, 1).number_format = "@"
    sh["A3"].options(index=False).value = shown
    sh.tables.add(source=sh["A3"].resize(len(shown) + 1, shown.shape[1]), name=sheet_name)
    sh.activate()
    return sh.name


@script(name="1. List Tables", button="[btn_list]MASTER!C12")
async def list_tables(book: xw.Book):
    """All tables and views in the database, with approximate row counts."""
    sql = """
        SELECT n.nspname AS schema, c.relname AS table_name,
               CASE c.relkind WHEN 'v' THEN 'view' WHEN 'm' THEN 'materialized view' ELSE 'table' END AS type,
               GREATEST(c.reltuples, 0)::bigint AS approx_rows,
               (SELECT count(*) FROM information_schema.columns ic
                 WHERE ic.table_schema = n.nspname AND ic.table_name = c.relname) AS columns
        FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
        WHERE c.relkind IN ('r', 'p', 'v', 'm') AND n.nspname NOT IN ('pg_catalog', 'information_schema')
          AND n.nspname NOT LIKE 'pg_toast%'
        ORDER BY 1, 2"""
    try:
        df = await neon_query(sql)
    except Exception as e:
        set_status(book, f"❌ {e}", ok=False)
        return
    write_result(book, "TABLES", "Tables and views in the database (approx_rows is Postgres' estimate)", df)
    set_status(book, f"✅ {len(df)} tables and views listed on the TABLES sheet. Type one in {CELLS['table']} for step 2.", ok=True)


@script(name="2. Get Table Data", button="[btn_get]MASTER!C17")
async def get_table_data(book: xw.Book):
    """First N or random N rows of a table into the DB_DATA sheet (Excel table DB_DATA)."""
    table = str(setting(book, "table") or "").strip()
    if not table:
        set_status(book, f"❌ Type a table name in {CELLS['table']} (run 1. List Tables to see them).", ok=False)
        return
    try:
        n = max(1, min(int(float(setting(book, "rows") or 1000)), MAX_ROWS))
    except (TypeError, ValueError):
        n = 1000
    random_rows = str(setting(book, "mode") or "").strip().lower().startswith("random")
    sql = f"SELECT * FROM {quote_ident(table)}" + (" ORDER BY random()" if random_rows else "") + f" LIMIT {n}"
    try:
        df = await neon_query(sql)
    except Exception as e:
        set_status(book, f"❌ {e}", ok=False)
        return
    write_result(book, "DB_DATA", f"{'Random' if random_rows else 'First'} {n:,} rows of {table}", df)
    set_status(book, f"✅ {len(df):,} rows of {table} on DB_DATA (Excel table DB_DATA, usable in step 3).", ok=True)


@script(name="3. Run SQL Query", button="[btn_sql]MASTER!C20")
async def run_custom_query(book: xw.Book):
    """Run the SQL in the yellow box; results to the QUERY_RESULT sheet."""
    sql = str(setting(book, "sql") or "").strip().rstrip(";")
    if not sql:
        set_status(book, f"❌ Type a SQL query in the yellow box ({CELLS['sql']}).", ok=False)
        return
    try:
        df = await neon_query(sql)
    except Exception as e:
        set_status(book, f"❌ {e}", ok=False)
        return
    write_result(book, "QUERY_RESULT", "Query result", df)
    set_status(book, f"✅ Query returned {len(df):,} rows: QUERY_RESULT sheet (Excel table QUERY_RESULT).", ok=True)


# -------------------------
# Analysis (pure pandas)
# -------------------------
def load_analysis_table(book: xw.Book) -> Tuple[Optional[pd.DataFrame], str, str, str, str]:
    name = str(setting(book, "analysis_table") or "").strip()
    id_col = str(setting(book, "id_col") or "").strip()
    target = str(setting(book, "target_col") or "").strip()
    _, table = find_table(book, name)
    if table is None:
        return None, name, id_col, target, f"Excel table '{name}' not found ({CELLS['analysis_table']})."
    df = table.range.options(pd.DataFrame, index=False).value
    lookup = {str(c).lower(): c for c in df.columns}
    id_col = lookup.get(id_col.lower(), "") if id_col else ""
    if target.lower() not in lookup:
        return None, name, id_col, target, f"Target column '{target}' is not in {name} ({CELLS['target_col']})."
    target = lookup[target.lower()]
    y = pd.to_numeric(df[target], errors="coerce")
    if not set(y.dropna().unique()) <= {0, 1} or y.isna().any():
        return None, name, id_col, target, f"Target '{target}' must be 0 / 1 with no blanks."
    df[target] = y.astype(int)
    return df, name, id_col, target, ""


def split_columns(df: pd.DataFrame, id_col: str, target: str) -> Tuple[list, list]:
    skip = {id_col, target} | {c for c in df.columns if str(c).lower() in EXCLUDE_FEATURES}
    num = [c for c in df.columns if c not in skip and pd.api.types.is_numeric_dtype(df[c])]
    cat = [c for c in df.columns if c not in skip and c not in num]
    return num, cat


def numeric_summary(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for c in cols:
        s = pd.to_numeric(df[c], errors="coerce")
        q = s.quantile([0.01, 0.25, 0.5, 0.75, 0.99])
        rows.append({"Variable": c, "Count": int(s.count()), "Missing": int(s.isna().sum()), "Mean": s.mean(),
                     "Std": s.std(), "Min": s.min(), "P1": q[0.01], "P25": q[0.25], "Median": q[0.5],
                     "P75": q[0.75], "P99": q[0.99], "Max": s.max()})
    return pd.DataFrame(rows)


def category_profile(df: pd.DataFrame, col: str, target: str, top: int = 30) -> pd.DataFrame:
    """Count, share and target rate (response rate) per category value."""
    s = df[col].fillna("(blank)").astype(str)
    g = df.assign(_v=s).groupby("_v")[target].agg(["count", "sum"]).sort_values("count", ascending=False)
    if len(g) > top:
        rest = g.iloc[top:].sum()
        g = pd.concat([g.iloc[:top], pd.DataFrame([rest], index=["(all other values)"])])
    return pd.DataFrame({"Value": g.index, "Count": g["count"].values,
                         "Share %": (g["count"].values / len(df) * 100).round(2),
                         "Responders": g["sum"].values,
                         "Response rate %": (g["sum"].values / g["count"].values * 100).round(2)})


def place(sheet: xw.Sheet, fig, anchor: str, name: str):
    fig.tight_layout()
    sheet.pictures.add(fig, name=name, update=True, anchor=sheet[anchor])
    plt.close(fig)


@script(name="4. Run EDA", button="[btn_eda]MASTER!C26")
def perform_eda(book: xw.Book):
    """Summary statistics, response rate by category, correlations and charts."""
    df, name, id_col, target, err = load_analysis_table(book)
    if df is None:
        set_status(book, f"❌ {err}", ok=False)
        return
    num, cat = split_columns(df, id_col, target)
    print(f"📊 EDA on {name}: {len(df):,} rows, {len(num)} numeric, {len(cat)} categorical, target {target}")
    overall = df[target].mean() * 100
    sh = new_sheet(book, f"{name}_EDA"[:31])
    sh["A1"].value = f"EDA for {name}: {len(df):,} rows, overall response rate {overall:.2f}%"
    sh["A1"].font.bold = True
    sh["A1"].font.size = 13
    row = 3
    if num:
        row = write_block(sh, row, "Numeric variables", numeric_summary(df, num))
    for c in cat:
        row = write_block(sh, row, f"{c}: count and response rate", category_profile(df, c, target))
    if len(num) > 1:
        corr = df[num + [target]].apply(pd.to_numeric, errors="coerce").corr().round(3)
        row = write_block(sh, row, f"Correlation matrix (numeric variables and {target})",
                          corr.reset_index().rename(columns={"index": "Variable"}))

    charts = new_sheet(book, f"{name}_EDA_CHARTS"[:31])
    charts["A1"].value = f"EDA charts for {name}"
    charts["A1"].font.bold = True
    slots = [f"{c}{r}" for r in range(3, 3000, 18) for c in ("A", "G", "M")]
    i = 0
    for c in num:
        s = pd.to_numeric(df[c], errors="coerce").dropna()
        if s.empty:
            continue
        fig, ax = plt.subplots(figsize=(3.4, 2.6))
        lo, hi = s.quantile(0.01), s.quantile(0.99)
        ax.hist(s.clip(lo, hi), bins=30, color="#9DC3E6", edgecolor="#5B7A99")
        ax.set_title(f"{c} (1st-99th pct)", fontsize=9, weight="bold")
        ax.grid(True, alpha=0.3)
        place(charts, fig, slots[i], f"hist_{c}")
        i += 1
    for c in cat:
        prof = category_profile(df, c, target, top=12)
        fig, ax = plt.subplots(figsize=(3.4, 2.6))
        ax.bar(range(len(prof)), prof["Response rate %"], color="#9DC3E6", edgecolor="#5B7A99")
        ax.axhline(overall, color=ORANGE, lw=1.2, ls="--", label="overall")
        ax.set_xticks(range(len(prof)), prof["Value"], rotation=45, ha="right", fontsize=7)
        ax.set_title(f"Response rate % by {c}", fontsize=9, weight="bold")
        ax.legend(fontsize=7)
        place(charts, fig, slots[i], f"rate_{c}")
        i += 1
    msg = f"✅ EDA done for {name}: {len(num)} numeric, {len(cat)} categorical, {i} charts. See {sh.name} and {charts.name}."
    print(msg)
    set_status(book, msg, ok=True)
    sh.activate()


def decile_table(probs: np.ndarray, actual: np.ndarray) -> pd.DataFrame:
    d = pd.DataFrame({"prob": probs, "actual": actual})
    d["Decile"] = pd.qcut(d["prob"].rank(method="first", ascending=False), 10, labels=False) + 1
    g = d.groupby("Decile").agg(Obs=("actual", "size"), Min_Score=("prob", "min"), Max_Score=("prob", "max"),
                                Avg_Score=("prob", "mean"), Responders=("actual", "sum")).reset_index()
    g["Response_Rate_%"] = (g["Responders"] / g["Obs"] * 100).round(2)
    g["Cum_Responders"] = g["Responders"].cumsum()
    g["Cum_Response_Capture_%"] = (g["Cum_Responders"] / g["Responders"].sum() * 100).round(2)
    g["Lift"] = (g["Response_Rate_%"] / (actual.mean() * 100)).round(2)
    return g


@script(name="5. Build Model", button="[btn_model]MASTER!C28")
async def score_and_deciles(book: xw.Book):
    """XGBoost response model: 70 / 30 split, Gini, deciles, ROC and gain charts, scores per row."""
    df, name, id_col, target, err = load_analysis_table(book)
    if df is None:
        set_status(book, f"❌ {err}", ok=False)
        return
    set_status(book, "⏳ Loading XGBoost and training the model ...", ok=True)
    await book.flush()
    from sklearn.metrics import confusion_matrix, f1_score, precision_score, recall_score, roc_auc_score, roc_curve
    from sklearn.model_selection import train_test_split
    from xgboost import XGBClassifier

    num, cat = split_columns(df, id_col, target)
    X = pd.get_dummies(df[num + cat].assign(**{c: df[c].fillna("(blank)").astype(str) for c in cat}),
                       columns=cat, drop_first=True, dtype=float)
    X = X.apply(pd.to_numeric, errors="coerce")
    y = df[target].to_numpy()
    X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.3, random_state=42, stratify=y)
    model = XGBClassifier(max_depth=3, n_estimators=100, learning_rate=0.1, eval_metric="logloss", random_state=42)
    model.fit(X_tr, y_tr)
    p_tr, p_te = model.predict_proba(X_tr)[:, 1], model.predict_proba(X_te)[:, 1]
    print(f"🌲 Trained on {len(X_tr):,} rows, {X.shape[1]} features; tested on {len(X_te):,}")

    def metrics(yt, p) -> dict[str, float]:
        pred = (p >= 0.5).astype(int)
        auc = roc_auc_score(yt, p)
        return {"AUC": auc, "Gini": 2 * auc - 1, "Accuracy at 0.5": float((pred == yt).mean()),
                "Precision at 0.5": precision_score(yt, pred, zero_division=0),
                "Recall at 0.5": recall_score(yt, pred, zero_division=0), "F1 at 0.5": f1_score(yt, pred, zero_division=0)}
    m_tr, m_te = metrics(y_tr, p_tr), metrics(y_te, p_te)
    perf = pd.DataFrame({"Metric": list(m_tr), "Train": [round(v, 4) for v in m_tr.values()],
                         "Test": [round(v, 4) for v in m_te.values()]})
    cms = []
    for label, yt, p in (("Train", y_tr, p_tr), ("Test", y_te, p_te)):
        cm = confusion_matrix(yt, (p >= 0.5).astype(int), labels=[0, 1])
        cms.append(pd.DataFrame({f"{label}: actual / predicted": ["Actual 0", "Actual 1"],
                                 "Predicted 0": cm[:, 0], "Predicted 1": cm[:, 1]}))
    imp = pd.DataFrame({"Feature": X.columns, "Importance (gain share)": model.feature_importances_})
    imp = imp.sort_values("Importance (gain share)", ascending=False).head(20).round(4)

    ev = new_sheet(book, f"{name}_MODEL"[:31])
    ev["A1"].value = (f"XGBoost model for {target} on {name}: {len(X_tr):,} train / {len(X_te):,} test rows, "
                      f"{X.shape[1]} features (max_depth=3, 100 trees)")
    ev["A1"].font.bold = True
    row = write_block(ev, 3, "Performance (Gini = 2 x AUC - 1)", perf)
    for cmdf in cms:
        row = write_block(ev, row, cmdf.columns[0].split(":")[0] + " confusion matrix at 0.5", cmdf)
    write_block(ev, row, "Top 20 features", imp)

    dec = new_sheet(book, f"{name}_DECILES"[:31])
    dec["A1"].value = "Deciles: rows ranked by score, decile 1 = highest scores"
    dec["A1"].font.bold = True
    row = write_block(dec, 3, "Train deciles", decile_table(p_tr, y_tr))
    write_block(dec, row, "Test deciles", decile_table(p_te, y_te))

    p_all = model.predict_proba(X)[:, 1]
    scores = pd.DataFrame({id_col or "Row": df[id_col] if id_col else np.arange(1, len(df) + 1),
                           target: y, "Score_Probability": p_all.round(6),
                           "Score_Decile": pd.qcut(pd.Series(p_all).rank(method="first", ascending=False), 10, labels=False) + 1,
                           "Sample": np.where(df.index.isin(X_te.index), "Test", "Train")})
    sc = new_sheet(book, f"{name}_SCORES"[:31])
    sc["A1"].value = f"Score per row (the input table {name} is not changed)"
    sc["A1"].font.bold = True
    sc["A3"].options(index=False).value = scores
    sc.tables.add(source=sc["A3"].resize(len(scores) + 1, scores.shape[1]))

    gr = new_sheet(book, f"{name}_MODEL_CHARTS"[:31])
    gr["A1"].value = "Model charts"
    gr["A1"].font.bold = True
    fig, ax = plt.subplots(figsize=(5, 4))
    for label, yt, p, col in (("Train", y_tr, p_tr, BLUE), ("Test", y_te, p_te, ORANGE)):
        fpr, tpr, _ = roc_curve(yt, p)
        ax.plot(fpr, tpr, color=col, label=f"{label} (AUC {roc_auc_score(yt, p):.3f})")
    ax.plot([0, 1], [0, 1], ls="--", color=GREY, label="Random")
    ax.set(title="ROC curve", xlabel="False positive rate", ylabel="True positive rate")
    ax.legend(fontsize=8)
    ax.grid(True, alpha=0.3)
    place(gr, fig, "A3", "ROC_Curve")
    fig, ax = plt.subplots(figsize=(5, 4))
    for label, yt, p, col in (("Train", y_tr, p_tr, BLUE), ("Test", y_te, p_te, ORANGE)):
        order = np.argsort(-p)
        cum = np.cumsum(yt[order]) / yt.sum()
        ax.plot(np.arange(1, len(p) + 1) / len(p), cum, color=col, label=label)
    ax.plot([0, 1], [0, 1], ls="--", color=GREY, label="Random")
    ax.set(title="Cumulative gain", xlabel="Share of rows (highest scores first)", ylabel="Share of responders")
    ax.legend(fontsize=8)
    ax.grid(True, alpha=0.3)
    place(gr, fig, "I3", "Gain_Curve")

    msg = (f"✅ Model done: test Gini {m_te['Gini']:.3f} (train {m_tr['Gini']:.3f}). "
           f"See {ev.name}, {dec.name}, {sc.name} and {gr.name}.")
    print(msg)
    set_status(book, msg, ok=True)
    ev.activate()
