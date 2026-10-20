# xlwings Lite: Coding Guidelines for AI Coders (Humans Welcome)

**Version 3.0 · 20 October 2026.** The latest copy is always at https://www.tigzig.com/files/xlwings/AI_CODER_INSTRUCTIONS.md. If your copy shows a lower version number, download the latest. Changes are listed under Version History at the end.

**Tested on:** xlwings Lite add-in build 1.0.0.104, xlwings 0.37.7, Pyodide 314.0.1 (Python 3.14.2, pandas 3.0.2, numpy 2.4.3), Excel desktop on Windows. Every platform behaviour marked **[tested]** was run in real Excel on that build. Items marked **[docs]** come from the official documentation and were not re-run here.

**Official docs:** https://lite.xlwings.org (machine-readable: https://lite.xlwings.org/llms-full.txt). If something here disagrees with a newer build, trust the build and the docs.

---

## 0. THE RULES (check every script against these)

1. **Write DataFrames with `.options(index=False)`.** `sheet["A1"].value = df` writes the pandas index as an extra first column. [tested]
2. **Re-runnable sheets: delete, then add, with a loop.** `"Name" in book.sheets` is ALWAYS `False`, even for a sheet that exists. Check names with `[s.name for s in book.sheets]` or loop over the sheets. [tested]
3. **Size new ranges with `.resize()`, never `.expand()`.** Reads come from a snapshot taken when the script started, so `.expand()` and `used_range` cannot see data written in the same run. [tested]
4. **Find tables by name across sheets** with the `find_table_in_workbook()` helper (section 5.2). `table.sheet` raises `AttributeError`; use `table.parent` if you have the table already. [tested]
5. **Table names are unique across the whole workbook.** If a user renamed an old output sheet, its table keeps its name, and the next run's `tables.add(name=...)` fails with a generic `InvalidArgument`. Omit `name=` unless the user asked for one, and when you do name a table, remove any existing table of that name first (section 5.3). [tested]
6. **Excel-side errors surface at flush, not at the call.** A `try/except` around `sheet.tables.add()` catches nothing unless you `await book.flush()` inside the `try`. That needs an `async def` script (section 3). [tested]
7. **Keep text as text.** A date-like string (`"2026-01-02"`) becomes a date and `"00123"` becomes `123`. Set `number_format = "@"` on the target range BEFORE writing (section 4.3). [tested]
8. **`np.inf` writes as a blank cell, silently.** NaN, None and `pd.NA` also write as blank. Replace infinities before writing if they matter. [tested]
9. **Write formulas through `.value`:** `rng.value = "=A2+A3"`. The `.formula` setter is not implemented. [tested]
10. **Formatting you can set:** `.color`, `.font.bold/italic/color/size/name`, `.number_format`, `Range.autofit()`. **You cannot:** `Sheet.autofit()`, `column_width`, `row_height`, `merge()`, `wrap_text`. **You cannot READ** any format (color, number format, font, formula, column width). [tested]
11. **Mark every user-runnable function with `@script`**, with `book: xw.Book` as its first argument. [docs]
12. **pandas 3 is the default for new workbooks** (pandas 2 still runs on older ones). Write code that works on both: no chained `inplace=True`, no `errors="ignore"` (section 11). [tested]
13. **Custom functions (`@func`): take cell inputs as `typing.Any` and convert inside**, because blank cells arrive as `0` and text in a numeric hint fails before your code runs (section 8.2). [docs]
14. **Activate the final output sheet** with `sheet.activate()` so the user sees the result. When the script runs from a sheet button, the add-in jumps back to the button's sheet afterwards (section 3.4). [tested]
15. **Print progress** at each major step. `print()` output appears in the Output pane when the script ends, or right away after `await book.flush()`. [tested]

---

## 1. Where Your Code Runs

- **Engine:** Pyodide (Python compiled to WebAssembly) inside the xlwings Lite task pane, in the browser sandbox. No local Python, no local file system except what the user shares (section 9). [docs]
- **Source files:** the editor holds `main.py` plus any other source files the user creates via the **Source files** menu. They live under `/src`, travel inside the workbook, and import each other normally. Scripts and functions in any workbook module show up in the Run dropdown. [docs]
- **Notebooks:** `*.nb.py` files run in cells like Jupyter (top-level `await` allowed). [docs]
- **Tests:** pytest runner for `test_*.py` / `*_test.py` files (add `pytest` to requirements). [docs]
- **Dependencies:** `requirements.txt` in the editor (section 10).
- **Output:** `print()` goes to the Output pane.
- **Trust:** a workbook the user didn't author asks "Do you trust the Authors of this workbook?" before any code runs. [docs]
- **Autostart:** the add-in can start with the workbook (Settings > Workbook). [docs]
- **App Mode:** hides the editor and shows only script buttons and forms for end users (Ctrl+Shift+M toggles). [docs]
- **Pyodide version** is per workbook (Settings > Workbook). 314.x = Python 3.14 + pandas 3; 0.27.5 = older stack with pandas 2. Changing it can change package major versions. [docs]

---

## 2. Scripts

### 2.1 Basic syntax
```python
import xlwings as xw
from xlwings import script

@script
def hello_world(book: xw.Book):
    sheet = book.sheets[0]
    sheet["A1"].value = "Hello xlwings!"
```

### 2.2 Script arguments and button label [docs]
Scripts may take extra arguments after `book`. In App Mode they render as a form (text box, number box, checkbox, dropdown for `Literal`, date picker for `datetime.date`). A script whose extra arguments lack defaults runs only in App Mode.
```python
import datetime as dt
from typing import Annotated, Literal

@script(name="Create Report")          # button label
def create_report(
    book: xw.Book,
    title: str = "Report",
    start: dt.date = dt.date(2026, 1, 1),        # needs xlwings >= 0.36.11
    rows: Annotated[int, {"label": "Number of rows"}] = 10,
    currency: Literal["USD", "EUR"] = "USD",
):
    ...
```
Arguments must be JSON-serializable. Keyword-only arguments (after `*`) and `**kwargs` are not supported.

### 2.3 Configuration [docs]
```python
@script(
    include=["Sheet1"],               # send only these sheets' content to Python
    exclude=["BigData"],              # or: send everything except these
    button="[mybutton]Sheet1!A1",     # sheet button (section 2.4)
    show_taskpane=True,               # open the task pane on button click
    name="My Button",                 # label in the Run dropdown / App Mode
)
```
By default the whole workbook's values are sent to Python before the script runs. For big workbooks use `include`/`exclude`, or the async API (section 3), to keep it fast.

### 2.4 Sheet buttons
1. Insert a shape, give it a name in the Name Box (e.g. `xlwings_button`).
2. Right-click > Link > Place in This Document > a cell hidden behind the shape (e.g. `B4`).
3. `@script(button="[xlwings_button]Sheet1!B4")`.

How it works: the add-in runs the script when that cell is SELECTED, then selects the cell below. [tested]
- Don't leave the linked cell selected when saving.
- Scripts that depend on the user's current selection don't work from a button (the click changes the selection).
- Because the add-in re-selects the cell under the button after the run, **an `activate()` at the end of a button-run script is undone**: Excel returns to the button's sheet. [tested]
- Excel on the web can't add hyperlinks to shapes, but buttons set up on desktop work on the web. [docs]

---

## 3. The Workbook Engine: Snapshot, Queue, Flush

This is the model behind most "why didn't that work" problems. [tested]

1. **Snapshot at start.** Before the script runs, xlwings Lite loads the workbook state. Every sync read (`.value`, `.expand()`, `used_range`, `book.sheets`) reflects that snapshot, not changes made during the run.
2. **Writes are queued.** Assignments are applied to Excel in order when the script ends (or on `await book.flush()`).
3. **Excel-side errors appear at flush time.** An invalid write (overlapping table, duplicate table name, unsupported operation) raises a `JsException` when the queue is applied, not on the line that queued it. In a plain `def` script that means the script fails at the end, and the traceback doesn't point at the line that caused it.

The async API gives you control:
```python
@script
async def my_script(book: xw.Book):
    sheet = book.sheets["Data"]
    sheet["A1"].value = "Hello"
    await book.flush()                         # apply queued writes now (and show prints)
    print(await sheet["A1"].get_value())       # read the CURRENT value from Excel -> 'Hello'
    await sheet.load()                         # refresh the snapshot for this sheet
    print(sheet["A1"].expand().address)        # now sees data written this run
```
- `await book.flush()`: apply queued writes now. Use it for live progress during long runs, before reading back what you wrote, before using a file produced by `to_png()`, and inside `try` blocks to catch Excel-side errors. [tested]
- `await rng.get_value()`: read one range live, without the snapshot. [tested]
- `await book.load()` / `await sheet.load()`: refresh the snapshot. [tested]
- `xw.BookAsync` type hint (instead of `xw.Book`): skips loading all values up front, for very large workbooks. You must then read with `await rng.get_value()`. [docs]
- `await book.sync()` is the deprecated name of `flush()`. [docs]

**Formulas compute in the same run:** write a formula, `await book.flush()`, then `await rng.get_value()` returns the result. [tested]

Catching an Excel-side error:
```python
@script
async def build(book: xw.Book):
    sheet = book.sheets["Report"]
    try:
        sheet.tables.add(source=sheet["A1"].resize(5, 3))
        await book.flush()                     # the error, if any, is raised HERE
    except Exception as e:
        print(f"Table creation failed: {e}")
```

---

## 4. Writing Data

### 4.1 DataFrames
```python
sheet["A1"].options(index=False).value = df            # always index=False
df = sheet["A1"].options(pd.DataFrame, index=False, expand="table").value  # read
```
Writing a DataFrame with mixed-type columns (e.g. `"< 600"` and `600` in one column), with `""` in a float column, or as a 2D Python list (`[["Label", 1.6], ["P", 0.8]]`) all write correctly on the tested build. Older builds failed on some of these. [tested]

Still good practice, for pandas reasons: use `None` (not `""`) for missing numbers so the column stays numeric, and convert genuinely mixed label columns with `.astype(str)` if you will sort or filter them.

### 4.2 What values become in Excel [tested]
| Python | Excel cell |
|---|---|
| `np.nan`, `None`, `pd.NA` | blank |
| `np.inf` | **blank (silent)** |
| `np.int64`, `Int64` | number |
| `bool` | TRUE / FALSE |
| `datetime.date`, `pd.Timestamp` | date / date-time |
| `"2026-01-02"` (string) | **converted to a date** |
| `"00123"` (string) | **converted to 123** |
| `"=A1+A2"` (string) | formula |

### 4.3 Keeping text as text [tested]
```python
rng = sheet["A2"].resize(len(df), 1)
rng.number_format = "@"                    # format as Text BEFORE writing
rng.options(transpose=True).value = df["ACCOUNT_ID"].tolist()
# Single cells: a leading apostrophe also works: sheet["A1"].value = "'00123"
```

### 4.4 Formulas
```python
sheet["C2"].value = "=A2*B2"     # works
# sheet["C2"].formula = "=A2*B2"  # NOT implemented (fails at flush)
```

---

## 5. Sheets and Tables

### 5.1 Re-runnable output sheets [tested]
```python
def fresh_sheet(book, name):
    for s in book.sheets:
        if s.name == name:
            s.delete()
    return book.sheets.add(name)

# WRONG: always False, even when the sheet exists
# if "Results" in book.sheets: ...
```

### 5.2 Finding a table [tested]
```python
from typing import Any, Optional, Tuple

def find_table_in_workbook(book: xw.Book, table_name: str) -> Tuple[Optional[xw.Sheet], Any]:
    """Search all sheets for a table. Returns (sheet, table) or (None, None)."""
    for sheet in book.sheets:
        if table_name in sheet.tables:          # membership by NAME works for tables
            return sheet, sheet.tables[table_name]
    return None, None

sheet, table = find_table_in_workbook(book, "Portfolio")
df = table.range.options(pd.DataFrame, index=False).value
```
- `table.sheet` -> `AttributeError`. `table.parent` returns the sheet.
- Useful table members: `table.range`, `table.data_body_range`, `table.update(df, index=False)` (rewrites the table and resizes it).

### 5.3 Creating tables [tested]
```python
sheet["A1"].options(index=False).value = df
rng = sheet["A1"].resize(df.shape[0] + 1, df.shape[1])   # explicit size, never .expand()
sheet.tables.add(source=rng)                             # let Excel name it (Table1, ...)
```
- **Overlap:** a table can't overlap another table (`InvalidArgument: A table can't overlap another table`).
- **Names are workbook-wide.** Naming is fine, including several named tables in a loop, as long as the names are new. When you must use a fixed name, clear it first:
```python
def drop_table_everywhere(book, name):
    for s in book.sheets:
        if name in s.tables:
            s.tables[name].delete()
```
- **Renamed output sheets:** a user renames `RESULTS` to `RESULTS_v1` to keep it. The table inside keeps its name. Next run, `tables.add(name="RES_TBL")` on the new `RESULTS` fails with `InvalidArgument: The argument is invalid or missing or has an incorrect format`, which doesn't mention the name. Use `drop_table_everywhere()` first, omit the name, or tell the user to copy sheets (Move or Copy > Create a copy) instead of renaming.

---

## 6. Formatting

```python
header = sheet["A1"].resize(1, df.shape[1])
header.color = "#4472C4"            # hex string or RGB tuple (68, 114, 196): both work [tested]
header.font.color = "#FFFFFF"       # always set a contrasting font colour with a fill
header.font.bold = True
sheet["B2"].resize(len(df), 1).number_format = "#,##0.00"
sheet["A1"].resize(df.shape[0] + 1, df.shape[1]).autofit()   # Range.autofit works [tested]
```
Not available on the tested build (they fail at flush with `JsException: TypeError: globalThis.callbacks[action.func] is not a function`): `Sheet.autofit()`, `column_width =`, `row_height =`, `merge()`, `unmerge()`, `wrap_text =`, `.formula =`, `autofill()`, `Sheet.copy()`, `Book.save()`, `app.screen_updating`. [tested]

Reading formats is not available: `.color`, `.number_format`, `.font.*`, `.formula`, `.column_width`, `.current_region`, `.table` raise `NotImplementedError` suggesting async getters (`await rng.get_color()` etc.), but those raise `AttributeError: getRangeData` on the tested build. Track formatting state in Python instead of reading it back. [tested]

Presentation conventions: clean column names (`df.rename(columns=...)`) before writing; formal Excel tables for summaries; a fill always paired with a readable font colour.

---

## 7. Pictures and Plots

```python
import matplotlib
matplotlib.use("agg")
import matplotlib.pyplot as plt

fig, ax = plt.subplots(figsize=(8, 5))
ax.plot(df["DATE"], df["CLOSE"])
sheet.pictures.add(fig, name="PriceChart", update=True, anchor=sheet["H2"])  # pass the figure directly
plt.close(fig)
```
- `update=True` with a fixed `name` replaces the picture on re-run. [tested]
- `Range.to_png("/data/range.png")` works; call `await book.flush()` before reading the file in the same script. [tested]
- Native Excel charts can't be created (`Charts.add()` is missing). [docs]
- Custom functions can return a Matplotlib figure (xlwings >= 0.36.8). [docs]

---

## 8. Custom Functions (`@func`)

### 8.1 Basics [docs]
```python
import pandas as pd
from xlwings import func, arg, ret

@func
def hello(name):
    return f"Hello {name}!"            # =HELLO("World") or =HELLO(A1)

@func
@arg("df", pd.DataFrame, index=False)
@ret(index=False)
def describe(df):
    return df.describe().reset_index()

@func
def myfunction(df: pd.DataFrame) -> pd.DataFrame:   # type hints work instead of decorators
    return df
```
Also available: `*args`, docstrings and `@arg(doc=...)` for Excel IntelliSense, `typing.Literal` parameters (dropdown while typing), `@func(name="MyName")`, namespaces (`__xlwings_func_namespace__`), dynamic arrays, volatile functions, object handles, streaming functions, `xw.Caller` (calling cell, xlwings >= 0.36.17), `xw.WithScript()` to run a script after the function returns (xlwings >= 0.36.17). See the Custom Functions page of the docs.

### 8.2 Inputs: blanks, text and the `Any` pattern [docs]
- Empty cells arrive as `0` (not `None`). To get `None`, the cell must contain `=""`.
- Integers arrive as `int`. Date cells need `dt.datetime` or `parse_dates`.
- With a specific hint like `float`, a cell holding text fails conversion before your code runs, and the cell shows `#VALUE!` (hover the cell's warning icon for the message).
- Robust pattern: accept `Any` and convert inside.
```python
from typing import Any

def _to_float(value: Any, default: float = 0.0) -> float:
    if value is None:
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default

@func
def double_it(value: Any):
    return _to_float(value) * 2
```

### 8.3 Outputs [docs]
- `np.nan` / `pd.NA` returned from a function show as `#NUM!` (scripts write them as blank).
- Return `"#N/A"`, `"#VALUE!"` etc. to produce real error cells.
- Functions are asynchronous: the cell shows `#BUSY!` while calculating.

### 8.4 Debugging `#VALUE!`
1. Does the function appear in Excel's autocomplete? If not, `main.py` (or a module) has an error: check the Problems pane / Output pane.
2. Use a probe function to see exactly what Python receives:
```python
@func
def debug_input(cell: Any) -> str:
    return f"type={type(cell).__name__} value={cell!r}"
```

---

## 9. Files [docs]

- `/data` is the working directory. Relative paths read and write there. Files appear under Menu > Files > Import/Export, where users can also upload files (drag and drop). They are deleted when the Office cache is cleared.
- **Local Folders** (Excel desktop on Windows only): the user mounts a folder from their computer at a path such as `/my_data`; code reads and writes real files there. Paths are case-sensitive: the default mount path is the folder name as typed, e.g. `/My_Data`, not `/my_data`. [tested]
- `/src` holds the workbook's source files (text only). Read one with `(Path(__file__).parent / "config.json").read_text()`.
- **Personal modules:** a folder mounted (or uploaded) as `/personal` loads its scripts and functions into every workbook, like `Personal.xlsb`; `personal/projects/<name>/` loads only for workbooks whose filename contains `<name>`. Restart the add-in after changes.

---

## 10. Dependencies

- List packages in `requirements.txt`, then restart the add-in (Menu > Restart). [docs]
- xlwings Lite records exact installed versions in `requirements-pinned.txt` (Pyodide 314+) and reinstalls them on every restart. Don't pin exact versions in `requirements.txt`; use an upper bound if needed (`xlwings<1.0.0`). [docs]
- Packages must be pure Python or built for Pyodide. Not available: anything needing `multiprocessing`/`threading`, PyTorch, direct TCP clients. [docs]
- Not installed by default on the tested build: `requests`, `httpx`, `polars`. [tested]
- Quick trial without editing requirements: `import micropip; await micropip.install("requests")` inside an async script works. For a finished tool, put the package in `requirements.txt`. [tested]
- Private wheels: a CORS-enabled URL in `requirements.txt`, or `/data/my-package.whl` after uploading. [docs]

---

## 11. pandas 3 (default for new workbooks)

pandas 3 uses Copy-on-Write. Two patterns that worked in pandas 2 now fail: [tested]
```python
# SILENTLY DOES NOTHING in pandas 3 (ChainedAssignmentError warning only):
df["col"].fillna(0, inplace=True)
# CORRECT in pandas 2 and 3:
df["col"] = df["col"].fillna(0)

# RAISES ValueError in pandas 3:
pd.to_numeric(s, errors="ignore")
df.apply(pd.to_numeric, errors="ignore")
# CORRECT: coerce, or convert only columns that fully parse
def to_numeric_where_possible(df):
    out = df.copy()
    for c in out.columns:
        conv = pd.to_numeric(out[c], errors="coerce")
        if conv.notna().sum() == out[c].notna().sum():
            out[c] = conv
    return out
```

---

## 12. Web Requests and Databases

### 12.1 CORS [docs]
Requests run from the user's browser, origin `https://addin.xlwings.org` (or the self-hosted domain). The server must send `Access-Control-Allow-Origin` for that origin. GitHub raw files work. If you don't control the server, use a CORS proxy you own. For IP allow-lists, allow the user's IP (requests don't go through xlwings servers).

### 12.2 Making requests [tested]
```python
from pyodide.http import pyfetch          # built in, no install needed

@script
async def fetch_data(book: xw.Book):
    r = await pyfetch("https://api.example.com/data", headers={"Accept": "application/json"})
    if not r.ok:
        raise RuntimeError(f"API error {r.status}")
    data = await r.json()                 # or: await r.text()
```
`requests` (sync) and `httpx` (async) also work once installed. `pd.read_csv("https://raw.githubusercontent.com/...")` works for CORS-enabled URLs. [docs]
Secrets go in environment variables (section 13), never in code.

Parsing a pipe-delimited text response (pandas 2 and 3):
```python
def parse_pipe_delimited(text: str) -> pd.DataFrame:
    lines = [ln for ln in text.strip().split("\n") if ln.strip()]
    if not lines:
        return pd.DataFrame()
    headers = [h.strip() for h in lines[0].split("|")]
    rows = [[c.strip() or None for c in ln.split("|")] for ln in lines[1:]]  # empty -> missing
    return to_numeric_where_possible(pd.DataFrame(rows, columns=headers))
```

### 12.3 Databases [docs]
No TCP sockets, so no direct PostgreSQL/MySQL connections. Options: a web API in front of the database (FastAPI, PostgREST, Supabase, Oracle ORDS, MySQL REST Service, NocoDB), or download a SQLite file and query it with `sqlite3`.

---

## 13. Environment Variables and Security [docs]

- Two scopes: **Local** (browser storage, all workbooks, deleted with the Office cache, back it up) and **Workbook** (saved inside the file; not for secrets). Local overrides Workbook. Restart the add-in after setting one.
- Read with `os.getenv("OPENAI_API_KEY")` and raise a clear error if missing.
- Local-scope secrets are readable by any workbook the user trusts, so only trust workbooks from known sources.

---

## 14. Analysis Conventions (house style)

These are our standing conventions for analysis scripts. They are style and correctness rules, not platform limits.

1. **Separate calculation from Excel I/O.** Calculation functions take and return DataFrames, with no `xw` objects. One final block writes everything. A failure in the maths then never leaves the workbook half-written, and errors are easy to place.
2. **Pivoted report tables: build from a list of dicts.** `.unstack()`/pivot fails or drops columns when a group is missing. Loop over variables and groups, `report_rows.append({...})`, then `pd.DataFrame(report_rows)`. A plain `groupby().agg().reset_index()` is fine when nothing is reshaped.
3. **Quantile bins: `pd.qcut(..., duplicates="drop")`.** Without it, a column with few unique values raises `ValueError: Bin edges must be unique`. [tested]
4. **Convert, then compute statistics, then impute, as separate steps:**
```python
df["X"] = pd.to_numeric(df["X"], errors="coerce")
if df["X"].isna().any():
    df["X"] = df["X"].fillna(df["X"].median())
# WRONG: pd.to_numeric(df["X"], errors="coerce").fillna(df["X"].median())
#        -> the median is taken from the ORIGINAL, unconverted column
```
5. **Balanced samples (test/control):** `sklearn.model_selection.train_test_split(..., stratify=...)`.
6. **Narrate** progress with `print()` and finish with a clear success line.
7. **Activate** the last output sheet (see rule 14 for the button caveat).

---

## 15. Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `JsException: InvalidArgument: The argument is invalid or missing...` at the end of a run | Duplicate table name (often from a renamed output sheet), or an invalid value in a write | `drop_table_everywhere()` before naming; omit `name=`; flush after each step to locate it |
| `InvalidArgument: A table can't overlap another table` | New table over an existing one | Delete the old sheet/table first |
| `JsException: TypeError: globalThis.callbacks[action.func] is not a function` | The operation isn't implemented in Lite (e.g. `column_width`, `merge`, `.formula =`) | Remove it (section 6) |
| `NotImplementedError: Reading ... synchronously isn't supported` | Reading a format/formula/width | Track the state in Python |
| `AttributeError: 'Table' object has no attribute 'sheet'` | `table.sheet` | `find_table_in_workbook()` or `table.parent` |
| A sheet "doesn't exist" check never fires | `"Name" in book.sheets` is always False | Loop over sheet names |
| Data written this run reads back empty / `.expand()` returns one cell | Snapshot taken at script start | `.resize()`; or `await book.flush()` then `await rng.get_value()` / `await sheet.load()` |
| Nothing appears in Excel until the script ends | Writes are queued | `await book.flush()` (async script) |
| `try/except` didn't catch an Excel error | Error raised at flush, after the `try` | Flush inside the `try` |
| `ValueError: invalid error value specified` | pandas 3 removed `errors="ignore"` | Section 11 |
| A `fillna(..., inplace=True)` had no effect | pandas 3 Copy-on-Write | Assign the result back |
| Leading zeros / IDs / codes changed in Excel | Auto-conversion of strings | `number_format = "@"` before writing |
| `ModuleNotFoundError` | Package not in `requirements.txt` | Add it and restart |
| CORS error on a request | Server doesn't allow `https://addin.xlwings.org` | Add the header, or use your own CORS proxy |
| Custom function shows `#VALUE!` | Type conversion failed, or the module has an error | Section 8.4 |

---

## 16. Reference Script Pattern

```python
import pandas as pd
import xlwings as xw
from xlwings import script


@script
async def summarise_portfolio(book: xw.Book):
    print("Loading data...")
    src_sheet, table = find_table_in_workbook(book, "Portfolio")
    if table is None:
        raise ValueError("Table 'Portfolio' not found. Create it via Insert > Table.")
    df = table.range.options(pd.DataFrame, index=False).value

    # PHASE 1: calculation, no Excel objects
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    summary = (
        df.groupby("Category", as_index=False)
        .agg(Count=("Value", "size"), Total=("Value", "sum"))
        .rename(columns={"Category": "Asset Class"})
    )
    print(f"Calculated {len(summary)} rows.")

    # PHASE 2: write
    out = fresh_sheet(book, "Summary")
    out["A1"].options(index=False).value = summary
    rng = out["A1"].resize(summary.shape[0] + 1, summary.shape[1])
    header = out["A1"].resize(1, summary.shape[1])
    header.color = "#4472C4"
    header.font.color = "#FFFFFF"
    header.font.bold = True
    try:
        out.tables.add(source=rng)
        await book.flush()                 # surface any Excel-side error here
    except Exception as e:
        print(f"Table creation failed: {e}")
    rng.autofit()
    out.activate()
    print("Done: results on sheet 'Summary'.")
```
(`find_table_in_workbook` from section 5.2, `fresh_sheet` from section 5.1.)

---

## Version History

- **3.0 (20 October 2026):** Rewritten and re-tested in real Excel against add-in build 1.0.0.104, xlwings 0.37.7 and pandas 3. Corrects several version 2 rules that no longer hold (RGB colours, mixed-type columns, 2D list writes, named tables in loops, single-argument scripts) and adds the snapshot/flush model, pandas 3 notes, text auto-conversion and the current API support table.
- **2.x (November 2025):** Compressed edition of the original guide.
