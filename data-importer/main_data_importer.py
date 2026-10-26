# xlwings Lite Data Importer v2 - import files from links or with your own tokens, into DuckDB or as-is
#
# =============================================================================
# CODE MAP
# =============================================================================
#   === SECTION: IMPORTS ===
#   === SECTION: OAUTH ===            Dropbox / Google Drive access tokens from refresh tokens
#   === SECTION: HELPERS ===
#   === SECTION: URL_ROUTING ===       proxy, Google Drive / Dropbox link conversion, file type detection
#   === SECTION: DATA_PROFILING ===    column statistics in DuckDB
#   === SECTION: SCHEMA_DISPLAY ===    result sheets PARQUET / CSV / JSON / DUCKDB
#   === SECTION: SHEET_LAYOUT ===      input cells and status messages
#   === SECTION: DOWNLOAD ===
#   === SECTION: PROCESS ===
#   === SECTION: SCRIPTS ===           the four import buttons
#   === SECTION: REMOTE_LOADER ===
#   === SECTION: SCRIPT_RUN_STATS ===  sample statistics (cricket / IMDB)
#
# =============================================================================
# SCRIPTS (buttons)
# =============================================================================
#   1. Link: Import to DuckDB    SHARE_LINK_ACCESS: B5 link
#   2. Link: Download As-Is      SHARE_LINK_ACCESS: B5 link, B8 file name, B9 folder
#   3. Token: Import to DuckDB   TOKEN_ACCESS: B5 provider, B6 path / ID / URL, B7 auth proxy (GitHub releases)
#   4. Token: Download As-Is     TOKEN_ACCESS: same, plus B9 file name, B10 folder
#   5. Run Stats (sample data)   uses the last imported file (cricket or IMDB samples)
#   Messages go to B12 on each sheet (blue = done, red = error).
#
# =============================================================================
# ENVIRONMENT VARIABLES (token access only; xlwings Lite: Settings > Environment Variables, scope Local)
# =============================================================================
#   Dropbox:      DROPBOX.REFRESH_TOKEN, DROPBOX.APP_KEY, DROPBOX.APP_SECRET
#   Google Drive: GDRIVE.REFRESH_TOKEN, GDRIVE.CLIENT_ID, GDRIVE.CLIENT_SECRET
#   GitHub:       GITHUB.PAT (personal access token with read access to the repo)
#
# Data files: .parquet, .duckdb, .sqlite / .db, .csv, .tsv, .txt, .pipe, .psv, .json
# As-Is download: any file type.


# === SECTION: IMPORTS ===
import xlwings as xw
from xlwings import script
import pandas as pd
from typing import Any, Tuple, Optional
from pyodide.http import pyfetch
import duckdb
import tempfile
import os
import time
import json


# === SECTION: OAUTH ===
# =============================================================================
# OAUTH TOKEN EXCHANGE FUNCTIONS
# =============================================================================

_dropbox_token = None
_gdrive_token = None


async def get_dropbox_token() -> str:
    """Get fresh Dropbox access token using refresh token from environment."""
    global _dropbox_token

    if _dropbox_token is not None:
        return _dropbox_token

    refresh_token = os.environ.get("DROPBOX.REFRESH_TOKEN")
    app_key = os.environ.get("DROPBOX.APP_KEY")
    app_secret = os.environ.get("DROPBOX.APP_SECRET")

    if not refresh_token:
        raise ValueError("DROPBOX.REFRESH_TOKEN not found in environment variables")
    if not app_key:
        raise ValueError("DROPBOX.APP_KEY not found in environment variables")
    if not app_secret:
        raise ValueError("DROPBOX.APP_SECRET not found in environment variables")

    print("   Getting Dropbox access token...")
    token_data = (
        f"grant_type=refresh_token"
        f"&refresh_token={refresh_token}"
        f"&client_id={app_key}"
        f"&client_secret={app_secret}"
    )

    response = await pyfetch(
        "https://api.dropbox.com/oauth2/token",
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        body=token_data
    )

    if not response.ok:
        error_text = await response.text()
        raise ValueError(f"Dropbox token exchange failed: {response.status} - {error_text[:200]}")

    token_json = await response.json()
    _dropbox_token = token_json["access_token"]
    print("   Dropbox access token obtained")

    return _dropbox_token


async def get_gdrive_token() -> str:
    """Get fresh Google Drive access token using refresh token from environment."""
    global _gdrive_token

    if _gdrive_token is not None:
        return _gdrive_token

    refresh_token = os.environ.get("GDRIVE.REFRESH_TOKEN")
    client_id = os.environ.get("GDRIVE.CLIENT_ID")
    client_secret = os.environ.get("GDRIVE.CLIENT_SECRET")

    if not refresh_token:
        raise ValueError("GDRIVE.REFRESH_TOKEN not found in environment variables")
    if not client_id:
        raise ValueError("GDRIVE.CLIENT_ID not found in environment variables")
    if not client_secret:
        raise ValueError("GDRIVE.CLIENT_SECRET not found in environment variables")

    print("   Getting Google Drive access token...")
    token_data = (
        f"grant_type=refresh_token"
        f"&refresh_token={refresh_token}"
        f"&client_id={client_id}"
        f"&client_secret={client_secret}"
    )

    response = await pyfetch(
        "https://oauth2.googleapis.com/token",
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        body=token_data
    )

    if not response.ok:
        error_text = await response.text()
        raise ValueError(f"Google token exchange failed: {response.status} - {error_text[:200]}")

    token_json = await response.json()
    _gdrive_token = token_json["access_token"]
    print("   Google Drive access token obtained")

    return _gdrive_token


# === SECTION: HELPERS ===
# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def find_table_in_workbook(book: xw.Book, table_name: str) -> Tuple[Optional[xw.Sheet], Any]:
    """Searches all sheets for a table and returns both the sheet and table objects."""
    for sheet in book.sheets:
        if table_name in sheet.tables:
            return sheet, sheet.tables[table_name]
    return None, None


def scalar(conn, query: str):
    """First value of the first row of a query."""
    row = conn.execute(query).fetchone()
    return row[0] if row else None


def ensure_sheet_exists(book: xw.Book, sheet_name: str) -> xw.Sheet:
    """Ensures a sheet exists, deleting any existing one first for re-runnability."""
    for s in book.sheets:
        if s.name == sheet_name:
            s.delete()
            break
    return book.sheets.add(name=sheet_name)


def get_parquet_temp_path() -> str:
    """Returns the standard temp path for downloaded parquet files."""
    return os.path.join(tempfile.gettempdir(), "imported_data.parquet")


def get_duckdb_temp_path() -> str:
    """Returns the standard temp path for downloaded DuckDB files."""
    return os.path.join(tempfile.gettempdir(), "imported_database.duckdb")


def get_delimited_temp_path() -> str:
    """Returns the standard temp path for downloaded delimited files."""
    return os.path.join(tempfile.gettempdir(), "imported_data.csv")


def get_json_temp_path() -> str:
    """Returns the standard temp path for downloaded JSON files."""
    return os.path.join(tempfile.gettempdir(), "imported_data.json")


def get_raw_temp_path(filename: str) -> str:
    """Returns the temp path for raw file downloads with user-specified filename."""
    return os.path.join(tempfile.gettempdir(), filename)


def get_last_state_path() -> str:
    """Returns the path for the last import state file."""
    return os.path.join(tempfile.gettempdir(), "last_import_state.json")


def save_last_import_state(source_sheet: str, source_url: str, file_type: str, file_path: Optional[str] = None) -> None:
    """
    Saves the state of the last import operation.
    This allows run_stats and test functions to know what was last imported.
    """
    state = {
        "source_sheet": source_sheet,
        "source_url": source_url,
        "file_type": file_type,
        "file_path": file_path,
        "import_time": time.strftime("%Y-%m-%dT%H:%M:%S")
    }
    try:
        state_path = get_last_state_path()
        with open(state_path, 'w') as f:
            json.dump(state, f)
    except Exception as e:
        print(f"   Warning: Could not save import state: {e}")


def get_last_import_state() -> Optional[dict]:
    """
    Retrieves the state of the last import operation.
    Returns None if no state file exists.
    """
    state_path = get_last_state_path()
    if not os.path.exists(state_path):
        return None
    try:
        with open(state_path, 'r') as f:
            state = json.load(f)
        return state
    except Exception:
        return None


# === SECTION: URL_ROUTING ===
# =============================================================================
# URL ROUTING AND FILE TYPE DETECTION
# =============================================================================

PUBLIC_PROXY_URL = "https://github-proxy-auth.tigzig.com"


def is_github_url(url: str) -> bool:
    """Check if URL is a GitHub URL."""
    url_lower = url.lower()
    return 'github.com' in url_lower or 'raw.githubusercontent.com' in url_lower


def is_github_release_url(url: str) -> bool:
    """Check if URL is a GitHub release download URL."""
    url_lower = url.lower()
    return 'github.com' in url_lower and '/releases/' in url_lower


def needs_proxy(url: str, is_private: bool = False) -> bool:
    """Determines if a URL needs to go through a CORS proxy."""
    url_lower = url.lower()

    if is_private:
        return is_github_url(url_lower)

    if 'github.com' in url_lower and '/releases/' in url_lower:
        return True

    if 'drive.google.com' in url_lower or 'drive.usercontent.google.com' in url_lower:
        return True

    if 'dropbox.com' in url_lower:
        return True

    return False


def build_public_proxy_url(url: str) -> str:
    """Builds the public CORS proxy URL."""
    import urllib.parse
    encoded_url = urllib.parse.quote(url, safe='')
    return f"{PUBLIC_PROXY_URL}/?url={encoded_url}"


def build_private_proxy_url(url: str, auth_proxy_base: str) -> str:
    """Builds the auth proxy URL for private GitHub repos."""
    import urllib.parse
    encoded_url = urllib.parse.quote(url, safe='')
    auth_proxy_base = auth_proxy_base.rstrip('/')
    return f"{auth_proxy_base}/?url={encoded_url}"


def convert_google_drive_url(url: str) -> Tuple[str, bool]:
    """Converts Google Drive sharing URLs to direct download URLs."""
    import re
    url_lower = url.lower()

    if 'drive.google.com' not in url_lower and 'drive.usercontent.google.com' not in url_lower:
        return url, False

    if 'drive.usercontent.google.com' in url_lower and 'confirm=t' in url_lower:
        return url, False

    file_id = None

    match = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if match:
        file_id = match.group(1)

    if not file_id:
        match = re.search(r'/open\?id=([a-zA-Z0-9_-]+)', url)
        if match:
            file_id = match.group(1)

    if not file_id:
        match = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
        if match:
            file_id = match.group(1)

    if file_id:
        direct_url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t"
        return direct_url, True

    return url, False


def extract_gdrive_file_id(input_str: str) -> Optional[str]:
    """Extracts Google Drive file ID from a URL or returns the input if already a file ID."""
    import re

    input_str = input_str.strip()

    if 'http' in input_str.lower() or 'drive.google.com' in input_str.lower():
        match = re.search(r'/file/d/([a-zA-Z0-9_-]+)', input_str)
        if match:
            return match.group(1)

        match = re.search(r'/open\?id=([a-zA-Z0-9_-]+)', input_str)
        if match:
            return match.group(1)

        match = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', input_str)
        if match:
            return match.group(1)

        return None

    if re.match(r'^[a-zA-Z0-9_-]+$', input_str) and len(input_str) > 10:
        return input_str

    return None


def convert_dropbox_url(url: str) -> Tuple[str, bool]:
    """Converts Dropbox sharing URLs to direct download URLs."""
    import urllib.parse
    url_lower = url.lower()

    if 'dropbox.com' not in url_lower:
        return url, False

    parsed = urllib.parse.urlparse(url)
    query_params = urllib.parse.parse_qs(parsed.query)

    if query_params.get('dl', ['0'])[0] == '1':
        return url, False

    query_params['dl'] = ['1']
    new_query = urllib.parse.urlencode(query_params, doseq=True)
    new_url = urllib.parse.urlunparse((
        parsed.scheme, parsed.netloc, parsed.path,
        parsed.params, new_query, parsed.fragment
    ))
    return new_url, True


def detect_file_type(url: str) -> Optional[str]:
    """Detects file type from URL."""
    import urllib.parse
    parsed = urllib.parse.urlparse(url)

    query_params = urllib.parse.parse_qs(parsed.query)
    if 'filetype' in query_params:
        ft = query_params['filetype'][0].lower()
        if ft in ('parquet', 'duckdb', 'csv', 'tsv', 'txt', 'pipe', 'psv', 'json'):
            return ft

    path_lower = parsed.path.lower()
    if '.' in path_lower:
        filename = path_lower.split('/')[-1]
        if filename.endswith('.parquet'):
            return 'parquet'
        elif filename.endswith('.duckdb'):
            return 'duckdb'
        elif filename.endswith('.csv'):
            return 'csv'
        elif filename.endswith('.tsv'):
            return 'tsv'
        elif filename.endswith('.txt'):
            return 'txt'
        elif filename.endswith('.pipe') or filename.endswith('.psv'):
            return 'pipe'
        elif filename.endswith('.json'):
            return 'json'

    return None


def detect_file_type_from_path(path: str) -> Optional[str]:
    """Detects file type from file path (for Dropbox/GDrive paths)."""
    path_lower = path.lower()
    if path_lower.endswith('.parquet'):
        return 'parquet'
    elif path_lower.endswith('.duckdb'):
        return 'duckdb'
    elif path_lower.endswith('.csv'):
        return 'csv'
    elif path_lower.endswith('.tsv'):
        return 'tsv'
    elif path_lower.endswith('.txt'):
        return 'txt'
    elif path_lower.endswith('.pipe') or path_lower.endswith('.psv'):
        return 'pipe'
    elif path_lower.endswith('.json'):
        return 'json'
    return None


def get_filename_from_content_disposition(header_value: str) -> Optional[str]:
    """Extracts filename from Content-Disposition header."""
    import re
    if not header_value:
        return None
    match = re.search(r'filename="([^"]+)"', header_value)
    if match:
        return match.group(1)
    match = re.search(r'filename=([^\s;]+)', header_value)
    if match:
        return match.group(1)
    return None


def detect_file_type_from_header(content_disposition: str) -> Optional[str]:
    """Detects file type from Content-Disposition header."""
    filename = get_filename_from_content_disposition(content_disposition)
    if filename:
        return detect_file_type_from_path(filename)
    return None


def detect_file_type_from_bytes(file_bytes: bytes) -> Optional[str]:
    """Detects file type from magic bytes (file signature)."""
    if len(file_bytes) < 20:
        return None

    try:
        text_start = file_bytes[:500].decode('utf-8', errors='ignore').lower().strip()
        if text_start.startswith('<!doctype') or text_start.startswith('<html') or '<html' in text_start[:200]:
            return 'html_error'
    except:
        pass

    if file_bytes[:4] == b'PAR1' or file_bytes[-4:] == b'PAR1':
        return 'parquet'

    if b'DUCK' in file_bytes[:20]:
        return 'duckdb'

    if b'SQLite format 3' in file_bytes[:16]:
        return 'duckdb'

    try:
        text_sample = file_bytes[:2048].decode('utf-8', errors='strict')
    except UnicodeDecodeError:
        return None

    first_char = text_sample.strip()[0] if text_sample.strip() else ''
    if first_char in ('{', '['):
        return 'json'

    if '\n' not in text_sample and '\r' not in text_sample:
        return 'txt'

    first_line = text_sample.split('\n')[0].split('\r')[0]
    pipe_count = first_line.count('|')
    tab_count = first_line.count('\t')
    comma_count = first_line.count(',')

    MIN_DELIMITER_COUNT = 2
    if pipe_count >= MIN_DELIMITER_COUNT:
        return 'pipe'
    elif tab_count >= MIN_DELIMITER_COUNT:
        return 'tsv'
    elif comma_count >= MIN_DELIMITER_COUNT:
        return 'csv'
    else:
        return 'txt'


def detect_delimiter_from_content(file_bytes: bytes) -> str:
    """Analyzes first line of text file to detect delimiter."""
    try:
        first_line = file_bytes.split(b'\n')[0].decode('utf-8', errors='ignore')
        counts = {
            '|': first_line.count('|'),
            '\t': first_line.count('\t'),
            ',': first_line.count(',')
        }
        if counts['|'] > 0:
            return '|'
        elif counts['\t'] > 0:
            return '\t'
        elif counts[','] > 0:
            return ','
        else:
            return ','
    except Exception:
        return ','


# === SECTION: DATA_PROFILING ===
# =============================================================================
# DATA PROFILING FUNCTIONS
# =============================================================================

NUMERIC_TYPES = {
    'INTEGER', 'BIGINT', 'SMALLINT', 'TINYINT', 'UBIGINT', 'UINTEGER', 'USMALLINT', 'UTINYINT',
    'DOUBLE', 'FLOAT', 'REAL', 'DECIMAL', 'NUMERIC', 'HUGEINT'
}


def is_numeric_type(column_type: str) -> bool:
    """Check if a DuckDB column type is numeric."""
    base_type = column_type.upper().split('(')[0].strip()
    return base_type in NUMERIC_TYPES


def generate_column_stats(conn, table_or_path: str, is_file: bool = False) -> pd.DataFrame:
    """Generate column statistics for all columns in a table or file."""
    if is_file:
        source = f"'{table_or_path}'"
    else:
        source = f'"{table_or_path}"'

    describe_result = conn.execute(f"DESCRIBE SELECT * FROM {source}").fetchall()
    total_rows = conn.execute(f"SELECT COUNT(*) FROM {source}").fetchone()[0]

    stats_data = []

    for row in describe_result:
        col_name = row[0]
        col_type = row[1]
        is_numeric = is_numeric_type(col_type)
        col_escaped = f'"{col_name}"'

        try:
            base_query = f"""
                SELECT
                    COUNT(*) - COUNT({col_escaped}) as null_count,
                    COUNT(DISTINCT {col_escaped}) as unique_count
                FROM {source}
            """
            base_stats = conn.execute(base_query).fetchone()
            null_count = base_stats[0]
            unique_count = base_stats[1]
            null_pct = round((null_count / total_rows * 100), 1) if total_rows > 0 else 0

            if is_numeric:
                try:
                    num_query = f"""
                        SELECT
                            MIN({col_escaped}),
                            MAX({col_escaped}),
                            ROUND(AVG({col_escaped}::DOUBLE), 2),
                            ROUND(MEDIAN({col_escaped}::DOUBLE), 2),
                            ROUND(PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY {col_escaped}::DOUBLE), 2),
                            ROUND(PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY {col_escaped}::DOUBLE), 2)
                        FROM {source}
                    """
                    num_stats = conn.execute(num_query).fetchone()
                    min_val = num_stats[0] if num_stats[0] is not None else '-'
                    max_val = num_stats[1] if num_stats[1] is not None else '-'
                    avg_val = num_stats[2] if num_stats[2] is not None else '-'
                    median_val = num_stats[3] if num_stats[3] is not None else '-'
                    p25_val = num_stats[4] if num_stats[4] is not None else '-'
                    p75_val = num_stats[5] if num_stats[5] is not None else '-'
                except:
                    min_val = max_val = avg_val = median_val = p25_val = p75_val = '-'
            else:
                min_val = max_val = avg_val = median_val = p25_val = p75_val = '-'

            stats_data.append({
                'column_name': col_name,
                'type': col_type,
                'nulls': null_count,
                'null%': null_pct,
                'unique': unique_count,
                'min': min_val,
                'max': max_val,
                'avg': avg_val,
                'median': median_val,
                'p25': p25_val,
                'p75': p75_val
            })

        except Exception as e:
            stats_data.append({
                'column_name': col_name,
                'type': col_type,
                'nulls': '-',
                'null%': '-',
                'unique': '-',
                'min': '-',
                'max': '-',
                'avg': '-',
                'median': '-',
                'p25': '-',
                'p75': '-'
            })

    return pd.DataFrame(stats_data)


# === SECTION: SCHEMA_DISPLAY ===
# =============================================================================
# SCHEMA DISPLAY FUNCTIONS
# =============================================================================

def display_parquet_schema(book: xw.Book, source_url: str, file_size_mb: float,
                           download_minutes: int, download_seconds: int, save_seconds: int = 0) -> bool:
    """Reads schema from downloaded parquet file and displays it on PARQUET sheet with data profiling."""
    temp_parquet_path = get_parquet_temp_path()
    print("\n   Reading Parquet schema and generating data profile...")
    duckdb_start_time = time.time()

    try:
        conn = duckdb.connect()

        count_query = f"SELECT COUNT(*) FROM '{temp_parquet_path}'"
        row_count = scalar(conn, count_query)
        print(f"   Total rows: {row_count:,}")

        print("   Generating column statistics...")
        stats_df = generate_column_stats(conn, temp_parquet_path, is_file=True)
        print(f"   Found {len(stats_df)} columns")

        sample_query = f"SELECT * FROM '{temp_parquet_path}' LIMIT 5"
        sample_df = conn.execute(sample_query).fetchdf()
        for col in sample_df.columns:
            sample_df[col] = sample_df[col].astype(str)
        conn.close()
        duckdb_duration = time.time() - duckdb_start_time
        duckdb_seconds = int(duckdb_duration)
        print(f"   DuckDB processing: {duckdb_seconds}s")

    except Exception as e:
        print(f"ERROR: Failed to read Parquet schema: {e}")
        return False

    print("\n   Writing results to PARQUET sheet...")
    excel_start_time = time.time()

    try:
        output_sheet = ensure_sheet_exists(book, 'PARQUET')
        output_sheet.range("A1").value = "Parquet Import Results"
        output_sheet.range("A1").font.bold = True
        output_sheet.range("A1").font.size = 14

        output_sheet.range("A2").value = f"Source: {source_url}"
        output_sheet.range("A3").value = f"File Size: {file_size_mb:.2f} MB | Rows: {row_count:,}"

        output_sheet.range("A4").font.size = 10

        current_row = 6

        output_sheet.range(f"A{current_row}").value = "DATA PROFILE (Schema + Statistics)"
        output_sheet.range(f"A{current_row}").font.bold = True
        output_sheet.range(f"A{current_row}").font.size = 12
        current_row += 1

        output_sheet.range(f"A{current_row}").options(index=False).value = stats_df

        stats_header = output_sheet.range(f"A{current_row}").resize(1, stats_df.shape[1])
        stats_header.color = '#4472C4'
        stats_header.font.color = '#FFFFFF'
        stats_header.font.bold = True

        current_row += stats_df.shape[0] + 3

        output_sheet.range(f"A{current_row}").value = "SAMPLE DATA (First 5 Rows)"
        output_sheet.range(f"A{current_row}").font.bold = True
        output_sheet.range(f"A{current_row}").font.size = 12
        output_sheet.range(f"A{current_row}").color = '#E2EFDA'
        current_row += 1

        output_sheet.range(f"A{current_row}").options(index=False).value = sample_df

        sample_header = output_sheet.range(f"A{current_row}").resize(1, sample_df.shape[1])
        sample_header.color = '#D9E1F2'
        sample_header.font.color = '#000000'
        sample_header.font.bold = True

        excel_duration = time.time() - excel_start_time
        excel_seconds = int(excel_duration)

        total_seconds = (download_minutes * 60) + download_seconds + save_seconds + duckdb_seconds + excel_seconds

        if download_minutes > 0:
            timing_text = f"Download: {download_minutes}m {download_seconds}s | Save: {save_seconds}s | DuckDB: {duckdb_seconds}s | Excel: {excel_seconds}s | Total: {total_seconds // 60}m {total_seconds % 60}s"
        else:
            timing_text = f"Download: {download_seconds}s | Save: {save_seconds}s | DuckDB: {duckdb_seconds}s | Excel: {excel_seconds}s | Total: {total_seconds}s"

        output_sheet.range("A4").value = timing_text

        output_sheet.activate()
        print(f"   Excel write: {excel_seconds}s")
        print("\n" + "=" * 60)
        print("PARQUET IMPORT COMPLETE!")
        print(f"Columns: {len(stats_df)} | Rows: {row_count:,}")
        print(timing_text)
        print("=" * 60)
        return True

    except Exception as e:
        print(f"ERROR: Failed to write results: {e}")
        return False


def display_delimited_schema(book: xw.Book, source_url: str, file_size_mb: float,
                             download_minutes: int, download_seconds: int, save_seconds: int,
                             delimiter: Optional[str], file_type: str) -> bool:
    """Imports delimited file into DuckDB for unified querying with data profiling."""
    temp_csv_path = get_delimited_temp_path()
    duck_db_path = get_duckdb_temp_path()

    type_display = {'csv': 'CSV', 'tsv': 'TSV', 'txt': 'TXT', 'pipe': 'Pipe-delimited'}.get(file_type, file_type.upper())
    delimiter_display = {',': 'comma', '\t': 'tab', '|': 'pipe'}.get(delimiter, repr(delimiter))

    print(f"\n   Converting {type_display} to DuckDB (delimiter: {delimiter_display})...")
    duckdb_start_time = time.time()

    try:
        if os.path.exists(duck_db_path):
            os.remove(duck_db_path)

        conn = duckdb.connect(duck_db_path)
        create_table_query = f"""
            CREATE TABLE imported_data AS
            SELECT * FROM read_csv_auto(
                '{temp_csv_path}',
                delim='{delimiter}',
                sample_size=-1,
                all_varchar=true
            )
        """
        conn.execute(create_table_query)

        row_count = scalar(conn, "SELECT COUNT(*) FROM imported_data")

        print("   Generating column statistics...")
        stats_df = generate_column_stats(conn, "imported_data", is_file=False)
        print(f"   Found {len(stats_df)} columns")

        sample_df = conn.execute("SELECT * FROM imported_data LIMIT 10").fetchdf()
        for col in sample_df.columns:
            sample_df[col] = sample_df[col].astype(str)

        conn.close()
        os.remove(temp_csv_path)

        duckdb_size_mb = os.path.getsize(duck_db_path) / (1024 * 1024)
        duckdb_duration = time.time() - duckdb_start_time
        duckdb_seconds = int(duckdb_duration)
        print(f"   DuckDB processing: {duckdb_seconds}s")

    except Exception as e:
        print(f"ERROR: Failed to read {type_display}: {e}")
        return False

    print(f"\n   Writing results to CSV sheet...")
    excel_start_time = time.time()

    try:
        output_sheet = ensure_sheet_exists(book, 'CSV')

        output_sheet.range("A1").value = f"{type_display} Import Results (converted to DuckDB)"
        output_sheet.range("A1").font.bold = True
        output_sheet.range("A1").font.size = 14

        output_sheet.range("A2").value = f"Source: {source_url}"
        output_sheet.range("A2").font.size = 10

        output_sheet.range("A3").value = f"Original {type_display}: {file_size_mb:.2f} MB | DuckDB: {duckdb_size_mb:.2f} MB | Rows: {row_count:,}"
        output_sheet.range("A3").font.size = 10

        output_sheet.range("A4").font.size = 10

        current_row = 6

        output_sheet.range(f"A{current_row}").value = "DATA PROFILE (Schema + Statistics)"
        output_sheet.range(f"A{current_row}").font.bold = True
        output_sheet.range(f"A{current_row}").font.size = 12
        current_row += 1

        output_sheet.range(f"A{current_row}").options(index=False).value = stats_df

        stats_header = output_sheet.range(f"A{current_row}").resize(1, stats_df.shape[1])
        stats_header.color = '#4472C4'
        stats_header.font.color = '#FFFFFF'
        stats_header.font.bold = True

        current_row += stats_df.shape[0] + 3

        output_sheet.range(f"A{current_row}").value = "SAMPLE DATA (First 10 Rows)"
        output_sheet.range(f"A{current_row}").font.bold = True
        output_sheet.range(f"A{current_row}").font.size = 12
        output_sheet.range(f"A{current_row}").color = '#E2EFDA'
        current_row += 1

        output_sheet.range(f"A{current_row}").options(index=False).value = sample_df

        sample_header = output_sheet.range(f"A{current_row}").resize(1, sample_df.shape[1])
        sample_header.color = '#D9E1F2'
        sample_header.font.color = '#000000'
        sample_header.font.bold = True

        excel_duration = time.time() - excel_start_time
        excel_seconds = int(excel_duration)

        total_seconds = (download_minutes * 60) + download_seconds + save_seconds + duckdb_seconds + excel_seconds

        if download_minutes > 0:
            timing_text = f"Download: {download_minutes}m {download_seconds}s | Save: {save_seconds}s | DuckDB: {duckdb_seconds}s | Excel: {excel_seconds}s | Total: {total_seconds // 60}m {total_seconds % 60}s"
        else:
            timing_text = f"Download: {download_seconds}s | Save: {save_seconds}s | DuckDB: {duckdb_seconds}s | Excel: {excel_seconds}s | Total: {total_seconds}s"

        output_sheet.range("A4").value = timing_text

        output_sheet.activate()
        print(f"   Excel write: {excel_seconds}s")
        print("\n" + "=" * 60)
        print(f"{type_display} -> DUCKDB CONVERSION COMPLETE!")
        print(f"Columns: {len(stats_df)} | Rows: {row_count:,}")
        print(timing_text)
        print("=" * 60)
        return True

    except Exception as e:
        print(f"ERROR: Failed to write results: {e}")
        return False


def display_json_schema(book: xw.Book, source_url: str, file_size_mb: float,
                        download_minutes: int, download_seconds: int, save_seconds: int = 0) -> bool:
    """Imports JSON file into DuckDB for unified querying."""
    temp_json_path = get_json_temp_path()
    duck_db_path = get_duckdb_temp_path()

    print(f"\n   Converting JSON to DuckDB...")
    duckdb_start_time = time.time()

    try:
        if os.path.exists(duck_db_path):
            os.remove(duck_db_path)

        conn = duckdb.connect(duck_db_path)
        create_table_query = f"""
            CREATE TABLE imported_data AS
            SELECT * FROM read_json_auto(
                '{temp_json_path}',
                auto_detect=true,
                maximum_object_size=10000000
            )
        """
        conn.execute(create_table_query)

        row_count = scalar(conn, "SELECT COUNT(*) FROM imported_data")

        print("   Generating column statistics...")
        stats_df = generate_column_stats(conn, "imported_data", is_file=False)
        print(f"   Found {len(stats_df)} columns")

        sample_df = conn.execute("SELECT * FROM imported_data LIMIT 10").fetchdf()
        for col in sample_df.columns:
            sample_df[col] = sample_df[col].astype(str)

        conn.close()
        os.remove(temp_json_path)

        duckdb_size_mb = os.path.getsize(duck_db_path) / (1024 * 1024)
        duckdb_duration = time.time() - duckdb_start_time
        duckdb_seconds = int(duckdb_duration)
        print(f"   DuckDB processing: {duckdb_seconds}s")

    except Exception as e:
        print(f"ERROR: Failed to convert JSON: {e}")
        return False

    print(f"\n   Writing results to JSON sheet...")
    excel_start_time = time.time()

    try:
        output_sheet = ensure_sheet_exists(book, 'JSON')

        output_sheet.range("A1").value = "JSON Import Results (converted to DuckDB)"
        output_sheet.range("A1").font.bold = True
        output_sheet.range("A1").font.size = 14

        output_sheet.range("A2").value = f"Source: {source_url}"
        output_sheet.range("A2").font.size = 10

        output_sheet.range("A3").value = f"Original JSON: {file_size_mb:.2f} MB | DuckDB: {duckdb_size_mb:.2f} MB | Rows: {row_count:,}"
        output_sheet.range("A3").font.size = 10

        output_sheet.range("A4").font.size = 10

        current_row = 6

        output_sheet.range(f"A{current_row}").value = "DATA PROFILE (Schema + Statistics)"
        output_sheet.range(f"A{current_row}").font.bold = True
        output_sheet.range(f"A{current_row}").font.size = 12
        current_row += 1

        output_sheet.range(f"A{current_row}").options(index=False).value = stats_df

        stats_header = output_sheet.range(f"A{current_row}").resize(1, stats_df.shape[1])
        stats_header.color = '#4472C4'
        stats_header.font.color = '#FFFFFF'
        stats_header.font.bold = True

        current_row += stats_df.shape[0] + 3

        output_sheet.range(f"A{current_row}").value = "SAMPLE DATA (First 10 Rows)"
        output_sheet.range(f"A{current_row}").font.bold = True
        output_sheet.range(f"A{current_row}").font.size = 12
        output_sheet.range(f"A{current_row}").color = '#E2EFDA'
        current_row += 1

        output_sheet.range(f"A{current_row}").options(index=False).value = sample_df

        sample_header = output_sheet.range(f"A{current_row}").resize(1, sample_df.shape[1])
        sample_header.color = '#D9E1F2'
        sample_header.font.color = '#000000'
        sample_header.font.bold = True

        excel_duration = time.time() - excel_start_time
        excel_seconds = int(excel_duration)

        total_seconds = (download_minutes * 60) + download_seconds + save_seconds + duckdb_seconds + excel_seconds

        if download_minutes > 0:
            timing_text = f"Download: {download_minutes}m {download_seconds}s | Save: {save_seconds}s | DuckDB: {duckdb_seconds}s | Excel: {excel_seconds}s | Total: {total_seconds // 60}m {total_seconds % 60}s"
        else:
            timing_text = f"Download: {download_seconds}s | Save: {save_seconds}s | DuckDB: {duckdb_seconds}s | Excel: {excel_seconds}s | Total: {total_seconds}s"

        output_sheet.range("A4").value = timing_text

        output_sheet.activate()
        print(f"   Excel write: {excel_seconds}s")
        print("\n" + "=" * 60)
        print(f"JSON -> DUCKDB CONVERSION COMPLETE!")
        print(f"Columns: {len(stats_df)} | Rows: {row_count:,}")
        print(timing_text)
        print("=" * 60)
        return True

    except Exception as e:
        print(f"ERROR: Failed to write results: {e}")
        return False


def display_sqlite_schema(book: xw.Book, source_url: str, file_size_mb: float,
                         download_minutes: int, download_seconds: int, save_seconds: int, db_path: str) -> bool:
    """Converts SQLite database to DuckDB format for unified querying."""
    import sqlite3

    print("\n   Converting SQLite to DuckDB format...")
    duckdb_start_time = time.time()

    try:
        sqlite_conn = sqlite3.connect(db_path)
        cursor = sqlite_conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        table_names = [row[0] for row in cursor.fetchall()]
        print(f"   Found {len(table_names)} tables: {table_names}")

        original_size_mb = os.path.getsize(db_path) / (1024 * 1024)
        table_info = []
        sample_data = {}
        table_dataframes = {}

        for table_name in table_names:
            df = pd.read_sql(f'SELECT * FROM "{table_name}"', sqlite_conn)
            table_dataframes[table_name] = df
            table_info.append({
                'Table Name': str(table_name),
                'Columns': int(len(df.columns)),
                'Rows': int(len(df))
            })
            sample_df = df.head(3).copy()
            for col in sample_df.columns:
                sample_df[col] = sample_df[col].astype(str)
            sample_data[table_name] = sample_df

        sqlite_conn.close()
        os.remove(db_path)

        duck_db_path = get_duckdb_temp_path()
        if os.path.exists(duck_db_path):
            os.remove(duck_db_path)

        duck_conn = duckdb.connect(duck_db_path)
        for table_name, df in table_dataframes.items():
            duck_conn.execute(f'CREATE TABLE "{table_name}" AS SELECT * FROM df')

        # Generate column statistics for each table
        table_stats: dict[str, Any] = {}
        for table_name in table_names:
            print(f"   Generating stats for table: {table_name}...")
            try:
                stats_df = generate_column_stats(duck_conn, table_name, is_file=False)
                table_stats[table_name] = stats_df
            except Exception as e:
                print(f"   Warning: Could not generate stats for {table_name}: {e}")
                table_stats[table_name] = None

        duck_conn.close()

        duckdb_size_mb = os.path.getsize(duck_db_path) / (1024 * 1024)
        schema_df = pd.DataFrame(table_info)
        duckdb_duration = time.time() - duckdb_start_time
        duckdb_seconds = int(duckdb_duration)
        print(f"   DuckDB processing: {duckdb_seconds}s")

    except Exception as e:
        print(f"ERROR: Failed to read SQLite: {e}")
        return False

    print("\n   Writing results to DUCKDB sheet...")
    excel_start_time = time.time()

    try:
        output_sheet = ensure_sheet_exists(book, 'DUCKDB')

        output_sheet.range("A1").value = "DuckDB Import Results (converted from SQLite)"
        output_sheet.range("A1").font.bold = True
        output_sheet.range("A1").font.size = 14

        output_sheet.range("A2").value = f"Source: {source_url}"
        output_sheet.range("A2").font.size = 10

        compression_ratio = (1 - duckdb_size_mb / original_size_mb) * 100 if original_size_mb > 0 else 0
        output_sheet.range("A3").value = f"Original SQLite: {file_size_mb:.2f} MB | DuckDB: {duckdb_size_mb:.2f} MB (saved {compression_ratio:.1f}%)"
        output_sheet.range("A3").font.size = 10

        output_sheet.range("A4").font.size = 10

        current_row = 6

        output_sheet.range(f"A{current_row}").value = "DATABASE SCHEMA"
        output_sheet.range(f"A{current_row}").font.bold = True
        output_sheet.range(f"A{current_row}").font.size = 12
        current_row += 1

        output_sheet.range(f"A{current_row}").options(index=False).value = schema_df

        header_range = output_sheet.range(f"A{current_row}:C{current_row}")
        header_range.color = '#4472C4'
        header_range.font.color = '#FFFFFF'
        header_range.font.bold = True

        current_row += schema_df.shape[0] + 3

        for table_name in table_names:
            # Table header
            output_sheet.range(f"A{current_row}").value = f"TABLE: {table_name}"
            output_sheet.range(f"A{current_row}").font.bold = True
            output_sheet.range(f"A{current_row}").font.size = 12
            output_sheet.range(f"A{current_row}").color = '#FFC000'
            current_row += 2

            # Column statistics
            if table_name in table_stats and table_stats[table_name] is not None:
                output_sheet.range(f"A{current_row}").value = "DATA PROFILE (Column Statistics)"
                output_sheet.range(f"A{current_row}").font.bold = True
                output_sheet.range(f"A{current_row}").font.size = 11
                current_row += 1

                stats_for_table = table_stats[table_name]
                output_sheet.range(f"A{current_row}").options(index=False).value = stats_for_table

                stats_header = output_sheet.range(f"A{current_row}").resize(1, stats_for_table.shape[1])
                stats_header.color = '#4472C4'
                stats_header.font.color = '#FFFFFF'
                stats_header.font.bold = True

                current_row += stats_for_table.shape[0] + 2

            # Sample data
            sample_df = sample_data.get(table_name)
            if sample_df is not None and not sample_df.empty:
                output_sheet.range(f"A{current_row}").value = "SAMPLE DATA (First 3 Rows)"
                output_sheet.range(f"A{current_row}").font.bold = True
                output_sheet.range(f"A{current_row}").font.size = 11
                output_sheet.range(f"A{current_row}").color = '#E2EFDA'
                current_row += 1

                output_sheet.range(f"A{current_row}").options(index=False).value = sample_df

                sample_header = output_sheet.range(f"A{current_row}").resize(1, sample_df.shape[1])
                sample_header.color = '#D9E1F2'
                sample_header.font.bold = True

                current_row += sample_df.shape[0] + 3
            else:
                output_sheet.range(f"A{current_row}").value = "(No data)"
                current_row += 2

        excel_duration = time.time() - excel_start_time
        excel_seconds = int(excel_duration)

        total_seconds = (download_minutes * 60) + download_seconds + save_seconds + duckdb_seconds + excel_seconds

        if download_minutes > 0:
            timing_text = f"Download: {download_minutes}m {download_seconds}s | Save: {save_seconds}s | DuckDB: {duckdb_seconds}s | Excel: {excel_seconds}s | Total: {total_seconds // 60}m {total_seconds % 60}s"
        else:
            timing_text = f"Download: {download_seconds}s | Save: {save_seconds}s | DuckDB: {duckdb_seconds}s | Excel: {excel_seconds}s | Total: {total_seconds}s"

        output_sheet.range("A4").value = timing_text

        output_sheet.activate()
        print(f"   Excel write: {excel_seconds}s")
        print("\n" + "=" * 60)
        print("SQLITE -> DUCKDB CONVERSION COMPLETE!")
        print(f"Tables imported: {len(schema_df)}")
        print(timing_text)
        print("=" * 60)
        return True

    except Exception as e:
        print(f"ERROR: Failed to write results: {e}")
        return False


def display_duckdb_schema(book: xw.Book, source_url: str, file_size_mb: float,
                          download_minutes: int, download_seconds: int, save_seconds: int = 0) -> bool:
    """Reads schema from downloaded DuckDB file and displays it on DUCKDB sheet."""
    temp_db_path = get_duckdb_temp_path()

    with open(temp_db_path, 'rb') as f:
        file_header = f.read(16)
        is_sqlite = b'SQLite format 3' in file_header

    if is_sqlite:
        print("\n   Detected SQLite database file")
        return display_sqlite_schema(book, source_url, file_size_mb, download_minutes, download_seconds, save_seconds, temp_db_path)

    print("\n   Reading DuckDB schema...")
    duckdb_start_time = time.time()

    try:
        conn = duckdb.connect(temp_db_path, read_only=True)
        print("   Connected to DuckDB")

        tables_query = """
            SELECT table_name
            FROM information_schema.tables
            WHERE table_schema = 'main'
            ORDER BY table_name
        """
        tables_result = conn.execute(tables_query).fetchall()
        table_names = [row[0] for row in tables_result]
        print(f"   Found {len(table_names)} tables: {table_names}")

        table_info = []
        table_stats: dict[str, Any] = {}
        sample_data = {}

        for table_name in table_names:
            row_count = scalar(conn, f'SELECT COUNT(*) FROM "{table_name}"')
            describe_result = conn.execute(f'DESCRIBE "{table_name}"').fetchall()
            table_info.append({
                'Table Name': str(table_name),
                'Columns': int(len(describe_result)),
                'Rows': int(row_count)
            })

            print(f"   Generating stats for table: {table_name}...")
            try:
                stats_df = generate_column_stats(conn, table_name, is_file=False)
                table_stats[table_name] = stats_df
            except Exception as e:
                print(f"   Warning: Could not generate stats for {table_name}: {e}")
                table_stats[table_name] = None

            try:
                sample_df = conn.execute(f'SELECT * FROM "{table_name}" LIMIT 3').fetchdf()
                for col in sample_df.columns:
                    sample_df[col] = sample_df[col].astype(str)
                sample_data[table_name] = sample_df
            except Exception as e:
                sample_data[table_name] = pd.DataFrame({'Error': [str(e)]})

        schema_df = pd.DataFrame(table_info)
        conn.close()
        duckdb_duration = time.time() - duckdb_start_time
        duckdb_seconds = int(duckdb_duration)
        print(f"   DuckDB processing: {duckdb_seconds}s")

    except Exception as e:
        print(f"ERROR: Failed to read DuckDB: {e}")
        return False

    print("\n   Writing results to DUCKDB sheet...")
    excel_start_time = time.time()

    try:
        output_sheet = ensure_sheet_exists(book, 'DUCKDB')

        output_sheet.range("A1").value = "DuckDB Import Results"
        output_sheet.range("A1").font.bold = True
        output_sheet.range("A1").font.size = 14

        output_sheet.range("A2").value = f"Source: {source_url}"
        output_sheet.range("A2").font.size = 10

        output_sheet.range("A3").value = f"File Size: ~{file_size_mb:.2f} MB"
        output_sheet.range("A3").font.size = 10

        output_sheet.range("A4").font.size = 10

        current_row = 6

        output_sheet.range(f"A{current_row}").value = "DATABASE SCHEMA (Tables Overview)"
        output_sheet.range(f"A{current_row}").font.bold = True
        output_sheet.range(f"A{current_row}").font.size = 12
        current_row += 1

        output_sheet.range(f"A{current_row}").options(index=False).value = schema_df

        header_range = output_sheet.range(f"A{current_row}:C{current_row}")
        header_range.color = '#4472C4'
        header_range.font.color = '#FFFFFF'
        header_range.font.bold = True

        current_row += schema_df.shape[0] + 3

        for table_name in table_names:
            output_sheet.range(f"A{current_row}").value = f"TABLE: {table_name}"
            output_sheet.range(f"A{current_row}").font.bold = True
            output_sheet.range(f"A{current_row}").font.size = 12
            output_sheet.range(f"A{current_row}").color = '#FFC000'
            current_row += 2

            if table_name in table_stats and table_stats[table_name] is not None:
                output_sheet.range(f"A{current_row}").value = "DATA PROFILE (Column Statistics)"
                output_sheet.range(f"A{current_row}").font.bold = True
                output_sheet.range(f"A{current_row}").font.size = 11
                current_row += 1

                stats_for_table = table_stats[table_name]
                output_sheet.range(f"A{current_row}").options(index=False).value = stats_for_table

                stats_header = output_sheet.range(f"A{current_row}").resize(1, stats_for_table.shape[1])
                stats_header.color = '#4472C4'
                stats_header.font.color = '#FFFFFF'
                stats_header.font.bold = True

                current_row += stats_for_table.shape[0] + 2

            sample_df = sample_data.get(table_name)
            if sample_df is not None and not sample_df.empty:
                output_sheet.range(f"A{current_row}").value = "SAMPLE DATA (First 3 Rows)"
                output_sheet.range(f"A{current_row}").font.bold = True
                output_sheet.range(f"A{current_row}").font.size = 11
                output_sheet.range(f"A{current_row}").color = '#E2EFDA'
                current_row += 1

                output_sheet.range(f"A{current_row}").options(index=False).value = sample_df

                sample_header = output_sheet.range(f"A{current_row}").resize(1, sample_df.shape[1])
                sample_header.color = '#D9E1F2'
                sample_header.font.bold = True

                current_row += sample_df.shape[0] + 3
            else:
                output_sheet.range(f"A{current_row}").value = "(No data)"
                current_row += 2

        excel_duration = time.time() - excel_start_time
        excel_seconds = int(excel_duration)

        total_seconds = (download_minutes * 60) + download_seconds + save_seconds + duckdb_seconds + excel_seconds

        if download_minutes > 0:
            timing_text = f"Download: {download_minutes}m {download_seconds}s | Save: {save_seconds}s | DuckDB: {duckdb_seconds}s | Excel: {excel_seconds}s | Total: {total_seconds // 60}m {total_seconds % 60}s"
        else:
            timing_text = f"Download: {download_seconds}s | Save: {save_seconds}s | DuckDB: {duckdb_seconds}s | Excel: {excel_seconds}s | Total: {total_seconds}s"

        output_sheet.range("A4").value = timing_text

        output_sheet.activate()
        print(f"   Excel write: {excel_seconds}s")
        print("\n" + "=" * 60)
        print("DUCKDB IMPORT COMPLETE!")
        print(f"Tables imported: {len(schema_df)}")
        print(timing_text)
        print("=" * 60)
        return True

    except Exception as e:
        print(f"ERROR: Failed to write results: {e}")
        return False



# === SECTION: SHEET_LAYOUT ===
# =============================================================================
# INPUT CELLS AND STATUS MESSAGES
# =============================================================================

LINK_SHEET = "SHARE_LINK_ACCESS"
TOKEN_SHEET = "TOKEN_ACCESS"
# SHARE_LINK_ACCESS: B5 link, B8 file name (as-is), B9 folder (as-is), B12 status
# TOKEN_ACCESS: B5 provider, B6 path / ID / URL, B7 auth proxy, B9 file name (as-is), B10 folder (as-is), B12 status
STATUS_CELL = {LINK_SHEET: "B12", TOKEN_SHEET: "B12"}
RESULT_SHEET = {'parquet': 'PARQUET', 'duckdb': 'DUCKDB', 'json': 'JSON'}


def cell_text(sheet: xw.Sheet, address: str) -> str:
    """Cell value as stripped text ('' when empty)."""
    value = sheet[address].value
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def set_status(sheet: xw.Sheet, message: str, ok: bool = True) -> None:
    """Writes a message in the sheet's status cell: blue when done, red on error."""
    cell = sheet[STATUS_CELL[sheet.name]]
    cell.value = message
    cell.font.color = '#1F4E79' if ok else '#C00000'
    cell.font.bold = True


def fail(sheet: xw.Sheet, message: str) -> None:
    print(f"\nERROR: {message}")
    set_status(sheet, f"ERROR: {message}", ok=False)


def clear_previous_imports() -> None:
    """Removes the temp files of the previous import, so run_stats never reads an old file."""
    for temp_path in [get_parquet_temp_path(), get_duckdb_temp_path(),
                      get_delimited_temp_path(), get_json_temp_path()]:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def size_text(n_bytes: int) -> str:
    return f"{n_bytes / (1024 * 1024):.2f} MB" if n_bytes >= 100 * 1024 else f"{n_bytes / 1024:.1f} KB"


def format_duration(seconds: float) -> Tuple[int, int]:
    return int(seconds // 60), int(seconds % 60)


# === SECTION: DOWNLOAD ===
# =============================================================================
# DOWNLOAD: SHAREABLE LINK OR TOKEN
# =============================================================================

async def download_link(url: str) -> Tuple[bytes, str, str]:
    """Downloads a public / shareable link. Google Drive, Dropbox and GitHub Releases go through
    the CORS proxy. Returns (bytes, final url, Content-Disposition header)."""
    url, converted = convert_google_drive_url(url)
    if converted:
        print("   Google Drive link converted to a direct download link")
    url, converted = convert_dropbox_url(url)
    if converted:
        print("   Dropbox link converted to a direct download link (dl=1)")

    if needs_proxy(url):
        fetch_url = build_public_proxy_url(url)
        print(f"   Via CORS proxy: {PUBLIC_PROXY_URL}")
    else:
        fetch_url = url
        print("   Direct download (no proxy)")

    response = await pyfetch(fetch_url, method="GET")
    if not response.ok:
        raise ValueError(f"download failed (HTTP {response.status}). Is the file shared publicly? "
                         f"For private files use TOKEN_ACCESS")
    content_disposition = response.headers.get('Content-Disposition', '') or ''
    return await response.bytes(), url, content_disposition


async def download_with_token(provider: str, file_input: str, auth_proxy_url: str) -> Tuple[bytes, str]:
    """Downloads a private file with your own credentials (environment variables).
    Returns (bytes, cleaned path / ID / URL)."""
    if provider == "Dropbox":
        if not file_input.startswith("/"):
            raise ValueError("Dropbox paths start with / (e.g. /folder/file.csv)")
        access_token = await get_dropbox_token()
        response = await pyfetch(
            "https://content.dropboxapi.com/2/files/download",
            method="POST",
            headers={
                "Authorization": f"Bearer {access_token}",
                "Dropbox-API-Arg": json.dumps({"path": file_input})
            }
        )
    elif provider == "Google Drive":
        file_id = extract_gdrive_file_id(file_input)
        if not file_id:
            raise ValueError("enter a Google Drive file ID or file link")
        file_input = file_id
        access_token = await get_gdrive_token()
        response = await pyfetch(
            f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media",
            method="GET",
            headers={"Authorization": f"Bearer {access_token}"}
        )
    elif provider == "GitHub":
        if not file_input.lower().startswith('http'):
            raise ValueError("GitHub needs the full file URL (https://github.com/... or https://raw.githubusercontent.com/...)")
        github_pat = os.environ.get("GITHUB.PAT")
        if not github_pat:
            raise ValueError("GITHUB.PAT is not set (xlwings Lite: Settings > Environment Variables)")
        headers = {"Authorization": f"token {github_pat}"}
        api_url = github_contents_api_url(file_input)
        if api_url:
            # repository file: the GitHub API accepts a token from the browser, raw.githubusercontent.com does not
            fetch_url = api_url
            headers["Accept"] = "application/vnd.github.raw"
            print("   Via the GitHub API (repository file, up to 100 MB)")
        elif auth_proxy_url:
            fetch_url = build_private_proxy_url(file_input, auth_proxy_url)
            print(f"   Via auth proxy: {auth_proxy_url}")
        elif is_github_release_url(file_input):
            raise ValueError("GitHub release files need the Auth Proxy URL in B7")
        else:
            fetch_url = file_input
        response = await pyfetch(fetch_url, method="GET", headers=headers)
    else:
        raise ValueError("choose Dropbox, Google Drive or GitHub in B5")

    if not response.ok:
        raise ValueError(f"{provider} download failed (HTTP {response.status}). Check the path and your credentials")
    return await response.bytes(), file_input


def github_contents_api_url(url: str) -> Optional[str]:
    """Maps a repository file link (raw.githubusercontent.com/owner/repo/branch/path or
    github.com/owner/repo/blob/branch/path) to the GitHub contents API. None for anything else."""
    import urllib.parse
    parsed = urllib.parse.urlparse(url)
    parts = [p for p in parsed.path.split('/') if p]
    host = parsed.netloc.lower()
    if host == 'raw.githubusercontent.com' and len(parts) >= 4:
        owner, repo, rest = parts[0], parts[1], parts[2:]
        if rest[:2] == ['refs', 'heads'] and len(rest) >= 4:
            rest = rest[2:]
    elif host in ('github.com', 'www.github.com') and len(parts) >= 5 and parts[2] == 'blob':
        owner, repo, rest = parts[0], parts[1], parts[3:]
    else:
        return None
    ref, path = rest[0], '/'.join(rest[1:])
    return f"https://api.github.com/repos/{owner}/{repo}/contents/{path}?ref={urllib.parse.quote(ref)}"


def read_token_inputs(sheet: xw.Sheet) -> Tuple[str, str, str]:
    provider = cell_text(sheet, "B5")
    file_input = cell_text(sheet, "B6")
    auth_proxy_url = cell_text(sheet, "B7")
    if provider not in ("Dropbox", "Google Drive", "GitHub"):
        raise ValueError("choose Dropbox, Google Drive or GitHub in B5")
    if not file_input:
        raise ValueError("enter the file path, ID or URL in B6")
    return provider, file_input, auth_proxy_url


# === SECTION: PROCESS ===
# =============================================================================
# PROCESS A DOWNLOADED FILE (DuckDB, schema, profile) OR SAVE IT AS-IS
# =============================================================================

def process_file(book: xw.Book, file_bytes: bytes, name_hint: str, source_display: str,
                 source_sheet: str, download_seconds: float, content_disposition: str = "") -> str:
    """Detects the file type, saves it to the temp folder, loads it into DuckDB and writes the
    schema and profile to a result sheet. Returns the result sheet name; raises on failure."""
    file_type = detect_file_type(name_hint) or detect_file_type_from_path(name_hint)
    if file_type:
        print(f"   File type from name: {file_type.upper()}")
    if not file_type and content_disposition:
        file_type = detect_file_type_from_header(content_disposition)
        if file_type:
            print(f"   File type from download header: {file_type.upper()}")
    if not file_type:
        file_type = detect_file_type_from_bytes(file_bytes)
        if file_type:
            print(f"   File type from content: {file_type.upper()}")
    if file_type == 'html_error':
        raise ValueError("received a web page, not a data file. Is the link private or not a direct file link?")
    if not file_type:
        raise ValueError("could not tell the file type. Data files: Parquet, DuckDB, SQLite, CSV, TSV, pipe, TXT, JSON. "
                         "For other files use the As-Is download")

    delimiter = None
    if file_type in ('csv', 'tsv', 'txt', 'pipe'):
        delimiter = detect_delimiter_from_content(file_bytes)
        print(f"   Delimiter: {repr(delimiter)}")

    temp_path = {'parquet': get_parquet_temp_path(), 'duckdb': get_duckdb_temp_path(),
                 'json': get_json_temp_path()}.get(file_type, get_delimited_temp_path())
    save_start = time.time()
    with open(temp_path, 'wb') as f:
        f.write(file_bytes)
    save_seconds = int(time.time() - save_start)
    save_last_import_state(source_sheet, source_display, file_type, temp_path)

    file_size_mb = len(file_bytes) / (1024 * 1024)
    minutes, seconds = format_duration(download_seconds)
    print(f"\nProcessing {file_type.upper()} file ({file_size_mb:.1f} MB)...")
    if file_type == 'parquet':
        ok = display_parquet_schema(book, source_display, file_size_mb, minutes, seconds, save_seconds)
    elif file_type == 'duckdb':
        ok = display_duckdb_schema(book, source_display, file_size_mb, minutes, seconds, save_seconds)
    elif file_type == 'json':
        ok = display_json_schema(book, source_display, file_size_mb, minutes, seconds, save_seconds)
    else:
        ok = display_delimited_schema(book, source_display, file_size_mb, minutes, seconds, save_seconds,
                                      delimiter, file_type)
    if ok is False:
        raise ValueError("the file downloaded but could not be read. See the console for details")
    return RESULT_SHEET.get(file_type, 'CSV')


def save_as_is(file_bytes: bytes, folder: str, filename: str) -> str:
    """Saves the downloaded bytes unchanged. Returns the full path."""
    if not filename or '.' not in filename:
        raise ValueError("enter a file name with extension (e.g. report.pdf)")
    if '/' in filename or '\\' in filename:
        raise ValueError("enter just the file name, and the folder in the next cell")
    name_part, ext_part = filename.rsplit('.', 1)
    filename = f"{name_part}.{ext_part.lower()}"
    folder = (folder or tempfile.gettempdir()).rstrip('/') or '/'
    if not os.path.isdir(folder):
        raise ValueError(f"folder {folder} not found. To save to your computer, add a folder in xlwings Lite "
                         f"(Files > Local Folders) and type its path, e.g. /data")
    path = os.path.join(folder, filename)
    with open(path, 'wb') as f:
        f.write(file_bytes)
    return path


# === SECTION: SCRIPTS ===
# =============================================================================
# SCRIPTS (one button each)
# =============================================================================

@script(name="1. Link: Import to DuckDB", button="[btn_link_import]SHARE_LINK_ACCESS!D5")
async def import_via_sharelink(book: xw.Book):
    """Imports a data file from a public or shareable link (SHARE_LINK_ACCESS!B5) into DuckDB."""
    sheet = book.sheets[LINK_SHEET]
    print("=" * 60)
    print("IMPORT FROM LINK")
    print("=" * 60)
    url = cell_text(sheet, "B5")
    if not url.lower().startswith("http"):
        fail(sheet, "paste a file link starting with http in B5")
        return
    set_status(sheet, "Downloading...")
    await book.flush()
    clear_previous_imports()
    try:
        start = time.time()
        file_bytes, url, content_disposition = await download_link(url)
        took = time.time() - start
        print(f"   Downloaded {len(file_bytes) / (1024 * 1024):.1f} MB in {took:.0f}s")
        result_sheet = process_file(book, file_bytes, url, url, LINK_SHEET, took, content_disposition)
    except Exception as e:
        fail(sheet, str(e))
        return
    set_status(sheet, f"Done: results on the {result_sheet} sheet ({time.time() - start:.0f}s)")


@script(name="2. Link: Download As-Is", button="[btn_link_raw]SHARE_LINK_ACCESS!D8")
async def import_raw_sharelink(book: xw.Book):
    """Downloads any file from a link without processing it (name in B8, folder in B9)."""
    sheet = book.sheets[LINK_SHEET]
    print("=" * 60)
    print("DOWNLOAD AS-IS FROM LINK")
    print("=" * 60)
    url = cell_text(sheet, "B5")
    if not url.lower().startswith("http"):
        fail(sheet, "paste a file link starting with http in B5")
        return
    set_status(sheet, "Downloading...")
    await book.flush()
    try:
        file_bytes, url, _ = await download_link(url)
        path = save_as_is(file_bytes, cell_text(sheet, "B9"), cell_text(sheet, "B8"))
    except Exception as e:
        fail(sheet, str(e))
        return
    save_last_import_state(LINK_SHEET, url, path.rsplit('.', 1)[-1], path)
    print(f"   Saved {size_text(len(file_bytes))} to {path}")
    set_status(sheet, f"Saved: {path} ({size_text(len(file_bytes))})")


@script(name="3. Token: Import to DuckDB", button="[btn_token_import]TOKEN_ACCESS!D6")
async def import_via_token(book: xw.Book):
    """Imports a private data file with your own token (TOKEN_ACCESS!B5:B7) into DuckDB."""
    sheet = book.sheets[TOKEN_SHEET]
    print("=" * 60)
    print("IMPORT WITH TOKEN")
    print("=" * 60)
    try:
        provider, file_input, auth_proxy_url = read_token_inputs(sheet)
    except ValueError as e:
        fail(sheet, str(e))
        return
    set_status(sheet, f"Downloading from {provider}...")
    await book.flush()
    clear_previous_imports()
    try:
        start = time.time()
        file_bytes, file_input = await download_with_token(provider, file_input, auth_proxy_url)
        took = time.time() - start
        print(f"   Downloaded {len(file_bytes) / (1024 * 1024):.1f} MB in {took:.0f}s")
        result_sheet = process_file(book, file_bytes, file_input, f"{provider}: {file_input}",
                                    TOKEN_SHEET, took)
    except Exception as e:
        fail(sheet, str(e))
        return
    set_status(sheet, f"Done: results on the {result_sheet} sheet ({time.time() - start:.0f}s)")


@script(name="4. Token: Download As-Is", button="[btn_token_raw]TOKEN_ACCESS!D9")
async def import_raw_token(book: xw.Book):
    """Downloads any private file with your own token, without processing (name in B9, folder in B10)."""
    sheet = book.sheets[TOKEN_SHEET]
    print("=" * 60)
    print("DOWNLOAD AS-IS WITH TOKEN")
    print("=" * 60)
    try:
        provider, file_input, auth_proxy_url = read_token_inputs(sheet)
    except ValueError as e:
        fail(sheet, str(e))
        return
    set_status(sheet, f"Downloading from {provider}...")
    await book.flush()
    try:
        file_bytes, file_input = await download_with_token(provider, file_input, auth_proxy_url)
        path = save_as_is(file_bytes, cell_text(sheet, "B10"), cell_text(sheet, "B9"))
    except Exception as e:
        fail(sheet, str(e))
        return
    save_last_import_state(TOKEN_SHEET, file_input, path.rsplit('.', 1)[-1], path)
    print(f"   Saved {size_text(len(file_bytes))} to {path}")
    set_status(sheet, f"Saved: {path} ({size_text(len(file_bytes))})")


# === SECTION: REMOTE_LOADER ===
# =============================================================================
# REMOTE MODULE LOADER
# =============================================================================

_loaded_modules = {}

GITHUB_RAW_BASE = "https://raw.githubusercontent.com/amararun/xlwings-lite-apps-codes-docs/main/remote_modules"


async def load_remote_module(module_name: str, github_url: Optional[str] = None) -> dict:
    """
    Fetch a Python module from GitHub, execute it, and cache in memory.
    """
    if module_name in _loaded_modules:
        print(f"   Using cached module: {module_name}")
        return _loaded_modules[module_name]

    if github_url is None:
        github_url = f"{GITHUB_RAW_BASE}/{module_name}.py"

    print(f"   Fetching module from GitHub...")
    print(f"   URL: {github_url}")

    try:
        response = await pyfetch(github_url)

        if response.status != 200:
            raise Exception(f"HTTP {response.status}: Failed to fetch module")

        code = await response.string()
        print(f"   Downloaded: {len(code):,} bytes")

        module_globals = {
            '__name__': module_name,
            '__file__': github_url,
        }
        exec(code, module_globals)

        _loaded_modules[module_name] = module_globals
        print(f"   Module '{module_name}' loaded and cached")

        return module_globals

    except Exception as e:
        print(f"   ERROR loading module: {e}")
        raise


# === SECTION: SCRIPT_RUN_STATS ===
# =============================================================================
# SCRIPT: RUN STATS
# =============================================================================

@script(name="5. Run Stats (sample data)", button="[btn_stats]SHARE_LINK_ACCESS!D14")
async def run_stats(book: xw.Book):
    """Statistics for the sample data sets: detects cricket or IMDB data in the last imported
    DuckDB / Parquet file and loads the matching module from GitHub (remote_modules)."""
    sheet = book.sheets[LINK_SHEET]
    print("=" * 60)
    print("RUN STATS (cricket or IMDB samples)")
    print("=" * 60)

    duck_db_path = get_duckdb_temp_path()
    parquet_path = get_parquet_temp_path()
    if os.path.exists(duck_db_path):
        data_source, data_path = 'duckdb', duck_db_path
    elif os.path.exists(parquet_path):
        data_source, data_path = 'parquet', parquet_path
    else:
        fail(sheet, "no imported DuckDB or Parquet file yet. Import a cricket or IMDB sample from FILES_URLS first")
        return
    print(f"   Data file: {data_path}")
    set_status(sheet, "Running stats...")
    await book.flush()

    try:
        if data_source == 'duckdb':
            conn = duckdb.connect(data_path, read_only=True)
            table_names = [t[0] for t in conn.execute(
                "SELECT table_name FROM information_schema.tables WHERE table_schema='main'").fetchall()]
            column_sets = [[c[0].lower() for c in conn.execute(f'DESCRIBE "{t}"').fetchall()] for t in table_names]
        else:
            conn = duckdb.connect()
            table_names = ['parquet_data']
            column_sets = [[c[0].lower() for c in conn.execute(f"DESCRIBE SELECT * FROM '{data_path}'").fetchall()]]
        conn.close()
        print(f"   Tables: {table_names}")

        is_imdb = 'title_basics' in table_names
        is_cricket = any({'match_type', 'striker', 'bowler'} <= set(cols) for cols in column_sets)

        if is_imdb and data_source == 'duckdb':
            print("\n   Detected: IMDB data")
            module = await load_remote_module('imdb_stats')
            module['run_imdb_stats'](book)
            set_status(sheet, "Done: IMDB stats on the IMDB_STATS and IMDB_CHARTS sheets")
        elif is_cricket:
            print("\n   Detected: cricket data")
            module = await load_remote_module('cricket_stats')
            module['run_cricket_stats'](book, data_source, data_path)
            set_status(sheet, "Done: cricket stats on the DOCS sheet and one sheet per match type (e.g. ODI, ODI_CHARTS)")
        else:
            fail(sheet, "Run Stats works on the cricket and IMDB samples only (FILES_URLS). Your data is on the result sheet")
    except Exception as e:
        fail(sheet, f"stats failed: {e}")
