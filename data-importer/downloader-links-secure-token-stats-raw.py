# xlwings Lite - Secure Data Downloader with Token Access + Raw Import
#
# =============================================================================
# CODE MAP (for AI navigation)
# =============================================================================
# Search for these markers to find sections:
#   === SECTION: IMPORTS ===
#   === SECTION: OAUTH ===
#   === SECTION: HELPERS ===
#   === SECTION: URL_ROUTING ===
#   === SECTION: DATA_PROFILING ===
#   === SECTION: SCHEMA_DISPLAY ===
#   === SECTION: SCRIPT_TOKEN_ACCESS ===
#   === SECTION: SCRIPT_SHARELINK ===
#   === SECTION: SCRIPT_RAW_IMPORT ===      (NEW)
#   === SECTION: SCRIPT_RAW_TESTS ===       (NEW)
#   === SECTION: REMOTE_LOADER ===
#   === SECTION: SCRIPT_RUN_STATS ===
#
# =============================================================================
# SCRIPT FUNCTIONS SUMMARY
# =============================================================================
# @script import_via_token      - TOKEN_ACCESS sheet (Dropbox/GDrive/GitHub OAuth)
# @script import_via_sharelink  - SHARE_LINK_ACCESS sheet (public/private URLs)
# @script import_raw_sharelink  - SHARE_LINK_ACCESS sheet (raw file download)  (NEW)
# @script import_raw_token      - TOKEN_ACCESS sheet (raw file download)       (NEW)
# @script test_last_image       - Test last imported image file                 (NEW)
# @script test_last_zip         - Test last imported ZIP file                   (NEW)
# @script test_last_pdf         - Test last imported PDF file                   (NEW)
# @script run_stats             - Auto-detect and run stats module
#
# =============================================================================
# SHEET INPUT LAYOUTS
# =============================================================================
#
# SHARE_LINK_ACCESS / MASTER sheet:
#   B5  - Source URL (GitHub, Google Drive, Dropbox, direct URLs)
#   B7  - Raw Mode Flag (1 = raw download, 0 or empty = process to DuckDB)  (NEW)
#   B8  - Output Filename with extension (required when B7=1)                (NEW)
#   B12 - Private Repo Flag (1 = use auth proxy for GitHub, 0 = public)
#   B17 - Auth Proxy URL (required when B12=1)
#   D8  - Output: Saved file path (written by import_raw_sharelink)         (NEW)
#
# TOKEN_ACCESS sheet (processed mode):
#   B5  - Storage Provider (dropdown: "Dropbox", "Google Drive", "GitHub")
#   B6  - File Path/ID (format depends on provider)
#   B7  - Auth Proxy URL (optional, only for GitHub private repos)
#
# TOKEN_ACCESS sheet (raw mode - import_raw_token):                            (NEW)
#   B5  - Storage Provider (dropdown: "Dropbox", "Google Drive", "GitHub")
#   B6  - File Path/ID (format depends on provider)
#   B7  - Auth Proxy URL (optional, only for GitHub private repos)
#   B9  - Raw Mode Flag (1 = raw download)                                     (NEW)
#   B10 - Output Filename with extension (required when B9=1)                  (NEW)
#   D10 - Output: Saved file path (written by import_raw_token)                (NEW)
#
# =============================================================================
# ENVIRONMENT VARIABLES
# =============================================================================
# For Dropbox:   DROPBOX.REFRESH_TOKEN, DROPBOX.APP_KEY, DROPBOX.APP_SECRET
# For GDrive:    GDRIVE.REFRESH_TOKEN, GDRIVE.CLIENT_ID, GDRIVE.CLIENT_SECRET
# For GitHub:    GITHUB.PAT (Personal Access Token with repo scope)
#
# =============================================================================
# SUPPORTED FILE TYPES (processed mode)
# =============================================================================
# .parquet, .duckdb, .csv, .tsv, .txt, .pipe, .psv, .json
#
# RAW MODE: Any file type - user specifies filename with extension
#

# === SECTION: IMPORTS ===
import xlwings as xw
from xlwings import script
import pandas as pd
from typing import Tuple, Optional
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


def _ensure_requests_patched():
    """Patch requests library for Pyodide/browser environment."""
    try:
        import pyodide_http
        pyodide_http.patch_all()
    except ImportError:
        pass


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

    print(f"   Dropbox App Key: {app_key}")
    print(f"   Refresh Token: {refresh_token[:20]}...")

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
    print(f"   Access token obtained: {_dropbox_token[:20]}...")

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

    print(f"   Google Client ID: {client_id[:30]}...")
    print(f"   Refresh Token: {refresh_token[:20]}...")

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
    print(f"   Access token obtained: {_gdrive_token[:20]}...")

    return _gdrive_token


# === SECTION: HELPERS ===
# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def find_table_in_workbook(book: xw.Book, table_name: str) -> Tuple[Optional[xw.Sheet], any]:
    """Searches all sheets for a table and returns both the sheet and table objects."""
    for sheet in book.sheets:
        if table_name in sheet.tables:
            return sheet, sheet.tables[table_name]
    return None, None


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


def save_last_import_state(source_sheet: str, source_url: str, file_type: str, file_path: str = None) -> None:
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
        row_count = conn.execute(count_query).fetchone()[0]
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
                             delimiter: str, file_type: str) -> bool:
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

        row_count = conn.execute("SELECT COUNT(*) FROM imported_data").fetchone()[0]

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

        row_count = conn.execute("SELECT COUNT(*) FROM imported_data").fetchone()[0]

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

        for table_name, sample_df in sample_data.items():
            output_sheet.range(f"A{current_row}").value = f"SAMPLE: {table_name}"
            output_sheet.range(f"A{current_row}").font.bold = True
            output_sheet.range(f"A{current_row}").font.size = 11
            output_sheet.range(f"A{current_row}").color = '#E2EFDA'
            current_row += 1

            if not sample_df.empty:
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
        table_stats = {}
        sample_data = {}

        for table_name in table_names:
            row_count = conn.execute(f'SELECT COUNT(*) FROM "{table_name}"').fetchone()[0]
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


# === SECTION: SCRIPT_TOKEN_ACCESS ===
# =============================================================================
# SCRIPT: IMPORT DATA FROM TOKEN ACCESS
# =============================================================================

@script
async def import_via_token(book: xw.Book):
    """
    Import data using OAuth tokens (Dropbox, Google Drive) or PAT (GitHub).

    Reads from TOKEN_ACCESS sheet:
        B5 - Storage Provider (dropdown: "Dropbox", "Google Drive", "GitHub")
        B6 - File Path/ID (format depends on provider)
        B7 - Auth Proxy URL (optional, only for GitHub private repos)
    """
    print("=" * 60)
    print("TOKEN ACCESS DATA IMPORTER")
    print("Supports: Dropbox OAuth | Google Drive OAuth | GitHub PAT")
    print("=" * 60)

    print("\n   Cleaning up previous imports...")
    cleanup_count = 0

    for temp_path in [get_parquet_temp_path(), get_duckdb_temp_path(),
                      get_delimited_temp_path(), get_json_temp_path()]:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
                cleanup_count += 1
            except Exception as e:
                print(f"   Warning: Could not remove {os.path.basename(temp_path)}: {e}")

    if cleanup_count > 0:
        print(f"   Removed {cleanup_count} previous temp file(s)")
    else:
        print(f"   No previous temp files found")

    print("\n[1/4] Reading configuration from TOKEN_ACCESS sheet...")

    try:
        token_sheet = book.sheets['TOKEN_ACCESS']

        provider = token_sheet.range("B5").value
        if not provider or not isinstance(provider, str):
            print("\nERROR: Storage Provider not selected in TOKEN_ACCESS!B5")
            token_sheet.range("D5").value = "ERROR: Select a provider from dropdown"
            token_sheet.range("D5").font.color = '#FF0000'
            token_sheet.range("D5").font.bold = True
            return

        provider = provider.strip()
        valid_providers = ["Dropbox", "Google Drive", "GitHub"]
        if provider not in valid_providers:
            print(f"\nERROR: Invalid provider '{provider}' in TOKEN_ACCESS!B5")
            token_sheet.range("D5").value = f"ERROR: Must be {', '.join(valid_providers)}"
            token_sheet.range("D5").font.color = '#FF0000'
            token_sheet.range("D5").font.bold = True
            return

        token_sheet.range("D5").value = ""
        token_sheet.range("D6").value = ""
        token_sheet.range("D7").value = ""

        print(f"   Provider: {provider}")

        file_input = token_sheet.range("B6").value
        if not file_input or not isinstance(file_input, str):
            print(f"\nERROR: File Path/ID not found in TOKEN_ACCESS!B6")
            token_sheet.range("D6").value = "ERROR: Enter file path or ID"
            token_sheet.range("D6").font.color = '#FF0000'
            token_sheet.range("D6").font.bold = True
            return

        file_input = file_input.strip()
        print(f"   File Path/ID: {file_input[:60]}{'...' if len(file_input) > 60 else ''}")

        validation_error = None
        if provider == "Dropbox":
            if not file_input.startswith("/"):
                validation_error = "Dropbox paths must start with / (e.g., /folder/file.ext)"
        elif provider == "Google Drive":
            extracted_id = extract_gdrive_file_id(file_input)
            if extracted_id:
                if extracted_id != file_input:
                    print(f"   Extracted file ID from URL: {extracted_id}")
                file_input = extracted_id
            else:
                validation_error = "Invalid Google Drive input. Provide a file ID or URL"
        elif provider == "GitHub":
            if not file_input.lower().startswith('http'):
                validation_error = "GitHub requires full URL (e.g., https://github.com/...)"

        if validation_error:
            print(f"\nERROR: Format validation failed")
            token_sheet.range("D6").value = f"ERROR: {validation_error}"
            token_sheet.range("D6").font.color = '#FF0000'
            token_sheet.range("D6").font.bold = True
            return

        auth_proxy_url = None
        if provider == "GitHub":
            auth_proxy_url = token_sheet.range("B7").value
            if auth_proxy_url and isinstance(auth_proxy_url, str):
                auth_proxy_url = auth_proxy_url.strip()
                if auth_proxy_url:
                    print(f"   Auth Proxy: {auth_proxy_url}")

    except KeyError:
        print("ERROR: TOKEN_ACCESS sheet not found!")
        return
    except Exception as e:
        print(f"ERROR: Could not read TOKEN_ACCESS sheet: {e}")
        return

    print("\n[2/4] Downloading file...")
    download_start_time = time.time()

    try:
        if provider == "Dropbox":
            print("   Mode: Dropbox OAuth")
            try:
                access_token = await get_dropbox_token()
            except ValueError as e:
                print(f"\nERROR: Dropbox authentication failed: {e}")
                token_sheet.range("D6").value = "ERROR: Check Dropbox credentials"
                token_sheet.range("D6").font.color = '#FF0000'
                token_sheet.range("D6").font.bold = True
                return

            print(f"   Downloading from Dropbox: {file_input}")
            response = await pyfetch(
                "https://content.dropboxapi.com/2/files/download",
                method="POST",
                headers={
                    "Authorization": f"Bearer {access_token}",
                    "Dropbox-API-Arg": json.dumps({"path": file_input})
                }
            )

            if not response.ok:
                error_text = await response.text()
                print(f"\nERROR: Dropbox download failed: {response.status}")
                token_sheet.range("D6").value = f"ERROR: Download failed ({response.status})"
                token_sheet.range("D6").font.color = '#FF0000'
                token_sheet.range("D6").font.bold = True
                return

            file_bytes = await response.bytes()
            source_display = f"Dropbox: {file_input}"

        elif provider == "Google Drive":
            print("   Mode: Google Drive OAuth")
            try:
                access_token = await get_gdrive_token()
            except ValueError as e:
                print(f"\nERROR: Google Drive authentication failed: {e}")
                token_sheet.range("D6").value = "ERROR: Check Google Drive credentials"
                token_sheet.range("D6").font.color = '#FF0000'
                token_sheet.range("D6").font.bold = True
                return

            download_url = f"https://www.googleapis.com/drive/v3/files/{file_input}?alt=media"
            print(f"   Downloading from Google Drive: {file_input}")
            response = await pyfetch(
                download_url,
                method="GET",
                headers={"Authorization": f"Bearer {access_token}"}
            )

            if not response.ok:
                error_text = await response.text()
                print(f"\nERROR: Google Drive download failed: {response.status}")
                token_sheet.range("D6").value = f"ERROR: Download failed ({response.status})"
                token_sheet.range("D6").font.color = '#FF0000'
                token_sheet.range("D6").font.bold = True
                return

            file_bytes = await response.bytes()
            source_display = f"Google Drive: {file_input}"

        elif provider == "GitHub":
            print("   Mode: GitHub PAT")
            github_pat = os.environ.get("GITHUB.PAT")
            if not github_pat:
                print("\nERROR: GITHUB.PAT not found in environment variables")
                token_sheet.range("D6").value = "ERROR: GITHUB.PAT not set"
                token_sheet.range("D6").font.color = '#FF0000'
                token_sheet.range("D6").font.bold = True
                return

            print(f"   GitHub PAT: {github_pat[:10]}...{github_pat[-4:]}")

            fetch_url = file_input
            fetch_headers = {"Authorization": f"token {github_pat}"}

            if auth_proxy_url:
                fetch_url = build_private_proxy_url(file_input, auth_proxy_url)
                print(f"   Using auth proxy: {auth_proxy_url}")
            elif is_github_release_url(file_input):
                print("\nERROR: Auth Proxy URL required for private GitHub access")
                token_sheet.range("D6").value = "ERROR: Auth Proxy URL required in B7"
                token_sheet.range("D6").font.color = '#FF0000'
                token_sheet.range("D6").font.bold = True
                return

            print(f"   Downloading from GitHub: {file_input[:80]}...")
            response = await pyfetch(fetch_url, method="GET", headers=fetch_headers)

            if not response.ok:
                error_text = await response.text()
                print(f"\nERROR: GitHub download failed: {response.status}")
                token_sheet.range("D6").value = f"ERROR: Download failed ({response.status})"
                token_sheet.range("D6").font.color = '#FF0000'
                token_sheet.range("D6").font.bold = True
                return

            file_bytes = await response.bytes()
            source_display = f"GitHub: {file_input[:80]}{'...' if len(file_input) > 80 else ''}"

        file_size_mb = len(file_bytes) / (1024 * 1024)
        download_end_time = time.time()
        download_duration = download_end_time - download_start_time
        download_minutes = int(download_duration // 60)
        download_seconds = int(download_duration % 60)

        print(f"   Downloaded: {file_size_mb:.2f} MB in {download_minutes}m {download_seconds}s")

    except Exception as e:
        print(f"\nERROR: Download failed: {e}")
        token_sheet.range("D6").value = f"ERROR: {str(e)[:50]}"
        token_sheet.range("D6").font.color = '#FF0000'
        token_sheet.range("D6").font.bold = True
        return

    print("\n[3/4] Detecting file type...")

    file_type = detect_file_type_from_path(file_input)
    if file_type:
        print(f"   File type detected from path: {file_type.upper()}")
    else:
        file_type = detect_file_type_from_bytes(file_bytes)
        if file_type:
            print(f"   File type detected from content: {file_type.upper()}")

    if file_type == 'html_error':
        print("\nERROR: Received HTML page instead of data file")
        token_sheet.range("D6").value = "ERROR: Private link? Check credentials"
        token_sheet.range("D6").font.color = '#FF0000'
        token_sheet.range("D6").font.bold = True
        return

    if not file_type:
        print("\nERROR: Could not determine file type")
        token_sheet.range("D6").value = "ERROR: Could not detect file type"
        token_sheet.range("D6").font.color = '#FF0000'
        token_sheet.range("D6").font.bold = True
        return

    delimiter = None
    if file_type in ('csv', 'tsv', 'txt', 'pipe'):
        delimiter = detect_delimiter_from_content(file_bytes)

    if file_type == 'parquet':
        temp_path = get_parquet_temp_path()
    elif file_type == 'duckdb':
        temp_path = get_duckdb_temp_path()
    elif file_type == 'json':
        temp_path = get_json_temp_path()
    else:
        temp_path = get_delimited_temp_path()

    save_start_time = time.time()
    with open(temp_path, 'wb') as f:
        f.write(file_bytes)
    save_duration = time.time() - save_start_time
    save_seconds = int(save_duration)
    print(f"   Saved to temp file in {save_seconds}s")

    print(f"\n[4/4] Processing {file_type.upper()} file...")

    save_last_import_state("TOKEN_ACCESS", file_input, file_type, temp_path)

    if file_type == 'parquet':
        display_parquet_schema(book, source_display, file_size_mb, download_minutes, download_seconds, save_seconds)
    elif file_type == 'duckdb':
        display_duckdb_schema(book, source_display, file_size_mb, download_minutes, download_seconds, save_seconds)
    elif file_type == 'json':
        display_json_schema(book, source_display, file_size_mb, download_minutes, download_seconds, save_seconds)
    else:
        display_delimited_schema(book, source_display, file_size_mb, download_minutes, download_seconds, save_seconds, delimiter, file_type)


# === SECTION: SCRIPT_SHARELINK ===
# =============================================================================
# SCRIPT: IMPORT DATA FROM SHAREABLE LINKS
# =============================================================================

@script
async def import_via_sharelink(book: xw.Book):
    """
    Import data from shareable links (backward compatible with existing sheets).

    Reads from MASTER or SHARE_LINK_ACCESS sheet:
        B5  - Source URL
        B12 - Private Repo Flag (1 = use auth proxy, else normal)
        B17 - Auth Proxy URL (required when B12=1)
    """
    print("=" * 60)
    print("SHAREABLE LINK DATA IMPORTER")
    print("Supports public/private URLs with optional auth proxy")
    print("=" * 60)

    print("\n   Cleaning up previous imports...")
    cleanup_count = 0

    for temp_path in [get_parquet_temp_path(), get_duckdb_temp_path(),
                      get_delimited_temp_path(), get_json_temp_path()]:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
                cleanup_count += 1
            except Exception as e:
                print(f"   Warning: Could not remove {os.path.basename(temp_path)}: {e}")

    if cleanup_count > 0:
        print(f"   Removed {cleanup_count} previous temp file(s)")
    else:
        print(f"   No previous temp files found")

    print("\n[1/4] Reading configuration from sheet...")

    sheet_name = None
    for name in ['SHARE_LINK_ACCESS', 'MASTER']:
        if name in [s.name for s in book.sheets]:
            sheet_name = name
            break

    if not sheet_name:
        print("ERROR: Neither SHARE_LINK_ACCESS nor MASTER sheet found!")
        return

    try:
        master_sheet = book.sheets[sheet_name]
        print(f"   Using sheet: {sheet_name}")

        data_url = master_sheet.range("B5").value
        if not data_url or not isinstance(data_url, str):
            print(f"ERROR: No URL found in {sheet_name}!B5")
            return

        data_url = data_url.strip()
        print(f"   URL: {data_url[:80]}{'...' if len(data_url) > 80 else ''}")

        private_flag = master_sheet.range("B12").value
        is_private = False
        if private_flag is not None:
            if isinstance(private_flag, (int, float)):
                is_private = int(private_flag) == 1
            elif isinstance(private_flag, str):
                is_private = private_flag.strip() == "1"

        print(f"   Private Repo: {'Yes' if is_private else 'No'}")

        auth_proxy_url = None
        if is_private:
            auth_proxy_url = master_sheet.range("B17").value
            if not auth_proxy_url or not isinstance(auth_proxy_url, str) or not auth_proxy_url.strip():
                print("\nERROR: Auth Proxy URL required for private repos!")
                return
            auth_proxy_url = auth_proxy_url.strip()
            print(f"   Auth Proxy: {auth_proxy_url}")

        github_pat = None
        if is_private:
            github_pat = os.environ.get("GITHUB.PAT")
            if not github_pat:
                print("\nERROR: GITHUB.PAT environment variable not found!")
                return
            print(f"   GitHub PAT: {github_pat[:10]}...{github_pat[-4:]}")

    except Exception as e:
        print(f"ERROR: Could not read sheet: {e}")
        return

    data_url, was_converted = convert_google_drive_url(data_url)
    if was_converted:
        print(f"   [Google Drive URL converted to direct download format]")

    data_url, was_converted = convert_dropbox_url(data_url)
    if was_converted:
        print(f"   [Dropbox URL converted to direct download format (dl=1)]")

    print("\n[2/4] Analyzing URL...")

    use_proxy = needs_proxy(data_url, is_private)

    if is_private and is_github_url(data_url):
        print(f"   Mode: Private GitHub (using auth proxy)")
    elif use_proxy:
        print(f"   Mode: Public with CORS proxy")
    else:
        print(f"   Mode: Direct (no proxy)")

    file_type = detect_file_type(data_url)
    if file_type:
        print(f"   File type detected from URL: {file_type.upper()}")
    else:
        print(f"   File type: Not in URL, will check headers and content")

    fetch_headers = {}

    if is_private and is_github_url(data_url):
        fetch_url = build_private_proxy_url(data_url, auth_proxy_url)
        fetch_headers["Authorization"] = f"token {github_pat}"
        print(f"   Fetch URL: {auth_proxy_url}/...")
    elif use_proxy:
        fetch_url = build_public_proxy_url(data_url)
        print(f"   Fetch URL: {PUBLIC_PROXY_URL}/...")
    else:
        fetch_url = data_url
        print(f"   Fetch URL: Direct")

    print("\n[3/4] Downloading file...")

    try:
        download_start_time = time.time()

        if fetch_headers:
            response = await pyfetch(fetch_url, method="GET", headers=fetch_headers)
        else:
            response = await pyfetch(fetch_url, method="GET")

        if not response.ok:
            print(f"ERROR: Download failed with status {response.status}")
            master_sheet.range("D5").value = "ERROR: Download failed. Private link? Use TOKEN_ACCESS"
            master_sheet.range("D5").font.color = '#FF0000'
            master_sheet.range("D5").font.bold = True
            return

        file_bytes = await response.bytes()
        file_size_mb = len(file_bytes) / (1024 * 1024)

        download_end_time = time.time()
        download_duration = download_end_time - download_start_time
        download_minutes = int(download_duration // 60)
        download_seconds = int(download_duration % 60)

        print(f"   Downloaded: {file_size_mb:.2f} MB in {download_minutes}m {download_seconds}s")

        if not file_type:
            try:
                content_disposition = response.headers.get('Content-Disposition', '')
                file_type = detect_file_type_from_header(content_disposition)
                if file_type:
                    print(f"   File type detected from header: {file_type.upper()}")
            except Exception as e:
                pass

        if not file_type:
            file_type = detect_file_type_from_bytes(file_bytes)
            if file_type:
                print(f"   File type detected from content: {file_type.upper()}")

        if file_type == 'html_error':
            print("\nERROR: Received HTML page instead of data file")
            master_sheet.range("D5").value = "ERROR: Private link? Use TOKEN_ACCESS with credentials"
            master_sheet.range("D5").font.color = '#FF0000'
            master_sheet.range("D5").font.bold = True
            return

        if not file_type:
            print("\nERROR: Could not determine file type")
            master_sheet.range("D5").value = "ERROR: Invalid file. Private link? Use TOKEN_ACCESS"
            master_sheet.range("D5").font.color = '#FF0000'
            master_sheet.range("D5").font.bold = True
            return

        delimiter = None
        if file_type in ('csv', 'tsv', 'txt', 'pipe'):
            expected_delimiter_map = {
                'csv': ',',
                'tsv': '\t',
                'pipe': '|',
                'txt': None
            }
            expected_delimiter = expected_delimiter_map.get(file_type)
            actual_delimiter = detect_delimiter_from_content(file_bytes)
            delimiter_name = {',': 'comma', '\t': 'tab', '|': 'pipe'}.get(actual_delimiter, repr(actual_delimiter))

            if expected_delimiter and actual_delimiter != expected_delimiter:
                expected_name = {',': 'comma', '\t': 'tab', '|': 'pipe'}.get(expected_delimiter, repr(expected_delimiter))
                print(f"   File extension suggests {expected_name}, but content uses {delimiter_name}")

            delimiter = actual_delimiter

        if file_type == 'parquet':
            temp_path = get_parquet_temp_path()
        elif file_type == 'duckdb':
            temp_path = get_duckdb_temp_path()
        elif file_type == 'json':
            temp_path = get_json_temp_path()
        else:
            temp_path = get_delimited_temp_path()

        save_start_time = time.time()
        with open(temp_path, 'wb') as f:
            f.write(file_bytes)
        save_duration = time.time() - save_start_time
        save_seconds = int(save_duration)
        print(f"   Saved to temp file in {save_seconds}s")

    except Exception as e:
        print(f"ERROR: Failed to download file: {e}")
        master_sheet.range("D5").value = "ERROR: Download failed. Private link? Use TOKEN_ACCESS"
        master_sheet.range("D5").font.color = '#FF0000'
        master_sheet.range("D5").font.bold = True
        return

    print(f"\n[4/4] Processing {file_type.upper()} file...")

    save_last_import_state(sheet_name, data_url, file_type, temp_path)

    try:
        if file_type == 'parquet':
            result = display_parquet_schema(book, data_url, file_size_mb, download_minutes, download_seconds, save_seconds)
        elif file_type == 'duckdb':
            result = display_duckdb_schema(book, data_url, file_size_mb, download_minutes, download_seconds, save_seconds)
        elif file_type == 'json':
            result = display_json_schema(book, data_url, file_size_mb, download_minutes, download_seconds, save_seconds)
        else:
            result = display_delimited_schema(book, data_url, file_size_mb, download_minutes, download_seconds, save_seconds, delimiter, file_type)

        if result is False:
            master_sheet.range("D5").value = "ERROR: File processing failed. Private link? Use TOKEN_ACCESS"
            master_sheet.range("D5").font.color = '#FF0000'
            master_sheet.range("D5").font.bold = True
    except Exception as e:
        print(f"ERROR: File processing failed: {e}")
        master_sheet.range("D5").value = "ERROR: File processing failed. Private link? Use TOKEN_ACCESS"
        master_sheet.range("D5").font.color = '#FF0000'
        master_sheet.range("D5").font.bold = True


# === SECTION: SCRIPT_RAW_IMPORT ===
# =============================================================================
# SCRIPT: RAW FILE IMPORT (NEW)
# =============================================================================

@script
async def import_raw_sharelink(book: xw.Book):
    """
    Raw file download - no processing, just save with user-specified filename.

    Reads from SHARE_LINK_ACCESS sheet:
        B5  - Source URL
        B7  - Raw Mode Flag (must be 1 to use this function)
        B8  - Output Filename with extension (e.g., report.pdf, image.png)

    Output:
        D8  - Saved file path (dark green text)

    Note: This function is for PUBLIC URLs only. For private repos, use TOKEN_ACCESS.
    """
    print("=" * 60)
    print("RAW FILE IMPORTER")
    print("Downloads any file type without processing")
    print("=" * 60)

    # Find the sheet
    sheet_name = None
    for name in ['SHARE_LINK_ACCESS', 'MASTER']:
        if name in [s.name for s in book.sheets]:
            sheet_name = name
            break

    if not sheet_name:
        print("ERROR: Neither SHARE_LINK_ACCESS nor MASTER sheet found!")
        return

    try:
        master_sheet = book.sheets[sheet_name]
        print(f"   Using sheet: {sheet_name}")

        # Clear previous output
        master_sheet.range("D8").value = ""

        # B5: Source URL
        data_url = master_sheet.range("B5").value
        if not data_url or not isinstance(data_url, str):
            print(f"ERROR: No URL found in {sheet_name}!B5")
            master_sheet.range("D8").value = "ERROR: Enter URL in B5"
            master_sheet.range("D8").font.color = '#FF0000'
            master_sheet.range("D8").font.bold = True
            return

        data_url = data_url.strip()
        print(f"   URL: {data_url[:80]}{'...' if len(data_url) > 80 else ''}")

        # B7: Raw Mode Flag
        raw_flag = master_sheet.range("B7").value
        is_raw_mode = False
        if raw_flag is not None:
            if isinstance(raw_flag, (int, float)):
                is_raw_mode = int(raw_flag) == 1
            elif isinstance(raw_flag, str):
                is_raw_mode = raw_flag.strip() == "1"

        if not is_raw_mode:
            print("\nERROR: Raw mode not enabled!")
            print("   Set B7 = 1 to enable raw mode")
            print("   Or use import_via_sharelink for processed imports")
            master_sheet.range("D8").value = "ERROR: Set B7=1 for raw mode"
            master_sheet.range("D8").font.color = '#FF0000'
            master_sheet.range("D8").font.bold = True
            return

        # B8: Output Filename
        output_filename = master_sheet.range("B8").value
        if not output_filename or not isinstance(output_filename, str):
            print("\nERROR: Output filename not found in B8")
            print("   Enter filename with extension (e.g., report.pdf, image.png)")
            master_sheet.range("D8").value = "ERROR: Enter filename in B8 (e.g., file.pdf)"
            master_sheet.range("D8").font.color = '#FF0000'
            master_sheet.range("D8").font.bold = True
            return

        output_filename = output_filename.strip()

        # Normalize extension to lowercase
        if '.' in output_filename:
            name_part, ext_part = output_filename.rsplit('.', 1)
            output_filename = f"{name_part}.{ext_part.lower()}"

        print(f"   Output filename: {output_filename}")

        # Get file extension for state tracking
        file_ext = output_filename.rsplit('.', 1)[-1].lower() if '.' in output_filename else 'unknown'

    except Exception as e:
        print(f"ERROR: Could not read sheet: {e}")
        return

    # URL conversion for GDrive/Dropbox
    data_url, was_converted = convert_google_drive_url(data_url)
    if was_converted:
        print(f"   [Google Drive URL converted]")

    data_url, was_converted = convert_dropbox_url(data_url)
    if was_converted:
        print(f"   [Dropbox URL converted]")

    # Check if proxy needed (public URLs only for raw import)
    use_proxy = needs_proxy(data_url, is_private=False)

    if use_proxy:
        fetch_url = build_public_proxy_url(data_url)
        print(f"   Using CORS proxy")
    else:
        fetch_url = data_url
        print(f"   Direct fetch")

    # Download
    print("\n   Downloading file...")
    download_start_time = time.time()

    try:
        response = await pyfetch(fetch_url, method="GET")

        if not response.ok:
            print(f"ERROR: Download failed with status {response.status}")
            master_sheet.range("D8").value = f"ERROR: Download failed ({response.status})"
            master_sheet.range("D8").font.color = '#FF0000'
            master_sheet.range("D8").font.bold = True
            return

        file_bytes = await response.bytes()
        file_size_mb = len(file_bytes) / (1024 * 1024)

        download_duration = time.time() - download_start_time
        download_seconds = int(download_duration)

        print(f"   Downloaded: {file_size_mb:.2f} MB in {download_seconds}s")

    except Exception as e:
        print(f"ERROR: Failed to download: {e}")
        master_sheet.range("D8").value = f"ERROR: {str(e)[:40]}"
        master_sheet.range("D8").font.color = '#FF0000'
        master_sheet.range("D8").font.bold = True
        return

    # Save file
    temp_path = get_raw_temp_path(output_filename)

    try:
        with open(temp_path, 'wb') as f:
            f.write(file_bytes)
        print(f"   Saved to: {temp_path}")
    except Exception as e:
        print(f"ERROR: Failed to save file: {e}")
        master_sheet.range("D8").value = f"ERROR: Save failed - {str(e)[:30]}"
        master_sheet.range("D8").font.color = '#FF0000'
        master_sheet.range("D8").font.bold = True
        return

    # Save state for test functions
    save_last_import_state(sheet_name, data_url, file_ext, temp_path)

    # Write success to D8 (dark green)
    master_sheet.range("D8").value = f"Saved: {temp_path}"
    master_sheet.range("D8").font.color = '#006400'  # Dark green
    master_sheet.range("D8").font.bold = True

    print("\n" + "=" * 60)
    print("RAW FILE IMPORT COMPLETE!")
    print(f"Filename: {output_filename}")
    print(f"Size: {file_size_mb:.2f} MB")
    print(f"Path: {temp_path}")
    print("=" * 60)


@script
async def import_raw_token(book: xw.Book):
    """
    Raw file download via OAuth/PAT - no processing, just save with user-specified filename.

    Reads from TOKEN_ACCESS sheet:
        B5  - Storage Provider (dropdown: "Dropbox", "Google Drive", "GitHub")
        B6  - File Path/ID (format depends on provider)
        B7  - Auth Proxy URL (optional, only for GitHub private repos)
        B9  - Raw Mode Flag (must be 1 to use this function)
        B10 - Output Filename with extension (e.g., report.pdf, image.png)

    Output:
        D10 - Saved file path (dark green text)

    Note: Uses OAuth tokens (Dropbox/GDrive) or PAT (GitHub) for authentication.
    """
    print("=" * 60)
    print("RAW FILE IMPORTER (TOKEN ACCESS)")
    print("Downloads any file type via OAuth/PAT without processing")
    print("=" * 60)

    try:
        token_sheet = book.sheets['TOKEN_ACCESS']
    except KeyError:
        print("ERROR: TOKEN_ACCESS sheet not found!")
        return

    # Clear previous output
    token_sheet.range("D10").value = ""

    # B5: Storage Provider
    provider = token_sheet.range("B5").value
    if not provider or not isinstance(provider, str):
        print("\nERROR: Storage Provider not selected in TOKEN_ACCESS!B5")
        token_sheet.range("D10").value = "ERROR: Select provider in B5"
        token_sheet.range("D10").font.color = '#FF0000'
        token_sheet.range("D10").font.bold = True
        return

    provider = provider.strip()
    valid_providers = ["Dropbox", "Google Drive", "GitHub"]
    if provider not in valid_providers:
        print(f"\nERROR: Invalid provider '{provider}' in TOKEN_ACCESS!B5")
        token_sheet.range("D10").value = f"ERROR: Must be {', '.join(valid_providers)}"
        token_sheet.range("D10").font.color = '#FF0000'
        token_sheet.range("D10").font.bold = True
        return

    print(f"   Provider: {provider}")

    # B6: File Path/ID
    file_input = token_sheet.range("B6").value
    if not file_input or not isinstance(file_input, str):
        print(f"\nERROR: File Path/ID not found in TOKEN_ACCESS!B6")
        token_sheet.range("D10").value = "ERROR: Enter file path/ID in B6"
        token_sheet.range("D10").font.color = '#FF0000'
        token_sheet.range("D10").font.bold = True
        return

    file_input = file_input.strip()
    print(f"   File Path/ID: {file_input[:60]}{'...' if len(file_input) > 60 else ''}")

    # Validate file input format per provider
    validation_error = None
    if provider == "Dropbox":
        if not file_input.startswith("/"):
            validation_error = "Dropbox paths must start with / (e.g., /folder/file.ext)"
    elif provider == "Google Drive":
        extracted_id = extract_gdrive_file_id(file_input)
        if extracted_id:
            if extracted_id != file_input:
                print(f"   Extracted file ID from URL: {extracted_id}")
            file_input = extracted_id
        else:
            validation_error = "Invalid Google Drive input. Provide a file ID or URL"
    elif provider == "GitHub":
        if not file_input.lower().startswith('http'):
            validation_error = "GitHub requires full URL (e.g., https://github.com/...)"

    if validation_error:
        print(f"\nERROR: Format validation failed")
        token_sheet.range("D10").value = f"ERROR: {validation_error[:40]}"
        token_sheet.range("D10").font.color = '#FF0000'
        token_sheet.range("D10").font.bold = True
        return

    # B7: Auth Proxy URL (GitHub only)
    auth_proxy_url = None
    if provider == "GitHub":
        auth_proxy_url = token_sheet.range("B7").value
        if auth_proxy_url and isinstance(auth_proxy_url, str):
            auth_proxy_url = auth_proxy_url.strip()
            if auth_proxy_url:
                print(f"   Auth Proxy: {auth_proxy_url}")

    # B9: Raw Mode Flag
    raw_flag = token_sheet.range("B9").value
    is_raw_mode = False
    if raw_flag is not None:
        if isinstance(raw_flag, (int, float)):
            is_raw_mode = int(raw_flag) == 1
        elif isinstance(raw_flag, str):
            is_raw_mode = raw_flag.strip() == "1"

    if not is_raw_mode:
        print("\nERROR: Raw mode not enabled!")
        print("   Set B9 = 1 to enable raw mode")
        print("   Or use import_via_token for processed imports")
        token_sheet.range("D10").value = "ERROR: Set B9=1 for raw mode"
        token_sheet.range("D10").font.color = '#FF0000'
        token_sheet.range("D10").font.bold = True
        return

    # B10: Output Filename
    output_filename = token_sheet.range("B10").value
    if not output_filename or not isinstance(output_filename, str):
        print("\nERROR: Output filename not found in B10")
        print("   Enter filename with extension (e.g., report.pdf, image.png)")
        token_sheet.range("D10").value = "ERROR: Enter filename in B10 (e.g., file.pdf)"
        token_sheet.range("D10").font.color = '#FF0000'
        token_sheet.range("D10").font.bold = True
        return

    output_filename = output_filename.strip()

    # Normalize extension to lowercase
    if '.' in output_filename:
        name_part, ext_part = output_filename.rsplit('.', 1)
        output_filename = f"{name_part}.{ext_part.lower()}"

    print(f"   Output filename: {output_filename}")

    # Get file extension for state tracking
    file_ext = output_filename.rsplit('.', 1)[-1].lower() if '.' in output_filename else 'unknown'

    # Download file
    print("\n   Downloading file...")
    download_start_time = time.time()

    try:
        if provider == "Dropbox":
            print("   Mode: Dropbox OAuth")
            try:
                access_token = await get_dropbox_token()
            except ValueError as e:
                print(f"\nERROR: Dropbox authentication failed: {e}")
                token_sheet.range("D10").value = "ERROR: Check Dropbox credentials"
                token_sheet.range("D10").font.color = '#FF0000'
                token_sheet.range("D10").font.bold = True
                return

            response = await pyfetch(
                "https://content.dropboxapi.com/2/files/download",
                method="POST",
                headers={
                    "Authorization": f"Bearer {access_token}",
                    "Dropbox-API-Arg": json.dumps({"path": file_input})
                }
            )

            if not response.ok:
                print(f"\nERROR: Dropbox download failed: {response.status}")
                token_sheet.range("D10").value = f"ERROR: Download failed ({response.status})"
                token_sheet.range("D10").font.color = '#FF0000'
                token_sheet.range("D10").font.bold = True
                return

            file_bytes = await response.bytes()
            source_url = f"dropbox:{file_input}"

        elif provider == "Google Drive":
            print("   Mode: Google Drive OAuth")
            try:
                access_token = await get_gdrive_token()
            except ValueError as e:
                print(f"\nERROR: Google Drive authentication failed: {e}")
                token_sheet.range("D10").value = "ERROR: Check GDrive credentials"
                token_sheet.range("D10").font.color = '#FF0000'
                token_sheet.range("D10").font.bold = True
                return

            download_url = f"https://www.googleapis.com/drive/v3/files/{file_input}?alt=media"
            response = await pyfetch(
                download_url,
                method="GET",
                headers={"Authorization": f"Bearer {access_token}"}
            )

            if not response.ok:
                print(f"\nERROR: Google Drive download failed: {response.status}")
                token_sheet.range("D10").value = f"ERROR: Download failed ({response.status})"
                token_sheet.range("D10").font.color = '#FF0000'
                token_sheet.range("D10").font.bold = True
                return

            file_bytes = await response.bytes()
            source_url = f"gdrive:{file_input}"

        elif provider == "GitHub":
            print("   Mode: GitHub PAT")
            github_pat = os.environ.get("GITHUB.PAT")
            if not github_pat:
                print("\nERROR: GITHUB.PAT not found in environment variables")
                token_sheet.range("D10").value = "ERROR: GITHUB.PAT not set"
                token_sheet.range("D10").font.color = '#FF0000'
                token_sheet.range("D10").font.bold = True
                return

            print(f"   GitHub PAT: {github_pat[:10]}...{github_pat[-4:]}")

            fetch_url = file_input
            fetch_headers = {"Authorization": f"token {github_pat}"}

            if auth_proxy_url:
                fetch_url = build_private_proxy_url(file_input, auth_proxy_url)
                print(f"   Using auth proxy")
            elif is_github_release_url(file_input):
                print("\nERROR: Auth Proxy URL required for private GitHub releases")
                token_sheet.range("D10").value = "ERROR: Set Auth Proxy in B7"
                token_sheet.range("D10").font.color = '#FF0000'
                token_sheet.range("D10").font.bold = True
                return

            response = await pyfetch(fetch_url, method="GET", headers=fetch_headers)

            if not response.ok:
                print(f"\nERROR: GitHub download failed: {response.status}")
                token_sheet.range("D10").value = f"ERROR: Download failed ({response.status})"
                token_sheet.range("D10").font.color = '#FF0000'
                token_sheet.range("D10").font.bold = True
                return

            file_bytes = await response.bytes()
            source_url = file_input

        file_size_mb = len(file_bytes) / (1024 * 1024)
        download_duration = time.time() - download_start_time
        download_seconds = int(download_duration)

        print(f"   Downloaded: {file_size_mb:.2f} MB in {download_seconds}s")

    except Exception as e:
        print(f"ERROR: Failed to download: {e}")
        token_sheet.range("D10").value = f"ERROR: {str(e)[:40]}"
        token_sheet.range("D10").font.color = '#FF0000'
        token_sheet.range("D10").font.bold = True
        return

    # Save file
    temp_path = get_raw_temp_path(output_filename)

    try:
        with open(temp_path, 'wb') as f:
            f.write(file_bytes)
        print(f"   Saved to: {temp_path}")
    except Exception as e:
        print(f"ERROR: Failed to save file: {e}")
        token_sheet.range("D10").value = f"ERROR: Save failed - {str(e)[:30]}"
        token_sheet.range("D10").font.color = '#FF0000'
        token_sheet.range("D10").font.bold = True
        return

    # Save state for test functions
    save_last_import_state("TOKEN_ACCESS", source_url, file_ext, temp_path)

    # Write success to D10 (dark green)
    token_sheet.range("D10").value = f"Saved: {temp_path}"
    token_sheet.range("D10").font.color = '#006400'  # Dark green
    token_sheet.range("D10").font.bold = True

    print("\n" + "=" * 60)
    print("RAW FILE IMPORT COMPLETE!")
    print(f"Provider: {provider}")
    print(f"Filename: {output_filename}")
    print(f"Size: {file_size_mb:.2f} MB")
    print(f"Path: {temp_path}")
    print("=" * 60)


# === SECTION: SCRIPT_RAW_TESTS ===
# =============================================================================
# SCRIPT: TEST FUNCTIONS FOR RAW IMPORTS (NEW)
# =============================================================================

@script
async def test_last_image(book: xw.Book):
    """
    Test the last imported image file (PNG, JPG, GIF).
    Displays dimensions and basic info.
    """
    print("=" * 60)
    print("IMAGE FILE TEST")
    print("=" * 60)

    state = get_last_import_state()
    if not state:
        print("ERROR: No previous import found!")
        print("   Run import_raw_sharelink first")
        return

    file_path = state.get('file_path')
    file_type = state.get('file_type', '').lower()

    if not file_path or not os.path.exists(file_path):
        print(f"ERROR: File not found: {file_path}")
        return

    print(f"   Testing file: {file_path}")
    print(f"   File type: {file_type}")

    # Check if it's an image type
    image_types = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp']
    if file_type not in image_types:
        print(f"\nWARNING: Last import was '{file_type}', not an image")
        print(f"   Expected one of: {', '.join(image_types)}")

    try:
        # Try to use PIL/Pillow (available in Pyodide)
        from PIL import Image

        img = Image.open(file_path)
        width, height = img.size
        mode = img.mode
        img_format = img.format

        print("\n   IMAGE INFO:")
        print(f"   - Dimensions: {width} x {height} pixels")
        print(f"   - Mode: {mode}")
        print(f"   - Format: {img_format}")
        print(f"   - File size: {os.path.getsize(file_path) / 1024:.1f} KB")

        img.close()

        print("\n" + "=" * 60)
        print("IMAGE TEST PASSED!")
        print("=" * 60)

    except ImportError:
        print("\nERROR: PIL/Pillow not available in this environment")
        print("   Cannot read image dimensions")

        # Fallback: just check file exists and size
        file_size = os.path.getsize(file_path)
        print(f"\n   FILE INFO (basic):")
        print(f"   - File exists: Yes")
        print(f"   - File size: {file_size / 1024:.1f} KB")

    except Exception as e:
        print(f"\nERROR: Failed to read image: {e}")


@script
async def test_last_zip(book: xw.Book):
    """
    Test the last imported ZIP file.
    Lists entries (max 50) and shows total count.
    """
    print("=" * 60)
    print("ZIP FILE TEST")
    print("=" * 60)

    state = get_last_import_state()
    if not state:
        print("ERROR: No previous import found!")
        print("   Run import_raw_sharelink first")
        return

    file_path = state.get('file_path')
    file_type = state.get('file_type', '').lower()

    if not file_path or not os.path.exists(file_path):
        print(f"ERROR: File not found: {file_path}")
        return

    print(f"   Testing file: {file_path}")
    print(f"   File type: {file_type}")

    if file_type != 'zip':
        print(f"\nWARNING: Last import was '{file_type}', not a ZIP file")

    try:
        import zipfile

        if not zipfile.is_zipfile(file_path):
            print("\nERROR: File is not a valid ZIP archive")
            return

        with zipfile.ZipFile(file_path, 'r') as zf:
            entries = zf.namelist()
            total_count = len(entries)

            # Calculate total uncompressed size
            total_size = sum(info.file_size for info in zf.infolist())

            print(f"\n   ZIP CONTENTS:")
            print(f"   - Total entries: {total_count}")
            print(f"   - Total uncompressed size: {total_size / (1024*1024):.2f} MB")
            print(f"   - Archive size: {os.path.getsize(file_path) / (1024*1024):.2f} MB")

            # Show first 50 entries
            print(f"\n   ENTRIES (first 50 of {total_count}):")
            for i, entry in enumerate(entries[:50]):
                info = zf.getinfo(entry)
                size_kb = info.file_size / 1024
                print(f"   {i+1:3}. {entry} ({size_kb:.1f} KB)")

            if total_count > 50:
                print(f"   ... and {total_count - 50} more entries")

        print("\n" + "=" * 60)
        print("ZIP TEST PASSED!")
        print("=" * 60)

    except Exception as e:
        print(f"\nERROR: Failed to read ZIP: {e}")


@script
async def test_last_pdf(book: xw.Book):
    """
    Test the last imported PDF file.
    Attempts basic validation.
    """
    print("=" * 60)
    print("PDF FILE TEST")
    print("=" * 60)

    state = get_last_import_state()
    if not state:
        print("ERROR: No previous import found!")
        print("   Run import_raw_sharelink first")
        return

    file_path = state.get('file_path')
    file_type = state.get('file_type', '').lower()

    if not file_path or not os.path.exists(file_path):
        print(f"ERROR: File not found: {file_path}")
        return

    print(f"   Testing file: {file_path}")
    print(f"   File type: {file_type}")

    if file_type != 'pdf':
        print(f"\nWARNING: Last import was '{file_type}', not a PDF file")

    try:
        file_size = os.path.getsize(file_path)

        # Check PDF magic bytes
        with open(file_path, 'rb') as f:
            header = f.read(8)

        is_valid_pdf = header.startswith(b'%PDF')

        print(f"\n   PDF INFO:")
        print(f"   - File size: {file_size / 1024:.1f} KB")
        print(f"   - Valid PDF header: {'Yes' if is_valid_pdf else 'No'}")

        if is_valid_pdf:
            # Try to extract PDF version
            version = header[5:8].decode('ascii', errors='ignore')
            print(f"   - PDF version: {version}")

            # Try to count pages (basic method - look for /Page objects)
            with open(file_path, 'rb') as f:
                content = f.read()
                page_count = content.count(b'/Type /Page') - content.count(b'/Type /Pages')
                if page_count > 0:
                    print(f"   - Estimated pages: ~{page_count}")

            print("\n" + "=" * 60)
            print("PDF TEST PASSED!")
            print("=" * 60)
        else:
            print("\nERROR: File does not have valid PDF header")
            print(f"   Header bytes: {header}")

    except Exception as e:
        print(f"\nERROR: Failed to read PDF: {e}")


# === SECTION: REMOTE_LOADER ===
# =============================================================================
# REMOTE MODULE LOADER
# =============================================================================

_loaded_modules = {}

GITHUB_RAW_BASE = "https://raw.githubusercontent.com/amararun/xlwings-lite-apps-codes-docs/main/remote_modules"


async def load_remote_module(module_name: str, github_url: str = None) -> dict:
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

@script
async def run_stats(book: xw.Book):
    """
    Universal statistics generator - auto-detects data type and loads remote module.

    Auto-detects:
    - IMDB data (title_basics table) -> loads imdb_stats.py from GitHub
    - Cricket data (match_type, striker columns) -> loads cricket_stats.py from GitHub

    PREREQUISITE: Run 'import_via_sharelink' or 'import_via_token' first.
    """
    print("=" * 60)
    print("STATISTICS GENERATOR (Remote Module Loader)")
    print("Auto-detects: Cricket stats or IMDB stats")
    print("=" * 60)

    # Check for DuckDB file
    duck_db_path = get_duckdb_temp_path()
    parquet_path = get_parquet_temp_path()

    if not os.path.exists(duck_db_path) and not os.path.exists(parquet_path):
        print("\nERROR: No data file found!")
        print("   Run import_via_sharelink or import_via_token first")
        return

    # Determine data source
    if os.path.exists(duck_db_path):
        print(f"\n   Found DuckDB file: {duck_db_path}")
        data_source = 'duckdb'
    else:
        print(f"\n   Found Parquet file: {parquet_path}")
        data_source = 'parquet'

    # Connect and detect data type
    try:
        if data_source == 'duckdb':
            conn = duckdb.connect(duck_db_path, read_only=True)
            tables = conn.execute("SELECT table_name FROM information_schema.tables WHERE table_schema='main'").fetchall()
            table_names = [t[0] for t in tables]
            print(f"   Tables found: {table_names}")
        else:
            conn = duckdb.connect()
            table_names = ['parquet_data']

        # Detect IMDB data
        is_imdb = 'title_basics' in table_names or any('imdb' in t.lower() for t in table_names)

        # Detect Cricket data
        is_cricket = False
        if data_source == 'duckdb':
            for table in table_names:
                try:
                    cols = conn.execute(f'DESCRIBE "{table}"').fetchall()
                    col_names = [c[0].lower() for c in cols]
                    if 'match_type' in col_names or 'striker' in col_names or 'batting_team' in col_names:
                        is_cricket = True
                        break
                except:
                    pass
        else:
            try:
                cols = conn.execute(f"DESCRIBE SELECT * FROM '{parquet_path}'").fetchall()
                col_names = [c[0].lower() for c in cols]
                if 'match_type' in col_names or 'striker' in col_names or 'batting_team' in col_names:
                    is_cricket = True
            except:
                pass

        conn.close()

        # Load appropriate module
        if is_imdb:
            print("\n   Detected: IMDB data")
            module = await load_remote_module('imdb_stats')
            if 'run_imdb_stats' in module:
                module['run_imdb_stats'](book)
            else:
                print("   ERROR: run_imdb_stats function not found in module")

        elif is_cricket:
            print("\n   Detected: Cricket data")
            module = await load_remote_module('cricket_stats')
            if 'run_cricket_stats' in module:
                module['run_cricket_stats'](book, 'duckdb' if data_source == 'duckdb' else 'parquet',
                                            duck_db_path if data_source == 'duckdb' else parquet_path)
            else:
                print("   ERROR: run_cricket_stats function not found in module")

        else:
            print("\n   Could not auto-detect data type")
            print("   Expected: IMDB data (title_basics table) or Cricket data (match_type/striker columns)")

    except Exception as e:
        print(f"\nERROR: {e}")
