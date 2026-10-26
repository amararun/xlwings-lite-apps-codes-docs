# xlwings Lite - Local File Patterns Prototype
#
# =============================================================================
# PURPOSE
# =============================================================================
# Demonstrates xlwings Lite local file system access (Feb 2026 feature)
# Phase 1: Download files from GitHub releases to mounted local directory
# Phase 2: Unzip and concatenate CSVs into single file
# Phase 3: Import CSV to parquet, display schema
# Phase 4: Run cricket stats (tables + charts)
# Phase 5: Generate HTML report with tables and charts
# Phase 6: Generate PDF report with ReportLab
# Phase 7: Save parquet file to local drive
# Phase 8: Send email with PDF and HTML attachments via Brevo
# Phase 9: Generate PowerPoint presentation with charts and tables
#
# =============================================================================
# REQUIREMENTS
# =============================================================================
# 1. Add a local folder in xlwings Lite (Files > Local Folders > Add Folder)
# 2. Type its path in MASTER!B5 (default /data)
#
# =============================================================================
# SHEET LAYOUT (MASTER sheet)
# =============================================================================
# B5      - Local folder (added under Files > Local Folders), e.g. /data
# B8:B16  - Status of steps 1-9 (buttons in column D): blue = done, red = error
# B18     - Email recipient, B19 - sender email (verified in Brevo), B20 - sender name
#
# ENVIRONMENT VARIABLES:
# BREVO_API_KEY - Brevo API key for sending emails (set in xlwings Lite settings)
#
# =============================================================================
# SCRIPT FUNCTIONS
# =============================================================================
# @script s1_download_zip       - Downloads odis_csv2.zip from GitHub releases
# @script s2_process_zip        - Unzips and concatenates CSVs into single file
# @script s3_import_parquet     - Imports CSV to parquet, shows schema
# @script s4_run_stats          - Runs cricket stats analysis with charts
# @script s5_generate_html      - Creates HTML report with tables and charts
# @script s6_generate_pdf       - Creates PDF report with ReportLab
# @script s7_save_parquet       - Saves parquet from browser memory to local drive
# @script s8_send_email         - Sends email with PDF/HTML attachments via Brevo
# @script s9_generate_pptx      - Creates PowerPoint presentation with charts and tables
#
# =============================================================================

import xlwings as xw
from xlwings import script
from pyodide.http import pyfetch
import os
import time
import urllib.parse
import zipfile
import io
import csv
import duckdb
import pandas as pd
import tempfile
import matplotlib.pyplot as plt
from PIL import Image
import base64
import json
from typing import Optional


# =============================================================================
# CONFIGURATION
# =============================================================================

# GitHub release URL for cricket data
GITHUB_RELEASE_URL = "https://github.com/amararun/xlwings-lite-apps-codes-docs/releases/download/v1.0/odis_csv2.zip"

# CORS proxy (required for GitHub release downloads in browser)
CORS_PROXY_URL = "https://github-proxy-auth.tigzig.com"

# Local folder: read from MASTER!B5 at the start of every step (default /data).
# Add the folder in xlwings Lite: Files > Local Folders.
LOCAL_MOUNT_PATH = "/data"

# Filenames
ZIP_FILENAME = "odis_csv2.zip"
OUTPUT_CSV_FILENAME = "odis_combined.csv"
OUTPUT_PARQUET_FILENAME = "odis_combined.parquet"
OUTPUT_HTML_FILENAME = "odi_report.html"
OUTPUT_PDF_FILENAME = "odi_report.pdf"
BATTING_CHART_FILENAME = "odi_batting_chart.png"
BOWLING_CHART_FILENAME = "odi_bowling_chart.png"
OUTPUT_PPTX_FILENAME = "odi_report.pptx"

# Match type for this dataset
MATCH_TYPE = "ODI"


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def build_proxy_url(url: str) -> str:
    """Build CORS proxy URL for GitHub downloads."""
    encoded_url = urllib.parse.quote(url, safe='')
    return f"{CORS_PROXY_URL}/?url={encoded_url}"


def scalar(conn, query: str):
    """First value of the first row of a query."""
    row = conn.execute(query).fetchone()
    return row[0] if row else None


def format_size(size_bytes: int) -> str:
    """Format bytes to human readable size."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"


def format_duration(seconds: float) -> str:
    """Format duration to human readable string."""
    if seconds < 60:
        return f"{seconds:.1f}s"
    else:
        mins = int(seconds // 60)
        secs = int(seconds % 60)
        return f"{mins}m {secs}s"


def get_master_sheet(book: xw.Book):
    """Returns the MASTER sheet and reads the local folder from MASTER!B5."""
    global LOCAL_MOUNT_PATH
    if 'MASTER' not in [s.name for s in book.sheets]:
        print("ERROR: MASTER sheet not found")
        return None
    master_sheet = book.sheets['MASTER']
    folder = str(master_sheet.range("B5").value or "").strip()
    LOCAL_MOUNT_PATH = "/" + folder.strip("/") if folder.strip("/") else "/data"
    return master_sheet


def write_status(sheet, cell: str, message: str, is_error: bool = False):
    """Write status message to cell with appropriate formatting."""
    sheet.range(cell).value = message
    if is_error:
        sheet.range(cell).font.color = '#C00000'  # Red
    else:
        sheet.range(cell).font.color = '#1F4E79'  # Blue
    sheet.range(cell).font.bold = True


def get_mounted_directories() -> list:
    """Get list of mounted directories (excluding system dirs)."""
    system_dirs = {'dev', 'proc', 'tmp', 'home', 'lib', 'lib64', 'usr', 'etc', 'sys', 'var', 'bin', 'sbin',
                   'src', 'share', 'opt', 'mnt', 'root', 'run', 'srv', 'boot', 'media'}
    try:
        root_contents = os.listdir('/')
        mounted = [d for d in root_contents if d not in system_dirs and os.path.isdir(f'/{d}')]
        return mounted
    except Exception:
        return []


def print_mount_info():
    """Print information about the mounted directory."""
    print("\n--- Mount Path Info ---")
    print(f"   Expected mount: {LOCAL_MOUNT_PATH}")

    # Check for other mounted directories
    mounted_dirs = get_mounted_directories()
    if mounted_dirs:
        print(f"   Detected mounts: {mounted_dirs}")

    if os.path.exists(LOCAL_MOUNT_PATH):
        print(f"   Status: FOUND")

        # List contents
        try:
            contents = os.listdir(LOCAL_MOUNT_PATH)
            print(f"   Contents: {len(contents)} items")
            if contents and len(contents) <= 10:
                for item in contents[:10]:
                    print(f"      - {item}")
        except Exception as e:
            print(f"   Contents: (unable to list: {e})")
    else:
        print(f"   Status: NOT FOUND")
        if mounted_dirs:
            print(f"   NOTE: Found other mounts: /{', /'.join(mounted_dirs)}")
            print("   Type one of these folders in MASTER!B5")
        else:
            print(f"   NOTE: No drives appear to be mounted")
            print(f"   Please mount your local drive via Files menu")

    print("-" * 25)


def get_mount_error_message() -> str:
    """Get appropriate error message based on mount status."""
    mounted_dirs = get_mounted_directories()

    if not mounted_dirs:
        return f"ERROR: No local folder added. In xlwings Lite: Files > Local Folders > Add Folder, then type its path in B5 (now {LOCAL_MOUNT_PATH})"
    else:
        return f"ERROR: Folder {LOCAL_MOUNT_PATH} not found. Added folders: /{', /'.join(mounted_dirs)}. Type one of them in B5."


# =============================================================================
# PHASE 1: DOWNLOAD
# =============================================================================

@script(name="1. Download Zip", button="[btn_s1]MASTER!D8")
async def s1_download_zip(book: xw.Book):
    """
    Download cricket zip file from GitHub releases to local mount.
    Status output: MASTER!B8
    """
    print("=" * 60)
    print("Phase 1: Download Zip")
    print("=" * 60)

    # Get sheet for status output
    master_sheet = get_master_sheet(book)
    if not master_sheet:
        print("ERROR: MASTER sheet not found!")
        return

    # Clear previous status
    master_sheet.range("B8").value = ""

    # Print mount info for debugging
    print_mount_info()

    # Check if mount path exists
    if not os.path.exists(LOCAL_MOUNT_PATH):
        error_msg = get_mount_error_message()
        print(f"\n{error_msg}")
        write_status(master_sheet, "B8", error_msg, is_error=True)
        return

    print(f"\n   Source: GitHub release v1.0")
    print(f"   File: {ZIP_FILENAME}")

    # Build proxy URL
    fetch_url = build_proxy_url(GITHUB_RELEASE_URL)
    print(f"\n   Using CORS proxy for GitHub download")

    # Download
    print("\n   Downloading...")
    download_start = time.time()

    try:
        response = await pyfetch(fetch_url, method="GET")

        if not response.ok:
            error_msg = f"ERROR: Download failed (HTTP {response.status})"
            print(f"\n{error_msg}")
            write_status(master_sheet, "B8", error_msg, is_error=True)
            return

        file_bytes = await response.bytes()
        download_duration = time.time() - download_start

        print(f"   Downloaded: {format_size(len(file_bytes))} in {format_duration(download_duration)}")

    except Exception as e:
        error_msg = f"ERROR: Download failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B8", error_msg, is_error=True)
        return

    # Save to local mount
    output_path = os.path.join(LOCAL_MOUNT_PATH, ZIP_FILENAME)
    print(f"\n   Saving to: {output_path}")

    try:
        with open(output_path, 'wb') as f:
            f.write(file_bytes)

        # Verify file was saved
        if os.path.exists(output_path):
            saved_size = os.path.getsize(output_path)
            print(f"   Verified: {format_size(saved_size)}")
        else:
            error_msg = "ERROR: File not found after save"
            print(f"\n{error_msg}")
            write_status(master_sheet, "B8", error_msg, is_error=True)
            return

    except Exception as e:
        error_msg = f"ERROR: Save failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B8", error_msg, is_error=True)
        return

    # Success - write to B5
    success_msg = f"Downloaded: {output_path} ({format_size(len(file_bytes))})"
    write_status(master_sheet, "B8", success_msg, is_error=False)

    print("\n" + "=" * 60)
    print("DOWNLOAD COMPLETE!")
    print(f"   Path: {output_path}")
    print(f"   Size: {format_size(len(file_bytes))}")
    print(f"   Time: {format_duration(download_duration)}")
    print("\nNext: click '2. Combine CSVs' to extract and combine CSVs")
    print("=" * 60)


# =============================================================================
# PHASE 2: UNZIP AND CONCATENATE
# =============================================================================

@script(name="2. Combine CSVs", button="[btn_s2]MASTER!D9")
async def s2_process_zip(book: xw.Book):
    """
    Unzip and concatenate all ball-by-ball CSVs into single file.
    Skips _info.csv files (metadata only).
    Status output: MASTER!B9
    """
    print("=" * 60)
    print("Phase 2: Process Zip")
    print("=" * 60)

    # Get sheet for status output
    master_sheet = get_master_sheet(book)
    if not master_sheet:
        print("ERROR: MASTER sheet not found!")
        return

    # Clear previous status
    master_sheet.range("B9").value = ""

    # Print mount info
    print_mount_info()

    # Check zip file exists
    zip_path = os.path.join(LOCAL_MOUNT_PATH, ZIP_FILENAME)
    if not os.path.exists(zip_path):
        error_msg = f"ERROR: Zip not found: {zip_path}"
        print(f"\n{error_msg}")
        print("Click '1. Download Zip' first")
        write_status(master_sheet, "B9", error_msg, is_error=True)
        return

    zip_size = os.path.getsize(zip_path)
    print(f"\n   Zip file: {zip_path}")
    print(f"   Size: {format_size(zip_size)}")

    # Process zip
    print("\n   Processing zip file...")
    process_start = time.time()

    try:
        all_rows = []
        header = None
        data_file_count = 0
        info_file_count = 0

        with zipfile.ZipFile(zip_path, 'r') as zf:
            file_list = zf.namelist()
            total_files = len(file_list)
            print(f"   Total files in zip: {total_files}")

            for filename in file_list:
                # Skip non-CSV files
                if not filename.endswith('.csv'):
                    continue

                # Skip _info.csv files (metadata)
                if filename.endswith('_info.csv'):
                    info_file_count += 1
                    continue

                # Skip README or other non-data files
                if 'README' in filename.upper():
                    continue

                # Process ball-by-ball CSV
                data_file_count += 1

                with zf.open(filename) as f:
                    content = f.read().decode('utf-8')
                    reader = csv.reader(io.StringIO(content))

                    for row_idx, row in enumerate(reader):
                        if row_idx == 0:
                            # Header row
                            if header is None:
                                header = row
                                all_rows.append(header)
                            continue
                        # Data row
                        all_rows.append(row)

                # Progress update every 500 files
                if data_file_count % 500 == 0:
                    print(f"      Processed {data_file_count} data files...")

        print(f"\n   Data files: {data_file_count}")
        print(f"   Info files skipped: {info_file_count}")
        print(f"   Total rows: {len(all_rows) - 1:,} (plus header)")

    except Exception as e:
        error_msg = f"ERROR: Zip processing failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B9", error_msg, is_error=True)
        return

    # Write combined CSV
    output_path = os.path.join(LOCAL_MOUNT_PATH, OUTPUT_CSV_FILENAME)
    print(f"\n   Writing combined CSV: {output_path}")

    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(all_rows)

        # Verify
        if os.path.exists(output_path):
            csv_size = os.path.getsize(output_path)
            print(f"   Written: {format_size(csv_size)}")
        else:
            error_msg = "ERROR: CSV not found after write"
            print(f"\n{error_msg}")
            write_status(master_sheet, "B9", error_msg, is_error=True)
            return

    except Exception as e:
        error_msg = f"ERROR: CSV write failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B9", error_msg, is_error=True)
        return

    process_duration = time.time() - process_start

    # Success - write to B6
    row_count = len(all_rows) - 1
    success_msg = f"Created: {output_path} ({row_count:,} rows, {format_size(csv_size)})"
    write_status(master_sheet, "B9", success_msg, is_error=False)

    print("\n" + "=" * 60)
    print("PROCESSING COMPLETE!")
    print(f"   Output: {output_path}")
    print(f"   Rows: {row_count:,}")
    print(f"   Size: {format_size(csv_size)}")
    print(f"   Time: {format_duration(process_duration)}")
    print(f"   Files processed: {data_file_count}")
    print(f"   Files skipped: {info_file_count} (_info.csv)")
    print("\nNext: click '3. Import to Parquet' to convert to parquet")
    print("=" * 60)


# =============================================================================
# PHASE 3: IMPORT CSV TO PARQUET
# =============================================================================

def ensure_sheet_exists(book: xw.Book, sheet_name: str) -> xw.Sheet:
    """Ensures a sheet exists, deleting any existing one first."""
    for s in book.sheets:
        if s.name == sheet_name:
            s.delete()
            break
    return book.sheets.add(name=sheet_name)


def get_parquet_temp_path() -> str:
    """Returns temp path for parquet file (browser memory, not local drive)."""
    return os.path.join(tempfile.gettempdir(), OUTPUT_PARQUET_FILENAME)


@script(name="3. Import to Parquet", button="[btn_s3]MASTER!D10")
async def s3_import_parquet(book: xw.Book):
    """
    Import CSV to parquet format with DuckDB.
    Parquet stays in browser memory (temp), not written to local drive.
    Adds match_type column and displays schema.
    Status output: MASTER!B10
    """
    print("=" * 60)
    print("Phase 3: Import to Parquet (Memory)")
    print("=" * 60)

    # Get sheet for status output
    master_sheet = get_master_sheet(book)
    if not master_sheet:
        print("ERROR: MASTER sheet not found!")
        return

    # Clear previous status
    master_sheet.range("B10").value = ""

    # Print mount info
    print_mount_info()

    # Check CSV exists
    csv_path = os.path.join(LOCAL_MOUNT_PATH, OUTPUT_CSV_FILENAME)
    if not os.path.exists(csv_path):
        error_msg = f"ERROR: CSV not found: {csv_path}"
        print(f"\n{error_msg}")
        print("Click '2. Combine CSVs' first")
        write_status(master_sheet, "B10", error_msg, is_error=True)
        return

    csv_size = os.path.getsize(csv_path)
    print(f"\n   CSV file: {csv_path}")
    print(f"   Size: {format_size(csv_size)}")

    # Convert to parquet with DuckDB - store in browser temp (not local drive)
    print("\n   Converting to parquet (browser memory)...")
    convert_start = time.time()

    # Parquet goes to temp directory (browser memory), NOT to local mount
    parquet_path = get_parquet_temp_path()

    try:
        conn = duckdb.connect()

        # Read CSV and add match_type column
        conn.execute(f"""
            CREATE TABLE cricket_data AS
            SELECT *, '{MATCH_TYPE}' as match_type
            FROM read_csv_auto('{csv_path}', header=true, all_varchar=true)
        """)

        # Get row count
        row_count = scalar(conn, "SELECT COUNT(*) FROM cricket_data")
        print(f"   Loaded: {row_count:,} rows")

        # Export to parquet in temp (browser memory)
        conn.execute(f"COPY cricket_data TO '{parquet_path}' (FORMAT 'parquet')")

        conn.close()

        # Verify parquet was created
        if os.path.exists(parquet_path):
            parquet_size = os.path.getsize(parquet_path)
            print(f"   Parquet size: {format_size(parquet_size)}")
        else:
            error_msg = "ERROR: Parquet conversion failed"
            print(f"\n{error_msg}")
            write_status(master_sheet, "B10", error_msg, is_error=True)
            return

    except Exception as e:
        error_msg = f"ERROR: Conversion failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B10", error_msg, is_error=True)
        return

    convert_duration = time.time() - convert_start

    # Display schema in new sheet
    print("\n   Creating SCHEMA sheet...")
    try:
        conn = duckdb.connect()
        conn.execute(f"CREATE VIEW data AS SELECT * FROM '{parquet_path}'")

        # Get schema
        schema_result = conn.execute("DESCRIBE data").fetchall()

        schema_sheet = ensure_sheet_exists(book, "SCHEMA")
        schema_sheet.range("A1").value = "ODI Cricket Data - Schema"
        schema_sheet.range("A1").font.bold = True
        schema_sheet.range("A1").font.size = 14

        schema_sheet.range("A2").value = f"Source: {parquet_path}"
        schema_sheet.range("A3").value = f"Rows: {row_count:,} | Size: {format_size(parquet_size)}"

        # Write schema table
        schema_sheet.range("A5").value = "Column"
        schema_sheet.range("B5").value = "Type"
        schema_sheet.range("A5:B5").font.bold = True
        schema_sheet.range("A5:B5").color = '#4472C4'
        schema_sheet.range("A5:B5").font.color = '#FFFFFF'

        for i, row in enumerate(schema_result):
            schema_sheet.range(f"A{6+i}").value = row[0]
            schema_sheet.range(f"B{6+i}").value = row[1]

        conn.close()
        print(f"   Schema: {len(schema_result)} columns")

    except Exception as e:
        print(f"   Schema sheet error: {e}")

    # Success - write to B7
    success_msg = f"Parquet imported to browser memory ({row_count:,} rows, {format_size(parquet_size)}) - ready for processing"
    write_status(master_sheet, "B10", success_msg, is_error=False)

    print("\n" + "=" * 60)
    print("IMPORT COMPLETE!")
    print(f"   Location: Browser memory (temp)")
    print(f"   Rows: {row_count:,}")
    print(f"   Size: {format_size(parquet_size)} (was {format_size(csv_size)} CSV)")
    print(f"   Compression: {(1 - parquet_size/csv_size)*100:.1f}% smaller")
    print(f"   Time: {format_duration(convert_duration)}")
    print("\n   NOTE: Parquet is in browser memory, not written to local drive yet")
    print("\nNext: click '4. Run Stats' for analysis and charts")
    print("=" * 60)


# =============================================================================
# PHASE 4: RUN STATS
# =============================================================================

NOT_BATTER_OUT = "('retired hurt', 'retired not out')"
NOT_BOWLER_WICKET = ("('run out', 'retired hurt', 'retired not out', 'retired out', "
                     "'obstructing the field', 'handled the ball', 'timed out')")


def get_batting_stats_query(match_type: str, limit: int = 15) -> str:
    """Top batters by runs, scorecard definitions: balls faced exclude wides, a dismissal counts at
    either end (run out as non-striker), super overs (innings 3+) left out."""
    return f"""
    WITH d AS (
        SELECT * FROM data
        WHERE match_type = '{match_type}' AND TRY_CAST(innings AS INTEGER) <= 2
    ),
    bat AS (
        SELECT striker AS Player,
            SUM(COALESCE(TRY_CAST(runs_off_bat AS INTEGER), 0)) AS Runs,
            SUM(CASE WHEN COALESCE(TRY_CAST(wides AS INTEGER), 0) = 0 THEN 1 ELSE 0 END) AS BF,
            SUM(CASE WHEN TRY_CAST(runs_off_bat AS INTEGER) = 4 THEN 1 ELSE 0 END) AS fours,
            SUM(CASE WHEN TRY_CAST(runs_off_bat AS INTEGER) = 6 THEN 1 ELSE 0 END) AS sixes
        FROM d GROUP BY striker
    ),
    crease AS (
        SELECT Player, COUNT(DISTINCT match_id) AS Mat
        FROM (SELECT match_id, striker AS Player FROM d UNION ALL SELECT match_id, non_striker AS Player FROM d)
        GROUP BY Player
    ),
    outs AS (
        SELECT player_dismissed AS Player, COUNT(*) AS Outs
        FROM d
        WHERE wicket_type IS NOT NULL AND wicket_type NOT IN {NOT_BATTER_OUT}
        GROUP BY player_dismissed
    )
    SELECT b.Player, c.Mat,
        CAST(b.Runs AS INTEGER) AS Runs,
        CAST(b.BF AS INTEGER) AS BF,
        CASE WHEN COALESCE(o.Outs, 0) > 0 THEN ROUND(b.Runs / o.Outs, 2) END AS "Avg",
        ROUND(b.Runs * 100.0 / NULLIF(b.BF, 0), 2) AS SR,
        CAST(b.fours AS INTEGER) AS "4s",
        CAST(b.sixes AS INTEGER) AS "6s"
    FROM bat b
    JOIN crease c ON c.Player = b.Player
    LEFT JOIN outs o ON o.Player = b.Player
    WHERE c.Mat >= 5
    ORDER BY Runs DESC, b.Player
    LIMIT {limit}
    """


def get_bowling_stats_query(match_type: str, limit: int = 15) -> str:
    """Top bowlers by wickets, scorecard definitions: overs from legal balls only (no wides or
    no-balls), runs conceded = off the bat + wides + no-balls, super overs left out."""
    return f"""
    WITH d AS (
        SELECT *,
            COALESCE(TRY_CAST(runs_off_bat AS INTEGER), 0) AS rb,
            COALESCE(TRY_CAST(wides AS INTEGER), 0) AS wd,
            COALESCE(TRY_CAST(noballs AS INTEGER), 0) AS nb
        FROM data
        WHERE match_type = '{match_type}' AND TRY_CAST(innings AS INTEGER) <= 2
    ),
    agg AS (
        SELECT bowler AS Player,
            COUNT(DISTINCT match_id) AS Mat,
            SUM(CASE WHEN wd = 0 AND nb = 0 THEN 1 ELSE 0 END) AS Balls,
            SUM(rb + wd + nb) AS Runs,
            SUM(CASE WHEN wicket_type IS NOT NULL AND wicket_type NOT IN {NOT_BOWLER_WICKET}
                     THEN 1 ELSE 0 END) AS Wkts,
            SUM(CASE WHEN wd = 0 AND nb = 0 AND rb = 0 THEN 1 ELSE 0 END) AS Dots
        FROM d GROUP BY bowler
    )
    SELECT Player, Mat,
        CAST(Balls // 6 AS DOUBLE) + (Balls % 6) / 10.0 AS Overs,
        CAST(Runs AS INTEGER) AS Runs,
        CAST(Wkts AS INTEGER) AS Wkts,
        ROUND(Runs * 6.0 / NULLIF(Balls, 0), 2) AS Econ,
        CASE WHEN Wkts > 0 THEN ROUND(Runs * 1.0 / Wkts, 2) END AS "Avg",
        ROUND(Dots * 100.0 / NULLIF(Balls, 0), 1) AS "Dot%"
    FROM agg
    WHERE Mat >= 5
    ORDER BY Wkts DESC, Player
    LIMIT {limit}
    """


def create_batting_chart(df: pd.DataFrame, match_type: str, num_players: int = 15) -> str:
    """Creates batting chart. Returns path to saved image."""
    chart_df = df.head(num_players).copy()
    players = [p[:12] + '..' if len(p) > 12 else p for p in chart_df['Player'].tolist()]
    x_pos = range(len(players))

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 6), sharex=True, gridspec_kw={'hspace': 0.1})

    # TOP: Runs
    runs = chart_df['Runs'].tolist()
    bars = ax1.bar(x_pos, runs, color='#4472C4', alpha=0.85, width=0.7)
    for bar, run in zip(bars, runs):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 50,
                f'{int(run):,}', ha='center', va='bottom', fontsize=7)
    ax1.set_title(f'Top {num_players} Batsmen - {match_type}', fontsize=12, fontweight='bold')
    ax1.set_ylabel('Runs', fontsize=10, fontweight='bold')
    ax1.grid(axis='y', alpha=0.3)

    # BOTTOM: SR & Avg
    sr = chart_df['SR'].tolist()
    avg = chart_df['Avg'].tolist()
    ax2.plot(x_pos, sr, 'o-', color='#FF6D00', linewidth=2, markersize=6, label='Strike Rate')
    ax2.plot(x_pos, avg, 's--', color='#1F4E79', linewidth=2, markersize=5, label='Average')
    ax2.set_ylabel('SR / Avg', fontsize=10, fontweight='bold')
    ax2.legend(loc='upper right', fontsize=9)
    ax2.grid(axis='y', alpha=0.3)
    ax2.set_xticks(x_pos)
    ax2.set_xticklabels(players, rotation=45, ha='right', fontsize=8)

    plt.tight_layout()
    chart_path = os.path.join(tempfile.gettempdir(), f"batting_chart_{match_type.lower()}.png")
    fig.savefig(chart_path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return chart_path


def create_bowling_chart(df: pd.DataFrame, match_type: str, num_players: int = 15) -> str:
    """Creates bowling chart. Returns path to saved image."""
    chart_df = df.head(num_players).copy()
    players = [p[:12] + '..' if len(p) > 12 else p for p in chart_df['Player'].tolist()]
    x_pos = range(len(players))

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 6), sharex=True, gridspec_kw={'hspace': 0.1})

    # TOP: Wickets
    wickets = chart_df['Wkts'].tolist()
    bars = ax1.bar(x_pos, wickets, color='#4472C4', alpha=0.85, width=0.7)
    for bar, wkt in zip(bars, wickets):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f'{int(wkt)}', ha='center', va='bottom', fontsize=7)
    ax1.set_title(f'Top {num_players} Bowlers - {match_type}', fontsize=12, fontweight='bold')
    ax1.set_ylabel('Wickets', fontsize=10, fontweight='bold')
    ax1.grid(axis='y', alpha=0.3)

    # BOTTOM: Econ & Avg
    econ = chart_df['Econ'].tolist()
    avg = chart_df['Avg'].tolist()
    ax2.plot(x_pos, econ, 'o-', color='#FF6D00', linewidth=2, markersize=6, label='Economy')
    ax2.plot(x_pos, avg, 's--', color='#1F4E79', linewidth=2, markersize=5, label='Average')
    ax2.set_ylabel('Econ / Avg', fontsize=10, fontweight='bold')
    ax2.legend(loc='upper right', fontsize=9)
    ax2.grid(axis='y', alpha=0.3)
    ax2.set_xticks(x_pos)
    ax2.set_xticklabels(players, rotation=45, ha='right', fontsize=8)

    plt.tight_layout()
    chart_path = os.path.join(tempfile.gettempdir(), f"bowling_chart_{match_type.lower()}.png")
    fig.savefig(chart_path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return chart_path


@script(name="4. Run Stats", button="[btn_s4]MASTER!D11")
async def s4_run_stats(book: xw.Book):
    """
    Run cricket statistics analysis and create charts.
    Status output: MASTER!B11
    """
    print("=" * 60)
    print("Phase 4: Run Stats")
    print("=" * 60)

    # Get sheet for status output
    master_sheet = get_master_sheet(book)
    if not master_sheet:
        print("ERROR: MASTER sheet not found!")
        return

    # Clear previous status
    master_sheet.range("B11").value = ""

    # Check parquet exists in browser memory (temp)
    parquet_path = get_parquet_temp_path()
    if not os.path.exists(parquet_path):
        error_msg = "ERROR: Parquet not in memory. Click '3. Import to Parquet' first"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B11", error_msg, is_error=True)
        return

    print(f"\n   Parquet: browser memory")

    # Connect and run queries
    print("\n   Loading data...")
    stats_start = time.time()

    try:
        conn = duckdb.connect()
        conn.execute(f"CREATE VIEW data AS SELECT * FROM '{parquet_path}'")

        row_count = scalar(conn, "SELECT COUNT(*) FROM data")
        print(f"   Rows: {row_count:,}")

        # Run batting stats
        print("\n   Generating batting stats...")
        batting_query = get_batting_stats_query(MATCH_TYPE, limit=15)
        batting_df = conn.execute(batting_query).fetchdf()
        print(f"   Top batsmen: {len(batting_df)}")

        # Run bowling stats
        print("\n   Generating bowling stats...")
        bowling_query = get_bowling_stats_query(MATCH_TYPE, limit=15)
        bowling_df = conn.execute(bowling_query).fetchdf()
        print(f"   Top bowlers: {len(bowling_df)}")

        conn.close()

    except Exception as e:
        error_msg = f"ERROR: Query failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B11", error_msg, is_error=True)
        return

    # Create stats sheet
    print("\n   Creating ODI stats sheet...")
    try:
        stats_sheet = ensure_sheet_exists(book, "ODI")

        stats_sheet.range("A1").value = f"ODI Player Statistics"
        stats_sheet.range("A1").font.bold = True
        stats_sheet.range("A1").font.size = 16

        stats_sheet.range("A2").value = f"Source: {parquet_path} | Generated: {time.strftime('%Y-%m-%d %H:%M')}"
        stats_sheet.range("A2").font.size = 10

        # Batting table
        stats_sheet.range("A4").value = "TOP 15 BATSMEN BY RUNS"
        stats_sheet.range("A4").font.bold = True
        stats_sheet.range("A4").font.size = 12
        stats_sheet.range("A4").font.color = '#1F4E79'

        stats_sheet.range("A5").options(index=False).value = batting_df

        # Format header
        header_range = stats_sheet.range("A5").resize(1, len(batting_df.columns))
        header_range.color = '#4472C4'
        header_range.font.color = '#FFFFFF'
        header_range.font.bold = True

        # Bowling table (below batting)
        bowling_start_row = 5 + len(batting_df) + 4
        stats_sheet.range(f"A{bowling_start_row}").value = "TOP 15 BOWLERS BY WICKETS"
        stats_sheet.range(f"A{bowling_start_row}").font.bold = True
        stats_sheet.range(f"A{bowling_start_row}").font.size = 12
        stats_sheet.range(f"A{bowling_start_row}").font.color = '#2E7D32'

        stats_sheet.range(f"A{bowling_start_row + 1}").options(index=False).value = bowling_df

        # Format header
        header_range = stats_sheet.range(f"A{bowling_start_row + 1}").resize(1, len(bowling_df.columns))
        header_range.color = '#2E7D32'
        header_range.font.color = '#FFFFFF'
        header_range.font.bold = True

        print(f"   ODI sheet created")

    except Exception as e:
        print(f"   Stats sheet error: {e}")

    # Create charts
    print("\n   Creating charts...")
    try:
        batting_chart_path = create_batting_chart(batting_df, MATCH_TYPE, num_players=15)
        bowling_chart_path = create_bowling_chart(bowling_df, MATCH_TYPE, num_players=15)

        charts_sheet = ensure_sheet_exists(book, "ODI_CHARTS")

        charts_sheet.range("A1").value = f"ODI Player Statistics - Charts"
        charts_sheet.range("A1").font.bold = True
        charts_sheet.range("A1").font.size = 16

        charts_sheet.range("A2").value = f"Generated: {time.strftime('%Y-%m-%d %H:%M')}"

        # Add charts
        scale_factor = 0.5

        batting_img = Image.open(batting_chart_path)
        batting_resized_path = os.path.join(tempfile.gettempdir(), "batting_chart_resized.png")
        batting_img.resize(
            (int(batting_img.width * scale_factor), int(batting_img.height * scale_factor)),
            Image.Resampling.LANCZOS
        ).save(batting_resized_path)

        charts_sheet.pictures.add(
            batting_resized_path,
            name="BattingChart",
            update=True,
            anchor=charts_sheet.range("A4")
        )

        bowling_img = Image.open(bowling_chart_path)
        bowling_resized_path = os.path.join(tempfile.gettempdir(), "bowling_chart_resized.png")
        bowling_img.resize(
            (int(bowling_img.width * scale_factor), int(bowling_img.height * scale_factor)),
            Image.Resampling.LANCZOS
        ).save(bowling_resized_path)

        charts_sheet.pictures.add(
            bowling_resized_path,
            name="BowlingChart",
            update=True,
            anchor=charts_sheet.range("L4")
        )

        print(f"   Charts created")

    except Exception as e:
        print(f"   Charts error: {e}")

    stats_duration = time.time() - stats_start

    # Success - write to B8
    success_msg = f"Stats complete: ODI sheet + ODI_CHARTS ({len(batting_df)} batsmen, {len(bowling_df)} bowlers)"
    write_status(master_sheet, "B11", success_msg, is_error=False)

    print("\n" + "=" * 60)
    print("STATS COMPLETE!")
    print(f"   Batting: Top {len(batting_df)} batsmen")
    print(f"   Bowling: Top {len(bowling_df)} bowlers")
    print(f"   Sheets: ODI, ODI_CHARTS")
    print(f"   Time: {format_duration(stats_duration)}")
    print("\nNext: click '5. HTML Report' to create HTML report")
    print("=" * 60)


# =============================================================================
# PHASE 5: GENERATE HTML REPORT
# =============================================================================

def create_batting_chart_for_html(df: pd.DataFrame, match_type: str, output_path: str, num_players: int = 15):
    """Creates batting chart and saves directly to local mount."""
    chart_df = df.head(num_players).copy()
    players = [p[:12] + '..' if len(p) > 12 else p for p in chart_df['Player'].tolist()]
    x_pos = range(len(players))

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 7), sharex=True, gridspec_kw={'hspace': 0.1})

    # TOP: Runs
    runs = chart_df['Runs'].tolist()
    bars = ax1.bar(x_pos, runs, color='#4472C4', alpha=0.85, width=0.7)
    for bar, run in zip(bars, runs):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 50,
                f'{int(run):,}', ha='center', va='bottom', fontsize=8)
    ax1.set_title(f'Top {num_players} Batsmen - {match_type}', fontsize=14, fontweight='bold')
    ax1.set_ylabel('Runs', fontsize=11, fontweight='bold')
    ax1.grid(axis='y', alpha=0.3)

    # BOTTOM: SR & Avg
    sr = chart_df['SR'].tolist()
    avg = chart_df['Avg'].tolist()
    ax2.plot(x_pos, sr, 'o-', color='#FF6D00', linewidth=2, markersize=6, label='Strike Rate')
    ax2.plot(x_pos, avg, 's--', color='#1F4E79', linewidth=2, markersize=5, label='Average')
    ax2.set_ylabel('SR / Avg', fontsize=11, fontweight='bold')
    ax2.legend(loc='upper right', fontsize=10)
    ax2.grid(axis='y', alpha=0.3)
    ax2.set_xticks(x_pos)
    ax2.set_xticklabels(players, rotation=45, ha='right', fontsize=9)

    plt.tight_layout()
    fig.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)


def create_bowling_chart_for_html(df: pd.DataFrame, match_type: str, output_path: str, num_players: int = 15):
    """Creates bowling chart and saves directly to local mount."""
    chart_df = df.head(num_players).copy()
    players = [p[:12] + '..' if len(p) > 12 else p for p in chart_df['Player'].tolist()]
    x_pos = range(len(players))

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 7), sharex=True, gridspec_kw={'hspace': 0.1})

    # TOP: Wickets
    wickets = chart_df['Wkts'].tolist()
    bars = ax1.bar(x_pos, wickets, color='#4472C4', alpha=0.85, width=0.7)
    for bar, wkt in zip(bars, wickets):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f'{int(wkt)}', ha='center', va='bottom', fontsize=8)
    ax1.set_title(f'Top {num_players} Bowlers - {match_type}', fontsize=14, fontweight='bold')
    ax1.set_ylabel('Wickets', fontsize=11, fontweight='bold')
    ax1.grid(axis='y', alpha=0.3)

    # BOTTOM: Econ & Avg
    econ = chart_df['Econ'].tolist()
    avg = chart_df['Avg'].tolist()
    ax2.plot(x_pos, econ, 'o-', color='#FF6D00', linewidth=2, markersize=6, label='Economy')
    ax2.plot(x_pos, avg, 's--', color='#1F4E79', linewidth=2, markersize=5, label='Average')
    ax2.set_ylabel('Econ / Avg', fontsize=11, fontweight='bold')
    ax2.legend(loc='upper right', fontsize=10)
    ax2.grid(axis='y', alpha=0.3)
    ax2.set_xticks(x_pos)
    ax2.set_xticklabels(players, rotation=45, ha='right', fontsize=9)

    plt.tight_layout()
    fig.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)


def get_html_table_styles() -> str:
    """Returns CSS styles for HTML tables."""
    return """
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #1a1a1a;
            border-bottom: 3px solid #4472C4;
            padding-bottom: 10px;
        }
        h2 {
            color: #1F4E79;
            margin-top: 30px;
        }
        h2.bowling {
            color: #2E7D32;
        }
        .meta {
            color: #666;
            font-size: 14px;
            margin-bottom: 20px;
        }
        .chart-container {
            margin: 20px 0;
            text-align: center;
        }
        .chart-container img {
            max-width: 100%;
            height: auto;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 15px 0;
            font-size: 14px;
        }
        th {
            background-color: #4472C4;
            color: white;
            padding: 12px 8px;
            text-align: left;
            font-weight: 600;
        }
        table.bowling th {
            background-color: #2E7D32;
        }
        td {
            padding: 10px 8px;
            border-bottom: 1px solid #ddd;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f0f0f0;
        }
        .number {
            text-align: right;
        }
    </style>
    """


@script(name="5. HTML Report", button="[btn_s5]MASTER!D12")
async def s5_generate_html(book: xw.Book):
    """
    Generate HTML report with stats tables and charts.
    Saves PNG charts and HTML file to local mount.
    Status output: MASTER!B12
    """
    print("=" * 60)
    print("Phase 5: Generate HTML Report")
    print("=" * 60)

    # Get sheet for status output
    master_sheet = get_master_sheet(book)
    if not master_sheet:
        print("ERROR: MASTER sheet not found!")
        return

    # Clear previous status
    master_sheet.range("B12").value = ""

    # Print mount info
    print_mount_info()

    # Check mount exists
    if not os.path.exists(LOCAL_MOUNT_PATH):
        error_msg = get_mount_error_message()
        print(f"\n{error_msg}")
        write_status(master_sheet, "B12", error_msg, is_error=True)
        return

    # Check parquet exists in browser memory
    parquet_path = get_parquet_temp_path()
    if not os.path.exists(parquet_path):
        error_msg = "ERROR: Parquet not in memory. Click '3. Import to Parquet' first"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B12", error_msg, is_error=True)
        return

    print(f"\n   Parquet: browser memory")

    # Generate stats
    print("\n   Generating statistics...")
    report_start = time.time()

    try:
        conn = duckdb.connect()
        conn.execute(f"CREATE VIEW data AS SELECT * FROM '{parquet_path}'")

        row_count = scalar(conn, "SELECT COUNT(*) FROM data")
        print(f"   Total rows: {row_count:,}")

        # Get batting stats
        batting_query = get_batting_stats_query(MATCH_TYPE, limit=15)
        batting_df = conn.execute(batting_query).fetchdf()
        print(f"   Batting: {len(batting_df)} players")

        # Get bowling stats
        bowling_query = get_bowling_stats_query(MATCH_TYPE, limit=15)
        bowling_df = conn.execute(bowling_query).fetchdf()
        print(f"   Bowling: {len(bowling_df)} players")

        conn.close()

    except Exception as e:
        error_msg = f"ERROR: Query failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B12", error_msg, is_error=True)
        return

    # Save charts to local mount
    print("\n   Creating charts...")
    batting_chart_path = os.path.join(LOCAL_MOUNT_PATH, BATTING_CHART_FILENAME)
    bowling_chart_path = os.path.join(LOCAL_MOUNT_PATH, BOWLING_CHART_FILENAME)

    try:
        create_batting_chart_for_html(batting_df, MATCH_TYPE, batting_chart_path)
        print(f"   Saved: {batting_chart_path}")

        create_bowling_chart_for_html(bowling_df, MATCH_TYPE, bowling_chart_path)
        print(f"   Saved: {bowling_chart_path}")

    except Exception as e:
        error_msg = f"ERROR: Chart creation failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B12", error_msg, is_error=True)
        return

    # Generate HTML using pandas to_html()
    print("\n   Building HTML report...")
    try:
        # Convert DataFrames to HTML tables
        batting_html = batting_df.to_html(
            index=False,
            classes='batting',
            border=0,
            justify='left'
        )

        bowling_html = bowling_df.to_html(
            index=False,
            classes='bowling',
            border=0,
            justify='left'
        )

        # Build complete HTML document
        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{MATCH_TYPE} Cricket Statistics Report</title>
    {get_html_table_styles()}
</head>
<body>
    <div class="container">
        <h1>{MATCH_TYPE} Cricket Statistics Report</h1>
        <p class="meta">Generated: {time.strftime('%Y-%m-%d %H:%M')} | Total Records: {row_count:,}</p>

        <h2>Top 15 Batsmen by Runs</h2>
        <div class="chart-container">
            <img src="{BATTING_CHART_FILENAME}" alt="Batting Statistics Chart">
        </div>
        {batting_html}

        <h2 class="bowling">Top 15 Bowlers by Wickets</h2>
        <div class="chart-container">
            <img src="{BOWLING_CHART_FILENAME}" alt="Bowling Statistics Chart">
        </div>
        {bowling_html}
    </div>
</body>
</html>
"""

        # Write HTML to local mount
        html_path = os.path.join(LOCAL_MOUNT_PATH, OUTPUT_HTML_FILENAME)
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        # Verify
        if os.path.exists(html_path):
            html_size = os.path.getsize(html_path)
            print(f"   Saved: {html_path} ({format_size(html_size)})")
        else:
            error_msg = "ERROR: HTML file not found after write"
            print(f"\n{error_msg}")
            write_status(master_sheet, "B12", error_msg, is_error=True)
            return

    except Exception as e:
        error_msg = f"ERROR: HTML generation failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B12", error_msg, is_error=True)
        return

    report_duration = time.time() - report_start

    # Success - write to B9
    success_msg = f"HTML report created: {html_path} ({format_size(html_size)})"
    write_status(master_sheet, "B12", success_msg, is_error=False)

    print("\n" + "=" * 60)
    print("HTML REPORT COMPLETE!")
    print(f"   HTML: {html_path}")
    print(f"   Batting chart: {batting_chart_path}")
    print(f"   Bowling chart: {bowling_chart_path}")
    print(f"   Time: {format_duration(report_duration)}")
    print("\nNext: click '6. PDF Report' to create PDF report")
    print("=" * 60)


# =============================================================================
# PHASE 6: GENERATE PDF REPORT WITH REPORTLAB
# =============================================================================

async def ensure_reportlab_installed():
    """Install reportlab via micropip if not already available."""
    try:
        import reportlab
        return True
    except ImportError:
        print("   Installing reportlab via micropip...")
        try:
            import micropip  # ty: ignore[unresolved-import]
            await micropip.install("reportlab")
            print("   ReportLab installed successfully")
            return True
        except Exception as e:
            print(f"   Failed to install reportlab: {e}")
            return False


def create_pdf_report(batting_df: pd.DataFrame, bowling_df: pd.DataFrame,
                      batting_chart_path: str, bowling_chart_path: str,
                      output_path: str, match_type: str, row_count: int):
    """Create PDF report using ReportLab."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
    from reportlab.lib.enums import TA_CENTER, TA_LEFT

    # Use landscape A4 for wider tables
    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(A4),
        rightMargin=1*cm,
        leftMargin=1*cm,
        topMargin=1*cm,
        bottomMargin=1*cm
    )

    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=20,
        spaceAfter=12,
        textColor=colors.HexColor('#1a1a1a'),
        alignment=TA_CENTER
    )
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        spaceBefore=20,
        spaceAfter=10,
        textColor=colors.HexColor('#1F4E79')
    )
    bowling_heading_style = ParagraphStyle(
        'BowlingHeading',
        parent=styles['Heading2'],
        fontSize=14,
        spaceBefore=20,
        spaceAfter=10,
        textColor=colors.HexColor('#2E7D32')
    )
    meta_style = ParagraphStyle(
        'Meta',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#666666'),
        alignment=TA_CENTER,
        spaceAfter=20
    )

    # Build document elements
    # Layout: Page 1 = Batting Chart, Page 2 = Batting Table, Page 3 = Bowling Chart, Page 4 = Bowling Table
    elements = []

    # === PAGE 1: Title + Batting Chart ===
    elements.append(Paragraph(f"{match_type} Cricket Statistics Report", title_style))
    elements.append(Paragraph(
        f"Generated: {time.strftime('%Y-%m-%d %H:%M')} | Total Records: {row_count:,}",
        meta_style
    ))
    elements.append(Paragraph("Top 15 Batsmen by Runs", heading_style))

    # Batting chart (if exists)
    if os.path.exists(batting_chart_path):
        img = Image(batting_chart_path, width=24*cm, height=12*cm)
        elements.append(img)

    # === PAGE 2: Batting Table ===
    elements.append(PageBreak())
    elements.append(Paragraph("Top 15 Batsmen by Runs - Data", heading_style))

    batting_data = [batting_df.columns.tolist()] + batting_df.values.tolist()
    batting_table = Table(batting_data, repeatRows=1)
    batting_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),  # Player names left-aligned
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
        ('TOPPADDING', (0, 1), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')]),
    ]))
    elements.append(batting_table)

    # === PAGE 3: Bowling Chart ===
    elements.append(PageBreak())
    elements.append(Paragraph("Top 15 Bowlers by Wickets", bowling_heading_style))

    # Bowling chart (if exists)
    if os.path.exists(bowling_chart_path):
        img = Image(bowling_chart_path, width=24*cm, height=12*cm)
        elements.append(img)

    # === PAGE 4: Bowling Table ===
    elements.append(PageBreak())
    elements.append(Paragraph("Top 15 Bowlers by Wickets - Data", bowling_heading_style))

    bowling_data = [bowling_df.columns.tolist()] + bowling_df.values.tolist()
    bowling_table = Table(bowling_data, repeatRows=1)
    bowling_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E7D32')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),  # Player names left-aligned
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
        ('TOPPADDING', (0, 1), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')]),
    ]))
    elements.append(bowling_table)

    # Build PDF
    doc.build(elements)


@script(name="6. PDF Report", button="[btn_s6]MASTER!D13")
async def s6_generate_pdf(book: xw.Book):
    """
    Generate PDF report with ReportLab.
    Uses chart PNGs from local mount and creates tables from data.
    Status output: MASTER!B13
    """
    print("=" * 60)
    print("Phase 6: Generate PDF Report")
    print("=" * 60)

    # Get sheet for status output
    master_sheet = get_master_sheet(book)
    if not master_sheet:
        print("ERROR: MASTER sheet not found!")
        return

    # Clear previous status
    master_sheet.range("B13").value = ""

    # Print mount info
    print_mount_info()

    # Check mount exists
    if not os.path.exists(LOCAL_MOUNT_PATH):
        error_msg = get_mount_error_message()
        print(f"\n{error_msg}")
        write_status(master_sheet, "B13", error_msg, is_error=True)
        return

    # Check parquet exists in browser memory
    parquet_path = get_parquet_temp_path()
    if not os.path.exists(parquet_path):
        error_msg = "ERROR: Parquet not in memory. Click '3. Import to Parquet' first"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B13", error_msg, is_error=True)
        return

    print(f"\n   Parquet: browser memory")

    # Ensure reportlab is installed
    print("\n   Checking ReportLab...")
    if not await ensure_reportlab_installed():
        error_msg = "ERROR: Failed to install ReportLab"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B13", error_msg, is_error=True)
        return

    # Generate stats
    print("\n   Generating statistics...")
    pdf_start = time.time()

    try:
        conn = duckdb.connect()
        conn.execute(f"CREATE VIEW data AS SELECT * FROM '{parquet_path}'")

        row_count = scalar(conn, "SELECT COUNT(*) FROM data")
        print(f"   Total rows: {row_count:,}")

        # Get batting stats
        batting_query = get_batting_stats_query(MATCH_TYPE, limit=15)
        batting_df = conn.execute(batting_query).fetchdf()
        print(f"   Batting: {len(batting_df)} players")

        # Get bowling stats
        bowling_query = get_bowling_stats_query(MATCH_TYPE, limit=15)
        bowling_df = conn.execute(bowling_query).fetchdf()
        print(f"   Bowling: {len(bowling_df)} players")

        conn.close()

    except Exception as e:
        error_msg = f"ERROR: Query failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B13", error_msg, is_error=True)
        return

    # Check if chart PNGs exist (from s5_generate_html)
    batting_chart_path = os.path.join(LOCAL_MOUNT_PATH, BATTING_CHART_FILENAME)
    bowling_chart_path = os.path.join(LOCAL_MOUNT_PATH, BOWLING_CHART_FILENAME)

    if not os.path.exists(batting_chart_path) or not os.path.exists(bowling_chart_path):
        print("\n   Chart PNGs not found, creating them...")
        try:
            create_batting_chart_for_html(batting_df, MATCH_TYPE, batting_chart_path)
            print(f"   Created: {batting_chart_path}")
            create_bowling_chart_for_html(bowling_df, MATCH_TYPE, bowling_chart_path)
            print(f"   Created: {bowling_chart_path}")
        except Exception as e:
            print(f"   Warning: Could not create charts - {e}")
            # Continue without charts

    # Create PDF
    print("\n   Building PDF report...")
    pdf_path = os.path.join(LOCAL_MOUNT_PATH, OUTPUT_PDF_FILENAME)

    try:
        create_pdf_report(
            batting_df=batting_df,
            bowling_df=bowling_df,
            batting_chart_path=batting_chart_path,
            bowling_chart_path=bowling_chart_path,
            output_path=pdf_path,
            match_type=MATCH_TYPE,
            row_count=row_count
        )

        # Verify
        if os.path.exists(pdf_path):
            pdf_size = os.path.getsize(pdf_path)
            print(f"   Saved: {pdf_path} ({format_size(pdf_size)})")
        else:
            error_msg = "ERROR: PDF file not found after creation"
            print(f"\n{error_msg}")
            write_status(master_sheet, "B13", error_msg, is_error=True)
            return

    except Exception as e:
        error_msg = f"ERROR: PDF creation failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B13", error_msg, is_error=True)
        return

    pdf_duration = time.time() - pdf_start

    # Success - write to B10
    success_msg = f"PDF report created: {pdf_path} ({format_size(pdf_size)})"
    write_status(master_sheet, "B13", success_msg, is_error=False)

    print("\n" + "=" * 60)
    print("PDF REPORT COMPLETE!")
    print(f"   PDF: {pdf_path}")
    print(f"   Size: {format_size(pdf_size)}")
    print(f"   Time: {format_duration(pdf_duration)}")
    print("\nNext: click '7. Save Parquet' to save parquet to local drive")
    print("=" * 60)


# =============================================================================
# PHASE 7: SAVE PARQUET TO LOCAL DRIVE
# =============================================================================

@script(name="7. Save Parquet", button="[btn_s7]MASTER!D14")
async def s7_save_parquet(book: xw.Book):
    """
    Save parquet file from browser memory to local drive.
    Status output: MASTER!B14
    """
    print("=" * 60)
    print("Phase 7: Save Parquet to Local Drive")
    print("=" * 60)

    # Get sheet for status output
    master_sheet = get_master_sheet(book)
    if not master_sheet:
        print("ERROR: MASTER sheet not found!")
        return

    # Clear previous status
    master_sheet.range("B14").value = ""

    # Print mount info
    print_mount_info()

    # Check mount exists
    if not os.path.exists(LOCAL_MOUNT_PATH):
        error_msg = get_mount_error_message()
        print(f"\n{error_msg}")
        write_status(master_sheet, "B14", error_msg, is_error=True)
        return

    # Check parquet exists in browser memory
    parquet_temp_path = get_parquet_temp_path()
    if not os.path.exists(parquet_temp_path):
        error_msg = "ERROR: Parquet not in memory. Click '3. Import to Parquet' first"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B14", error_msg, is_error=True)
        return

    temp_size = os.path.getsize(parquet_temp_path)
    print(f"\n   Source: browser memory (temp)")
    print(f"   Size: {format_size(temp_size)}")

    # Copy parquet to local mount
    print("\n   Copying parquet to local drive...")
    save_start = time.time()

    local_parquet_path = os.path.join(LOCAL_MOUNT_PATH, OUTPUT_PARQUET_FILENAME)

    try:
        # Read from temp and write to local mount
        with open(parquet_temp_path, 'rb') as src:
            parquet_data = src.read()

        with open(local_parquet_path, 'wb') as dst:
            dst.write(parquet_data)

        # Verify
        if os.path.exists(local_parquet_path):
            saved_size = os.path.getsize(local_parquet_path)
            print(f"   Saved: {local_parquet_path} ({format_size(saved_size)})")
        else:
            error_msg = "ERROR: Parquet file not found after save"
            print(f"\n{error_msg}")
            write_status(master_sheet, "B14", error_msg, is_error=True)
            return

    except Exception as e:
        error_msg = f"ERROR: Save failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B14", error_msg, is_error=True)
        return

    save_duration = time.time() - save_start

    # Success - write to B11
    success_msg = f"Parquet saved: {local_parquet_path} ({format_size(saved_size)})"
    write_status(master_sheet, "B14", success_msg, is_error=False)

    print("\n" + "=" * 60)
    print("PARQUET SAVE COMPLETE!")
    print(f"   Path: {local_parquet_path}")
    print(f"   Size: {format_size(saved_size)}")
    print(f"   Time: {format_duration(save_duration)}")
    print("\nNext: click '8. Send Email' to send report via email")
    print("=" * 60)


# =============================================================================
# PHASE 8: SEND EMAIL WITH ATTACHMENTS VIA BREVO
# =============================================================================

# Brevo API configuration
# IMPORTANT: Store your API key securely - read from Excel cell or environment
BREVO_API_URL = "https://api.brevo.com/v3/smtp/email"
SENDER_EMAIL = ""  # read from MASTER!B19 (a sender verified in your Brevo account)
SENDER_NAME = ""   # read from MASTER!B20


async def send_email_brevo(
    api_key: str,
    to_email: str,
    subject: str,
    html_content: str,
    attachments: Optional[list] = None
) -> dict:
    """
    Send email via Brevo API.

    Args:
        api_key: Brevo API key
        to_email: Recipient email address
        subject: Email subject
        html_content: HTML body content
        attachments: List of dicts with 'name' and 'content' (base64 encoded)

    Returns:
        dict with 'success' and 'message_id' or 'error'
    """
    email_data = {
        "sender": {"email": SENDER_EMAIL, "name": SENDER_NAME},
        "to": [{"email": to_email}],
        "subject": subject,
        "htmlContent": html_content
    }

    if attachments:
        email_data["attachment"] = attachments

    try:
        response = await pyfetch(
            BREVO_API_URL,
            method="POST",
            headers={
                "Content-Type": "application/json",
                "api-key": api_key
            },
            body=json.dumps(email_data)
        )

        if response.ok:
            result = await response.json()
            return {"success": True, "message_id": result.get("messageId", "sent")}
        else:
            error_text = await response.text()
            return {"success": False, "error": f"HTTP {response.status}: {error_text[:100]}"}

    except Exception as e:
        return {"success": False, "error": str(e)[:100]}


def get_email_html_template(match_type: str) -> str:
    """Generate HTML email body."""
    return f"""
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #000000; line-height: 1.5;">
    Hi,<br><br>
    Please find attached the {match_type} Cricket Statistics Report.<br><br>
    <b>Attachments:</b><br>
    - PDF Report (odi_report.pdf)<br>
    - HTML Report (odi_report.html)<br><br>
    Best regards,<br>
    {SENDER_NAME}<br><br>
    ---<br>
    <span style="color: #666; font-size: 12px;">Generated by xlwings Lite - Local File Patterns Prototype</span>
    </div>
    """


@script(name="8. Send Email", button="[btn_s8]MASTER!D15")
async def s8_send_email(book: xw.Book):
    """
    Send email with PDF and HTML report as attachments via Brevo.
    Reads recipient email from MASTER!A13.
    Reads Brevo API key from BREVO_API_KEY environment variable.
    Status output: MASTER!B15
    """
    print("=" * 60)
    print("Phase 8: Send Email via Brevo")
    print("=" * 60)

    # Get sheet for status output
    master_sheet = get_master_sheet(book)
    if not master_sheet:
        print("ERROR: MASTER sheet not found!")
        return

    # Clear previous status
    master_sheet.range("B15").value = ""

    # Read recipient email from A13
    to_email = master_sheet.range("B18").value
    if not to_email or '@' not in str(to_email):
        error_msg = "ERROR: No valid email in B18. Enter the recipient email address."
        print(f"\n{error_msg}")
        write_status(master_sheet, "B15", error_msg, is_error=True)
        return

    to_email = str(to_email).strip()

    global SENDER_EMAIL, SENDER_NAME
    SENDER_EMAIL = str(master_sheet.range("B19").value or "").strip()
    SENDER_NAME = str(master_sheet.range("B20").value or "").strip() or SENDER_EMAIL
    if '@' not in SENDER_EMAIL:
        error_msg = "ERROR: Enter the sender email in B19 (a sender verified in your Brevo account)."
        print(f"\n{error_msg}")
        write_status(master_sheet, "B15", error_msg, is_error=True)
        return
    print(f"\n   Recipient: {to_email}")

    # Read Brevo API key from environment variable
    api_key = os.environ.get("BREVO_API_KEY")
    if not api_key:
        error_msg = "ERROR: BREVO_API_KEY environment variable not set."
        print(f"\n{error_msg}")
        print("   Set it in xlwings Lite: Settings > Environment Variables > BREVO_API_KEY")
        write_status(master_sheet, "B15", error_msg, is_error=True)
        return

    print(f"   API Key: {api_key[:10]}...{api_key[-4:]}")

    # Check mount exists
    if not os.path.exists(LOCAL_MOUNT_PATH):
        error_msg = get_mount_error_message()
        print(f"\n{error_msg}")
        write_status(master_sheet, "B15", error_msg, is_error=True)
        return

    # Check PDF exists
    pdf_path = os.path.join(LOCAL_MOUNT_PATH, OUTPUT_PDF_FILENAME)
    if not os.path.exists(pdf_path):
        error_msg = f"ERROR: PDF not found: {pdf_path}. Click '6. PDF Report' first."
        print(f"\n{error_msg}")
        write_status(master_sheet, "B15", error_msg, is_error=True)
        return

    # Check HTML exists
    html_path = os.path.join(LOCAL_MOUNT_PATH, OUTPUT_HTML_FILENAME)
    if not os.path.exists(html_path):
        error_msg = f"ERROR: HTML not found: {html_path}. Click '5. HTML Report' first."
        print(f"\n{error_msg}")
        write_status(master_sheet, "B15", error_msg, is_error=True)
        return

    # Prepare attachments
    print("\n   Preparing attachments...")
    email_start = time.time()

    try:
        # Read and encode PDF
        with open(pdf_path, 'rb') as f:
            pdf_content = base64.b64encode(f.read()).decode('utf-8')
        pdf_size = os.path.getsize(pdf_path)
        print(f"   PDF: {format_size(pdf_size)}")

        # Read and encode HTML
        with open(html_path, 'rb') as f:
            html_content = base64.b64encode(f.read()).decode('utf-8')
        html_size = os.path.getsize(html_path)
        print(f"   HTML: {format_size(html_size)}")

        attachments = [
            {"name": OUTPUT_PDF_FILENAME, "content": pdf_content},
            {"name": OUTPUT_HTML_FILENAME, "content": html_content}
        ]

    except Exception as e:
        error_msg = f"ERROR: Failed to read attachments - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B15", error_msg, is_error=True)
        return

    # Send email
    print("\n   Sending email via Brevo...")
    subject = f"{MATCH_TYPE} Cricket Statistics Report"
    body_html = get_email_html_template(MATCH_TYPE)

    result = await send_email_brevo(
        api_key=api_key,
        to_email=to_email,
        subject=subject,
        html_content=body_html,
        attachments=attachments
    )

    email_duration = time.time() - email_start

    if result["success"]:
        success_msg = f"Email sent to {to_email} (PDF + HTML attached)"
        write_status(master_sheet, "B15", success_msg, is_error=False)
        print(f"\n   SUCCESS: {success_msg}")
        print(f"   Message ID: {result.get('message_id', 'N/A')}")
    else:
        error_msg = f"ERROR: {result.get('error', 'Unknown error')}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B15", error_msg, is_error=True)
        return

    print("\n" + "=" * 60)
    print("EMAIL SENT!")
    print(f"   To: {to_email}")
    print(f"   Subject: {subject}")
    print(f"   Attachments: {OUTPUT_PDF_FILENAME}, {OUTPUT_HTML_FILENAME}")
    print(f"   Time: {format_duration(email_duration)}")
    print("\n   Next: s9_generate_pptx to create PowerPoint")
    print("=" * 60)


# =============================================================================
# PHASE 9: POWERPOINT GENERATION
# =============================================================================

async def ensure_pptx_installed():
    """Install python-pptx via micropip if not already available."""
    try:
        import pptx
        print("   python-pptx already installed")
    except ImportError:
        print("   Installing python-pptx via micropip...")
        import micropip  # ty: ignore[unresolved-import]
        await micropip.install("python-pptx")
        print("   python-pptx installed successfully")


@script(name="9. PowerPoint", button="[btn_s9]MASTER!D16")
async def s9_generate_pptx(book: xw.Book):
    """
    Generate PowerPoint presentation with charts and tables.
    Page 1: Top 15 Batsmen - chart on left, table on right
    Page 2: Top 15 Bowlers - table on left, chart on right
    Status output: MASTER!B16
    """
    print("=" * 60)
    print("Phase 9: Generate PowerPoint Presentation")
    print("=" * 60)

    # Get sheet for status output
    master_sheet = get_master_sheet(book)
    if not master_sheet:
        print("ERROR: MASTER sheet not found!")
        return

    # Clear previous status
    master_sheet.range("B16").value = ""

    # Check mount exists
    if not os.path.exists(LOCAL_MOUNT_PATH):
        error_msg = get_mount_error_message()
        print(f"\n{error_msg}")
        write_status(master_sheet, "B16", error_msg, is_error=True)
        return

    # Check charts exist
    batting_chart_path = os.path.join(LOCAL_MOUNT_PATH, BATTING_CHART_FILENAME)
    bowling_chart_path = os.path.join(LOCAL_MOUNT_PATH, BOWLING_CHART_FILENAME)

    if not os.path.exists(batting_chart_path):
        error_msg = f"ERROR: Batting chart not found. Click '4. Run Stats' first."
        print(f"\n{error_msg}")
        write_status(master_sheet, "B16", error_msg, is_error=True)
        return

    if not os.path.exists(bowling_chart_path):
        error_msg = f"ERROR: Bowling chart not found. Click '4. Run Stats' first."
        print(f"\n{error_msg}")
        write_status(master_sheet, "B16", error_msg, is_error=True)
        return

    # Check parquet exists in browser temp
    parquet_path = os.path.join(tempfile.gettempdir(), OUTPUT_PARQUET_FILENAME)
    if not os.path.exists(parquet_path):
        error_msg = f"ERROR: Parquet not found. Click '3. Import to Parquet' first."
        print(f"\n{error_msg}")
        write_status(master_sheet, "B16", error_msg, is_error=True)
        return

    # Install python-pptx if needed
    await ensure_pptx_installed()

    # Import pptx modules
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_LINE_DASH_STYLE

    print("\n   Loading cricket data...")
    pptx_start = time.time()

    # Connect to DuckDB and get data
    conn = duckdb.connect()
    conn.execute(f"CREATE VIEW cricket AS SELECT * FROM read_parquet('{parquet_path}')")

    # Get batting data (using correct column names from cricket data)
    batting_query = """
        SELECT striker as player, batting_team as country,
               CAST(SUM(TRY_CAST(runs_off_bat AS INTEGER)) AS INTEGER) as total_runs,
               COUNT(DISTINCT match_id) as matches,
               ROUND(CAST(SUM(TRY_CAST(runs_off_bat AS INTEGER)) AS DOUBLE) /
                     NULLIF(COUNT(DISTINCT match_id), 0), 1) as avg_runs
        FROM cricket
        WHERE runs_off_bat IS NOT NULL AND TRY_CAST(innings AS INTEGER) <= 2
        GROUP BY striker, batting_team
        HAVING COUNT(DISTINCT match_id) >= 5
        ORDER BY total_runs DESC
        LIMIT 15
    """
    batting_df = conn.execute(batting_query).fetchdf()

    # Get bowling data (using correct column names from cricket data)
    bowling_query = """
        SELECT bowler as player, bowling_team as country,
               SUM(CASE WHEN wicket_type IS NOT NULL
                        AND wicket_type NOT IN ('run out', 'retired hurt', 'retired not out', 'retired out',
                                                'obstructing the field', 'handled the ball', 'timed out')
                        THEN 1 ELSE 0 END) as total_wickets,
               COUNT(DISTINCT match_id) as matches,
               ROUND(SUM(CASE WHEN wicket_type IS NOT NULL
                              AND wicket_type NOT IN ('run out', 'retired hurt', 'retired not out', 'retired out',
                                                      'obstructing the field', 'handled the ball', 'timed out')
                              THEN 1 ELSE 0 END) * 1.0 /
                     NULLIF(COUNT(DISTINCT match_id), 0), 2) as avg_wickets
        FROM cricket
        WHERE TRY_CAST(innings AS INTEGER) <= 2
        GROUP BY bowler, bowling_team
        HAVING COUNT(DISTINCT match_id) >= 5
        ORDER BY total_wickets DESC
        LIMIT 15
    """
    bowling_df = conn.execute(bowling_query).fetchdf()
    conn.close()

    print(f"   Batting records: {len(batting_df)}")
    print(f"   Bowling records: {len(bowling_df)}")

    # Create presentation
    print("\n   Creating PowerPoint presentation...")
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9 widescreen
    prs.slide_height = Inches(7.5)

    # Define layout positions
    header_top = Inches(0.3)
    header_height = Inches(0.5)
    line_top = Inches(0.85)
    subtitle_top = Inches(1.0)
    content_top = Inches(1.5)
    content_height = Inches(5.5)

    # Colors
    NAVY_BLUE = RGBColor(0x1a, 0x2a, 0x56)  # Dark navy blue
    HEADER_BG = RGBColor(0x2c, 0x3e, 0x50)  # Dark blue-grey for table header
    ALT_ROW_BG = RGBColor(0xf8, 0xf9, 0xfa)  # Light grey for alternating rows
    WHITE = RGBColor(0xff, 0xff, 0xff)
    BLACK = RGBColor(0x00, 0x00, 0x00)

    # Helper function to style table
    def style_table(table, num_rows, num_cols):
        """Apply professional executive styling to table."""
        for row_idx in range(num_rows):
            for col_idx in range(num_cols):
                cell = table.cell(row_idx, col_idx)
                # Header row styling
                if row_idx == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = HEADER_BG
                    para = cell.text_frame.paragraphs[0]
                    para.font.color.rgb = WHITE
                    para.font.bold = True
                    para.font.size = Pt(10)
                else:
                    # Alternating row colors
                    cell.fill.solid()
                    if row_idx % 2 == 0:
                        cell.fill.fore_color.rgb = ALT_ROW_BG
                    else:
                        cell.fill.fore_color.rgb = WHITE
                    para = cell.text_frame.paragraphs[0]
                    para.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
                    para.font.size = Pt(9)

    # Slide 1: Batting - Chart Left, Table Right
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank slide

    # Header - Left aligned, dark navy blue
    header1 = slide1.shapes.add_textbox(Inches(0.5), header_top, Inches(12), header_height)
    tf1 = header1.text_frame
    p1 = tf1.paragraphs[0]
    p1.text = f"{MATCH_TYPE} Cricket Statistics"
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.font.color.rgb = NAVY_BLUE
    p1.alignment = PP_ALIGN.LEFT

    # Horizontal line below header
    line1 = slide1.shapes.add_connector(1, Inches(0.5), line_top, Inches(12.8), line_top)
    line1.line.color.rgb = BLACK
    line1.line.width = Pt(1.5)

    # Subtitle - Left aligned
    sub1 = slide1.shapes.add_textbox(Inches(0.5), subtitle_top, Inches(12), Inches(0.4))
    tf_sub1 = sub1.text_frame
    p_sub1 = tf_sub1.paragraphs[0]
    p_sub1.text = "Top 15 Batsmen"
    p_sub1.font.size = Pt(18)
    p_sub1.font.bold = True
    p_sub1.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    p_sub1.alignment = PP_ALIGN.LEFT

    # Batting chart on left
    slide1.shapes.add_picture(batting_chart_path, Inches(0.3), content_top, width=Inches(6.3))
    print("   Added batting chart to slide 1")

    # Batting table on right
    rows = len(batting_df) + 1
    cols = 4
    table1 = slide1.shapes.add_table(rows, cols, Inches(6.8), content_top, Inches(6.2), Inches(5.0)).table

    # Set column widths
    table1.columns[0].width = Inches(2.0)  # Player
    table1.columns[1].width = Inches(1.3)  # Country
    table1.columns[2].width = Inches(1.3)  # Total Runs
    table1.columns[3].width = Inches(1.0)  # Average

    # Header row
    headers1 = ["Player", "Country", "Total Runs", "Runs / Match"]
    for i, header in enumerate(headers1):
        cell = table1.cell(0, i)
        cell.text = header

    # Data rows
    for idx, (_, row) in enumerate(batting_df.iterrows()):
        table1.cell(idx + 1, 0).text = str(row['player'])[:25]
        table1.cell(idx + 1, 1).text = str(row['country'])[:10]
        table1.cell(idx + 1, 2).text = f"{int(row['total_runs']):,}"
        table1.cell(idx + 1, 3).text = str(row['avg_runs'])

    # Apply professional styling
    style_table(table1, rows, cols)
    print("   Added batting table to slide 1")

    # Slide 2: Bowling - Table Left, Chart Right
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank slide

    # Header - Left aligned, dark navy blue
    header2 = slide2.shapes.add_textbox(Inches(0.5), header_top, Inches(12), header_height)
    tf2 = header2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = f"{MATCH_TYPE} Cricket Statistics"
    p2.font.size = Pt(32)
    p2.font.bold = True
    p2.font.color.rgb = NAVY_BLUE
    p2.alignment = PP_ALIGN.LEFT

    # Horizontal line below header
    line2 = slide2.shapes.add_connector(1, Inches(0.5), line_top, Inches(12.8), line_top)
    line2.line.color.rgb = BLACK
    line2.line.width = Pt(1.5)

    # Subtitle - Left aligned
    sub2 = slide2.shapes.add_textbox(Inches(0.5), subtitle_top, Inches(12), Inches(0.4))
    tf_sub2 = sub2.text_frame
    p_sub2 = tf_sub2.paragraphs[0]
    p_sub2.text = "Top 15 Bowlers"
    p_sub2.font.size = Pt(18)
    p_sub2.font.bold = True
    p_sub2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    p_sub2.alignment = PP_ALIGN.LEFT

    # Bowling table on left
    rows2 = len(bowling_df) + 1
    cols2 = 4
    table2 = slide2.shapes.add_table(rows2, cols2, Inches(0.3), content_top, Inches(6.2), Inches(5.0)).table

    # Set column widths
    table2.columns[0].width = Inches(2.0)  # Player
    table2.columns[1].width = Inches(1.3)  # Country
    table2.columns[2].width = Inches(1.5)  # Total Wickets
    table2.columns[3].width = Inches(1.0)  # Average

    # Header row
    headers2 = ["Player", "Country", "Wickets", "Wkts / Match"]
    for i, header in enumerate(headers2):
        cell = table2.cell(0, i)
        cell.text = header

    # Data rows
    for idx, (_, row) in enumerate(bowling_df.iterrows()):
        table2.cell(idx + 1, 0).text = str(row['player'])[:25]
        table2.cell(idx + 1, 1).text = str(row['country'])[:10]
        table2.cell(idx + 1, 2).text = f"{int(row['total_wickets']):,}"
        table2.cell(idx + 1, 3).text = str(row['avg_wickets'])

    # Apply professional styling
    style_table(table2, rows2, cols2)

    print("   Added bowling table to slide 2")

    # Bowling chart on right
    slide2.shapes.add_picture(bowling_chart_path, Inches(6.8), content_top, width=Inches(6.3))
    print("   Added bowling chart to slide 2")

    # Save presentation
    output_path = os.path.join(LOCAL_MOUNT_PATH, OUTPUT_PPTX_FILENAME)
    print(f"\n   Saving to: {output_path}")

    try:
        prs.save(output_path)

        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            pptx_duration = time.time() - pptx_start
            success_msg = f"PowerPoint created: {OUTPUT_PPTX_FILENAME} ({format_size(file_size)})"
            write_status(master_sheet, "B16", success_msg, is_error=False)
            print(f"\n   SUCCESS: {success_msg}")
        else:
            error_msg = "ERROR: File not found after save"
            print(f"\n{error_msg}")
            write_status(master_sheet, "B16", error_msg, is_error=True)
            return

    except Exception as e:
        error_msg = f"ERROR: Save failed - {str(e)[:50]}"
        print(f"\n{error_msg}")
        write_status(master_sheet, "B16", error_msg, is_error=True)
        return

    print("\n" + "=" * 60)
    print("POWERPOINT CREATED!")
    print(f"   File: {output_path}")
    print(f"   Slides: 2")
    print(f"   - Slide 1: Top 15 Batsmen (chart + table)")
    print(f"   - Slide 2: Top 15 Bowlers (table + chart)")
    print(f"   Time: {format_duration(pptx_duration)}")
    print("\n   All phases complete!")
    print("=" * 60)
