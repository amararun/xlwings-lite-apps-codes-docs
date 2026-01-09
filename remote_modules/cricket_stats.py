# Cricket Statistics Module - Remote Loadable
#
# This module is loaded remotely by main-duckdb-remote.py from GitHub.
# Contains all Cricket player stats generation logic (tables + charts).
#
# Usage: Fetched at runtime, cached in memory for the session.

import xlwings as xw
import pandas as pd
import duckdb
import tempfile
import os
import time
import matplotlib.pyplot as plt
from PIL import Image


def get_duckdb_temp_path() -> str:
    """Returns the standard temp path for downloaded DuckDB files."""
    return os.path.join(tempfile.gettempdir(), "imported_database.duckdb")


def get_parquet_temp_path() -> str:
    """Returns the standard temp path for downloaded parquet files."""
    return os.path.join(tempfile.gettempdir(), "imported_data.parquet")


def ensure_sheet_exists(book: xw.Book, sheet_name: str) -> xw.Sheet:
    """Ensures a sheet exists, deleting any existing one first for re-runnability."""
    for s in book.sheets:
        if s.name == sheet_name:
            s.delete()
            break
    return book.sheets.add(name=sheet_name)


def get_batting_stats_query(table_name: str, match_type: str, limit: int = 30) -> str:
    """Returns SQL query for top batsmen by runs for a specific match type."""
    return f"""
    SELECT
        striker as Player,
        COUNT(DISTINCT match_id) as Mat,
        COUNT(DISTINCT match_id) as Inns,
        CAST(SUM(TRY_CAST(runs_off_bat AS INTEGER)) AS INTEGER) as Runs,
        COUNT(*) as BF,
        SUM(CASE
            WHEN wicket_type IS NOT NULL
                AND player_dismissed = striker
                AND wicket_type NOT IN ('retired hurt', 'retired not out')
            THEN 1 ELSE 0
        END) as Outs,
        COUNT(DISTINCT match_id) - SUM(CASE
            WHEN wicket_type IS NOT NULL
                AND player_dismissed = striker
                AND wicket_type NOT IN ('retired hurt', 'retired not out')
            THEN 1 ELSE 0
        END) as NO,
        CASE
            WHEN SUM(CASE WHEN wicket_type IS NOT NULL AND player_dismissed = striker
                          AND wicket_type NOT IN ('retired hurt', 'retired not out')
                          THEN 1 ELSE 0 END) > 0
            THEN ROUND(CAST(SUM(TRY_CAST(runs_off_bat AS INTEGER)) AS DOUBLE) /
                       SUM(CASE WHEN wicket_type IS NOT NULL AND player_dismissed = striker
                                AND wicket_type NOT IN ('retired hurt', 'retired not out')
                                THEN 1 ELSE 0 END), 2)
            ELSE CAST(SUM(TRY_CAST(runs_off_bat AS INTEGER)) AS DOUBLE)
        END as Avg,
        ROUND(CAST(SUM(TRY_CAST(runs_off_bat AS INTEGER)) AS DOUBLE) / NULLIF(COUNT(*), 0) * 100, 2) as SR,
        SUM(CASE WHEN TRY_CAST(runs_off_bat AS INTEGER) = 4 THEN 1 ELSE 0 END) as "4s",
        SUM(CASE WHEN TRY_CAST(runs_off_bat AS INTEGER) = 6 THEN 1 ELSE 0 END) as "6s",
        CASE
            WHEN SUM(TRY_CAST(runs_off_bat AS INTEGER)) > 0
            THEN ROUND((SUM(CASE WHEN TRY_CAST(runs_off_bat AS INTEGER) = 4 THEN 4 ELSE 0 END) +
                        SUM(CASE WHEN TRY_CAST(runs_off_bat AS INTEGER) = 6 THEN 6 ELSE 0 END)) * 100.0 /
                       SUM(TRY_CAST(runs_off_bat AS INTEGER)), 1)
            ELSE 0
        END as "Bnd%"
    FROM "{table_name}"
    WHERE match_type = '{match_type}'
    GROUP BY striker
    HAVING COUNT(DISTINCT match_id) >= 5
    ORDER BY Runs DESC
    LIMIT {limit}
    """


def get_bowling_stats_query(table_name: str, match_type: str, limit: int = 30) -> str:
    """Returns SQL query for top bowlers by wickets for a specific match type."""
    return f"""
    SELECT
        bowler as Player,
        COUNT(DISTINCT match_id) as Mat,
        ROUND(COUNT(*) / 6.0, 1) as Overs,
        CAST(SUM(COALESCE(TRY_CAST(runs_off_bat AS INTEGER), 0) +
                 COALESCE(TRY_CAST(wides AS INTEGER), 0) +
                 COALESCE(TRY_CAST(noballs AS INTEGER), 0)) AS INTEGER) as Runs,
        SUM(CASE
            WHEN wicket_type IS NOT NULL
                AND wicket_type NOT IN ('run out', 'retired hurt', 'retired not out',
                                        'retired out', 'obstructing the field')
            THEN 1 ELSE 0
        END) as Wkts,
        ROUND(CAST(SUM(COALESCE(TRY_CAST(runs_off_bat AS INTEGER), 0) +
                       COALESCE(TRY_CAST(wides AS INTEGER), 0) +
                       COALESCE(TRY_CAST(noballs AS INTEGER), 0)) AS DOUBLE) / NULLIF(COUNT(*) / 6.0, 0), 2) as Econ,
        CASE
            WHEN SUM(CASE WHEN wicket_type IS NOT NULL
                          AND wicket_type NOT IN ('run out', 'retired hurt', 'retired not out',
                                                  'retired out', 'obstructing the field')
                          THEN 1 ELSE 0 END) > 0
            THEN ROUND(CAST(SUM(COALESCE(TRY_CAST(runs_off_bat AS INTEGER), 0) +
                               COALESCE(TRY_CAST(wides AS INTEGER), 0) +
                               COALESCE(TRY_CAST(noballs AS INTEGER), 0)) AS DOUBLE) /
                       SUM(CASE WHEN wicket_type IS NOT NULL
                                AND wicket_type NOT IN ('run out', 'retired hurt', 'retired not out',
                                                        'retired out', 'obstructing the field')
                                THEN 1 ELSE 0 END), 2)
            ELSE 0
        END as Avg,
        CASE
            WHEN SUM(CASE WHEN wicket_type IS NOT NULL
                          AND wicket_type NOT IN ('run out', 'retired hurt', 'retired not out',
                                                  'retired out', 'obstructing the field')
                          THEN 1 ELSE 0 END) > 0
            THEN ROUND(CAST(COUNT(*) AS DOUBLE) /
                       SUM(CASE WHEN wicket_type IS NOT NULL
                                AND wicket_type NOT IN ('run out', 'retired hurt', 'retired not out',
                                                        'retired out', 'obstructing the field')
                                THEN 1 ELSE 0 END), 1)
            ELSE 0
        END as SR,
        ROUND(SUM(CASE WHEN COALESCE(TRY_CAST(runs_off_bat AS INTEGER), 0) = 0
                       AND COALESCE(TRY_CAST(wides AS INTEGER), 0) = 0
                       AND COALESCE(TRY_CAST(noballs AS INTEGER), 0) = 0 THEN 1 ELSE 0 END) *
              100.0 / NULLIF(COUNT(*), 0), 1) as "Dot%"
    FROM "{table_name}"
    WHERE match_type = '{match_type}'
    GROUP BY bowler
    HAVING COUNT(DISTINCT match_id) >= 5
    ORDER BY Wkts DESC
    LIMIT {limit}
    """


def write_stats_table(sheet: xw.Sheet, df: pd.DataFrame, start_row: int,
                      title: str, header_color: str = None, header_font_color: str = None) -> int:
    """Writes a DataFrame as a formatted Excel table. Returns the next available row."""
    sheet.range(f"A{start_row}").value = title
    sheet.range(f"A{start_row}").font.bold = True
    sheet.range(f"A{start_row}").font.size = 11
    sheet.range(f"A{start_row}").font.color = '#1F4E79'
    start_row += 1

    sheet.range(f"A{start_row}").options(index=False).value = df

    table_range = sheet.range(f"A{start_row}").resize(df.shape[0] + 1, df.shape[1])
    try:
        sheet.tables.add(source=table_range)
    except Exception as e:
        print(f"      Table creation note: {e}")

    header_range = sheet.range(f"A{start_row}").resize(1, df.shape[1])
    header_range.color = '#4472C4'
    header_range.font.color = '#FFFFFF'
    header_range.font.bold = True

    return start_row + df.shape[0] + 3


def create_batting_chart(df: pd.DataFrame, match_type: str, num_players: int = 15) -> str:
    """Creates a compound VERTICAL batting chart. Returns path to saved chart image."""
    chart_df = df.head(num_players).copy()
    players = [p[:12] + '..' if len(p) > 12 else p for p in chart_df['Player'].tolist()]
    x_pos = range(len(players))

    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(8, 6),
                                        height_ratios=[2, 1, 1],
                                        sharex=True,
                                        gridspec_kw={'hspace': 0})

    # TOP: Runs
    runs = chart_df['Runs'].tolist()
    bars = ax1.bar(x_pos, runs, color='#4472C4', alpha=0.85, label='Runs', width=0.7)
    for bar, run in zip(bars, runs):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 50,
                f'{int(run):,}', ha='center', va='bottom', fontsize=7, color='#333333')
    ax1.set_title(f'Top {num_players} Batsmen - {match_type}', fontsize=12, fontweight='bold', pad=10)
    ax1.set_ylabel('Runs', fontsize=10, fontweight='bold', color='#333333')
    ax1.legend(loc='upper right', fontsize=9)
    ax1.grid(axis='y', alpha=0.3)
    ax1.set_xticklabels([])

    # MIDDLE: Strike Rate & Average
    sr = chart_df['SR'].tolist()
    avg = chart_df['Avg'].tolist()
    ax2.plot(x_pos, sr, 'o-', color='#FF6D00', linewidth=2, markersize=6, label='Strike Rate')
    ax2.plot(x_pos, avg, 's--', color='#2E7D32', linewidth=2, markersize=5, label='Average')
    ax2.set_ylabel('SR / Avg', fontsize=10, fontweight='bold', color='#333333')
    ax2.legend(loc='upper right', fontsize=9)
    ax2.grid(axis='y', alpha=0.3)
    ax2.set_xticklabels([])

    # BOTTOM: 4s and 6s
    fours = chart_df['4s'].tolist()
    sixes = chart_df['6s'].tolist()
    ax3.bar(x_pos, fours, color='#26A69A', alpha=0.85, label='4s', width=0.7)
    ax3.bar(x_pos, sixes, bottom=fours, color='#EF5350', alpha=0.85, label='6s', width=0.7)
    ax3.set_ylabel('Boundaries', fontsize=10, fontweight='bold', color='#333333')
    ax3.legend(loc='upper right', fontsize=9)
    ax3.grid(axis='y', alpha=0.3)
    ax3.set_xticks(x_pos)
    ax3.set_xticklabels(players, rotation=45, ha='right', fontsize=8)

    plt.tight_layout()
    temp_dir = tempfile.gettempdir()
    chart_path = os.path.join(temp_dir, f"batting_chart_{match_type.lower()}.png")
    fig.savefig(chart_path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return chart_path


def create_bowling_chart(df: pd.DataFrame, match_type: str, num_players: int = 15) -> str:
    """Creates a compound VERTICAL bowling chart. Returns path to saved chart image."""
    chart_df = df.head(num_players).copy()
    players = [p[:12] + '..' if len(p) > 12 else p for p in chart_df['Player'].tolist()]
    x_pos = range(len(players))

    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(8, 6),
                                        height_ratios=[2, 1, 1],
                                        sharex=True,
                                        gridspec_kw={'hspace': 0})

    # TOP: Wickets
    wickets = chart_df['Wkts'].tolist()
    bars = ax1.bar(x_pos, wickets, color='#2E7D32', alpha=0.85, label='Wickets', width=0.7)
    for bar, wkt in zip(bars, wickets):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f'{int(wkt)}', ha='center', va='bottom', fontsize=7, color='#333333')
    ax1.set_title(f'Top {num_players} Bowlers - {match_type}', fontsize=12, fontweight='bold', pad=10)
    ax1.set_ylabel('Wickets', fontsize=10, fontweight='bold', color='#333333')
    ax1.legend(loc='upper right', fontsize=9)
    ax1.grid(axis='y', alpha=0.3)
    ax1.set_xticklabels([])

    # MIDDLE: Economy & Average
    econ = chart_df['Econ'].tolist()
    avg = chart_df['Avg'].tolist()
    ax2.plot(x_pos, econ, 'o-', color='#FF6D00', linewidth=2, markersize=6, label='Economy')
    ax2.plot(x_pos, avg, 's--', color='#9C27B0', linewidth=2, markersize=5, label='Average')
    ax2.set_ylabel('Econ / Avg', fontsize=10, fontweight='bold', color='#333333')
    ax2.legend(loc='upper right', fontsize=9)
    ax2.grid(axis='y', alpha=0.3)
    ax2.set_xticklabels([])

    # BOTTOM: Dot Ball Percentage
    dot_pct = chart_df['Dot%'].tolist()
    colors = ['#26A69A' if d >= 50 else '#FF9800' if d >= 40 else '#EF5350' for d in dot_pct]
    bars3 = ax3.bar(x_pos, dot_pct, color=colors, alpha=0.85, width=0.7)
    for bar, pct in zip(bars3, dot_pct):
        ax3.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f'{pct:.0f}%', ha='center', va='bottom', fontsize=7, color='#333333')
    ax3.set_ylabel('Dot%', fontsize=10, fontweight='bold', color='#333333')
    ax3.axhline(y=50, color='#26A69A', linestyle='--', alpha=0.5, linewidth=1)
    ax3.set_ylim(0, max(dot_pct) * 1.15)
    ax3.grid(axis='y', alpha=0.3)
    ax3.set_xticks(x_pos)
    ax3.set_xticklabels(players, rotation=45, ha='right', fontsize=8)

    plt.tight_layout()
    temp_dir = tempfile.gettempdir()
    chart_path = os.path.join(temp_dir, f"bowling_chart_{match_type.lower()}.png")
    fig.savefig(chart_path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return chart_path


def run_cricket_stats(book: xw.Book, file_type: str, file_path: str):
    """
    Generate Cricket player statistics tables and charts.

    Args:
        book: xlwings Book object
        file_type: 'parquet' or 'duckdb'
        file_path: Path to the data file
    """
    print("=" * 60)
    print("CRICKET STATISTICS GENERATOR (Remote Module)")
    print("=" * 60)

    # -------------------------------------------------------------------------
    # PHASE 1: CONNECT TO DATA
    # -------------------------------------------------------------------------
    print("\n[1/5] Loading data into DuckDB...")

    conn = None
    tables_to_process = []

    try:
        if file_type == 'parquet':
            conn = duckdb.connect()
            view_name = "parquet_data"
            conn.execute(f"CREATE VIEW {view_name} AS SELECT * FROM '{file_path}'")
            print(f"   Created view: {view_name}")
            tables_to_process.append((view_name, "Parquet file"))

        elif file_type == 'duckdb':
            conn = duckdb.connect(file_path, read_only=True)
            print("   Connected to DuckDB database")

            tables_query = """
                SELECT table_name
                FROM information_schema.tables
                WHERE table_schema = 'main'
                ORDER BY table_name
            """
            tables_result = conn.execute(tables_query).fetchall()
            table_names = [row[0] for row in tables_result]
            print(f"   Found {len(table_names)} tables: {table_names}")

            for table_name in table_names:
                tables_to_process.append((table_name, f"DuckDB: {table_name}"))

    except Exception as e:
        print(f"ERROR: Failed to load data: {e}")
        if conn:
            conn.close()
        return

    # -------------------------------------------------------------------------
    # PHASE 2: VALIDATE TABLES
    # -------------------------------------------------------------------------
    print(f"\n[2/5] Validating {len(tables_to_process)} table(s)...")

    valid_tables = []
    required_columns = ['match_type', 'match_id', 'striker', 'bowler', 'runs_off_bat']

    for table_name, source_label in tables_to_process:
        try:
            describe_query = f'DESCRIBE "{table_name}"'
            describe_result = conn.execute(describe_query).fetchall()
            column_names = [row[0].lower() for row in describe_result]

            missing = [col for col in required_columns if col not in column_names]

            if missing:
                print(f"   SKIP: {table_name} - missing columns: {missing}")
            else:
                count_query = f'SELECT COUNT(*) FROM "{table_name}"'
                row_count = conn.execute(count_query).fetchone()[0]
                print(f"   OK: {table_name} ({row_count:,} rows)")
                valid_tables.append((table_name, source_label))

        except Exception as e:
            print(f"   SKIP: {table_name} - error: {e}")

    if not valid_tables:
        print("\nERROR: No valid tables found with required columns!")
        conn.close()
        return

    # -------------------------------------------------------------------------
    # PHASE 3: GENERATE STATS
    # -------------------------------------------------------------------------
    print("\n[3/5] Generating statistics...")

    temp_dir = tempfile.gettempdir()
    sheets_created = 0

    for table_name, source_label in valid_tables:
        print(f"\n   Processing: {source_label}")

        # Get match types
        match_types_query = f"""
            SELECT DISTINCT match_type, COUNT(*) as cnt
            FROM "{table_name}"
            GROUP BY match_type
            ORDER BY cnt DESC
        """
        match_types_result = conn.execute(match_types_query).fetchall()
        match_types = [row[0] for row in match_types_result]

        print(f"   Found {len(match_types)} match types")

        all_stats = {}
        for match_type in match_types:
            try:
                batting_query = get_batting_stats_query(table_name, match_type, limit=30)
                batting_df = conn.execute(batting_query).fetchdf()

                bowling_query = get_bowling_stats_query(table_name, match_type, limit=30)
                bowling_df = conn.execute(bowling_query).fetchdf()

                all_stats[match_type] = {'batting': batting_df, 'bowling': bowling_df}
                print(f"      {match_type}: {len(batting_df)} batsmen, {len(bowling_df)} bowlers")

            except Exception as e:
                print(f"      {match_type}: ERROR - {e}")
                all_stats[match_type] = {'batting': pd.DataFrame(), 'bowling': pd.DataFrame()}

        # -------------------------------------------------------------------------
        # PHASE 4: CREATE DOCS SHEET
        # -------------------------------------------------------------------------
        docs_exists = any(s.name == 'DOCS' for s in book.sheets)
        if not docs_exists:
            print("\n[4/5] Creating DOCS sheet...")
            try:
                docs_sheet = ensure_sheet_exists(book, 'DOCS')
                docs_sheet.range("A1").value = "Cricket Statistics - Formula Documentation"
                docs_sheet.range("A1").font.bold = True
                docs_sheet.range("A1").font.size = 16

                docs_sheet.range("A2").value = f"Source: {source_label} | Generated: {time.strftime('%Y-%m-%d %H:%M')} (Remote Module)"
                docs_sheet.range("A2").font.size = 10

                current_row = 4

                # Batting docs
                docs_sheet.range(f"A{current_row}").value = "BATTING METRICS"
                docs_sheet.range(f"A{current_row}").font.bold = True
                docs_sheet.range(f"A{current_row}").font.size = 14
                docs_sheet.range(f"A{current_row}").font.color = '#1F4E79'
                current_row += 2

                batting_docs = pd.DataFrame([
                    ["Mat", "Matches played as a batsman", "COUNT(DISTINCT match_id)"],
                    ["Runs", "Total runs scored off the bat", "SUM(runs_off_bat)"],
                    ["BF", "Balls faced", "COUNT(*) of all deliveries faced"],
                    ["Avg", "Batting average", "Runs / Outs"],
                    ["SR", "Strike rate", "(Runs / BF) * 100"],
                    ["4s", "Fours hit", "COUNT where runs_off_bat = 4"],
                    ["6s", "Sixes hit", "COUNT where runs_off_bat = 6"],
                ], columns=["Metric", "Description", "Formula"])
                docs_sheet.range(f"A{current_row}").options(index=False).value = batting_docs
                current_row += len(batting_docs) + 4

                # Bowling docs
                docs_sheet.range(f"A{current_row}").value = "BOWLING METRICS"
                docs_sheet.range(f"A{current_row}").font.bold = True
                docs_sheet.range(f"A{current_row}").font.size = 14
                docs_sheet.range(f"A{current_row}").font.color = '#2E7D32'
                current_row += 2

                bowling_docs = pd.DataFrame([
                    ["Mat", "Matches bowled in", "COUNT(DISTINCT match_id)"],
                    ["Overs", "Overs bowled", "Balls / 6"],
                    ["Wkts", "Wickets taken", "Excludes run outs, retired"],
                    ["Econ", "Economy rate", "Runs / Overs"],
                    ["Avg", "Bowling average", "Runs / Wickets"],
                    ["Dot%", "Dot ball percentage", "Dot balls / Total balls * 100"],
                ], columns=["Metric", "Description", "Formula"])
                docs_sheet.range(f"A{current_row}").options(index=False).value = bowling_docs

                sheets_created += 1
            except Exception as e:
                print(f"   DOCS error: {e}")

        # -------------------------------------------------------------------------
        # PHASE 5: WRITE STATS AND CHARTS
        # -------------------------------------------------------------------------
        print("\n[5/5] Writing results to Excel...")

        for match_type, stats in all_stats.items():
            sheet_name = match_type[:31]
            print(f"   Creating sheet: {sheet_name}")

            try:
                output_sheet = ensure_sheet_exists(book, sheet_name)

                output_sheet.range("A1").value = f"{match_type} Player Statistics"
                output_sheet.range("A1").font.bold = True
                output_sheet.range("A1").font.size = 16

                output_sheet.range("A2").value = f"Source: {source_label} | Generated: {time.strftime('%Y-%m-%d %H:%M')} (Remote Module)"
                output_sheet.range("A2").font.size = 12

                current_row = 5

                batting_df = stats['batting']
                if not batting_df.empty:
                    current_row = write_stats_table(output_sheet, batting_df, current_row,
                                                   f"TOP 30 BATSMEN BY RUNS - {match_type}")

                bowling_df = stats['bowling']
                if not bowling_df.empty:
                    current_row = write_stats_table(output_sheet, bowling_df, current_row,
                                                   f"TOP 30 BOWLERS BY WICKETS - {match_type}")

                sheets_created += 1

                # Create charts
                try:
                    batting_chart_path = create_batting_chart(batting_df, match_type, num_players=15)
                    bowling_chart_path = create_bowling_chart(bowling_df, match_type, num_players=15)

                    charts_sheet_name = f'{sheet_name}_CHARTS'[:31]
                    charts_sheet = ensure_sheet_exists(book, charts_sheet_name)

                    charts_sheet.range("A1").value = f"{match_type} Player Statistics - Charts"
                    charts_sheet.range("A1").font.bold = True
                    charts_sheet.range("A1").font.size = 16

                    charts_sheet.range("A2").value = f"Generated: {time.strftime('%Y-%m-%d %H:%M')} (Remote Module)"
                    charts_sheet.range("A2").font.size = 12

                    charts_sheet.range("A3").value = "Showcasing analytics & chart generation. Formulas under validation."
                    charts_sheet.range("A3").font.size = 12

                    scale_factor = 0.5

                    batting_img = Image.open(batting_chart_path)
                    batting_resized_path = os.path.join(temp_dir, f"batting_chart_{match_type.lower()}_resized.png")
                    batting_img.resize(
                        (int(batting_img.width * scale_factor), int(batting_img.height * scale_factor)),
                        Image.LANCZOS
                    ).save(batting_resized_path)

                    charts_sheet.pictures.add(
                        batting_resized_path,
                        name="BattingChart",
                        update=True,
                        anchor=charts_sheet.range("A5")
                    )

                    bowling_img = Image.open(bowling_chart_path)
                    bowling_resized_path = os.path.join(temp_dir, f"bowling_chart_{match_type.lower()}_resized.png")
                    bowling_img.resize(
                        (int(bowling_img.width * scale_factor), int(bowling_img.height * scale_factor)),
                        Image.LANCZOS
                    ).save(bowling_resized_path)

                    charts_sheet.pictures.add(
                        bowling_resized_path,
                        name="BowlingChart",
                        update=True,
                        anchor=charts_sheet.range("K5")
                    )

                    sheets_created += 1
                    print(f"      Charts created: {charts_sheet_name}")

                except Exception as chart_error:
                    print(f"      Charts error: {chart_error}")

            except Exception as e:
                print(f"   ERROR: {e}")

    conn.close()

    print("\n" + "=" * 60)
    print("CRICKET STATS COMPLETE! (Remote Module)")
    print(f"Sheets created: {sheets_created}")
    print("=" * 60)
