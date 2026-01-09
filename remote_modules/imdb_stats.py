# IMDB Statistics Module - Remote Loadable
#
# This module is loaded remotely by main-duckdb-remote.py from GitHub.
# Contains all IMDB stats generation logic (tables + charts).
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


def ensure_sheet_exists(book: xw.Book, sheet_name: str) -> xw.Sheet:
    """
    Ensures a sheet exists, deleting any existing one first for re-runnability.
    Returns the sheet object.
    """
    for s in book.sheets:
        if s.name == sheet_name:
            s.delete()
            break
    return book.sheets.add(name=sheet_name)


def write_stats_table(sheet: xw.Sheet, df: pd.DataFrame, start_row: int,
                      title: str, header_color: str = None, header_font_color: str = None) -> int:
    """
    Writes a DataFrame as a formatted Excel table.
    Returns the next available row after the table.
    """
    # Write title as plain bold text
    sheet.range(f"A{start_row}").value = title
    sheet.range(f"A{start_row}").font.bold = True
    sheet.range(f"A{start_row}").font.size = 11
    sheet.range(f"A{start_row}").font.color = '#1F4E79'
    start_row += 1

    # Write DataFrame
    sheet.range(f"A{start_row}").options(index=False).value = df

    # Create table with explicit range sizing
    table_range = sheet.range(f"A{start_row}").resize(
        df.shape[0] + 1,
        df.shape[1]
    )
    try:
        sheet.tables.add(source=table_range)
    except Exception as e:
        print(f"      Table creation note: {e}")

    # Return next row position (data rows + header + title + 2 blank rows)
    return start_row + df.shape[0] + 3


def run_imdb_stats(book: xw.Book):
    """
    Generate IMDB statistics tables and charts.

    Creates: IMDB_STATS sheet with all tables + IMDB_CHARTS sheet with 6 charts.
    Uses same patterns as player stats for formatting and chart creation.
    """
    print("=" * 60)
    print("IMDB STATISTICS GENERATOR (Remote Module)")
    print("=" * 60)

    # -------------------------------------------------------------------------
    # PHASE 1: CONNECT TO DUCKDB
    # -------------------------------------------------------------------------
    print("\n[1/4] Connecting to DuckDB...")

    temp_duckdb_path = get_duckdb_temp_path()

    if not os.path.exists(temp_duckdb_path):
        print("ERROR: No DuckDB file found. Run import_data first.")
        return

    conn = duckdb.connect(temp_duckdb_path, read_only=True)

    # Verify IMDB tables exist
    tables = [t[0] for t in conn.execute("SHOW TABLES").fetchall()]
    if "title_basics" not in tables:
        print("ERROR: Not IMDB data. title_basics table not found.")
        conn.close()
        return

    print(f"   Connected. Tables: {tables}")

    # -------------------------------------------------------------------------
    # PHASE 2: GENERATE ALL TABLES ON ONE SHEET
    # -------------------------------------------------------------------------
    print("\n[2/4] Generating summary tables...")

    stats_sheet = ensure_sheet_exists(book, "IMDB_STATS")
    stats_sheet.range("A1").value = "IMDB Database Statistics"
    stats_sheet.range("A1").font.bold = True
    stats_sheet.range("A1").font.size = 16
    stats_sheet.range("A1").font.color = '#1F4E79'

    stats_sheet.range("A2").value = f"Generated: {time.strftime('%Y-%m-%d %H:%M')} (Remote Module)"
    stats_sheet.range("A2").font.size = 10

    current_row = 4

    # --- Table 1: Overall Summary ---
    print("   Creating Database Summary...")
    summary_data = []

    total = conn.execute("SELECT COUNT(*) FROM title_basics").fetchone()[0]
    summary_data.append(["Total Titles", f"{total:,}"])

    type_counts = conn.execute("""
        SELECT titleType, COUNT(*) as cnt
        FROM title_basics
        GROUP BY titleType
        ORDER BY cnt DESC
    """).fetchall()
    for ttype, cnt in type_counts[:6]:
        summary_data.append([f"  {ttype}", f"{cnt:,}"])

    rated = conn.execute("SELECT COUNT(*) FROM title_ratings").fetchone()[0]
    summary_data.append(["Titles with Ratings", f"{rated:,}"])

    people = conn.execute("SELECT COUNT(*) FROM name_basics").fetchone()[0]
    summary_data.append(["Total People", f"{people:,}"])

    if "title_principals" in tables:
        principals = conn.execute("SELECT COUNT(*) FROM title_principals").fetchone()[0]
        summary_data.append(["Cast/Crew Records", f"{principals:,}"])

    avg_rating = conn.execute("SELECT ROUND(AVG(averageRating), 2) FROM title_ratings").fetchone()[0]
    summary_data.append(["Average Rating", str(avg_rating)])

    summary_df = pd.DataFrame(summary_data, columns=["Metric", "Value"])
    current_row = write_stats_table(stats_sheet, summary_df, current_row, "Database Summary")

    # --- Table 2: Top 10 Rated Movies ---
    print("   Creating Top Rated Movies...")
    top_rated = conn.execute("""
        SELECT
            tb.primaryTitle as Title,
            tb.startYear as Year,
            tr.averageRating as Rating,
            tr.numVotes as Votes,
            tb.genres as Genres
        FROM title_basics tb
        JOIN title_ratings tr ON tb.tconst = tr.tconst
        WHERE tb.titleType = 'movie'
          AND tr.numVotes > 50000
        ORDER BY tr.averageRating DESC, tr.numVotes DESC
        LIMIT 10
    """).fetchdf()

    current_row = write_stats_table(stats_sheet, top_rated, current_row,
                                     "Top 10 Highest Rated Movies (50K+ votes)")

    # --- Table 3: Movies by Decade ---
    print("   Creating Titles by Decade...")
    by_decade = conn.execute("""
        SELECT
            (startYear / 10) * 10 as Decade,
            COUNT(*) as Total,
            COUNT(CASE WHEN titleType = 'movie' THEN 1 END) as Films,
            COUNT(CASE WHEN titleType = 'tvSeries' THEN 1 END) as TV_Series
        FROM title_basics
        WHERE startYear IS NOT NULL AND startYear >= 1900 AND startYear <= 2025
        GROUP BY (startYear / 10) * 10
        ORDER BY Decade
    """).fetchall()

    decade_df = pd.DataFrame(
        [[f"{d[0]}s", d[1], d[2], d[3]] for d in by_decade],
        columns=["Decade", "Total", "Movies", "TV Series"]
    )
    current_row = write_stats_table(stats_sheet, decade_df, current_row, "Titles by Decade")

    # --- Table 4: Top Genres ---
    print("   Creating Top Genres...")
    top_genres = conn.execute("""
        WITH genre_split AS (
            SELECT UNNEST(STRING_SPLIT(genres, ',')) as genre
            FROM title_basics
            WHERE genres IS NOT NULL AND genres != ''
        )
        SELECT TRIM(genre) as Genre, COUNT(*) as Count
        FROM genre_split
        WHERE TRIM(genre) != ''
        GROUP BY TRIM(genre)
        ORDER BY Count DESC
        LIMIT 15
    """).fetchdf()

    current_row = write_stats_table(stats_sheet, top_genres, current_row, "Top 15 Genres")

    # -------------------------------------------------------------------------
    # PHASE 3: GENERATE CHARTS
    # -------------------------------------------------------------------------
    print("\n[3/4] Generating charts...")

    temp_dir = tempfile.gettempdir()

    chart_sheet = ensure_sheet_exists(book, "IMDB_CHARTS")
    chart_sheet.range("A1").value = "IMDB Data Visualizations"
    chart_sheet.range("A1").font.bold = True
    chart_sheet.range("A1").font.size = 16

    chart_sheet.range("A2").value = f"Generated: {time.strftime('%Y-%m-%d %H:%M')} (Remote Module)"
    chart_sheet.range("A2").font.size = 10

    chart_sheet.range("A3").value = "Showcasing analytics & chart generation capabilities."
    chart_sheet.range("A3").font.size = 10

    fig_width, fig_height = 8, 6
    scale_factor = 0.5
    charts_created = 0

    # --- Chart 1: Movies by Decade (Bar) ---
    print("   Chart 1: Movies by Decade...")
    try:
        decades = [f"{d[0]}s" for d in by_decade if d[0] >= 1920]
        counts = [d[1] for d in by_decade if d[0] >= 1920]

        fig, ax = plt.subplots(figsize=(fig_width, fig_height))
        ax.bar(decades, counts, color="#4472C4", alpha=0.85, edgecolor="black", linewidth=0.5)
        ax.set_xlabel("Decade", fontsize=10, fontweight="bold")
        ax.set_ylabel("Number of Titles", fontsize=10, fontweight="bold")
        ax.set_title("Titles by Decade", fontsize=12, fontweight="bold")
        ax.tick_params(axis="x", rotation=45)
        ax.grid(axis="y", alpha=0.3)
        plt.tight_layout()

        chart_path = os.path.join(temp_dir, "imdb_chart_decade.png")
        fig.savefig(chart_path, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)

        img = Image.open(chart_path)
        resized_path = os.path.join(temp_dir, "imdb_chart_decade_resized.png")
        img.resize(
            (int(img.width * scale_factor), int(img.height * scale_factor)),
            Image.LANCZOS
        ).save(resized_path)

        chart_sheet.pictures.add(
            resized_path,
            name="chart_decade",
            update=True,
            anchor=chart_sheet.range("A5")
        )
        charts_created += 1
    except Exception as e:
        print(f"      Error: {e}")

    # --- Chart 2: Title Types (Pie) ---
    print("   Chart 2: Title Types Distribution...")
    try:
        types = [t[0] for t in type_counts[:5]]
        type_vals = [t[1] for t in type_counts[:5]]
        other_count = sum(t[1] for t in type_counts[5:])
        if other_count > 0:
            types.append("Other")
            type_vals.append(other_count)

        fig, ax = plt.subplots(figsize=(fig_width, fig_height))
        colors = plt.cm.Set3(range(len(types)))
        ax.pie(type_vals, labels=types, autopct="%1.1f%%", colors=colors, startangle=90)
        ax.set_title("Title Type Distribution", fontsize=12, fontweight="bold")
        plt.tight_layout()

        chart_path = os.path.join(temp_dir, "imdb_chart_types.png")
        fig.savefig(chart_path, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)

        img = Image.open(chart_path)
        resized_path = os.path.join(temp_dir, "imdb_chart_types_resized.png")
        img.resize(
            (int(img.width * scale_factor), int(img.height * scale_factor)),
            Image.LANCZOS
        ).save(resized_path)

        chart_sheet.pictures.add(
            resized_path,
            name="chart_types",
            update=True,
            anchor=chart_sheet.range("K5")
        )
        charts_created += 1
    except Exception as e:
        print(f"      Error: {e}")

    # --- Chart 3: Rating Distribution (Histogram) ---
    print("   Chart 3: Rating Distribution...")
    try:
        ratings_dist = conn.execute("""
            SELECT FLOOR(averageRating) as rating_bucket, COUNT(*) as cnt
            FROM title_ratings
            GROUP BY FLOOR(averageRating)
            ORDER BY rating_bucket
        """).fetchall()

        buckets = [int(r[0]) for r in ratings_dist]
        r_counts = [r[1] for r in ratings_dist]

        fig, ax = plt.subplots(figsize=(fig_width, fig_height))
        ax.bar(buckets, r_counts, color="#FF6D00", alpha=0.85, edgecolor="black", linewidth=0.5, width=0.8)
        ax.set_xlabel("Rating", fontsize=10, fontweight="bold")
        ax.set_ylabel("Number of Titles", fontsize=10, fontweight="bold")
        ax.set_title("Rating Distribution", fontsize=12, fontweight="bold")
        ax.set_xticks(range(1, 11))
        ax.grid(axis="y", alpha=0.3)
        plt.tight_layout()

        chart_path = os.path.join(temp_dir, "imdb_chart_ratings.png")
        fig.savefig(chart_path, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)

        img = Image.open(chart_path)
        resized_path = os.path.join(temp_dir, "imdb_chart_ratings_resized.png")
        img.resize(
            (int(img.width * scale_factor), int(img.height * scale_factor)),
            Image.LANCZOS
        ).save(resized_path)

        chart_sheet.pictures.add(
            resized_path,
            name="chart_ratings",
            update=True,
            anchor=chart_sheet.range("A25")
        )
        charts_created += 1
    except Exception as e:
        print(f"      Error: {e}")

    # --- Chart 4: Top Genres (Horizontal Bar) ---
    print("   Chart 4: Top Genres...")
    try:
        genres_list = top_genres['Genre'].tolist()[:10][::-1]
        genre_counts = top_genres['Count'].tolist()[:10][::-1]

        fig, ax = plt.subplots(figsize=(fig_width, fig_height))
        ax.barh(genres_list, genre_counts, color="#2E7D32", alpha=0.85, edgecolor="black", linewidth=0.5)
        ax.set_xlabel("Number of Titles", fontsize=10, fontweight="bold")
        ax.set_title("Top 10 Genres", fontsize=12, fontweight="bold")
        ax.grid(axis="x", alpha=0.3)
        plt.tight_layout()

        chart_path = os.path.join(temp_dir, "imdb_chart_genres.png")
        fig.savefig(chart_path, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)

        img = Image.open(chart_path)
        resized_path = os.path.join(temp_dir, "imdb_chart_genres_resized.png")
        img.resize(
            (int(img.width * scale_factor), int(img.height * scale_factor)),
            Image.LANCZOS
        ).save(resized_path)

        chart_sheet.pictures.add(
            resized_path,
            name="chart_genres",
            update=True,
            anchor=chart_sheet.range("K25")
        )
        charts_created += 1
    except Exception as e:
        print(f"      Error: {e}")

    # --- Chart 5: Movies per Year 2000-2024 (Line) ---
    print("   Chart 5: Movies per Year (2000-2024)...")
    try:
        yearly = conn.execute("""
            SELECT startYear, COUNT(*) as cnt
            FROM title_basics
            WHERE titleType = 'movie' AND startYear >= 2000 AND startYear <= 2024
            GROUP BY startYear
            ORDER BY startYear
        """).fetchall()

        years = [y[0] for y in yearly]
        year_counts = [y[1] for y in yearly]

        fig, ax = plt.subplots(figsize=(fig_width, fig_height))
        ax.plot(years, year_counts, marker="o", color="#9C27B0", linewidth=2, markersize=5)
        ax.fill_between(years, year_counts, alpha=0.3, color="#9C27B0")
        ax.set_xlabel("Year", fontsize=10, fontweight="bold")
        ax.set_ylabel("Number of Movies", fontsize=10, fontweight="bold")
        ax.set_title("Movies Released per Year (2000-2024)", fontsize=12, fontweight="bold")
        ax.grid(True, alpha=0.3)
        plt.tight_layout()

        chart_path = os.path.join(temp_dir, "imdb_chart_yearly.png")
        fig.savefig(chart_path, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)

        img = Image.open(chart_path)
        resized_path = os.path.join(temp_dir, "imdb_chart_yearly_resized.png")
        img.resize(
            (int(img.width * scale_factor), int(img.height * scale_factor)),
            Image.LANCZOS
        ).save(resized_path)

        chart_sheet.pictures.add(
            resized_path,
            name="chart_yearly",
            update=True,
            anchor=chart_sheet.range("A45")
        )
        charts_created += 1
    except Exception as e:
        print(f"      Error: {e}")

    # --- Chart 6: Average Rating by Title Type (Bar) ---
    print("   Chart 6: Avg Rating by Title Type...")
    try:
        avg_by_type = conn.execute("""
            SELECT tb.titleType, ROUND(AVG(tr.averageRating), 2) as avg_rating
            FROM title_basics tb
            JOIN title_ratings tr ON tb.tconst = tr.tconst
            GROUP BY tb.titleType
            HAVING COUNT(*) > 1000
            ORDER BY avg_rating DESC
            LIMIT 8
        """).fetchall()

        ttypes = [t[0] for t in avg_by_type]
        avg_ratings = [float(t[1]) for t in avg_by_type]

        fig, ax = plt.subplots(figsize=(fig_width, fig_height))
        bars = ax.bar(ttypes, avg_ratings, color="#26A69A", alpha=0.85, edgecolor="black", linewidth=0.5)
        ax.set_xlabel("Title Type", fontsize=10, fontweight="bold")
        ax.set_ylabel("Average Rating", fontsize=10, fontweight="bold")
        ax.set_title("Average Rating by Title Type", fontsize=12, fontweight="bold")
        ax.set_ylim(0, 10)
        ax.tick_params(axis="x", rotation=45)
        ax.grid(axis="y", alpha=0.3)

        for bar, val in zip(bars, avg_ratings):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                   f"{val:.1f}", ha="center", va="bottom", fontsize=9, fontweight="bold")

        plt.tight_layout()

        chart_path = os.path.join(temp_dir, "imdb_chart_avgrating.png")
        fig.savefig(chart_path, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)

        img = Image.open(chart_path)
        resized_path = os.path.join(temp_dir, "imdb_chart_avgrating_resized.png")
        img.resize(
            (int(img.width * scale_factor), int(img.height * scale_factor)),
            Image.LANCZOS
        ).save(resized_path)

        chart_sheet.pictures.add(
            resized_path,
            name="chart_avgrating",
            update=True,
            anchor=chart_sheet.range("K45")
        )
        charts_created += 1
    except Exception as e:
        print(f"      Error: {e}")

    # -------------------------------------------------------------------------
    # PHASE 4: FINALIZE
    # -------------------------------------------------------------------------
    print("\n[4/4] Finalizing...")

    conn.close()
    stats_sheet.activate()

    print("\n" + "=" * 60)
    print("IMDB STATS COMPLETE! (Remote Module)")
    print(f"Sheets: IMDB_STATS (all tables), IMDB_CHARTS (6 charts)")
    print(f"Charts created: {charts_created}")
    print("=" * 60)
