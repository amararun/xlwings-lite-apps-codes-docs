import xlwings as xw
import pandas as pd
import numpy as np
import re
import textwrap
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from typing import Tuple, Optional
from xlwings import script

# -------------------------
# Helper Functions (Following AI Coder Guidelines)
# -------------------------

def find_table_in_workbook(book: xw.Book, table_name: str) -> Tuple[Optional[xw.Sheet], Optional[object]]:
    """
    Searches all sheets for a table and returns both the sheet and table objects.
    Returns: (xw.Sheet, table_object) or (None, None) if not found.
    """
    for sheet in book.sheets:
        if table_name in sheet.tables:
            return sheet, sheet.tables[table_name]
    return None, None

def create_canonical_name(df: pd.DataFrame) -> pd.Series:
    """Create canonical name from available name columns"""
    cols_lower = {c.lower(): c for c in df.columns}
    preferred = cols_lower.get("company_name_std") or cols_lower.get("company_name")
    fallback = cols_lower.get("instrument_name") or cols_lower.get("instrument")
    
    def choose(row):
        val = row.get(preferred)
        if preferred and pd.notna(val) and val != "":
            return val
        val = row.get(fallback)
        if fallback and pd.notna(val) and val != "":
            return val
        for name_col_lower in ["company_name_std", "company_name", "instrument_name", "instrument"]:
            actual_col = cols_lower.get(name_col_lower)
            if not actual_col: continue
            val = row.get(actual_col)
            if pd.notna(val) and val != "":
                return val
        return ""
    return df.apply(choose, axis=1)

def standardize_display_name(name: str) -> str:
    """Standardize display name"""
    if pd.isna(name):
        return ""
    s = str(name).strip()
    s = re.sub(r"\s+", " ", s)
    return s

def standardize_key(name: str) -> str:
    """Create standardized key for grouping"""
    s = str(name).lower()
    s = re.sub(r"[^a-z0-9]", "", s)
    return s

def create_export_df(summary_df: pd.DataFrame, monthA: pd.Timestamp, monthB: pd.Timestamp) -> pd.DataFrame:
    """
    Transforms the summary DataFrame into the export format with standardized p1/p2 column names.
    P1 (monthA) is the earlier period, P2 (monthB) is the later period.
    """
    monthA_b = monthA.strftime('%b')
    monthB_b = monthB.strftime('%b')
    monthA_Y = monthA.strftime('%Y')
    monthB_Y = monthB.strftime('%Y')

    rename_map = {
        "Name": "company_name",
        f"Mkt. Val {monthA_b} - %": "mkt_val_p1_pct",
        f"Mkt. Val {monthB_b} - %": "mkt_val_p2_pct",
        "MV % Change": "mv_pct_change",
        "Qty % Change": "qty_pct_change",
        f"Num of MF {monthA_b} - MV": "num_mf_p1_mv",
        f"Num of MF {monthB_b} - MV": "num_mf_p2_mv",
        f"Num of MF {monthA_b} - Qty": "num_mf_p1_qty",
        f"Num of MF {monthB_b} - Qty": "num_mf_p2_qty",
        f"Mkt Val. {monthA_b} {monthA_Y}": "mkt_val_p1",
        f"Mkt Val. {monthB_b} {monthB_Y}": "mkt_val_p2",
        f"Qty {monthA_b} {monthA_Y}": "qty_p1",
        f"Qty {monthB_b} {monthB_Y}": "qty_p2",
        "Comment": "comment"
    }

    export_df = summary_df.rename(columns=rename_map)
    
    final_cols = [
        "company_name", "mkt_val_p1_pct", "mkt_val_p2_pct", "mv_pct_change", "qty_pct_change",
        "num_mf_p1_mv", "num_mf_p2_mv", "num_mf_p1_qty", "num_mf_p2_qty",
        "mkt_val_p1", "mkt_val_p2", "qty_p1", "qty_p2", "comment"
    ]
    
    existing_cols = [col for col in final_cols if col in export_df.columns]
    
    return export_df[existing_cols]

# -------------------------
# Chart Generation Functions
# -------------------------

def process_data_for_charts(df: pd.DataFrame, monthA: pd.Timestamp, monthB: pd.Timestamp) -> list:
    """
    Transforms the summary DataFrame into a list of dictionaries suitable for chart generation.
    """
    monthA_str = monthA.strftime('%b')
    monthB_str = monthB.strftime('%b')
    
    col_mv_pct_A = f"Mkt. Val {monthA_str} - %"
    col_mv_pct_B = f"Mkt. Val {monthB_str} - %"
    col_num_mf_A = f"Num of MF {monthA_str} - MV"
    col_num_mf_B = f"Num of MF {monthB_str} - MV"
    
    chart_data = []
    for _, row in df.iterrows():
        chart_data.append({
            'name': row['Name'],
            'mvChange': row['MV % Change'] * 100 if pd.notna(row['MV % Change']) else 0,
            'qtyChange': row['Qty % Change'] * 100 if pd.notna(row['Qty % Change']) else 0,
            'numofmf_A': row[col_num_mf_A],
            'numofmf_B': row[col_num_mf_B],
            'marketValuePct_A': row[col_mv_pct_A] * 100 if pd.notna(row[col_mv_pct_A]) else 0,
            'marketValuePct_B': row[col_mv_pct_B] * 100 if pd.notna(row[col_mv_pct_B]) else 0,
            'comment': row.get('Comment', '')
        })
    return chart_data

def run_chart_generation(book: xw.Book, summary_df: pd.DataFrame, monthA: pd.Timestamp, monthB: pd.Timestamp):
    """
    Main logic to create and insert the two side-by-side top 10 holdings charts.
    """
    chart_sheet_name = "CHART_TOPHOLD"
    if chart_sheet_name in [s.name for s in book.sheets]:
        try:
            book.sheets[chart_sheet_name].delete()
            print(f"🧹 Deleted existing '{chart_sheet_name}' sheet.")
        except Exception as e:
            print(f"⚠️ Could not delete existing sheet '{chart_sheet_name}': {e}")
    
    chart_sheet = book.sheets.add(name=chart_sheet_name)
    print(f"📄 Created new '{chart_sheet_name}' sheet.")

    def _generate_one_chart(month_to_chart: pd.Timestamp, anchor: str):
        print(f"📈 Preparing Top 10 data for {month_to_chart.strftime('%B %Y')}:")
        
        month_col = f"Mkt. Val {month_to_chart.strftime('%b')} - %"
        chart_df = summary_df.copy()
        chart_df = chart_df[chart_df[month_col].fillna(0) > 0].sort_values(month_col, ascending=False).head(10)
        
        if len(chart_df) > 0:
            other_month = monthB if month_to_chart == monthA else monthA
            
            processed_data = process_data_for_charts(chart_df, monthA, monthB)
            fig = create_single_compound_chart(processed_data, display_period=month_to_chart, other_period=other_month)
            save_and_insert_chart(fig, chart_sheet, f"Top10_Holdings_{month_to_chart.strftime('%b')}", anchor=anchor)
        else:
            print(f"⚠️ No data available to generate chart for {month_to_chart.strftime('%B %Y')}.")

    _generate_one_chart(month_to_chart=monthB, anchor="A1")
    _generate_one_chart(month_to_chart=monthA, anchor="L1")
    
    chart_sheet.activate()

def create_single_compound_chart(data: list, display_period: pd.Timestamp, other_period: pd.Timestamp):
    """
    Creates a single compound chart for a given period, consisting of three vertically stacked subplots.
    """
    fig, axes = plt.subplots(3, 1, figsize=(7, 9), height_ratios=[3, 2, 2.5], facecolor='white')
    plt.subplots_adjust(hspace=0)

    companies = [item['name'][:15] + '...' if len(item['name']) > 15 else item['name'] for item in data]
    mv_changes = [item['mvChange'] for item in data]
    qty_changes = [item['qtyChange'] for item in data]
    funds_A = [item['numofmf_A'] for item in data]
    funds_B = [item['numofmf_B'] for item in data]
    market_values_A = [item['marketValuePct_A'] for item in data]
    market_values_B = [item['marketValuePct_B'] for item in data]
    comments = [item['comment'] for item in data]

    earlier_month = min(display_period, other_period)
    later_month = max(display_period, other_period)

    outlier_text = create_percentage_change_chart(axes[0], companies, mv_changes, qty_changes)
    create_fund_count_chart(axes[1], companies, funds_A, funds_B, earlier_month, later_month)
    
    market_values_to_display = market_values_B if display_period > other_period else market_values_A
    create_market_value_chart(axes[2], companies, market_values_to_display, comments, display_period.strftime('%b'))

    axes[0].set_title(f'Top 10 Holdings - {display_period.strftime("%B %Y")}', 
                      fontsize=14, fontweight='bold', pad=25)

    if outlier_text:
        wrapped_outlier_text = textwrap.fill(outlier_text, width=100)
        fig.text(0.5, 0.01, wrapped_outlier_text, ha='center', va='bottom', fontsize=8, color='black',
                 bbox=dict(boxstyle='round,pad=0.4', facecolor='#FFEBEE', edgecolor='#E57373'))

    plt.tight_layout(rect=[0, 0.05, 1, 0.95])
    return fig

def save_and_insert_chart(fig, output_sheet: xw.Sheet, chart_name: str, anchor: str = "A1"):
    """
    Save chart to temporary file and insert into Excel.
    """
    import tempfile
    import os
    
    try:
        temp_dir = tempfile.gettempdir()
        chart_path = os.path.join(temp_dir, f"{chart_name}.png")
        
        fig.savefig(chart_path, dpi=100, bbox_inches='tight', facecolor='white', edgecolor='none')
        print(f"📊 Chart saved to: {chart_path}")
        
        anchor_cell = output_sheet[anchor]
        output_sheet.pictures.add(chart_path, name=chart_name, update=True, anchor=anchor_cell, format="png")
        print(f"✅ Chart '{chart_name}' inserted into Excel successfully")
        
        os.remove(chart_path)
        plt.close(fig)
        
    except Exception as e:
        print(f"❌ ERROR inserting chart '{chart_name}': {e}")
        if 'fig' in locals(): plt.close(fig)

def create_percentage_change_chart(ax, companies, mv_changes, qty_changes):
    """
    TOP CHART: Grouped bar chart with outlier text box.
    """
    mv_capped = [max(-50, min(125, val)) if pd.notna(val) else 0 for val in mv_changes]
    qty_capped = [max(-50, min(125, val)) if pd.notna(val) else 0 for val in qty_changes]
    
    x = np.arange(len(companies))
    width = 0.35
    
    ax.bar(x - width/2, mv_capped, width, label='MV % Change', color='#f59e0b', alpha=0.9)
    ax.bar(x + width/2, qty_capped, width, label='Qty % Change', color='#059669', alpha=0.9)
    
    outliers = []
    for i, (mv_orig, qty_orig, name) in enumerate(zip(mv_changes, qty_changes, companies)):
        if pd.notna(mv_orig) and (mv_orig > 125 or mv_orig < -50):
            outliers.append(f"{name}: MV {mv_orig:+.0f}%")
            bar_height = mv_capped[i]
            marker_y = bar_height + 8 if bar_height > 0 else bar_height - 18
            ax.text(x[i] - width/2, marker_y, '⚠', color='#B91C1C', fontsize=12, ha='center', va='center')

        if pd.notna(qty_orig) and (qty_orig > 125 or qty_orig < -50):
            outliers.append(f"{name}: Qty {qty_orig:+.0f}%")
            bar_height = qty_capped[i]
            marker_y = bar_height + 8 if bar_height > 0 else bar_height - 18
            ax.text(x[i] + width/2, marker_y, '⚠', color='#B91C1C', fontsize=12, ha='center', va='center')

    ax.set_ylabel('% Change', fontweight='bold')
    ax.set_ylim(-60, 140)
    ax.axhline(y=0, color='#d1d5db', linewidth=1, zorder=1)
    ax.grid(False)
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.2), ncol=2, frameon=False, fontsize=10)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f'{y:+.0f}%'))
    
    for spine in ['top', 'right', 'bottom']:
        ax.spines[spine].set_visible(False)
    ax.spines['left'].set_color('#d1d5db')
    ax.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom=False)
    ax.tick_params(axis='y', colors='#525252')
    
    if outliers:
        return f"🔺 OUTLIERS (Capped): { ', '.join(outliers)}"
    return ""

def create_fund_count_chart(ax, companies, funds_A, funds_B, monthA: pd.Timestamp, monthB: pd.Timestamp):
    """
    MIDDLE CHART: Dumbbell chart for fund count comparison.
    """
    x = np.arange(len(companies))
    
    for i in range(len(companies)):
        if funds_B[i] != funds_A[i]:
            ax.plot([x[i], x[i]], [funds_A[i], funds_B[i]], color='#9CA3AF', linewidth=2, zorder=1)
        else:
            ax.plot([x[i] - 0.2, x[i] + 0.2], [funds_A[i], funds_A[i]], color='#374151', linewidth=2, solid_capstyle='butt')

    ax.scatter(x, funds_A, marker='s', s=80, color='#DC2626', zorder=3)
    ax.scatter(x, funds_B, marker='o', s=80, color='#2563EB', zorder=3)

    ax.set_ylabel('Number of Funds', fontweight='bold')
    ax.grid(False)
    ax.locator_params(axis='y', nbins=4, integer=True)
    
    legend_elements = [
        plt.Line2D([0], [0], marker='s', color='w', markerfacecolor='#DC2626', markersize=8, label=monthA.strftime('%b %Y')),
        plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='#2563EB', markersize=8, label=monthB.strftime('%b %Y')),
        plt.Line2D([0], [0], color='#374151', lw=2, label='No Change')
    ]
    ax.legend(handles=legend_elements, loc='upper center', bbox_to_anchor=(0.5, 1.3), ncol=3, frameon=False, fontsize=10)

    for spine in ['top', 'right', 'bottom']:
        ax.spines[spine].set_visible(False)
    ax.spines['left'].set_color('#d1d5db')
    ax.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom=False)
    ax.tick_params(axis='y', colors='#525252')

def create_market_value_chart(ax, companies, market_values, comments, period_str):
    """
    BOTTOM CHART: Market value percentage bars.
    """
    x = np.arange(len(companies))
    
    color = '#2563EB' if period_str.lower() == 'aug' else '#DC2626'
    
    bars = ax.bar(x, market_values, color=color, alpha=0.9)

    # Add comment annotations
    for i, bar in enumerate(bars):
        comment = comments[i]
        if pd.notna(comment) and comment.strip() != "":
            y_val = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2.0, y_val + 0.4, 
                    str(comment)[:7], 
                    ha='center', va='bottom', fontsize=9, color='black', fontweight='bold',
                    bbox=dict(boxstyle='round,pad=0.4', facecolor='#FFEBEE', edgecolor='#E57373'))

    ax.set_ylabel('Market Value %', fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(companies, rotation=45, ha='right', fontsize=10, color='#525252')
    ax.grid(False)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f'{y:.1f}%'))
    
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
    ax.spines['left'].set_color('#d1d5db')
    ax.spines['bottom'].set_color('#d1d5db')
    ax.tick_params(axis='y', colors='#525252')
    ax.tick_params(axis='x', colors='#525252')
    
# -------------------------
# Pipeline Steps
# -------------------------

def step_validation_totals(df: pd.DataFrame, monthA: str, monthB: str) -> pd.DataFrame:
    """
    Produce totals table: scheme_name | metric | monthA_total | monthB_total
    """
    print("📊 Creating validation totals...")
    
    df["MARKET_VALUE"] = pd.to_numeric(df.get("MARKET_VALUE", 0), errors="coerce").fillna(0.0)
    df["QUANTITY"] = pd.to_numeric(df.get("QUANTITY", 0), errors="coerce").fillna(0.0)

    agg = df.groupby(["SCHEME_NAME", "MONTH_END"], dropna=False).agg(
        market_value_sum=("MARKET_VALUE", "sum"),
        quantity_sum=("QUANTITY", "sum")
    ).reset_index()

    mv_pivot = agg.pivot(index="SCHEME_NAME", columns="MONTH_END", values="market_value_sum").fillna(0).reset_index()
    qty_pivot = agg.pivot(index="SCHEME_NAME", columns="MONTH_END", values="quantity_sum").fillna(0).reset_index()

    def format_month_col(col_name):
        if pd.isna(col_name) or col_name == 'SCHEME_NAME':
            return col_name
        try:
            if isinstance(col_name, str):
                dt = pd.to_datetime(col_name)
            else:
                dt = col_name
            return dt.strftime('%d-%m-%Y')
        except:
            return str(col_name)

    mv_pivot.columns = [format_month_col(col) for col in mv_pivot.columns]
    qty_pivot.columns = [format_month_col(col) for col in qty_pivot.columns]

    monthA_formatted = format_month_col(monthA)
    monthB_formatted = format_month_col(monthB)

    rows = []
    for _, r in mv_pivot.iterrows():
        rows.append({"metric": f"{r['SCHEME_NAME']} - Market Value", monthA_formatted: r.get(monthA_formatted, 0.0), monthB_formatted: r.get(monthB_formatted, 0.0)})
    for _, r in qty_pivot.iterrows():
        rows.append({"metric": f"{r['SCHEME_NAME']} - Quantity", monthA_formatted: r.get(monthA_formatted, 0.0), monthB_formatted: r.get(monthB_formatted, 0.0)})

    return pd.DataFrame(rows)

def step_isin_mapping(df: pd.DataFrame, cut_n: int = 7) -> pd.DataFrame:
    """
    Build ISIN mapping: one row per ISIN, with an exception tag.
    """
    print("🗺️ Creating ISIN mapping with exception tags...")
    
    df = df.copy()
    
    df["canonical_name_raw"] = create_canonical_name(df)
    d2 = df.drop_duplicates(subset=["ISIN"], keep="first").copy().reset_index(drop=True)
    d2["standardized_name_display"] = d2["canonical_name_raw"].apply(standardize_display_name)
    d2["standardized_name_key"] = d2["standardized_name_display"].apply(standardize_key)
    d2["name_cut"] = d2["standardized_name_key"].str.slice(0, cut_n)
    
    # --- Exception Tagging Logic ---
    exception_counts = d2.groupby("name_cut").agg(
        name_count=('standardized_name_display', 'nunique'),
        isin_count=('ISIN', 'nunique')
    ).reset_index()
    
    problem_cuts = exception_counts[
        (exception_counts['name_count'] > 1) | 
        (exception_counts['isin_count'] > 1)
    ]['name_cut'].tolist()

    d2['exception_tag'] = np.where(d2['name_cut'].isin(problem_cuts), 1, 0)
    d2["comments"] = ""
    # --- End of Exception Logic ---

    output_cols = ["ISIN", "standardized_name_display", "name_cut", "exception_tag", "comments"]
    out = d2[output_cols]
    
    out = out.sort_values(by=["exception_tag", "name_cut", "standardized_name_display"], ascending=[False, True, True])
    
    print(f"✅ ISIN mapping created. Tagged {d2['exception_tag'].sum()} ISINs as exceptions.")
    
    return out

def step_summary_by_standardized_name(df: pd.DataFrame, isin_map_df: pd.DataFrame, monthA: str, monthB: str) -> pd.DataFrame:
    """
    Final summary by standardized_name_display (grouped using ISIN mapping).
    """
    print("📈 Creating final summary...")
    
    df = df.copy()
    df["MARKET_VALUE"] = pd.to_numeric(df.get("MARKET_VALUE",0), errors="coerce").fillna(0.0)
    df["QUANTITY"] = pd.to_numeric(df.get("QUANTITY",0), errors="coerce").fillna(0.0)

    merged = df.merge(isin_map_df, left_on="ISIN", right_on="ISIN", how="left")

    merged["standardized_name_display"] = merged["standardized_name_display"].fillna(merged.get("COMPANY_NAME_STD","")).fillna(merged.get("INSTRUMENT_NAME", ""))

    monthA_friendly = f"{monthA.strftime('%b')}_{monthA.strftime('%Y')}"
    monthB_friendly = f"{monthB.strftime('%b')}_{monthB.strftime('%Y')}"
    month_map = {monthA: monthA_friendly, monthB: monthB_friendly}
    merged["month_friendly"] = merged["MONTH_END"].map(month_map).fillna(merged["MONTH_END"])

    agg = merged.groupby(["standardized_name_display", "month_friendly"], dropna=False).agg(
        mv_sum=("MARKET_VALUE","sum"),
        qty_sum=("QUANTITY","sum")
    ).reset_index()

    mv_pivot = agg.pivot(index="standardized_name_display", columns="month_friendly", values="mv_sum").fillna(0).reset_index()
    qty_pivot = agg.pivot(index="standardized_name_display", columns="month_friendly", values="qty_sum").fillna(0).reset_index()

    for c in [monthA_friendly, monthB_friendly]:
        if c not in mv_pivot.columns:
            mv_pivot[c] = 0.0
        if c not in qty_pivot.columns:
            qty_pivot[c] = 0.0

    summary = mv_pivot.merge(qty_pivot, on="standardized_name_display", suffixes=('_mv','_qty'))

    summary = summary.rename(columns={
        monthA_friendly: f"{monthA_friendly}_mv", 
        monthB_friendly: f"{monthB_friendly}_mv", 
        f"{monthA_friendly}_qty": f"{monthA_friendly}_qty", 
        f"{monthB_friendly}_qty": f"{monthB_friendly}_qty"
    })
    
    total_monthA_mv = summary[f"{monthA_friendly}_mv"].sum()
    total_monthB_mv = summary[f"{monthB_friendly}_mv"].sum()

    summary["mv_monthA_pct"] = summary[f"{monthA_friendly}_mv"] / (total_monthA_mv if total_monthA_mv != 0 else 1.0)
    summary["mv_monthB_pct"] = summary[f"{monthB_friendly}_mv"] / (total_monthB_mv if total_monthB_mv != 0 else 1.0)

    summary[f"{monthA_friendly}_qty"] = pd.to_numeric(summary.get(f"{monthA_friendly}_qty", 0.0), errors="coerce").fillna(0.0)
    summary[f"{monthB_friendly}_qty"] = pd.to_numeric(summary.get(f"{monthB_friendly}_qty", 0.0), errors="coerce").fillna(0.0)

    scheme_agg = merged.groupby(["standardized_name_display","SCHEME_NAME","month_friendly"], dropna=False).agg(mv_sum=("MARKET_VALUE","sum"), qty_sum=("QUANTITY","sum")).reset_index()
    mf_count_mv_monthA = scheme_agg[(scheme_agg["month_friendly"]==monthA_friendly) & (scheme_agg["mv_sum"]>0)].groupby("standardized_name_display")["SCHEME_NAME"].nunique().rename("mf_count_mv_monthA")
    mf_count_mv_monthB = scheme_agg[(scheme_agg["month_friendly"]==monthB_friendly) & (scheme_agg["mv_sum"]>0)].groupby("standardized_name_display")["SCHEME_NAME"].nunique().rename("mf_count_mv_monthB")
    mf_count_qty_monthA = scheme_agg[(scheme_agg["month_friendly"]==monthA_friendly) & (scheme_agg["qty_sum"]>0)].groupby("standardized_name_display")["SCHEME_NAME"].nunique().rename("mf_count_qty_monthA")
    mf_count_qty_monthB = scheme_agg[(scheme_agg["month_friendly"]==monthB_friendly) & (scheme_agg["qty_sum"]>0)].groupby("standardized_name_display")["SCHEME_NAME"].nunique().rename("mf_count_qty_monthB")

    summary = summary.merge(mf_count_mv_monthA, on="standardized_name_display", how="left")
    summary = summary.merge(mf_count_mv_monthB, on="standardized_name_display", how="left")
    summary = summary.merge(mf_count_qty_monthA, on="standardized_name_display", how="left")
    summary = summary.merge(mf_count_qty_monthB, on="standardized_name_display", how="left")
    
    for col in ["mf_count_mv_monthA","mf_count_mv_monthB","mf_count_qty_monthA","mf_count_qty_monthB"]:
        summary[col] = summary[col].fillna(0).astype(int)

    def pct_change_frac(new, old):
        try:
            new = float(new); old = float(old)
        except:
            return np.nan
        if old == 0:
            return np.nan
        return new/old - 1.0

    summary["mv_pct_change"] = summary.apply(lambda r: pct_change_frac(r.get(f"{monthB_friendly}_mv",0), r.get(f"{monthA_friendly}_mv",0)), axis=1)
    summary["qty_pct_change"] = summary.apply(lambda r: pct_change_frac(r.get(f"{monthB_friendly}_qty",0), r.get(f"{monthA_friendly}_qty",0)), axis=1)

    if "standardized_name_display" in isin_map_df.columns and "comments" in isin_map_df.columns:
        cm = isin_map_df[["standardized_name_display","comments"]].drop_duplicates().groupby("standardized_name_display")["comments"].apply(lambda s: "; ".join(sorted(set([v for v in s if v and str(v).strip()])))).reset_index().rename(columns={"comments":"Comment"})
        summary = summary.merge(cm, on="standardized_name_display", how="left")
        summary["Comment"] = summary["Comment"].fillna("")
    else:
        summary["Comment"] = ""

    final_cols = [
        "standardized_name_display",
        "mv_monthB_pct", "mv_monthA_pct", "mv_pct_change", "qty_pct_change",
        "mf_count_mv_monthB", "mf_count_mv_monthA", "mf_count_qty_monthB", "mf_count_qty_monthA",
        f"{monthB_friendly}_mv", f"{monthA_friendly}_mv", f"{monthB_friendly}_qty", f"{monthA_friendly}_qty",
        "Comment"
    ]
    final = summary[[col for col in final_cols if col in summary.columns]].copy()

    final = final.rename(columns={
        "standardized_name_display":"Name",
        "mv_monthA_pct":f"Mkt. Val {monthA.strftime('%b')} - %",
        "mv_monthB_pct":f"Mkt. Val {monthB.strftime('%b')} - %",
        "mv_pct_change":"MV % Change",
        "qty_pct_change":"Qty % Change",
        "mf_count_mv_monthA":f"Num of MF {monthA.strftime('%b')} - MV",
        "mf_count_mv_monthB":f"Num of MF {monthB.strftime('%b')} - MV",
        "mf_count_qty_monthA":f"Num of MF {monthA.strftime('%b')} - Qty",
        "mf_count_qty_monthB":f"Num of MF {monthB.strftime('%b')} - Qty",
        f"{monthA_friendly}_mv":f"Mkt Val. {monthA.strftime('%b')} {monthA.strftime('%Y')}",
        f"{monthB_friendly}_mv":f"Mkt Val. {monthB.strftime('%b')} {monthB.strftime('%Y')}",
        f"{monthA_friendly}_qty":f"Qty {monthA.strftime('%b')} {monthA.strftime('%Y')}",
        f"{monthB_friendly}_qty":f"Qty {monthB.strftime('%b')} {monthB.strftime('%Y')}"
    })
    return final

# -------------------------
# Main Script Function
# -------------------------

@script(button="[run_stage1_btn]Control!J19")
def stage1_full_pipeline(book: xw.Book):
    """
    Stage 1: Full MF Holdings Analysis Pipeline
    """
    run_full_pipeline(book, run_mode="full")

@script(button="[run_stage2_btn]Control!J24")
def stage2_summary_update(book: xw.Book):
    """
    Stage 2: Update Summary Analysis Only
    """
    run_full_pipeline(book, run_mode="stage2")

def run_full_pipeline(book: xw.Book, run_mode: str = "full"):
    """
    Main pipeline function with configurable run modes.
    """
    print("🚀 Starting MF Holdings Summary Processing...")
    print(f"📋 Running in {run_mode.upper()} mode")
    
    try:
        print("📋 Loading data from T_DATA table...")
        source_sheet, data_table = find_table_in_workbook(book, 'T_DATA')
        
        if source_sheet is None or data_table is None:
            print("❌ Error: T_DATA table not found in workbook")
            return
            
        print(f"✅ Found T_DATA table on sheet: {source_sheet.name}")
        
        df = data_table.range.options(pd.DataFrame, index=False).value
        print(f"📊 Loaded {len(df)} records from T_DATA table")
        
        unique_months = df['MONTH_END'].dropna().unique()
        print(f"📅 Found months: {list(unique_months)}")
        
        if len(unique_months) < 2:
            print("❌ Error: Need at least 2 months of data for comparison")
            return
            
        sorted_months = sorted(unique_months, reverse=True)
        monthB = sorted_months[0]
        monthA = sorted_months[1]

        print(f"📊 Comparing {monthA} vs {monthB}")
        print(f"ℹ️ Month detection is now set to use the two most recent dates in the data.")

        df = df[df['MONTH_END'].isin([monthA, monthB])]
        print(f"📊 Filtered data to {len(df)} records for the two selected months.")
        
        print(f"🔄 Processing data through pipeline (mode: {run_mode})...")
        
        if run_mode == "full":
            print("📊 Creating validation totals...")
            validation_df = step_validation_totals(df, monthA, monthB)
            
            print("🗺️ Creating ISIN mapping...")
            isin_mapping_df = step_isin_mapping(df)
            
        elif run_mode == "stage2":
            # --- Original ISIN Mapping Logic ---
            print("📋 Loading existing ISIN mapping from sheet...")
            try:
                print("🔍 Looking for ISIN_MAPPING sheet...")
                isin_sheet = book.sheets['ISIN_MAPPING']
                print(f"✅ Found ISIN_MAPPING sheet: {isin_sheet.name}")
                
                if len(isin_sheet.tables) > 0:
                    isin_table = isin_sheet.tables[0]
                else:
                    print("❌ No tables found in ISIN_MAPPING sheet")
                    return
                
                isin_mapping_df = isin_table.range.options(pd.DataFrame, index=False).value
                print(f"✅ Loaded existing ISIN mapping with {len(isin_mapping_df)} records")
                
            except Exception as e:
                print(f"❌ Error loading ISIN mapping: {e}")
                import traceback
                print(f"🔍 Debug - Full error traceback: {traceback.format_exc()}")
                return
        
        print("📈 Creating final summary...")
        summary_df = step_summary_by_standardized_name(df, isin_mapping_df, monthA, monthB)
        
        print("✅ All data processing complete!")
        
        print("📊 Generating top 10 holdings charts...")
        run_chart_generation(book, summary_df, monthA, monthB)
        
        print("📄 Creating output sheets...")
        
        def create_formatted_sheet(sheet_name: str, df: pd.DataFrame, title: str):
            if sheet_name in [s.name for s in book.sheets]:
                try:
                    book.sheets[sheet_name].delete()
                    print(f"🧹 Deleted existing '{sheet_name}' sheet")
                except Exception as e:
                    print(f"⚠️ Could not delete existing sheet '{sheet_name}': {e}")
            
            new_sheet = book.sheets.add(name=sheet_name, after=source_sheet)
            
            new_sheet["A1"].value = title
            new_sheet["A1"].font.bold = True
            new_sheet["A1"].font.size = 14
            new_sheet["A1"].color = '#E6F3FF'
            new_sheet["A1"].font.color = '#000000'
            
            if df is not None and not df.empty:
                new_sheet["A3"].options(index=False).value = df
            
                try:
                    table_range = new_sheet["A3"].resize(df.shape[0] + 1, df.shape[1])
                    new_sheet.tables.add(source=table_range)
                    print(f"✅ Created formatted table in '{sheet_name}'")
                except Exception as e:
                    print(f"⚠️ Could not create table in '{sheet_name}': {e}")
            
                header_range = new_sheet["A3"].resize(1, df.shape[1])
                header_range.color = '#F0F0F0'
                header_range.font.color = '#000000'
                header_range.font.bold = True
            
            return new_sheet
        
        sheets_created = []
        
        monthA_title = monthA.strftime('%d-%m-%Y')
        monthB_title = monthB.strftime('%d-%m-%Y')
        
        if run_mode == "full":
            val_sheet = create_formatted_sheet(
                "VALIDATIONS", 
                validation_df, 
                f"MF Holdings Validation Totals ({monthB_title} vs {monthA_title})"
            )
            
            if validation_df is not None and not validation_df.empty:
                try:
                    for col_idx in range(1, validation_df.shape[1]):
                        col_range = val_sheet[f"{chr(65 + col_idx)}4:{chr(65 + col_idx)}{validation_df.shape[0] + 3}"]
                        col_range.number_format = '#,##0'
                    print("✅ Applied number formatting to VALIDATIONS")
                except Exception as e:
                    print(f"⚠️ Could not apply number formatting: {e}")
            sheets_created.append(val_sheet)
            
            isin_sheet = create_formatted_sheet(
                "ISIN_MAPPING", 
                isin_mapping_df, 
                "ISIN to Standardized Name Mapping"
            )
            sheets_created.append(isin_sheet)
            
        elif run_mode == "stage2":
            print("ℹ️ Stage 2 mode: Reusing existing sheets, only updating Summary Analysis")
            for sheet_name in ["VALIDATIONS", "ISIN_MAPPING"]:
                try:
                    sheets_created.append(book.sheets[sheet_name])
                    print(f"✅ Found existing sheet: {sheet_name} (reusing, not recreating)")
                except (KeyError, ValueError):
                    print(f"⚠️ Could not find existing sheet: {sheet_name}")
        
        summary_sheet = create_formatted_sheet(
            "SUMMARY", 
            summary_df, 
            f"MF Holdings Summary Analysis ({monthB_title} vs {monthA_title})"
        )
        
        if summary_df is not None and not summary_df.empty:
            try:
                percentage_cols = [1, 2, 3, 4]
                for col_idx in percentage_cols:
                    if col_idx < summary_df.shape[1]:
                        col_range = summary_sheet[f"{chr(65 + col_idx)}4:{chr(65 + col_idx)}{summary_df.shape[0] + 3}"]
                        col_range.number_format = '0.0%'
                
                integer_cols = [5, 6, 7, 8]
                for col_idx in integer_cols:
                    if col_idx < summary_df.shape[1]:
                        col_range = summary_sheet[f"{chr(65 + col_idx)}4:{chr(65 + col_idx)}{summary_df.shape[0] + 3}"]
                        col_range.number_format = '0'
                
                big_number_cols = [9, 10, 11, 12]
                for col_idx in big_number_cols:
                    if col_idx < summary_df.shape[1]:
                        col_range = summary_sheet[f"{chr(65 + col_idx)}4:{chr(65 + col_idx)}{summary_df.shape[0] + 3}"]
                        col_range.number_format = '#,##0'
                
                print("✅ Applied number formatting to SUMMARY")
            except Exception as e:
                print(f"⚠️ Could not apply number formatting to summary: {e}")
        
        sheets_created.append(summary_sheet)
        
        # --- Create exportDat sheet ---
        print("📄 Creating exportDat sheet...")
        try:
            export_df = create_export_df(summary_df, monthA, monthB)
            export_sheet_name = "exportDat"

            if export_sheet_name in [s.name for s in book.sheets]:
                book.sheets[export_sheet_name].delete()
            
            # Default to adding after summary_sheet if chart sheet not found
            after_sheet = summary_sheet
            try:
                after_sheet = book.sheets['CHART_TOPHOLD']
            except (KeyError, ValueError):
                print("⚠️ 'CHART_TOPHOLD' not found. Placing 'exportDat' after 'SUMMARY'.")

            export_sheet = book.sheets.add(name=export_sheet_name, after=after_sheet)

            if not export_df.empty:
                # Write data WITH header, but no index
                export_sheet["A1"].options(index=False).value = export_df
                print(f"✅ Wrote {len(export_df)} records to '{export_sheet_name}' sheet (with header).")
            else:
                print(f"⚠️ Export DataFrame was empty. Nothing to write to '{export_sheet_name}'.")
        except Exception as e:
            print(f"❌ ERROR creating exportDat sheet: {e}")
        # --- End of exportDat block ---

        summary_sheet.activate()
        
        print(f"🎉 Successfully processed {len(sheets_created)} analysis sheets!")
        print("✅ MF Holdings Summary Processing Complete!")
        
    except Exception as e:
        print(f"❌ Error during processing: {e}")
        import traceback
        print(f"📋 Full error details: {traceback.format_exc()}")
