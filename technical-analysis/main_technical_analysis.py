import xlwings as xw
import pandas as pd
from finta import TA
import requests
from xlwings import script
from datetime import datetime
import os
import matplotlib.pyplot as plt
import tempfile
import base64
from PIL import Image
import time
import asyncio

# FastAPI CORS Configuration for Excel/xlwings:
# When setting up a FastAPI server to work with xlwings in Excel, use these origins:
#
# from fastapi.middleware.cors import CORSMiddleware
#
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=[
#         "https://addin.xlwings.org",    # Main xlwings add-in domain - THIS IS THE PRIMARY ONE NEEDED
#     ],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )
#
# Note: The actual request comes from the Excel WebView2 browser component via the
# xlwings add-in hosted at addin.xlwings.org, NOT from Excel/Office domains directly.

async def _execute_technicals_flow(book: xw.Book):
    """Internal function to create daily & weekly technicals, and generate a single consolidated chart sheet."""
    print("🔵 STARTING _execute_technicals_flow (Shared Logic)")
    master_sheet = book.sheets["MASTER"]
    
    ticker = str(master_sheet["B4"].value).strip().upper() if master_sheet["B4"].value else None
    daily_start_date_raw = master_sheet["D4"].value
    daily_end_date_raw = master_sheet["E4"].value
    weekly_start_date_raw = master_sheet["D5"].value
    weekly_end_date_raw = master_sheet["E5"].value
    
    if not all([ticker, daily_start_date_raw, daily_end_date_raw, weekly_start_date_raw, weekly_end_date_raw]):
        master_sheet["B8"].value = "Please enter all required parameters in the MASTER sheet."
        return None, None, None, None

    daily_start_date = convert_excel_date(daily_start_date_raw)
    daily_end_date = convert_excel_date(daily_end_date_raw)
    weekly_start_date = convert_excel_date(weekly_start_date_raw)
    weekly_end_date = convert_excel_date(weekly_end_date_raw)

    # --- Process Daily Data ---
    print(f"\n▶ Processing Daily Data for {ticker}")
    daily_api_url = f"https://yfin.hosting.tigzig.com/get-all-prices/?tickers={ticker}&start_date={daily_start_date}&end_date={daily_end_date}"
    daily_response = requests.get(daily_api_url)
    if not daily_response.ok:
        master_sheet["B8"].value = "Daily data service temporarily unavailable."
        return None, None, None, None

    daily_data = daily_response.json()
    if not isinstance(daily_data, dict) or daily_data.get("error"):
        error_msg = daily_data.get("error") if isinstance(daily_data, dict) else "Invalid daily response format"
        master_sheet["B8"].value = f"Daily data error: {error_msg}"
        return None, None, None, None

    daily_rows = [dict(d[ticker], Date=date) for date, d in daily_data.items() if ticker in d]
    daily_df = pd.DataFrame(daily_rows)
    daily_df.columns = [col.lower() for col in daily_df.columns]
    daily_df['date'] = pd.to_datetime(daily_df['date'])
    daily_df = daily_df.sort_values('date')

    daily_display_df = calculate_technical_indicators(daily_df.copy())
    daily_display_df.rename(columns=lambda c: c.upper(), inplace=True)
    
    # Write daily data to sheet
    print("\nWriting daily data to PRICES_DAILY sheet...")
    if "PRICES_DAILY" in [s.name for s in book.sheets]:
        prices_daily_sheet = book.sheets["PRICES_DAILY"]
        prices_daily_sheet.clear()
    else:
        prices_daily_sheet = book.sheets.add(name="PRICES_DAILY", after=master_sheet)
    
    prices_daily_sheet["A1"].value = f"Daily Prices for {ticker} from {pd.to_datetime(daily_start_date).strftime('%d%b%Y')} to {pd.to_datetime(daily_end_date).strftime('%d%b%Y')}"
    prices_daily_sheet["A2"].options(index=False).value = daily_display_df
    prices_daily_sheet.tables.add(source=prices_daily_sheet["A2"].resize(daily_display_df.shape[0] + 1, daily_display_df.shape[1]), name="PricesDaily")
    print("✓ Daily data written successfully.")

    # Create daily chart image
    daily_chart_path = create_chart(daily_display_df, ticker, "Technical Analysis Charts", "Daily")

    # --- Process Weekly Data ---
    print(f"\n▶ Processing Weekly Data for {ticker}")
    weekly_api_url = f"https://yfin.hosting.tigzig.com/get-all-prices/?tickers={ticker}&start_date={weekly_start_date}&end_date={weekly_end_date}"
    weekly_response = requests.get(weekly_api_url)
    if not weekly_response.ok:
        master_sheet["B8"].value = "Weekly data service temporarily unavailable."
        return None, None, None, None

    weekly_data = weekly_response.json()
    if not isinstance(weekly_data, dict) or weekly_data.get("error"):
        error_msg = weekly_data.get("error") if isinstance(weekly_data, dict) else "Invalid weekly response format"
        master_sheet["B8"].value = f"Weekly data error: {error_msg}"
        return None, None, None, None

    weekly_rows = [dict(d[ticker], Date=date) for date, d in weekly_data.items() if ticker in d]
    weekly_df = pd.DataFrame(weekly_rows)
    weekly_df['Date'] = pd.to_datetime(weekly_df['Date'])
    weekly_df = weekly_df.sort_values('Date').resample('W-FRI', on='Date').agg({
        'Open': 'first', 'High': 'max', 'Low': 'min', 'Close': 'last', 'Volume': 'sum'
    }).dropna().reset_index()

    weekly_display_df = calculate_technical_indicators(weekly_df.copy())
    weekly_display_df.rename(columns=lambda c: c.upper(), inplace=True)

    # Write weekly data to sheet
    print("\nWriting weekly data to PRICES_WEEKLY sheet...")
    if "PRICES_WEEKLY" in [s.name for s in book.sheets]:
        prices_weekly_sheet = book.sheets["PRICES_WEEKLY"]
        prices_weekly_sheet.clear()
    else:
        prices_weekly_sheet = book.sheets.add(name="PRICES_WEEKLY", after=prices_daily_sheet)

    prices_weekly_sheet["A1"].value = f"Weekly Prices for {ticker} from {pd.to_datetime(weekly_start_date).strftime('%d%b%Y')} to {pd.to_datetime(weekly_end_date).strftime('%d%b%Y')}"
    prices_weekly_sheet["A2"].options(index=False).value = weekly_display_df
    prices_weekly_sheet.tables.add(source=prices_weekly_sheet["A2"].resize(weekly_display_df.shape[0] + 1, weekly_display_df.shape[1]), name="PricesWeekly")
    print("✓ Weekly data written successfully.")

    # Create weekly chart image
    weekly_chart_path = create_chart(weekly_display_df, ticker, "Technical Analysis Charts", "Weekly")

    # --- Consolidate Charts ---
    if not daily_chart_path or not weekly_chart_path:
        master_sheet["B8"].value = "Chart creation failed. Cannot consolidate."
        return None, None, None, None

    print("\n▶ Consolidating charts into a single sheet...")
    if "CHARTS" in [s.name for s in book.sheets]:
        charts_sheet = book.sheets["CHARTS"]
        charts_sheet.clear()
    else:
        charts_sheet = book.sheets.add(name="CHARTS", after=prices_weekly_sheet)

    try:
        for pic in list(charts_sheet.pictures):
            if pic.name in ["DailyChart", "WeeklyChart"]:
                pic.delete()
    except Exception as e:
        print(f"Warning: Could not delete existing pictures: {e}")

    temp_dir = tempfile.gettempdir()
    for path, name, anchor in [(daily_chart_path, "DailyChart", "A3"), (weekly_chart_path, "WeeklyChart", "J3")]:
        try:
            img = Image.open(path)
            resized_path = os.path.join(temp_dir, f"resized_{os.path.basename(path)}")
            img.resize((int(img.width * 0.5), int(img.height * 0.5)), Image.LANCZOS).save(resized_path)
            charts_sheet.pictures.add(resized_path, name=name, update=True, anchor=charts_sheet[anchor])
            print(f"✓ Placed {name} onto CHARTS sheet.")
        except Exception as e:
            print(f"❌ ERROR placing {name}: {e}")
            master_sheet["B8"].value = f"Error placing {name}: {e}"
            return None, None, None, None

    charts_sheet.activate()
    print("\n🟢 SUCCESS: Technical analysis complete. Final charts are on the 'CHARTS' sheet.")
    print("🟢 ENDING _execute_technicals_flow (Shared Logic)")
    
    return daily_display_df, weekly_display_df, daily_chart_path, weekly_chart_path


@script(button="[btn_tech]MASTER!H4", show_taskpane=True)
async def create_technicals(book: xw.Book):
    """Create daily & weekly technicals, and generate a single consolidated chart sheet."""
    await _execute_technicals_flow(book)

def combine_charts(daily_path, weekly_path, daily_start, daily_end, weekly_start, weekly_end):
    """Combine daily and weekly charts into a single side-by-side image for PDF output."""
    # Read the images
    daily_img = plt.imread(daily_path)
    weekly_img = plt.imread(weekly_path)
    
    # Convert Excel dates to readable format
    try:
        def convert_excel_date(date_val):
            if isinstance(date_val, (datetime, pd.Timestamp)):
                return date_val
            # Convert Excel serial number to datetime
            return pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(date_val))
        
        # Convert dates for daily chart
        daily_start_date = convert_excel_date(daily_start)
        daily_end_date = convert_excel_date(daily_end)
        
        # Convert dates for weekly chart
        weekly_start_date = convert_excel_date(weekly_start)
        weekly_end_date = convert_excel_date(weekly_end)
        
        # Format dates for display
        daily_start_str = daily_start_date.strftime('%d %b %Y')
        daily_end_str = daily_end_date.strftime('%d %b %Y')
        weekly_start_str = weekly_start_date.strftime('%d %b %Y')
        weekly_end_str = weekly_end_date.strftime('%d %b %Y')
        
        print("\nDate Ranges:")
        print(f"Daily: {daily_start_str} to {daily_end_str}")
        print(f"Weekly: {weekly_start_str} to {weekly_end_str}")
        
    except Exception as e:
        print(f"[ERROR] Error processing dates: {str(e)}")
        return None
    
    # Create a new figure with appropriate size
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(24, 12))
    
    # Display images
    ax1.imshow(daily_img)
    ax2.imshow(weekly_img)
    
    # Remove axes
    ax1.axis('off')
    ax2.axis('off')
    
    # Add titles with date ranges on single line
    ax1.set_title(f'Daily Chart ({daily_start_str} to {daily_end_str})', fontsize=14, fontweight='bold', pad=10)
    ax2.set_title(f'Weekly Chart ({weekly_start_str} to {weekly_end_str})', fontsize=14, fontweight='bold', pad=10)
    
    # Adjust layout
    plt.tight_layout()
    
    # Save combined figure
    temp_dir = tempfile.gettempdir()
    combined_path = os.path.join(temp_dir, "combined_technical_chart.png")
    fig.savefig(combined_path, dpi=150, bbox_inches='tight')
    plt.close(fig)
    
    return combined_path

@script(button="[btn_ai]MASTER!H8", show_taskpane=True)
async def get_technical_analysis_from_gemini(book: xw.Book):
    """Get technical analysis from Gemini API using separate charts for analysis but combined chart for PDF."""
    print("🔵 STARTING get_technical_analysis_from_gemini")

    # Run the shared technicals flow to get data, sheets, and chart paths.
    daily_data, weekly_data, daily_chart_path, weekly_chart_path = await _execute_technicals_flow(book)

    # Check if the technicals flow was successful. If not, the data will be None.
    if daily_data is None:
        print("❌ ERROR: The prerequisite technical analysis flow failed. Check logs above. Aborting Gemini analysis.")
        # The helper function already wrote a specific error to the MASTER sheet, so we just need to stop.
        return
    
    print(f"✓ Technicals flow successful. Received dataframes and chart paths.")
    
    # Debug flag - set to False to prevent HTML printing in logs
    DEBUG_HTML = False
    
    # Get the required sheets
    master_sheet = book.sheets["MASTER"]
    
    # Clear previous results
    master_sheet["A11:B12"].clear_contents()
    
    # Read model name from MASTER sheet
    model_name = master_sheet["B7"].value
    
    # Attempt to get API key from environment variable first
    api_key = os.getenv("GEMINI_API_KEY")
    
    # If environment variable is not set or is empty, fall back to Excel cell
    if not api_key:
        api_key = master_sheet["B8"].value
        print("⚠️ Using API key from Excel cell (B8) as environment variable 'GEMINI_API_KEY' was not found or was empty.")
    else:
        print("✅ Using API key from environment variable 'GEMINI_API_KEY'.")
    
    # Get ticker and date parameters from MASTER sheet
    ticker = str(master_sheet["B4"].value).strip().upper() if master_sheet["B4"].value else None
    daily_start_date = master_sheet["D4"].value
    daily_end_date = master_sheet["E4"].value
    weekly_start_date = master_sheet["D5"].value
    weekly_end_date = master_sheet["E5"].value
    
    if not all([model_name, api_key, ticker, daily_start_date, daily_end_date, weekly_start_date, weekly_end_date]):
        master_sheet["A11"].value = "Error"
        master_sheet["A12"].value = "Please enter all required parameters (model name, API key, ticker, start dates, end dates)"
        return
    
    # The chart paths and data are now passed directly from the helper function.
    print(f"Using daily chart at path: {daily_chart_path}")
    print(f"Using weekly chart at path: {weekly_chart_path}")
    
    # Verify the paths are different
    if daily_chart_path == weekly_chart_path:
        master_sheet["A11"].value = "Error"
        master_sheet["A12"].value = "Error: Daily and weekly chart paths are the same"
        return
    
    # Upload both charts separately for Gemini
    try:
        print("\nUploading charts to server...")
        
        # Upload daily chart
        daily_files = {
            'file': ('daily_chart.png', open(daily_chart_path, 'rb'), 'image/png')
        }
        daily_upload_response = requests.post(
            "https://mdtopdf.hosting.tigzig.com/api/upload-image",
            files=daily_files
        )
        
        if not daily_upload_response.ok:
            error_msg = f"Failed to upload daily image: {daily_upload_response.status_code}"
            print(f"ERROR: {error_msg}")
            master_sheet["A11"].value = "Error"
            master_sheet["A12"].value = error_msg
            return
        
        daily_upload_data = daily_upload_response.json()
        daily_image_path = daily_upload_data['image_path']
        print(f"Daily image uploaded successfully. Path: {daily_image_path}")
        
        # Upload weekly chart
        weekly_files = {
            'file': ('weekly_chart.png', open(weekly_chart_path, 'rb'), 'image/png')
        }
        weekly_upload_response = requests.post(
            "https://mdtopdf.hosting.tigzig.com/api/upload-image",
            files=weekly_files
        )
        
        if not weekly_upload_response.ok:
            error_msg = f"Failed to upload weekly image: {weekly_upload_response.status_code}"
            print(f"ERROR: {error_msg}")
            master_sheet["A11"].value = "Error"
            master_sheet["A12"].value = error_msg
            return
        
        weekly_upload_data = weekly_upload_response.json()
        weekly_image_path = weekly_upload_data['image_path']
        print(f"Weekly image uploaded successfully. Path: {weekly_image_path}")
        
        # Create combined chart with date ranges from original parameters
        print("\nCreating combined chart for PDF...")
        combined_chart_path = combine_charts(daily_chart_path, weekly_chart_path, 
                                          daily_start=daily_start_date, daily_end=daily_end_date,
                                          weekly_start=weekly_start_date, weekly_end=weekly_end_date)
        
        # Upload combined chart
        combined_files = {
            'file': ('combined_chart.png', open(combined_chart_path, 'rb'), 'image/png')
        }
        combined_upload_response = requests.post(
            "https://mdtopdf.hosting.tigzig.com/api/upload-image",
            files=combined_files
        )
        
        if not combined_upload_response.ok:
            error_msg = f"Failed to upload combined image: {combined_upload_response.status_code}"
            print(f"ERROR: {error_msg}")
            master_sheet["A11"].value = "Error"
            master_sheet["A12"].value = error_msg
            return
        
        combined_upload_data = combined_upload_response.json()
        combined_image_path = combined_upload_data['image_path']
        print(f"Combined image uploaded successfully. Path: {combined_image_path}")
        
    except Exception as e:
        error_msg = f"Error uploading images: {str(e)}"
        print(f"ERROR: {error_msg}")
        master_sheet["A11"].value = "Error"
        master_sheet["A12"].value = error_msg
        return
    
    # Data is now passed in directly from the helper function.
    # No need to read from sheets.

    # Get the latest data points
    latest_daily = daily_data.iloc[-1]
    latest_weekly = weekly_data.iloc[-1]
    
    # Get last 20 rows for additional data to send to Gemini (keeping all columns)
    last_20_days = daily_data.tail(20)
    last_20_weeks = weekly_data.tail(20)
    
    # Format the last 20 rows data as markdown tables for Gemini analysis
    # This is separate from the HTML table that's used for display
    def format_data_for_analysis(df, title):
        # Convert DataFrame to markdown table string with clear header
        header = f"### {title} (Last 20 rows)\n"
        # Make sure dates are formatted nicely
        df_copy = df.copy()
        if 'DATE' in df_copy.columns:
            df_copy['DATE'] = pd.to_datetime(df_copy['DATE']).dt.strftime('%Y-%m-%d')
        
        # Create markdown table rows
        rows = []
        # Header row
        rows.append("| " + " | ".join(str(col) for col in df_copy.columns) + " |")
        # Separator row
        rows.append("| " + " | ".join(["---"] * len(df_copy.columns)) + " |")
        # Data rows
        for _, row in df_copy.iterrows():
            formatted_row = []
            for val in row:
                if isinstance(val, (int, float)):
                    # Format numbers with 2 decimal places
                    formatted_row.append(f"{val:.2f}" if isinstance(val, float) else str(val))
                else:
                    formatted_row.append(str(val))
            rows.append("| " + " | ".join(formatted_row) + " |")
        
        return header + "\n".join(rows)
    
    # Create formatted data tables for Gemini analysis
    daily_data_for_analysis = format_data_for_analysis(last_20_days, "Daily Price & Technical Data")
    weekly_data_for_analysis = format_data_for_analysis(last_20_weeks, "Weekly Price & Technical Data")
    
    # Create tables with last 5 days of data for both daily and weekly
    last_5_days = daily_data.tail(5)[['DATE', 'CLOSE', 'EMA_26', 'ROC_14', 'RSI_14']]
    last_5_weeks = weekly_data.tail(5)[['DATE', 'CLOSE', 'EMA_26', 'ROC_14', 'RSI_14']]
    
    # Create a version of the HTML table that's less likely to leak
    # by avoiding direct concatenation in f-strings
    table_html_parts = []
    
    # Add opening wrapper div
    table_html_parts.append('<div style="display: flex; justify-content: space-between;">')
    
    # Daily table - construct part by part
    table_html_parts.append('<div style="width: 48%; display: inline-block;">')
    table_html_parts.append('<table style="border-collapse: collapse; width: 100%; font-size: 7pt;">')
    table_html_parts.append('<thead><tr>')
    
    # Headers - separate each header to avoid DAILY and CLOSE leaking
    headers = ["DAILY", "CLOSE", "EMA-26", "ROC", "RSI"]
    for header in headers:
        table_html_parts.append(f'<th style="border: 0.25pt solid #000; padding: 2pt; text-align: center;">{header}</th>')
    
    table_html_parts.append('</tr></thead><tbody>')
    
    # Add daily rows
    for _, row in last_5_days.iterrows():
        date = pd.to_datetime(row['DATE'])
        date_str = date.strftime('%d-%b')
        table_html_parts.append('<tr>')
        table_html_parts.append(f'<td style="border: 0.25pt solid #000; padding: 2pt; text-align: center;">{date_str}</td>')
        table_html_parts.append(f'<td style="border: 0.25pt solid #000; padding: 2pt; text-align: right;">{row["CLOSE"]:.1f}</td>')
        table_html_parts.append(f'<td style="border: 0.25pt solid #000; padding: 2pt; text-align: right;">{row["EMA_26"]:.1f}</td>')
        table_html_parts.append(f'<td style="border: 0.25pt solid #000; padding: 2pt; text-align: right;">{row["ROC_14"]:.1f}</td>')
        table_html_parts.append(f'<td style="border: 0.25pt solid #000; padding: 2pt; text-align: right;">{int(row["RSI_14"])}</td>')
        table_html_parts.append('</tr>')
    
    table_html_parts.append('</tbody></table></div>')
    
    # Weekly table - construct part by part
    table_html_parts.append('<div style="width: 48%; display: inline-block;">')
    table_html_parts.append('<table style="border-collapse: collapse; width: 100%; font-size: 7pt;">')
    table_html_parts.append('<thead><tr>')
    
    # Headers
    headers = ["WEEKLY", "CLOSE", "EMA-26", "ROC", "RSI"]
    for header in headers:
        table_html_parts.append(f'<th style="border: 0.25pt solid #000; padding: 2pt; text-align: center;">{header}</th>')
    
    table_html_parts.append('</tr></thead><tbody>')
    
    # Add weekly rows
    for _, row in last_5_weeks.iterrows():
        date = pd.to_datetime(row['DATE'])
        date_str = date.strftime('%d-%b')
        table_html_parts.append('<tr>')
        table_html_parts.append(f'<td style="border: 0.25pt solid #000; padding: 2pt; text-align: center;">{date_str}</td>')
        table_html_parts.append(f'<td style="border: 0.25pt solid #000; padding: 2pt; text-align: right;">{row["CLOSE"]:.1f}</td>')
        table_html_parts.append(f'<td style="border: 0.25pt solid #000; padding: 2pt; text-align: right;">{row["EMA_26"]:.1f}</td>')
        table_html_parts.append(f'<td style="border: 0.25pt solid #000; padding: 2pt; text-align: right;">{row["ROC_14"]:.1f}</td>')
        table_html_parts.append(f'<td style="border: 0.25pt solid #000; padding: 2pt; text-align: right;">{int(row["RSI_14"])}</td>')
        table_html_parts.append('</tr>')
    
    table_html_parts.append('</tbody></table></div>')
    
    # Close the wrapper div
    table_html_parts.append('</div>')
    
    # Join all parts only when needed for the API call, not for printing
    table_section = ''.join(table_html_parts)
    
    # Convert both charts to base64 for Gemini API
    try:
        with open(daily_chart_path, "rb") as daily_file:
            daily_chart_base64 = base64.b64encode(daily_file.read()).decode('utf-8')
        with open(weekly_chart_path, "rb") as weekly_file:
            weekly_chart_base64 = base64.b64encode(weekly_file.read()).decode('utf-8')
    except Exception as e:
        print(f"[ERROR] Error reading chart images: {str(e)}")
        master_sheet["A11"].value = "Error"
        master_sheet["A12"].value = "Error reading chart images"
        return
    
    # Build the prompt parts
    prompt_parts = []
    prompt_parts.append("""
    [SYSTEM INSTRUCTIONS]
    Your analysis is for professional use, so prioritize clarity, precision, and actionable insights. You will receive two types of data:

    1. REPORT STRUCTURE DATA: Pre-formatted HTML tables showing the last 5 rows of data
       - These tables are part of the final report structure
       - They MUST be preserved exactly as provided
       - They appear right after the chart image

    2. REFERENCE DATA: Additional 20 rows of data in markdown format
       - This data is PROVIDED ONLY FOR YOUR ANALYSIS
       - DO NOT include this data in the final report
       - Use it to inform your analysis in sections 1-6

    **CRITICAL: REQUIRED REPORT STRUCTURE**
    The final report must follow this exact structure - no additions or modifications:

        # Integrated Technical Analysis
        ## [TICKER_SYMBOL]
        ## Daily and Weekly Charts
        ![Combined Technical Analysis](charts/[CHART_FILENAME])
        [PRESERVE EXISTING HTML TABLES HERE - DO NOT MODIFY]
        ### 1. Price Action and Trend Analysis
        **Daily:** [your analysis]
        **Weekly:** [your analysis]
        **Confirmation/Divergence:** [your analysis]
        [CONTINUE WITH SECTIONS 2-6 AS SPECIFIED]

    **MANDATORY FORMATTING RULES**
    1. Keep the report structure exactly as shown above
    2. DO NOT add any new sections or data tables
    3. DO NOT modify or remove existing HTML tables
    4. Use markdown only for your analysis in sections 1-6
    5. The 20-row reference data tables MUST NOT appear in the final report
    6. Keep exactly one blank line between sections

    **ANALYSIS REQUIREMENTS**
    - Use the 20-row reference data to inform your analysis
    - Write your analysis ONLY in sections 1-6
    - Keep analysis concise and actionable
    - Focus on technical insights and patterns

    **WORD COUNT LIMITS**
    - Follow the word count limits specified in each section
    - Focus on actionable insights
    - No generic statements

    Remember: The 20-row data tables are for your reference ONLY. They should NOT appear in the final report structure.
    """)
    
    prompt_parts.append(f"# {ticker}")
    prompt_parts.append("## Daily and Weekly Charts")
    
    prompt_parts.append(f"\n![Combined Technical Analysis](charts/{combined_image_path})")
    
    # Insert the table HTML
    prompt_parts.append("\n" + table_section)
    
    # Add a page break before the analysis sections
    prompt_parts.append('<div style="page-break-after: always;"></div>')
    
    # Continue with the rest of the prompt
    prompt_parts.append("""
    ### 1. Price Action and Trend Analysis
    Apply your comprehensive knowledge; **the examples provided below are illustrative, not exhaustive.** Word count limit: 200-250 words

    **Daily:** [Analyze the daily trend's character (e.g., strong, maturing, range-bound). Identify the current phase (e.g., impulse wave, corrective pullback, consolidation). Describe the recent sequence of highs and lows.]

    **Weekly:** [Analyze the primary trend on the weekly chart. e.g Is it well-established? Is it showing signs of acceleration or deceleration? etc. Place the recent daily action into the context of this longer-term trend.]

    **Confirmation/Divergence:** [Synthesize the timeframes. e.g Is the daily action a simple pause (consolidation) in a strong weekly uptrend? Or are there early warnings on the daily chart (e.g., a lower high) that challenge the weekly trend? etc ]



    ### 2. Support and Resistance Levels
    Apply your comprehensive knowledge; **the examples provided below are illustrative, not exhaustive.** Word count limit: 200-250 words

    **Daily Levels:**
    - **Horizontal S/R:** [Identify key horizontal support (e.g previous swing lows, consolidation bottoms) and resistance (e.g previous swing highs, consolidation tops) levels. Be specific with price zones.]
    - **Dynamic S/R:** [Analyze the EMAs and Bollinger Bands as areas of potential dynamic support or resistance.]

    **Weekly Levels:**
    - **Horizontal S/R:** [Identify the most significant, long-term horizontal support and resistance zones from the weekly chart.]
    - **Dynamic S/R:** [Analyze the weekly EMAs and Bollinger Bands as major trend-following support levels.]

    **Level Alignment:** [Discuss the interaction of levels. e.g Is a key daily resistance level just below a major weekly resistance? Does a daily support level coincide with the weekly etc] 


    ### 3. Technical Indicator Analysis
    Apply your comprehensive knowledge; **the examples provided below are illustrative, not exhaustive.** Word count limit: 200-250 words

    **Daily Indicators:**
    - **EMAs (12 & 26):** [Analyze the crossover status, the spread between the EMAs (indicating momentum), and their role as dynamic support/resistance etc]
    - **MACD:** [Analyze the MACD line vs. the signal line, its position relative to the zero line, and the momentum trajectory shown by the histogram. Look for divergences with price.]
    - **RSI & ROC:** [Analyze RSI levels (overbought/oversold context, support/resistance flips ) , ROC for momentum speed. Crucially, identify any **bullish or bearish divergences** against recent price highs/lows.]
    - **Bollinger Bands:** [Analyze the price's position relative to the bands. e.g Is it "walking the band" (strong trend)? Note the width of the bands – are they contracting (Bollinger Squeeze, indicating potential for a volatile move) or expanding?]

    **Weekly Indicators:**
    - **EMAs (12 & 26):** [Same this as Daily but now with Weekly Chart]
    - **MACD:** [Same this as Daily but now with Weekly Chart]
    - **RSI & ROC:** [Same this as Daily but now with Weekly Chart]


    ### 4. Pattern Recognition
    Apply your comprehensive knowledge; **the examples provided below are illustrative, not exhaustive.** Word count limit: 200-250 words

    **Daily Patterns:** [Identify classic chart patterns (e.g., triangles, double tops, rising tops and bottoms, flags, pennants, channels, wedges etc)]

    **Weekly Patterns:** [Identify larger, multi-month patterns on the weekly chart. Note the overall market structure.]

    **Pattern Alignment:** [How do the daily and weekly patter align? e.g Does the daily pattern (e.g., a bull flag) fit within the context of the larger weekly uptrend? This alignment provides a higher-probability trade setup.]

    ### 5. Volume Analysis
    Apply your comprehensive knowledge; **the examples provided below are illustrative, not exhaustive.** Word count limit: 200-250 words

    **Daily Volume:** [Analyze the volume trend. Correlate volume with price action. Is volume increasing on up-days and decreasing on down-days (bullish confirmation)? Note any high-volume spikes and sharp drpos and what they signify (e.g., capitulation, breakout etc).]

    **Weekly Volume:** [Analyze the weekly volume bars in context. Does volume confirm the primary trend? e.g Is there a significant drop-off in volume that suggests waning conviction?]

    **Volume Trends:** [Summarize the volume picture. Is participation generally increasing or decreasing, and what does this imply for the sustainability of the current trend?]


    ### 6. Technical Outlook
    Apply your comprehensive knowledge; **the examples provided below are illustrative, not exhaustive.** This is the most important section. Synthesize all the above points into a coherent thesis. Word count limit:250-300 words

    **Primary Scenario (Base Case):** [Based on the weight of the evidence, describe the most likely path for the price in the short-to-medium term and why you think so. Mention specific price targets for this scenario. Don't just share the scenaior and price targets, essential to share the reasoning behind the scenario synthesizing the detailed analysis of the charts and data tables. .]

    - **Confirmation:** [e.g what specific price action (e.g., "a decisive daily close above the [X] resistance on high volume") would confirm the primary bullish hypthesis?]

    - **Invalidation:** [e.g What specific price action (e.g., "a break below the [Y] support and the 26-day EMA") would invalidate the primary bullish thesis and suggest a deeper correction towards the next support at [Z]?]
    """)
    
    # Add the technical data
    prompt_parts.append(f"""
    Current Technical Data:
    **Daily Data**:
    - Close: {latest_daily['CLOSE']} | EMA_12: {latest_daily['EMA_12']:.2f} | EMA_26: {latest_daily['EMA_26']:.2f}
    - MACD: {latest_daily['MACD_12_26']:.2f} | Signal: {latest_daily['MACD_SIGNAL_9']:.2f}
    - RSI: {latest_daily['RSI_14']:.2f} | BB Upper: {latest_daily['BBANDS_UPPER_20_2']:.2f} | BB Lower: {latest_daily['BBANDS_LOWER_20_2']:.2f}

    **Weekly Data**:
    - Close: {latest_weekly['CLOSE']} | EMA_12: {latest_weekly['EMA_12']:.2f} | EMA_26: {latest_weekly['EMA_26']:.2f}
    - MACD: {latest_weekly['MACD_12_26']:.2f} | Signal: {latest_weekly['MACD_SIGNAL_9']:.2f}
    - RSI: {latest_weekly['RSI_14']:.2f} | BB Upper: {latest_weekly['BBANDS_UPPER_20_2']:.2f} | BB Lower: {latest_weekly['BBANDS_LOWER_20_2']:.2f}
    
    Below you will find the last 20 rows of data for both daily and weekly timeframes. These are provided as supporting information for your chart analysis. Note the different date patterns to distinguish daily from weekly data:
    - Daily data: Consecutive trading days
    - Weekly data: Weekly intervals, typically Friday closing prices
    """)
    
    prompt_parts.append(daily_data_for_analysis)
    prompt_parts.append(weekly_data_for_analysis)
    
    prompt_parts.append("""
    IMPORTANT:
    1. Follow the EXACT markdown structure and formatting shown above
    2. Use bold (**) for timeframe headers as shown
    3. Maintain consistent section ordering
    4. Ensure each section has Daily, Weekly, and Confirmation/Alignment analysis
    5. Keep the analysis concise but comprehensive
    6. Focus primarily on chart analysis, using the data tables as supporting information only
    7. Analyze the complete timeframe shown in the charts, not just the last 20 rows of data
    """)
    
    # Join the prompt parts
    prompt = ''.join(prompt_parts)
    
    print("\nSending request to Gemini API...")
    
    # Prepare API payload with both full-size images and text
    payload = {
        "contents": [{
            "parts": [
                {
                    "inline_data": {
                        "mime_type": "image/png",
                        "data": daily_chart_base64
                    }
                },
                {
                    "inline_data": {
                        "mime_type": "image/png",
                        "data": weekly_chart_base64
                    }
                },
                {
                    "text": prompt
                }
            ]
        }],
        "generationConfig": {
            "temperature": 0.7,
            "maxOutputTokens": 7500
        }
    }
    
    # Make API call to Gemini
    api_url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
    
    try:
        print("\nCalling Gemini API...")
        
        # When logging, don't print the actual payload content
        if DEBUG_HTML:
            print(f"API URL: {api_url}")
            print("Payload includes: 2 images and prompt text")
        
        response = requests.post(
            api_url,
            headers={"Content-Type": "application/json"},
            json=payload
        )
        
        if response.status_code == 200:
            response_json = response.json()
            if 'candidates' in response_json:
                analysis = response_json['candidates'][0]['content']['parts'][0]['text']
                
                print("\nRaw Gemini Response (first 500 chars):")
                print("=" * 80)
                print(analysis[:500])
                print("=" * 80)
                
                # Add disclaimer
                disclaimer_note = """
                
                #### Important Disclaimer

This report is generated using AI based on live technical data and is provided for informational
purposes only. It is not investment advice, investment analysis, or formal research. While care has
been taken to ensure accuracy, outputs should be verified and are intended to support- not replace
sound human judgment. The tool also demonstrates an automated pipeline from data pull and transformations
				to final formatted report.

"""
                final_markdown = f"{disclaimer_note}{analysis}"
                
                # Convert to PDF and save URL
                try:
                    pdf_api_url = "https://mdtopdf.hosting.tigzig.com/text-input"
                    print("\nConverting markdown to PDF...")
                    print(f"Using combined image for PDF: {combined_image_path}")
                    
                    pdf_response = requests.post(
                        pdf_api_url,
                        headers={"Content-Type": "application/json", "Accept": "application/json"},
                        json={"text": final_markdown, "image_path": combined_image_path}
                    )
                    
                    print(f"Status Code: {pdf_response.status_code}")
                    # Don't print raw response text as it could contain HTML
                    if DEBUG_HTML:
                        print(f"Response Text: {pdf_response.text}")
                    else:
                        print("Response received. URLs processed.")
                    
                    response_data = pdf_response.json()
                    
                    # Clear only specific cells where we'll update content
                    master_sheet["A11"].clear_contents()  # PDF Report URL header
                    master_sheet["A12"].clear_contents()  # PDF URL
                    master_sheet["A16"].clear_contents()  # HTML Report URL header
                    master_sheet["A17"].clear_contents()  # HTML URL
                    
                    # Handle PDF URL - don't touch the Open Link cell
                    master_sheet["A11"].value = "PDF Report URL"
                    master_sheet["A12"].value = response_data["pdf_url"]
                    master_sheet["A12"].add_hyperlink(response_data["pdf_url"])
                    
                    # Handle HTML URL - don't touch the Open Link cell
                    master_sheet["A16"].value = "HTML Report URL"
                    master_sheet["A17"].value = response_data["html_url"]
                    master_sheet["A17"].add_hyperlink(response_data["html_url"])
                    
                    print("🟢 URLs saved successfully!")
                except Exception as e:
                    print(f"[ERROR] Error processing response: {str(e)}")
            else:
                master_sheet["A11"].value = "Error"
                master_sheet["A12"].value = "No analysis generated"
        else:
            master_sheet["A11"].value = "Error"
            master_sheet["A12"].value = f"API call failed: {response.status_code}"
    except Exception as e:
        master_sheet["A11"].value = "Error"
        master_sheet["A12"].value = f"Error: {str(e)}"
    
    print("🟢 ENDING get_technical_analysis_from_gemini")

def create_chart(df, ticker, title, frequency):
    """Create a chart as an image file and return its path."""
    print(f"\nCreating {frequency} chart image using matplotlib...")
    
    # Create matplotlib figure with three subplots
    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(8, 8), 
                                       height_ratios=[2, 1, 1], 
                                       sharex=True, 
                                       gridspec_kw={'hspace': 0})
    
    # Create a twin axis for volume
    ax1v = ax1.twinx()
    
    # Plot on the first subplot (price chart)
    ax1.plot(df['DATE'], df['CLOSE'], label='Close Price', color='black', linewidth=1.5, alpha=0.7)
    ax1.plot(df['DATE'], df['BBANDS_UPPER_20_2'], label='BB Upper', color='gray', linestyle='--', linewidth=1, alpha=0.7)
    ax1.plot(df['DATE'], df['BBANDS_MIDDLE_20_2'], label='BB Middle', color='gray', linestyle=':', linewidth=1, alpha=0.7)
    ax1.plot(df['DATE'], df['BBANDS_LOWER_20_2'], label='BB Lower', color='gray', linestyle='--', linewidth=1, alpha=0.7)
    ax1.plot(df['DATE'], df['EMA_12'], label='EMA-12', color='blue', linewidth=2)
    ax1.plot(df['DATE'], df['EMA_26'], label='EMA-26', color='red', linewidth=2)
    
    # Add volume bars with improved scaling
    df['price_change'] = df['CLOSE'].diff()
    volume_colors = ['#26A69A' if val >= 0 else '#EF5350' for val in df['price_change']]
    bar_width = (df['DATE'].iloc[-1] - df['DATE'].iloc[0]).days / len(df) * 0.8
    price_range = df['CLOSE'].max() - df['CLOSE'].min()
    volume_scale_factor = price_range * 0.2 / df['VOLUME'].max()
    normalized_volume = df['VOLUME'] * volume_scale_factor
    ax1v.bar(df['DATE'], normalized_volume, width=bar_width, color=volume_colors, alpha=0.3)
    ax1v.set_ylabel('Volume', fontsize=10, color='gray')
    ax1v.set_yticklabels([])
    ax1v.tick_params(axis='y', length=0)
    ax1v.set_ylim(0, price_range * 0.3)
    
    ax1.set_title(f"{ticker} - Price with EMAs and Bollinger Bands ({frequency})", fontsize=14, fontweight='bold', pad=10, loc='center')
    ax1.set_ylabel('Price', fontsize=12)
    ax1.legend(loc='upper left', fontsize=10)
    ax1.grid(True, alpha=0.2)
    ax1.set_xticklabels([])
    
    # Plot on the second subplot (MACD)
    macd_hist = df['MACD_12_26'] - df['MACD_SIGNAL_9']
    colors = ['#26A69A' if val >= 0 else '#EF5350' for val in macd_hist]
    bar_width = (df['DATE'].iloc[-1] - df['DATE'].iloc[0]).days / len(df) * 0.8
    ax2.bar(df['DATE'], macd_hist, color=colors, alpha=0.85, label='MACD Histogram', width=bar_width)
    ax2.plot(df['DATE'], df['MACD_12_26'], label='MACD', color='#2962FF', linewidth=1.5)
    ax2.plot(df['DATE'], df['MACD_SIGNAL_9'], label='Signal', color='#FF6D00', linewidth=1.5)
    ax2.axhline(y=0, color='gray', linestyle='-', linewidth=0.8, alpha=0.3)
    ax2.set_title(f'MACD (12,26,9) - {frequency}', fontsize=12, fontweight='bold', loc='center')
    ax2.set_ylabel('MACD', fontsize=12)
    ax2.legend(loc='upper left', fontsize=10)
    ax2.grid(True, alpha=0.2)
    ax2.set_xticklabels([])
    
    # Plot on the third subplot (RSI and ROC)
    ax3.plot(df['DATE'], df['RSI_14'], label='RSI (14)', color='#2962FF', linewidth=1.5)
    ax3_twin = ax3.twinx()
    ax3_twin.plot(df['DATE'], df['ROC_14'], label='ROC (14)', color='#FF6D00', linewidth=1.5)
    ax3.axhline(y=70, color='#EF5350', linestyle='--', linewidth=0.8, alpha=0.3)
    ax3.axhline(y=30, color='#26A69A', linestyle='--', linewidth=0.8, alpha=0.3)
    ax3.axhline(y=50, color='gray', linestyle='-', linewidth=0.8, alpha=0.2)
    ax3_twin.axhline(y=0, color='gray', linestyle='-', linewidth=0.8, alpha=0.3)
    ax3.set_ylim(0, 100)
    ax3.set_title(f'RSI & ROC - {frequency}', fontsize=12, fontweight='bold', loc='center')
    ax3.set_ylabel('RSI', fontsize=12, color='#2962FF')
    ax3_twin.set_ylabel('ROC', fontsize=12, color='#FF6D00')
    ax3.tick_params(axis='y', labelcolor='#2962FF')
    ax3_twin.tick_params(axis='y', labelcolor='#FF6D00')
    lines1, labels1 = ax3.get_legend_handles_labels()
    lines2, labels2 = ax3_twin.get_legend_handles_labels()
    ax3.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=10)
    ax3.grid(True, alpha=0.2)
    
    # Format x-axis dates
    first_date = df['DATE'].iloc[0]
    last_date = df['DATE'].iloc[-1]
    date_range = last_date - first_date
    num_ticks = min(8, len(df)) if date_range.days <= 30 else 8 if date_range.days <= 90 else 10
    tick_indices = [0] + list(range(len(df) // (num_ticks - 2), len(df) - 1, len(df) // (num_ticks - 2)))[:num_ticks-2] + [len(df) - 1]
    ax3.set_xticks([df['DATE'].iloc[i] for i in tick_indices])
    date_format = '%Y-%m-%d' if date_range.days > 30 else '%m-%d'
    tick_labels = [df['DATE'].iloc[i].strftime(date_format) for i in tick_indices]
    ax3.set_xticklabels(tick_labels, rotation=45, ha='right')
    
    plt.tight_layout()
    
    # Save compound figure to temporary file and return the path
    temp_dir = tempfile.gettempdir()
    chart_filename = f"{ticker}_{frequency.lower()}_technical_chart.png"
    temp_path = os.path.join(temp_dir, chart_filename)
    fig.savefig(temp_path, dpi=150, bbox_inches='tight')
    plt.close(fig)
    
    print(f"{frequency} chart image created at: {temp_path}")
    return temp_path

def convert_excel_date(excel_date):
    """Convert Excel date to YYYY-MM-DD format."""
    if isinstance(excel_date, (datetime, pd.Timestamp)):
        return excel_date.strftime("%Y-%m-%d")
    else:
        # Convert Excel serial number to datetime
        date = pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(excel_date))
        return date.strftime("%Y-%m-%d")


def calculate_technical_indicators(df):
    """Calculate technical indicators for a DataFrame."""
    print("\nCalculating indicators...")
    
    # 1. EMA - Exponential Moving Average (12 days)
    print("Calculating EMA-12...")
    df['EMA_12'] = TA.EMA(df, 12)
    
    # 2. EMA - Exponential Moving Average (26 days)
    print("Calculating EMA-26...")
    df['EMA_26'] = TA.EMA(df, 26)
    
    # 3. RSI - Relative Strength Index (14 periods)
    print("Calculating RSI...")
    df['RSI_14'] = TA.RSI(df)
    
    # 4. ROC - Rate of Change (14 periods)
    print("Calculating ROC...")
    df['ROC_14'] = TA.ROC(df, 14)
    
    # 5. MACD - Moving Average Convergence Divergence (12/26)
    print("Calculating MACD...")
    macd = TA.MACD(df)  # Using default 12/26 periods
    if isinstance(macd, pd.DataFrame):
        df['MACD_12_26'] = macd['MACD']
        df['MACD_SIGNAL_9'] = macd['SIGNAL']
    
    # 6. Bollinger Bands (20 periods, 2 standard deviations)
    print("Calculating Bollinger Bands...")
    bb = TA.BBANDS(df)
    if isinstance(bb, pd.DataFrame):
        df['BBANDS_UPPER_20_2'] = bb['BB_UPPER']
        df['BBANDS_MIDDLE_20_2'] = bb['BB_MIDDLE']
        df['BBANDS_LOWER_20_2'] = bb['BB_LOWER']
    
    # 7. Stochastic Oscillator
    print("Calculating Stochastic Oscillator...")
    try:
        stoch = TA.STOCH(df)
        if isinstance(stoch, pd.DataFrame):
            df['STOCH_K'] = stoch['K']
            df['STOCH_D'] = stoch['D']
    except Exception as stoch_error:
        print(f"Stochastic calculation error: {str(stoch_error)}")
    
    # 8. Williams %R
    print("Calculating Williams %R...")
    df['WILLIAMS_R'] = TA.WILLIAMS(df)

    return df