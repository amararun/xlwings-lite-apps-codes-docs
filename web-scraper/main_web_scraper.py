import xlwings as xw
from xlwings import script
import requests
import json
import os
import traceback
import time
from datetime import datetime
import sys
import pandas as pd

@script
def scrape_urls_from_list(book: xw.Book):
    """
    Excel-based web scraper that fetches data for URLs listed in the URL_LIST sheet.
    Results are formatted into an Excel table in the DATA sheet.
    Column definitions are read from COLUMN_INPUTS sheet.
    """
    print("--- SCRIPT FUNCTION ENTERED ---")
    # Print start message
    start_time = datetime.now()
    print("="*50)
    print(f"⭐ BATCH URL SCRAPER STARTED AT: {start_time.strftime('%Y-%m-%d %H:%M:%S')} ⭐")
    print("="*50)
    
    try:
        # Combined step summary approach - print step header only
        print("\n[1/7] READING PARAMETERS")

        # Check for MASTER sheet first
        if "MASTER" not in [s.name for s in book.sheets]:
            raise ValueError("Required sheet 'MASTER' not found. Please ensure the configuration sheet exists and is named correctly.")
        
        master_sheet = book.sheets["MASTER"]
        print("  • 'MASTER' sheet found.")

        # Attempt to get Jina API key from environment variable first
        jina_api_key = os.getenv("JINA_API_KEY")
        if not jina_api_key:
            jina_api_key = master_sheet["B4"].value
            print("  ⚠️ Using Jina API key from Excel cell (B4) as environment variable 'JINA_API_KEY' was not found or was empty.")
        else:
            print("  ✅ Using Jina API key from environment variable 'JINA_API_KEY'.")

        # Read Gemini model from Excel
        gemini_model = master_sheet["B5"].value

        # Attempt to get Gemini API key from environment variable first
        gemini_api_key = os.getenv("GEMINI_API_KEY")
        if not gemini_api_key:
            gemini_api_key = master_sheet["B6"].value
            print("  ⚠️ Using Gemini API key from Excel cell (B6) as environment variable 'GEMINI_API_KEY' was not found or was empty.")
        else:
            print("  ✅ Using Gemini API key from environment variable 'GEMINI_API_KEY'.")

        # Read additional parameters with defaults if not present
        request_delay = master_sheet["B8"].value or 2
        retry_delay = master_sheet["B9"].value or 5
        max_retries = master_sheet["B10"].value or 3
        request_timeout = master_sheet["B11"].value or 30
        max_output_tokens = master_sheet["B13"].value or 40000
        thinking_budget = master_sheet["B14"].value  # None if blank - optional parameter
        temperature = master_sheet["B15"].value if master_sheet["B15"].value is not None else 0.0
        topP = master_sheet["B16"].value if master_sheet["B16"].value is not None else 0.8

        # Log parameter values for debugging
        print("  • Configuration Parameters:")
        print(f"    - Request Delay (B8): {request_delay} seconds (delay between processing each URL)")
        print(f"    - Retry Delay (B9): {retry_delay} seconds (delay before retrying failed API calls)")
        print(f"    - Max Retries (B10): {max_retries} attempts (maximum retry attempts for API calls)")
        print(f"    - Request Timeout (B11): {request_timeout} SECONDS (maximum wait time for API response)")
        print(f"    - Max Output Tokens (B13): {max_output_tokens:,} tokens (maximum tokens Gemini can generate)")

        # Log thinking budget
        if thinking_budget is not None:
            print(f"    - Thinking Budget (B14): {int(thinking_budget):,} tokens (for Gemini 2.5+ models)")
            if thinking_budget == 0:
                print(f"      → Thinking DISABLED (fastest, cheapest, may reduce accuracy)")
            elif thinking_budget == -1:
                print(f"      → DYNAMIC thinking (model decides budget automatically)")
            else:
                print(f"      → FIXED thinking budget (balance quality vs. cost)")
        else:
            print(f"    - Thinking Budget (B14): Not set (using model default - dynamic for 2.5 Flash)")
            print(f"      💡 Tip: Set to 0 (no thinking), -1 (dynamic), or 512-24576 (fixed budget)")

        # Temperature validation and auto-clamp
        original_temp = temperature
        if temperature < 0.0:
            temperature = 0.0
            print(f"    - ⚠️ Temperature {original_temp} too low, clamped to 0.0")
        elif temperature > 2.0:
            temperature = 2.0
            print(f"    - ⚠️ Temperature {original_temp} too high, clamped to 2.0")
        print(f"    - Temperature (B15): {temperature} (range: 0.0-2.0, higher = more random)")

        # TopP validation and auto-clamp
        original_topP = topP
        if topP < 0.0:
            topP = 0.0
            print(f"    - ⚠️ TopP {original_topP} too low, clamped to 0.0")
        elif topP > 1.0:
            topP = 1.0
            print(f"    - ⚠️ TopP {original_topP} too high, clamped to 1.0")
        print(f"    - TopP (B16): {topP} (range: 0.0-1.0, higher = more diverse)")

        print(f"    ⚠️ IMPORTANT: Request Timeout is in SECONDS, not minutes!")
        if request_timeout < 60:
            print(f"    ⚠️ WARNING: Timeout is only {request_timeout} seconds - may be too short for large pages")
            print(f"    💡 Recommended: Set to at least 120 seconds (2 minutes) for pages with many items")
        if max_output_tokens < 10000:
            print(f"    ⚠️ WARNING: Max Output Tokens is only {max_output_tokens:,} - may be too low for large extractions")
            print(f"    💡 Recommended: Set to at least 40,000 for gemini-2.0-flash or 65,000 for gemini-2.5-flash")

        # Validate required parameters
        if not jina_api_key:
            raise ValueError("Jina API key not found. Please set 'JINA_API_KEY' environment variable or enter it in cell B4.")
        if not gemini_model:
            raise ValueError("Gemini model not specified in cell B5. Please enter a valid model name.")
        if not gemini_api_key:
            raise ValueError("Gemini API key not found. Please set 'GEMINI_API_KEY' environment variable or enter it in cell B6.")
        
        # Step 2: Read column definitions
        print("\n[2/7] READING COLUMN DEFINITIONS")
        try:
            column_sheet = book.sheets["COLUMN_INPUTS"]
        except:
            raise ValueError("COLUMN_INPUTS sheet not found. Please create this sheet with your column definitions.")
        
        # Try to read custom instructions from cell D2 if available
        custom_instructions = column_sheet["D2"].value
        if custom_instructions:
            print(f"  • Custom instructions found")
        
        # Read column definitions from A3 downward
        column_fields = []
        column_descriptions = []
        
        # Get all values from columns A and B starting at row 3
        names_range = column_sheet.range("A3").expand('down')
        names = names_range.value
        
        # Handle single cell case
        if not isinstance(names, list):
            names = [names]
            
        # Get descriptions (may be shorter than names list)
        descriptions_range = column_sheet.range(f"B3:B{2+len(names)}")
        descriptions = descriptions_range.value
        
        # Handle single cell case for descriptions
        if not isinstance(descriptions, list):
            descriptions = [descriptions]
            
        # Process each column name and add valid ones
        for i, name in enumerate(names):
            # Skip empty names
            if not name:
                continue
                
            # Clean the name (convert to string and strip whitespace)
            clean_name = str(name).strip()
            if not clean_name:
                continue
                
            # Add the valid column name
            column_fields.append(clean_name)
            
            # Add corresponding description if available
            if i < len(descriptions) and descriptions[i]:
                column_descriptions.append(str(descriptions[i]).strip())
            else:
                column_descriptions.append("")  # Empty description
        
        # Final validation for columns
        if not column_fields:
            raise ValueError("No valid column definitions found in COLUMN_INPUTS sheet. Please add column definitions starting from cell A3.")
        
        print(f"  • Found {len(column_fields)} column definitions")
        
        # Step 3: Get URLs from URL_LIST sheet
        print("\n[3/7] READING URLS")
        try:
            url_sheet = book.sheets["URL_LIST"]
        except:
            raise ValueError("URL_LIST sheet not found. Please create this sheet with URLs starting from cell A2.")

        # Read URLs from column A starting from row 2
        url_range = url_sheet.range("A2").expand('down')
        urls_raw = url_range.value

        # Simplified empty check and URL processing
        if not urls_raw:
            raise ValueError("No URLs found in URL_LIST sheet. Please add URLs starting from cell A2.")

        # Ensure urls_raw is a list
        urls_raw = [urls_raw] if isinstance(urls_raw, str) else urls_raw

        # Read STATUS column (B) - same number of rows as URLs
        status_range = url_sheet.range(f"B2:B{1 + len(urls_raw)}")
        status_values = status_range.value

        # Ensure status_values is a list
        if not isinstance(status_values, list):
            status_values = [status_values]

        # Build list of (url, row_number, status) tuples for processing
        url_data = []
        for i, url in enumerate(urls_raw):
            if url and isinstance(url, str) and url.strip():
                row_num = i + 2  # Rows start at 2 in Excel
                status = status_values[i] if i < len(status_values) else None
                url_data.append((url.strip(), row_num, status))

        if not url_data:
            raise ValueError("No valid URLs found in URL_LIST sheet. Please add URLs starting from cell A2.")

        # Filter out URLs with DONE or SKIP status
        urls_to_process = [
            (url, row_num) for url, row_num, status in url_data
            if status not in ['✅ DONE', '⏭️ SKIP']
        ]

        skipped_count = len(url_data) - len(urls_to_process)
        print(f"  • Found {len(url_data)} total URLs")
        print(f"  • Skipping {skipped_count} URLs (already DONE or marked SKIP)")
        print(f"  • Processing {len(urls_to_process)} URLs")

        # Step 4: Prepare DATA sheet - simplified with less status output
        print("\n[4/7] PREPARING DATA SHEET")
        try:
            data_sheet = book.sheets["DATA"]
            data_sheet.clear()
        except:
            data_sheet = book.sheets.add("DATA")

        # Set up headers
        headers = ["URL"] + column_fields
        data_sheet["A1"].value = headers
        header_range = data_sheet["A1"].resize(1, len(headers))
        header_range.color = "#A7D9AB"  # Light green color

        # Prepare ERROR_LOG sheet
        try:
            error_log_sheet = book.sheets["ERROR_LOG"]
            # Clear only from row 2 onwards to preserve manual title in row 1
            error_log_sheet.range("A2:ZZ10000").clear_contents()
        except:
            error_log_sheet = book.sheets.add("ERROR_LOG")

        # Set up ERROR_LOG headers (starting from row 2 to preserve manual title in row 1)
        error_log_headers = ["TIMESTAMP", "URL", "ERROR_TYPE", "ERROR_MESSAGE"]
        error_log_sheet["A2"].value = error_log_headers
        error_log_header_range = error_log_sheet["A2"].resize(1, len(error_log_headers))
        error_log_header_range.color = "#FFB3B3"  # Light red color
        error_log_row = 3  # Track current row in error log (start data from row 3)

        # Step 5: Process URLs - consolidated progress reporting
        print(f"\n[5/7] PROCESSING {len(urls_to_process)} URLS")

        all_items_data = []
        current_row = 2
        total_items = 0
        successful_urls = 0
        total_api_calls = 0

        # Enhanced tracking for dashboard metrics
        url_processing_times = []  # List of (url, duration_seconds) tuples
        jina_response_times = []  # List of Jina API response times in seconds
        gemini_response_times = []  # List of Gemini API response times in seconds
        total_input_tokens = 0
        total_output_tokens = 0
        items_per_url = []  # List of item counts per URL
        urls_with_retries = 0  # Count of URLs that needed retries
        first_url_timestamp = None
        last_url_timestamp = None

        # Helper function to handle errors and log to both DATA and ERROR_LOG sheets
        def log_error(url, error_type, error_message):
            nonlocal current_row, error_log_row
            # Add error row to DATA sheet
            error_row = [url] + [error_message] + [""] * (len(headers) - 2)
            data_sheet[f"A{current_row}"].value = error_row
            current_row += 1
            # Log to ERROR_LOG sheet
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            error_log_sheet[f"A{error_log_row}"].value = [timestamp, url, error_type, error_message]
            error_log_row += 1

        # Simplified URL progress reporting
        for i, (url, url_row_num) in enumerate(urls_to_process):
            print(f"\n--- Processing URL {i+1}/{len(urls_to_process)}: {url} ---")

            # Track URL processing start time
            url_start_time = time.time()
            if first_url_timestamp is None:
                first_url_timestamp = datetime.now()

            success, markdown_content, jina_time = scrape_url(url, jina_api_key, request_timeout)
            total_api_calls += 1  # Jina API call
            jina_response_times.append(jina_time)

            if not success or not markdown_content:
                print("    ❌ ERROR: Failed to scrape URL with Jina API")
                log_error(url, "Jina Scrape Failure", "Failed to scrape URL with Jina API")
                # Update STATUS to ERROR with timestamp (write both columns at once)
                url_sheet.range(f"B{url_row_num}:C{url_row_num}").value = ['❌ ERROR', datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
                # Track URL processing time even for failures
                url_processing_time = time.time() - url_start_time
                url_processing_times.append((url, url_processing_time))
                last_url_timestamp = datetime.now()
                continue

            success, structured_data, gemini_time, input_tok, output_tok, had_retry = extract_structured_data(
                markdown_content, gemini_api_key, gemini_model, column_fields,
                column_descriptions, max_retries, retry_delay, request_timeout, max_output_tokens, thinking_budget, temperature, topP, custom_instructions
            )
            total_api_calls += 1  # Gemini API call
            gemini_response_times.append(gemini_time)
            total_input_tokens += input_tok
            total_output_tokens += output_tok
            if had_retry:
                urls_with_retries += 1

            if not success:
                # Gemini API actually failed
                print("    ❌ ERROR: Gemini API failed to extract data")
                log_error(url, "Gemini API Failure", "Gemini API failed after all retries")
                # Update STATUS to ERROR with timestamp (write both columns at once)
                url_sheet.range(f"B{url_row_num}:C{url_row_num}").value = ['❌ ERROR', datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
                # Track URL processing time even for failures
                url_processing_time = time.time() - url_start_time
                url_processing_times.append((url, url_processing_time))
                last_url_timestamp = datetime.now()
                continue

            if not structured_data:
                # API succeeded but returned empty data (likely 404 or no matching content)
                print("    ❌ ERROR: No data found (empty result) - URL may be invalid or content doesn't match criteria")
                log_error(url, "No Data Found", "API succeeded but returned empty result - URL may be invalid or page has no matching content")
                # Update STATUS to ERROR with timestamp (write both columns at once)
                url_sheet.range(f"B{url_row_num}:C{url_row_num}").value = ['❌ ERROR', datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
                # Track URL processing time and zero items
                url_processing_time = time.time() - url_start_time
                url_processing_times.append((url, url_processing_time))
                items_per_url.append(0)
                last_url_timestamp = datetime.now()
                continue

            # Process extracted data
            for item in structured_data:
                row_data = [url] + [item.get(field, "") for field in column_fields]
                data_sheet[f"A{current_row}"].value = row_data
                current_row += 1
                all_items_data.append(row_data)

            total_items += len(structured_data)
            successful_urls += 1
            items_per_url.append(len(structured_data))
            print(f"    - Successfully processed URL, found {len(structured_data)} items.")

            # Track URL processing time
            url_processing_time = time.time() - url_start_time
            url_processing_times.append((url, url_processing_time))
            last_url_timestamp = datetime.now()

            # Update STATUS to DONE with timestamp (write both columns at once)
            url_sheet.range(f"B{url_row_num}:C{url_row_num}").value = ['✅ DONE', datetime.now().strftime('%Y-%m-%d %H:%M:%S')]

            # Add delay between URLs
            if i < len(urls_to_process) - 1:
                time.sleep(request_delay)
        
        # Step 6: Format data table - simplified logging
        print("\n[6/7] FORMATTING DATA")

        if all_items_data:
            try:
                table_range = data_sheet["A1"].resize(len(all_items_data) + 1, len(headers))
                try:
                    table = data_sheet.tables.add(table_range, "ExportedData")
                except Exception:
                    pass  # Table formatting is optional
            except Exception:
                pass  # Continue even if formatting fails
            
            print(f"  • Processed {len(all_items_data)} items from {len(urls_to_process)} URLs")
        else:
            print("  • No data collected")

        # Step 7: Create/Update DASHBOARD sheet
        print("\n[7/7] UPDATING DASHBOARD")
        try:
            dashboard_sheet = book.sheets["DASHBOARD"]
            # Clear only from row 2 onwards to preserve manual title in row 1
            dashboard_sheet.range("A2:ZZ10000").clear_contents()
        except:
            dashboard_sheet = book.sheets.add("DASHBOARD")

        # Helper function to format time duration
        def format_duration(seconds_float):
            hours = int(seconds_float // 3600)
            minutes = int((seconds_float % 3600) // 60)
            secs = int(seconds_float % 60)
            return f"{hours:02d} hr: {minutes:02d} min: {secs:02d} sec"

        # Calculate comprehensive metrics
        total_duration = (datetime.now() - start_time).total_seconds()
        end_time = datetime.now()
        success_rate = (successful_urls / len(urls_to_process) * 100) if len(urls_to_process) > 0 else 0
        error_rate = ((len(urls_to_process) - successful_urls) / len(urls_to_process) * 100) if len(urls_to_process) > 0 else 0
        avg_items_per_url = (len(all_items_data) / successful_urls) if successful_urls > 0 else 0

        # Performance metrics
        avg_time_per_url = (sum([t for _, t in url_processing_times]) / len(url_processing_times)) if url_processing_times else 0
        fastest_url_time = min([t for _, t in url_processing_times]) if url_processing_times else 0
        slowest_url_time = max([t for _, t in url_processing_times]) if url_processing_times else 0
        avg_jina_time = (sum(jina_response_times) / len(jina_response_times)) if jina_response_times else 0
        avg_gemini_time = (sum(gemini_response_times) / len(gemini_response_times)) if gemini_response_times else 0

        # Data quality metrics
        urls_with_zero_items = items_per_url.count(0) if items_per_url else 0
        max_items_single_url = max(items_per_url) if items_per_url else 0
        non_zero_items = [count for count in items_per_url if count > 0]
        min_items_single_url = min(non_zero_items) if non_zero_items else 0

        # Token metrics
        total_tokens = total_input_tokens + total_output_tokens
        avg_tokens_per_url = (total_tokens / len(urls_to_process)) if len(urls_to_process) > 0 else 0
        avg_tokens_per_item = (total_tokens / len(all_items_data)) if len(all_items_data) > 0 else 0

        # URL status breakdown
        urls_pending = sum(1 for _, _, status in url_data if not status or status.strip() == '')
        urls_skipped = sum(1 for _, _, status in url_data if status == '⏭️ SKIP')
        urls_remaining = len(url_data) - len(urls_to_process) - urls_skipped
        completion_pct = ((len(urls_to_process)) / len(url_data) * 100) if len(url_data) > 0 else 0

        # Build dashboard metrics organized by section
        dashboard_rows = []

        # --- A. PERFORMANCE & SPEED METRICS ---
        dashboard_rows.append({"Metric": "--- PERFORMANCE & SPEED ---", "Value": ""})
        dashboard_rows.append({"Metric": "Total Run Time", "Value": format_duration(total_duration)})
        dashboard_rows.append({"Metric": "Average Time per URL", "Value": format_duration(avg_time_per_url)})
        dashboard_rows.append({"Metric": "Fastest URL Processed", "Value": format_duration(fastest_url_time)})
        dashboard_rows.append({"Metric": "Slowest URL Processed", "Value": format_duration(slowest_url_time)})
        dashboard_rows.append({"Metric": "Average Jina API Response Time", "Value": format_duration(avg_jina_time)})
        dashboard_rows.append({"Metric": "Average Gemini API Response Time", "Value": format_duration(avg_gemini_time)})
        dashboard_rows.append({"Metric": "", "Value": ""})  # Blank row separator

        # --- B. URL PROCESSING SUMMARY ---
        dashboard_rows.append({"Metric": "--- URL PROCESSING SUMMARY ---", "Value": ""})
        dashboard_rows.append({"Metric": "Total URLs in List", "Value": len(url_data)})
        dashboard_rows.append({"Metric": "URLs Processed (This Run)", "Value": len(urls_to_process)})
        dashboard_rows.append({"Metric": "Successful URLs", "Value": successful_urls})
        dashboard_rows.append({"Metric": "Failed URLs", "Value": len(urls_to_process) - successful_urls})
        dashboard_rows.append({"Metric": "Error Rate %", "Value": f"{error_rate:.1f}%"})
        dashboard_rows.append({"Metric": "", "Value": ""})  # Blank row separator

        # --- C. DATA EXTRACTION METRICS ---
        dashboard_rows.append({"Metric": "--- DATA EXTRACTION METRICS ---", "Value": ""})
        dashboard_rows.append({"Metric": "Total Items Extracted", "Value": len(all_items_data)})
        dashboard_rows.append({"Metric": "Average Items per URL", "Value": f"{avg_items_per_url:.1f}"})
        dashboard_rows.append({"Metric": "URLs with Zero Items", "Value": urls_with_zero_items})
        dashboard_rows.append({"Metric": "Max Items from Single URL", "Value": max_items_single_url})
        dashboard_rows.append({"Metric": "Min Items from Single URL (excl. 0)", "Value": min_items_single_url})
        dashboard_rows.append({"Metric": "", "Value": ""})  # Blank row separator

        # --- D. ERROR ANALYSIS ---
        dashboard_rows.append({"Metric": "--- ERROR ANALYSIS ---", "Value": ""})
        dashboard_rows.append({"Metric": "URLs with Retries", "Value": urls_with_retries})
        dashboard_rows.append({"Metric": "", "Value": ""})  # Blank row separator

        # --- E. TOKEN & COST TRACKING ---
        dashboard_rows.append({"Metric": "--- TOKEN & COST TRACKING ---", "Value": ""})
        dashboard_rows.append({"Metric": "Total Input Tokens", "Value": f"{total_input_tokens:,}"})
        dashboard_rows.append({"Metric": "Total Output Tokens", "Value": f"{total_output_tokens:,}"})
        dashboard_rows.append({"Metric": "Total Tokens", "Value": f"{total_tokens:,}"})
        dashboard_rows.append({"Metric": "Average Tokens per URL", "Value": f"{avg_tokens_per_url:.0f}"})
        dashboard_rows.append({"Metric": "Average Tokens per Item", "Value": f"{avg_tokens_per_item:.0f}"})
        dashboard_rows.append({"Metric": "", "Value": ""})  # Blank row separator

        # --- F. URL STATUS BREAKDOWN ---
        dashboard_rows.append({"Metric": "--- URL STATUS BREAKDOWN ---", "Value": ""})
        dashboard_rows.append({"Metric": "URLs Pending (No Status)", "Value": urls_pending})
        dashboard_rows.append({"Metric": "URLs Skipped", "Value": urls_skipped})
        dashboard_rows.append({"Metric": "URLs Remaining to Process", "Value": urls_remaining})
        dashboard_rows.append({"Metric": "Completion %", "Value": f"{completion_pct:.1f}%"})
        dashboard_rows.append({"Metric": "", "Value": ""})  # Blank row separator

        # --- G. TEMPORAL & TIMESTAMPS ---
        dashboard_rows.append({"Metric": "--- TEMPORAL & TIMESTAMPS ---", "Value": ""})
        dashboard_rows.append({"Metric": "First URL Processed At", "Value": first_url_timestamp.strftime('%Y-%m-%d %H:%M:%S') if first_url_timestamp else "N/A"})
        dashboard_rows.append({"Metric": "Last URL Processed At", "Value": last_url_timestamp.strftime('%Y-%m-%d %H:%M:%S') if last_url_timestamp else "N/A"})
        dashboard_rows.append({"Metric": "Total API Calls Made", "Value": total_api_calls})

        # Convert to DataFrame
        dashboard_df = pd.DataFrame(dashboard_rows)

        # Write to sheet using .options(index=False) to exclude index column
        # Starting from row 2 to preserve manual title in row 1
        start_cell = dashboard_sheet["A2"]
        start_cell.options(pd.DataFrame, index=False).value = dashboard_df

        # Format as table
        try:
            table_range = start_cell.resize(dashboard_df.shape[0] + 1, dashboard_df.shape[1])
            dashboard_sheet.tables.add(source=table_range)
            print("  • Dashboard formatted as table")
        except Exception as e:
            print(f"  ⚠️ Warning: Could not format dashboard as table. Error: {e}")

        # Auto-fit columns
        try:
            dashboard_sheet.autofit("columns")
            print("  • Auto-fitted columns")
        except Exception as e:
            print(f"  ⚠️ Warning: Could not auto-fit columns. Error: {e}")

        # Color the header row
        dashboard_sheet.range("A2:B2").color = "#4A90E2"  # Blue header

        print("  • Dashboard updated successfully")

        # Simplified summary report
        print("\n" + "="*50)
        print(f"✅ COMPLETED in {total_duration:.1f}s | URLs Processed: {len(urls_to_process)} | Items: {len(all_items_data)}")
        print("="*50)

    except Exception as e:
        print("\n" + "="*50)
        print(f"❌ ERROR: {str(e)}")
        print(traceback.format_exc())
        print("="*50)
        return False
    
    return True

def scrape_url(url, jina_api_key, request_timeout):
    """
    Scrape a URL using Jina API.

    Args:
        url (str): The URL to scrape
        jina_api_key (str): Jina API key
        request_timeout (int): Request timeout in seconds

    Returns:
        tuple: (bool, str, float) - Success status, markdown content if successful, response time in seconds
    """
    print(f"    - Scraping URL with Jina...")
    try:
        jina_url = f"https://r.jina.ai/{url}"
        headers = {
            "Authorization": f"Bearer {jina_api_key}",
            "X-Engine": "browser",
            "X-Return-Format": "markdown"
        }

        print(f"      > Calling Jina API endpoint...")
        jina_start_time = time.time()
        response = requests.get(jina_url, headers=headers, timeout=request_timeout)
        jina_response_time = time.time() - jina_start_time

        if response.status_code == 200:
            content = response.text
            print("      > Jina API call successful, content received.")
            return True, content, jina_response_time
        else:
            print(f"      > Jina API returned status {response.status_code}. Response: {response.text}")
            return False, None, jina_response_time
            
    except Exception as e:
        print(f"      > Jina API request failed. Error: {e}")
        return False, None, 0.0

def extract_structured_data(markdown_content, gemini_api_key, gemini_model, column_fields, column_descriptions, max_retries, retry_delay, request_timeout, max_output_tokens, thinking_budget, temperature, topP, custom_instructions):
    """
    Extract structured data from markdown content using Google Gemini API.

    Args:
        markdown_content (str): The markdown content to process
        gemini_api_key (str): Gemini API key
        gemini_model (str): Gemini model name
        column_fields (list): List of field names to extract
        column_descriptions (list): List of field descriptions
        max_retries (int): Maximum number of retries
        retry_delay (int): Delay between retries in seconds
        request_timeout (int): Request timeout in seconds
        max_output_tokens (int): Maximum output tokens for Gemini response
        thinking_budget (int or None): Thinking budget for Gemini 2.5+ models (None = use model default)
        temperature (float): Temperature for randomness control (0.0-2.0)
        topP (float): Top-P nucleus sampling parameter (0.0-1.0)
        custom_instructions (str): Custom instructions for data extraction

    Returns:
        tuple: (bool, list, float, int, int, bool) - Success status, structured data (if successful),
               response time in seconds, input tokens, output tokens, whether retry was needed
    """
    print("    - Extracting structured data with Gemini...")
    
    # Construct the field descriptions for the prompt
    field_descriptions = ""
    for i, (field, desc) in enumerate(zip(column_fields, column_descriptions)):
        field_descriptions += f"{i+1}. {field}: {desc}\n"
    
    # Prepare prompt for Gemini API
    prompt_start = f"""
    You are a web scraping expert tasked with FILTERING data based on specific criteria.
    """
    
    # Add custom instructions if available
    if custom_instructions:
        prompt_start += f"""
    >>>>>> CRITICAL FILTERING INSTRUCTIONS - YOU MUST FOLLOW THESE <<<<<<
    
    {custom_instructions}
    
    I REPEAT: ONLY extract and return items that match the above criteria.
    ALL OTHER ITEMS MUST BE EXCLUDED from your response.
    This is the most important part of your task.
    """
    
    # Create sanitized field names and add to schema
    sanitized_fields = {}
    for field in column_fields:
        # Sanitize field name for the schema (remove spaces, lowercase)
        schema_field = field.lower().replace(" ", "_").replace("-", "_")
        sanitized_fields[field] = schema_field
    
    # Simplified JSON schema example construction
    field_examples = [f'        "{schema_field}": "string with {field}"' 
                      for field, schema_field in sanitized_fields.items()]
    
    # Create the JSON schema example with both examples in one go
    json_schema_example = "[\n    {\n" + \
                          ",\n".join(field_examples) + \
                          "\n    },\n    {\n" + \
                          ",\n".join([f'        "{schema_field}": "string with next {field}"' 
                                    for field, schema_field in sanitized_fields.items()]) + \
                          "\n    },\n    ...and so on for ALL entries\n]"
    
    # Complete the prompt with the schema example
    prompt = prompt_start + f"""
    {markdown_content}
    
    This is a web scraping project where we need to extract specific items listed on the page.
    
    For EACH qualifying item on the page, extract these specific fields:
    {field_descriptions}
    
    Return the data as a JSON array where each object has this structure:
    {json_schema_example}
    
    DO NOT miss any qualifying item. If any field is not found for a particular item, use null or empty string as appropriate.
    """
    
    # Add a final reminder about filtering if custom instructions are provided
    if custom_instructions:
        prompt += f"""
    
    FINAL REMINDER: {custom_instructions}
    ONLY include matching items in your response. Filter out everything else.
    """

    # Temperature and topP are now passed as parameters from Excel (B15 and B16)
    # and have already been validated/clamped in the main function

    # Prepare API payload
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": temperature,
            "maxOutputTokens": max_output_tokens,
            "topP": topP,
            "response_mime_type": "application/json",
            "response_schema": {
                "type": "ARRAY",
                "items": {
                    "type": "OBJECT",
                    "properties": {},
                    "propertyOrdering": list(sanitized_fields.values()),  # Add property ordering
                    "required": list(sanitized_fields.values())  # Make all fields required
                }
            }
        }
    }
    
    # Add sanitized field names to schema
    for field, schema_field in sanitized_fields.items():
        payload["generationConfig"]["response_schema"]["items"]["properties"][schema_field] = {"type": "STRING"}

    # Add thinking config if specified (optional - only for Gemini 2.5+ models)
    if thinking_budget is not None:
        # Check if model supports thinking (2.5+ models only, not 2.0)
        model_supports_thinking = "2.0" not in gemini_model.lower()

        if model_supports_thinking:
            # Model-specific thinking budget limits (min, max)
            # Source: https://cloud.google.com/vertex-ai/generative-ai/docs/thinking
            model_name_lower = gemini_model.lower()

            if "flash-lite" in model_name_lower or "flashlite" in model_name_lower:
                min_budget, max_budget = 512, 24576
                model_variant = "Flash-Lite"
            elif "2.5-pro" in model_name_lower or "2_5-pro" in model_name_lower:
                min_budget, max_budget = 128, 32768
                model_variant = "2.5 Pro"
            elif "flash" in model_name_lower:
                # Covers: 2.5-flash, flash-latest, etc.
                min_budget, max_budget = 1, 24576
                model_variant = "Flash"
            else:
                # Default to Flash limits for unknown 2.5+ models
                min_budget, max_budget = 1, 24576
                model_variant = "Unknown (using Flash defaults)"

            # Auto-clamp thinking budget to valid range (unless 0 or -1 which are special)
            original_budget = int(thinking_budget)
            clamped_budget = original_budget

            # Special values: 0 (disabled) and -1 (dynamic) are always valid
            if original_budget not in [0, -1]:
                if original_budget < min_budget:
                    clamped_budget = min_budget
                    print(f"      > ⚠️ Thinking Budget {original_budget:,} is below minimum for {model_variant}")
                    print(f"      > 🔧 Auto-clamped to minimum: {clamped_budget:,} tokens")
                elif original_budget > max_budget:
                    clamped_budget = max_budget
                    print(f"      > ⚠️ Thinking Budget {original_budget:,} exceeds maximum for {model_variant}")
                    print(f"      > 🔧 Auto-clamped to maximum: {clamped_budget:,} tokens")

            # Apply the (potentially clamped) thinking budget
            payload["generationConfig"]["thinkingConfig"] = {
                "thinkingBudget": clamped_budget
            }

            # Log final thinking budget
            if clamped_budget == 0:
                print(f"      > Thinking Budget: {clamped_budget:,} tokens (DISABLED)")
            elif clamped_budget == -1:
                print(f"      > Thinking Budget: DYNAMIC (model decides, range: {min_budget:,}-{max_budget:,})")
            else:
                print(f"      > Thinking Budget: {clamped_budget:,} tokens (valid range: {min_budget:,}-{max_budget:,})")

        else:
            print(f"      > ⚠️ Thinking Budget set to {int(thinking_budget):,} but SKIPPED - {gemini_model} does not support thinking")
            print(f"      > 💡 Thinking is only supported by Gemini 2.5+ models (2.5-flash, 2.5-flash-lite, 2.5-pro)")
            print(f"      > To use thinking, switch to gemini-2.5-flash or leave B14 blank for {gemini_model}")

    # Make API request to Gemini with retries
    print(f"      > Calling Gemini API ({gemini_model}) with {max_retries} retries...")

    # Log the API endpoint (without the key for security)
    gemini_url = f"https://generativelanguage.googleapis.com/v1beta/models/{gemini_model}:generateContent"
    api_key_preview = gemini_api_key[:8] + "..." if len(gemini_api_key) > 8 else "***"
    print(f"        - API Endpoint: {gemini_url}")
    print(f"        - API Key: {api_key_preview}")

    # Track response time and tokens
    gemini_response_time = 0.0
    input_tokens = 0
    output_tokens = 0
    needed_retry = False

    for attempt in range(max_retries):
        print(f"        - Attempt {attempt + 1}/{max_retries}...")
        try:
            gemini_start_time = time.time()
            response = requests.post(
                f"{gemini_url}?key={gemini_api_key}",
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=request_timeout
            )
            gemini_response_time = time.time() - gemini_start_time

            if response.status_code == 200:
                print("        - Gemini API call successful (Status 200).")
                response_data = response.json()

                # Extract token usage metadata
                if 'usageMetadata' in response_data:
                    usage = response_data['usageMetadata']
                    input_tokens = usage.get('promptTokenCount', 0)
                    # Output tokens = total - prompt (some models include thoughtsTokenCount)
                    total_tokens_from_api = usage.get('totalTokenCount', 0)
                    output_tokens = total_tokens_from_api - input_tokens
                    print(f"        - Token usage: Input={input_tokens}, Output={output_tokens}, Total={total_tokens_from_api}")

                # Log the full response structure for debugging
                print(f"        - Full API response structure:")
                print(f"          {json.dumps(response_data, indent=10)[:2000]}...")  # First 2000 chars

                if 'candidates' in response_data and len(response_data['candidates']) > 0:
                    candidate = response_data['candidates'][0]
                    print(f"        - Candidate structure: {list(candidate.keys())}")

                    # Check finishReason to detect truncated responses
                    finish_reason = candidate.get('finishReason', 'UNKNOWN')
                    print(f"        - Finish reason: {finish_reason}")

                    if finish_reason != 'STOP':
                        print(f"        - ❌ ERROR: Response did not complete normally")
                        print(f"        - Finish reason: {finish_reason}")

                        if finish_reason == 'MAX_TOKENS':
                            print(f"        - Response was truncated due to MAX_TOKENS limit")
                            print(f"        - Current maxOutputTokens: {max_output_tokens:,}")
                            print(f"        - Suggestion: This page has too many items to extract in one response")
                            print(f"        - Input tokens: {input_tokens}, Output tokens so far: {output_tokens}")
                            print(f"        - Either increase maxOutputTokens (cell B13) or reduce the content to scrape")
                        elif finish_reason == 'LENGTH':
                            print(f"        - Response exceeded length limit")
                            print(f"        - This page may have too many items to extract at once")
                        elif finish_reason == 'SAFETY':
                            print(f"        - Response was blocked by safety filters")
                        else:
                            print(f"        - Unknown finish reason: {finish_reason}")

                        # Fall through to retry
                        if attempt < max_retries - 1:
                            print(f"        - Retrying in {retry_delay} seconds...")
                            time.sleep(retry_delay)
                        continue

                    # Check if content exists and has parts
                    if 'content' in candidate:
                        content = candidate['content']
                        print(f"        - Content structure: {list(content.keys())}")

                        if 'parts' in content and len(content['parts']) > 0:
                            structured_data_text = content['parts'][0]['text']
                        else:
                            print(f"        - ❌ ERROR: 'parts' key not found in content or parts is empty")
                            print(f"        - Content value: {content}")
                            # Fall through to retry
                            if attempt < max_retries - 1:
                                print(f"        - Retrying in {retry_delay} seconds...")
                                time.sleep(retry_delay)
                            continue
                    else:
                        print(f"        - ❌ ERROR: 'content' key not found in candidate")
                        print(f"        - Candidate: {candidate}")
                        # Fall through to retry
                        if attempt < max_retries - 1:
                            print(f"        - Retrying in {retry_delay} seconds...")
                            time.sleep(retry_delay)
                        continue
                    
                    # Parse the JSON - simplified handling
                    try:
                        # Handle markdown code blocks in the response
                        if "```" in structured_data_text:
                            print("        - Found markdown code block in response, attempting to extract JSON.")
                            # Extract content between code block markers
                            structured_data_text = structured_data_text.split("```")[1]
                            # Remove potential language identifier (like 'json')
                            if not structured_data_text.startswith("{") and not structured_data_text.startswith("["):
                                structured_data_text = structured_data_text.split("\n", 1)[1]
                            structured_data_text = structured_data_text.strip()
                        
                        print("        - Parsing JSON data...")
                        structured_data = json.loads(structured_data_text)

                        # Create reverse mapping and transform data more efficiently
                        reverse_mapping = {schema_field: orig_field for orig_field, schema_field in sanitized_fields.items()}

                        # Transform using dictionary comprehension for better efficiency
                        transformed_data = [
                            {reverse_mapping.get(field_name, field_name): value
                             for field_name, value in item.items()}
                            for item in structured_data
                        ]

                        print(f"        - Successfully parsed and transformed data. Found {len(transformed_data)} items.")
                        # Track if retry was needed (if we're not on the first attempt)
                        if attempt > 0:
                            needed_retry = True
                        return True, transformed_data, gemini_response_time, input_tokens, output_tokens, needed_retry
                    except json.JSONDecodeError as e:
                        print(f"        - ❌ ERROR: Failed to parse JSON from Gemini response")
                        print(f"        - JSON Error: {e}")

                        # Check if error is due to truncation
                        error_msg = str(e).lower()
                        if 'unterminated' in error_msg or 'unexpected end' in error_msg or 'expecting' in error_msg:
                            print(f"        - ⚠️ This looks like a TRUNCATED response (incomplete JSON)")
                            print(f"        - The response may have hit MAX_TOKENS even though finishReason was STOP")
                            print(f"        - Try increasing maxOutputTokens in the payload configuration")

                        print(f"        - Raw text length: {len(structured_data_text)} characters")
                        print(f"        - Raw text that failed parsing (last 500 chars):")
                        print(f"          ...{structured_data_text[-500:]}")
                        # Fall through to retry
                else:
                    print("        - WARNING: Gemini response had no 'candidates'.")
                    print(f"        - Full response: {json.dumps(response_data, indent=2)}")
            else:
                # Enhanced error logging
                print(f"        - ❌ ERROR: Gemini API returned status {response.status_code}")
                print(f"        - Model used: {gemini_model}")

                # Try to parse error response as JSON
                try:
                    error_data = response.json()
                    print(f"        - Error details (JSON):")
                    print(f"          {json.dumps(error_data, indent=10)}")

                    # Extract specific error message if available
                    if 'error' in error_data:
                        error_obj = error_data['error']
                        if 'message' in error_obj:
                            print(f"        - Error message: {error_obj['message']}")
                        if 'status' in error_obj:
                            print(f"        - Error status: {error_obj['status']}")
                        if 'code' in error_obj:
                            print(f"        - Error code: {error_obj['code']}")
                except:
                    # If not JSON, show raw text
                    print(f"        - Raw error response (first 500 chars):")
                    print(f"          {response.text[:500]}")

            # Only retry if this wasn't the last attempt
            if attempt < max_retries - 1:
                needed_retry = True
                print(f"        - Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)

        except Exception as e:
            print(f"        - ❌ ERROR: Gemini API request attempt {attempt + 1} failed")
            print(f"        - Model used: {gemini_model}")
            print(f"        - Exception type: {type(e).__name__}")
            print(f"        - Exception message: {str(e)}")

            # Only retry if this wasn't the last attempt
            if attempt < max_retries - 1:
                needed_retry = True
                print(f"        - Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)

    print(f"      > ❌ Gemini API calls failed after all {max_retries} retries.")
    print(f"      > Model that failed: {gemini_model}")
    print(f"      > Please check the console logs above for detailed error messages.")
    return False, None, gemini_response_time, input_tokens, output_tokens, needed_retry 