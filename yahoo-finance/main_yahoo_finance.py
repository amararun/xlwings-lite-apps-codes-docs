import xlwings as xw
import pandas as pd
from finta import TA
import requests
from xlwings import script
from datetime import datetime
import os

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

@script
def get_prices(book: xw.Book):
    """Fetch all historical price data (Open, High, Low, Close, Volume, etc.) for a ticker."""
    print("▶ STARTING get_prices ◀")
    
    # Get the PRICES sheet
    prices_sheet = book.sheets["PRICES"]
    
    # Read ticker, start_date, and end_date from PRICES sheet
    ticker = str(prices_sheet["B3"].value).strip().upper() if prices_sheet["B3"].value else None
    start_date_raw = prices_sheet["D3"].value
    end_date_raw = prices_sheet["F3"].value
    
    if not all([ticker, start_date_raw, end_date_raw]):
        prices_sheet["B8"].value = "Please enter ticker (B3), start date (D3), and end date (F3)"
        return
    
    # Convert Excel dates to YYYY-MM-DD format
    if isinstance(start_date_raw, (datetime, pd.Timestamp)):
        start_date = start_date_raw.strftime("%Y-%m-%d")
    else:
        # Convert Excel serial number to datetime
        start_date = pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(start_date_raw))
        start_date = start_date.strftime("%Y-%m-%d")
    
    if isinstance(end_date_raw, (datetime, pd.Timestamp)):
        end_date = end_date_raw.strftime("%Y-%m-%d")
    else:
        # Convert Excel serial number to datetime
        end_date = pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(end_date_raw))
        end_date = end_date.strftime("%Y-%m-%d")
    
    print(f"📊 Fetching prices for ticker: {ticker}")
    
    # Use the get-all-prices endpoint
    api_url = f"https://yfin.hosting.tigzig.com/get-all-prices/?tickers={ticker}&start_date={start_date}&end_date={end_date}"
    print(f"🔗 API URL: {api_url}")
    
    # Make API request
    response = requests.get(api_url)
    print(f"📥 Response status: {response.status_code}")
    
    if response.ok:
        # Parse JSON response
        data = response.json()
        
        if isinstance(data, dict) and not data.get("error"):
            # Convert nested JSON to DataFrame
            rows = []
            for date, ticker_data in data.items():
                if ticker in ticker_data:
                    row = ticker_data[ticker]
                    row['Date'] = date  # Add date to the row
                    rows.append(row)
            
            # Create DataFrame with all price columns
            df = pd.DataFrame(rows)
            
            # Reorder columns to put Date first
            cols = ['Date'] + [col for col in df.columns if col != 'Date']
            df = df[cols]
            
            # Clear existing content
            prices_sheet["B7:Z1000"].clear_contents()
            
            # Write headers in row 7
            prices_sheet["B7"].value = df.columns.tolist()
            
            # Write data starting from row 8
            prices_sheet["B8"].value = df.values.tolist()
            
            # Format as table including headers
            data_range = prices_sheet["B7"].resize(len(df) + 1, len(df.columns))
            prices_sheet.tables.add(data_range)
            
            print("✅ Headers and data pasted successfully")
            print(f"Headers: {df.columns.tolist()}")
        else:
            error_msg = data.get("error") if isinstance(data, dict) else "Invalid response format"
            prices_sheet["B8"].value = f"N/A - {error_msg}"
    else:
        prices_sheet["B8"].value = "N/A - Service temporarily unavailable"
    
    print("✓ ENDING get_prices ✓")

@script
def get_profit_loss(book: xw.Book):
    """Fetch income statement data for a ticker and display in PL sheet."""
    print("▶ STARTING get_profit_loss ◀")
    
    # Get the PL sheet
    pl_sheet = book.sheets["PL"]
    
    # Read ticker from B3 and format it
    ticker = str(pl_sheet["B3"].value).strip().upper() if pl_sheet["B3"].value else None
    if not ticker:
        pl_sheet["B8"].value = "Please enter a ticker symbol in cell B3"
        return
    
    print(f"📊 Fetching income statement data for ticker: {ticker}")
    
    # Use the Excel-specific endpoint
    api_url = f"https://yfin.hosting.tigzig.com/excel/get-income-statement/?tickers={ticker}"
    print(f"🔗 API URL: {api_url}")
    
    # Make API request
    response = requests.get(api_url)
    print(f"📥 Response status: {response.status_code}")
    
    if response.ok:
        # Parse JSON response
        data = response.json()
        
        if ticker in data:
            ticker_data = data[ticker]
            
            if 'error' not in ticker_data:
                # Clear existing content
                pl_sheet["B7:Z100"].clear_contents()
                
                # Get dates and metrics data
                dates = ticker_data['dates']
                metrics_data = ticker_data['data']
                
                # Create headers list starting with 'Metric'
                headers = ['Metric'] + dates
                
                # Write headers
                pl_sheet["B7"].value = headers
                
                # Create data rows
                data_rows = []
                for row_data in metrics_data:
                    row = [row_data['metric']]  # Start with metric name
                    for date in dates:
                        row.append(row_data.get(date))  # Add values for each date
                    data_rows.append(row)
                
                # Write data rows
                pl_sheet["B8"].value = data_rows
                
                # Format as table including headers
                table_range = pl_sheet["B7"].resize(len(data_rows) + 1, len(headers))
                pl_sheet.tables.add(table_range)
                
                print(f"✅ Successfully processed {len(metrics_data)} rows of data")
            else:
                pl_sheet["B8"].value = f"N/A - {ticker_data['error']}"
        else:
            pl_sheet["B8"].value = "N/A - No data found for ticker"
    else:
        pl_sheet["B8"].value = "N/A - Service temporarily unavailable"
    
    print("✓ ENDING get_profit_loss ✓")

@script
def get_balance_sheet(book: xw.Book):
    """Fetch balance sheet data for a ticker and display in BS sheet."""
    print("▶ STARTING get_balance_sheet ◀")
    
    # Get the BS sheet
    bs_sheet = book.sheets["BS"]
    
    # Read ticker from B3 and format it
    ticker = str(bs_sheet["B3"].value).strip().upper() if bs_sheet["B3"].value else None
    if not ticker:
        bs_sheet["B8"].value = "Please enter a ticker symbol in cell B3"
        return
    
    print(f"📊 Fetching balance sheet data for ticker: {ticker}")
    
    # Use the Excel-specific endpoint
    api_url = f"https://yfin.hosting.tigzig.com/excel/get-balance-sheet/?tickers={ticker}"
    print(f"🔗 API URL: {api_url}")
    
    # Make API request
    response = requests.get(api_url)
    print(f"📥 Response status: {response.status_code}")
    
    if response.ok:
        # Parse JSON response
        data = response.json()
        
        if ticker in data:
            ticker_data = data[ticker]
            
            if 'error' not in ticker_data:
                # Clear existing content
                bs_sheet["B7:Z100"].clear_contents()
                
                # Get dates and metrics data
                dates = ticker_data['dates']
                metrics_data = ticker_data['data']
                
                # Create headers list starting with 'Metric'
                headers = ['Metric'] + dates
                
                # Write headers
                bs_sheet["B7"].value = headers
                
                # Create data rows
                data_rows = []
                for row_data in metrics_data:
                    row = [row_data['metric']]  # Start with metric name
                    for date in dates:
                        row.append(row_data.get(date))  # Add values for each date
                    data_rows.append(row)
                
                # Write data rows
                bs_sheet["B8"].value = data_rows
                
                # Format as table including headers
                table_range = bs_sheet["B7"].resize(len(data_rows) + 1, len(headers))
                bs_sheet.tables.add(table_range)
                
                print(f"✅ Successfully processed {len(metrics_data)} rows of data")
            else:
                bs_sheet["B8"].value = f"N/A - {ticker_data['error']}"
        else:
            bs_sheet["B8"].value = "N/A - No data found for ticker"
    else:
        bs_sheet["B8"].value = "N/A - Service temporarily unavailable"
    
    print("✓ ENDING get_balance_sheet ✓")

@script
def get_cash_flow(book: xw.Book):
    """Fetch cash flow data for a ticker and display in CF sheet."""
    print("▶ STARTING get_cash_flow ◀")
    
    # Get the CF sheet
    cf_sheet = book.sheets["CF"]
    
    # Read ticker from B3 and format it
    ticker = str(cf_sheet["B3"].value).strip().upper() if cf_sheet["B3"].value else None
    if not ticker:
        cf_sheet["B8"].value = "Please enter a ticker symbol in cell B3"
        return
    
    print(f"📊 Fetching cash flow data for ticker: {ticker}")
    
    # Use the Excel-specific endpoint
    api_url = f"https://yfin.hosting.tigzig.com/excel/get-cash-flow/?tickers={ticker}"
    print(f"🔗 API URL: {api_url}")
    
    # Make API request
    response = requests.get(api_url)
    print(f"📥 Response status: {response.status_code}")
    
    if response.ok:
        # Parse JSON response
        data = response.json()
        
        if ticker in data:
            ticker_data = data[ticker]
            
            if 'error' not in ticker_data:
                # Clear existing content
                cf_sheet["B7:Z100"].clear_contents()
                
                # Get dates and metrics data
                dates = ticker_data['dates']
                metrics_data = ticker_data['data']
                
                # Create headers list starting with 'Metric'
                headers = ['Metric'] + dates
                
                # Write headers
                cf_sheet["B7"].value = headers
                
                # Create data rows
                data_rows = []
                for row_data in metrics_data:
                    row = [row_data['metric']]  # Start with metric name
                    for date in dates:
                        row.append(row_data.get(date))  # Add values for each date
                    data_rows.append(row)
                
                # Write data rows
                cf_sheet["B8"].value = data_rows
                
                # Format as table including headers
                table_range = cf_sheet["B7"].resize(len(data_rows) + 1, len(headers))
                cf_sheet.tables.add(table_range)
                
                print(f"✅ Successfully processed {len(metrics_data)} rows of data")
            else:
                cf_sheet["B8"].value = f"N/A - {ticker_data['error']}"
        else:
            cf_sheet["B8"].value = "N/A - No data found for ticker"
    else:
        cf_sheet["B8"].value = "N/A - Service temporarily unavailable"
    
    print("✓ ENDING get_cash_flow ✓")

@script
def get_quarterly(book: xw.Book):
    """Fetch quarterly income statement data for a ticker and display in QTLY sheet."""
    print("▶ STARTING get_quarterly ◀")
    
    # Get the QTLY sheet
    qtly_sheet = book.sheets["QTLY"]
    
    # Read ticker from B3 and format it
    ticker = str(qtly_sheet["B3"].value).strip().upper() if qtly_sheet["B3"].value else None
    if not ticker:
        qtly_sheet["B8"].value = "Please enter a ticker symbol in cell B3"
        return
    
    print(f"📊 Fetching quarterly income statement data for ticker: {ticker}")
    
    # Use the Excel-specific endpoint
    api_url = f"https://yfin.hosting.tigzig.com/excel/get-quarterly-income-statement/?tickers={ticker}"
    print(f"🔗 API URL: {api_url}")
    
    # Make API request
    response = requests.get(api_url)
    print(f"📥 Response status: {response.status_code}")
    
    if response.ok:
        # Parse JSON response
        data = response.json()
        
        if ticker in data:
            ticker_data = data[ticker]
            
            if 'error' not in ticker_data:
                # Clear existing content
                qtly_sheet["B7:Z100"].clear_contents()
                
                # Get dates and metrics data
                dates = ticker_data['dates']
                metrics_data = ticker_data['data']
                
                # Create headers list starting with 'Metric'
                headers = ['Metric'] + dates
                
                # Write headers
                qtly_sheet["B7"].value = headers
                
                # Create data rows
                data_rows = []
                for row_data in metrics_data:
                    row = [row_data['metric']]  # Start with metric name
                    for date in dates:
                        row.append(row_data.get(date))  # Add values for each date
                    data_rows.append(row)
                
                # Write data rows
                qtly_sheet["B8"].value = data_rows
                
                # Format as table including headers
                table_range = qtly_sheet["B7"].resize(len(data_rows) + 1, len(headers))
                qtly_sheet.tables.add(table_range)
                
                print(f"✅ Successfully processed {len(metrics_data)} rows of data")
            else:
                qtly_sheet["B8"].value = f"N/A - {ticker_data['error']}"
        else:
            qtly_sheet["B8"].value = "N/A - No data found for ticker"
    else:
        qtly_sheet["B8"].value = "N/A - Service temporarily unavailable"
    
    print("✓ ENDING get_quarterly ✓")

@script
def get_profile(book: xw.Book):
    """Fetch detailed company information for a ticker and display in PROFILE sheet."""
    print("▶ STARTING get_profile ◀")
    
    # Get the PROFILE sheet
    profile_sheet = book.sheets["PROFILE"]
    
    # Read ticker from B3 and format it
    ticker = str(profile_sheet["B3"].value).strip().upper() if profile_sheet["B3"].value else None
    if not ticker:
        profile_sheet["B8"].value = "Please enter a ticker symbol in cell B3"
        return
    
    print(f"📊 Fetching detailed information for ticker: {ticker}")
    
    # Use the new detailed info endpoint
    api_url = f"https://yfin.hosting.tigzig.com/get-detailed-info/?tickers={ticker}"
    print(f"🔗 API URL: {api_url}")
    
    # Make API request
    response = requests.get(api_url)
    print(f"📥 Response status: {response.status_code}")
    
    if response.ok:
        # Parse JSON response
        data = response.json()
        
        if ticker in data:
            ticker_data = data[ticker]
            
            if 'error' not in ticker_data:
                # Clear existing content
                profile_sheet["B7:Z1000"].clear_contents()
                
                # Process main company information
                main_info = ticker_data['main_info']
                
                # Define table sections and their fields
                sections = {
                    "1. Company Information": [
                        "address1", "address2", "city", "zip", "country", "phone", "fax", "website",
                        "industry", "industryKey", "industryDisp", "sector", "sectorKey", "sectorDisp",
                        "symbol", "longName", "shortName", "displayName", "typeDisp", "quoteType",
                        "language", "region", "market", "exchange", "exchangeTimezoneName",
                        "exchangeTimezoneShortName", "fullExchangeName", "messageBoardId", "irWebsite",
                        "corporateActions", "executiveTeam", "maxAge"
                    ],
                    "2. Stock Price & Market Activity": [
                        "currentPrice", "previousClose", "open", "dayLow", "dayHigh",
                        "regularMarketPrice", "regularMarketPreviousClose", "regularMarketOpen",
                        "regularMarketDayLow", "regularMarketDayHigh", "regularMarketChange",
                        "regularMarketChangePercent", "regularMarketVolume", "volume",
                        "averageVolume", "averageVolume3Month", "averageVolume10days",
                        "averageDailyVolume10Day", "regularMarketDayRange", "bid", "ask",
                        "bidSize", "askSize", "priceHint", "preMarketPrice", "preMarketTime",
                        "preMarketChange", "preMarketChangePercent", "marketState",
                        "hasPrePostMarketData", "quoteSourceName", "triggerable",
                        "customPriceAlertConfidence", "gmtOffSetMilliseconds", "exchangeDataDelayedBy"
                    ],
                    "3. Valuation & Ratios": [
                        "marketCap", "enterpriseValue", "priceToSalesTrailing12Months",
                        "priceToBook", "trailingPE", "forwardPE", "trailingPegRatio", "beta",
                        "enterpriseToRevenue", "enterpriseToEbitda", "payoutRatio", "debtToEquity",
                        "recommendationMean", "recommendationKey", "numberOfAnalystOpinions",
                        "averageAnalystRating", "priceEpsCurrentYear"
                    ],
                    "4. Dividend Information": [
                        "dividendRate", "dividendYield", "trailingAnnualDividendRate",
                        "trailingAnnualDividendYield", "fiveYearAvgDividendYield",
                        "lastDividendValue", "lastDividendDate", "dividendDate", "exDividendDate"
                    ],
                    "5. Earnings & Financials": [
                        "trailingEps", "forwardEps", "epsTrailingTwelveMonths", "epsForward",
                        "epsCurrentYear", "netIncomeToCommon", "totalRevenue", "revenuePerShare",
                        "grossProfits", "ebitda", "operatingCashflow", "freeCashflow",
                        "totalCash", "totalCashPerShare", "totalDebt", "bookValue",
                        "earningsGrowth", "revenueGrowth", "earningsQuarterlyGrowth",
                        "returnOnAssets", "returnOnEquity", "grossMargins", "operatingMargins",
                        "ebitdaMargins", "quickRatio", "currentRatio", "financialCurrency"
                    ],
                    "6. Share Statistics": [
                        "sharesOutstanding", "impliedSharesOutstanding", "floatShares",
                        "sharesShort", "sharesShortPriorMonth", "sharesShortPreviousMonthDate",
                        "dateShortInterest", "sharesPercentSharesOut", "shortPercentOfFloat",
                        "shortRatio", "heldPercentInsiders", "heldPercentInstitutions"
                    ],
                    "7. Analyst Price Targets": [
                        "targetHighPrice", "targetLowPrice", "targetMeanPrice", "targetMedianPrice"
                    ],
                    "8. Technical Indicators": [
                        "fiftyDayAverage", "fiftyDayAverageChange", "fiftyDayAverageChangePercent",
                        "twoHundredDayAverage", "twoHundredDayAverageChange",
                        "twoHundredDayAverageChangePercent", "fiftyTwoWeekLow", "fiftyTwoWeekHigh",
                        "fiftyTwoWeekLowChange", "fiftyTwoWeekLowChangePercent",
                        "fiftyTwoWeekHighChange", "fiftyTwoWeekHighChangePercent",
                        "fiftyTwoWeekRange", "fiftyTwoWeekChange", "fiftyTwoWeekChangePercent",
                        "SandP52WeekChange", "52WeekChange"
                    ],
                    "9. Earnings Call / Fiscal Calendar": [
                        "lastFiscalYearEnd", "nextFiscalYearEnd", "mostRecentQuarter",
                        "earningsTimestamp", "earningsTimestampStart", "earningsTimestampEnd",
                        "earningsCallTimestampStart", "earningsCallTimestampEnd",
                        "isEarningsDateEstimate"
                    ],
                    "10. Other": [
                        "lastSplitFactor", "lastSplitDate", "firstTradeDateMilliseconds",
                        "esgPopulated", "cryptoTradeable", "tradeable", "sourceInterval"
                    ]
                }
                
                # Start from row 7
                current_row = 7
                
                # Process each section
                for section_name, fields in sections.items():
                    # Write section header with light green background
                    header_cell = profile_sheet.cells(current_row, 2)
                    header_cell.value = section_name
                    header_range = profile_sheet.range(f"B{current_row}:C{current_row}")
                    header_range.color = "#A7D9AB"  # Light green color
                    current_row += 1
                    
                    # Write table headers
                    profile_sheet.cells(current_row, 2).value = "Metric"
                    profile_sheet.cells(current_row, 3).value = "Value"
                    current_row += 1
                    
                    # Track the start row for this section's data
                    section_start_row = current_row
                    
                    # Write section data
                    for field in fields:
                        if field in main_info and main_info[field] is not None:
                            profile_sheet.cells(current_row, 2).value = field
                            profile_sheet.cells(current_row, 3).value = main_info[field]
                            current_row += 1
                    
                    # Format as table
                    if current_row > section_start_row:
                        table_range = profile_sheet.range(f"B{section_start_row-1}:C{current_row-1}")
                        profile_sheet.tables.add(table_range)
                    
                    # Add spacing between sections
                    current_row += 2
                
                # Add spacing before officers table
                current_row += 2
                
                # Process company officers information
                officers = ticker_data['officers']
                if officers:
                    # Write officers section header with light green background
                    header_cell = profile_sheet.cells(current_row, 2)
                    header_cell.value = "11. Key Executives"
                    header_range = profile_sheet.range(f"B{current_row}:E{current_row}")
                    header_range.color = "#A7D9AB"  # Light green color
                    current_row += 1
                    
                    # Write officers table headers
                    profile_sheet.cells(current_row, 2).value = "Name"
                    profile_sheet.cells(current_row, 3).value = "Title"
                    profile_sheet.cells(current_row, 4).value = "Age"
                    profile_sheet.cells(current_row, 5).value = "Total Pay"
                    
                    # Track the start row for officers table
                    officers_start_row = current_row
                    current_row += 1
                    
                    for officer in officers:
                        profile_sheet.cells(current_row, 2).value = officer.get('name', 'N/A')
                        profile_sheet.cells(current_row, 3).value = officer.get('title', 'N/A')
                        profile_sheet.cells(current_row, 4).value = officer.get('age', 'N/A')
                        profile_sheet.cells(current_row, 5).value = officer.get('totalPay', 'N/A')
                        current_row += 1
                    
                    # Format officers as table including the headers row
                    if current_row > officers_start_row:
                        table_range = profile_sheet.range(f"B{officers_start_row}:E{current_row-1}")
                        profile_sheet.tables.add(table_range)
                
                print(f"✅ Successfully processed detailed information")
            else:
                profile_sheet["B8"].value = f"N/A - {ticker_data['error']}"
        else:
            profile_sheet["B8"].value = "N/A - No data found for ticker"
    else:
        profile_sheet["B8"].value = "N/A - Service temporarily unavailable"
    
    print("✓ ENDING get_profile ✓")

@script
def create_technicals(book: xw.Book):
    """Create technical indicators for a ticker using price data."""
    print("▶ STARTING create_technicals ◀")
    
    # Get the PRICES sheet
    prices_sheet = book.sheets["PRICES"]
    
    # Read ticker, start_date, and end_date from PRICES sheet
    ticker = str(prices_sheet["B3"].value).strip().upper() if prices_sheet["B3"].value else None
    start_date_raw = prices_sheet["D3"].value
    end_date_raw = prices_sheet["F3"].value
    
    if not all([ticker, start_date_raw, end_date_raw]):
        prices_sheet["B8"].value = "Please enter ticker (B3), start date (D3), and end date (F3)"
        return
    
    # Convert Excel dates to YYYY-MM-DD format
    if isinstance(start_date_raw, (datetime, pd.Timestamp)):
        start_date = start_date_raw.strftime("%Y-%m-%d")
    else:
        # Convert Excel serial number to datetime
        start_date = pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(start_date_raw))
        start_date = start_date.strftime("%Y-%m-%d")
    
    if isinstance(end_date_raw, (datetime, pd.Timestamp)):
        end_date = end_date_raw.strftime("%Y-%m-%d")
    else:
        # Convert Excel serial number to datetime
        end_date = pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(end_date_raw))
        end_date = end_date.strftime("%Y-%m-%d")
    
    print(f"📊 Fetching prices for ticker: {ticker}")
    
    # Use the get-all-prices endpoint
    api_url = f"https://yfin.hosting.tigzig.com/get-all-prices/?tickers={ticker}&start_date={start_date}&end_date={end_date}"
    print(f"🔗 API URL: {api_url}")
    
    # Make API request
    response = requests.get(api_url)
    print(f"📥 Response status: {response.status_code}")
    
    if response.ok:
        # Parse JSON response
        data = response.json()
        
        if isinstance(data, dict) and not data.get("error"):
            # Convert nested JSON to DataFrame
            rows = []
            for date, ticker_data in data.items():
                if ticker in ticker_data:
                    row = ticker_data[ticker]
                    row['Date'] = date  # Add date to the row
                    rows.append(row)
            
            # Create DataFrame with all price columns
            df = pd.DataFrame(rows)
            
            # Debug print DataFrame info
            print("\n📊 DataFrame Info:")
            print(df.info())
            print("\n📊 DataFrame Head:")
            print(df.head())
            
            # Reorder columns to put Date first and ensure column names match finta requirements
            # First convert to lowercase for finta compatibility
            df.columns = [col.lower() for col in df.columns]
            cols = ['date'] + [col for col in df.columns if col != 'date']
            df = df[cols]
            
            # Convert Date to datetime for proper sorting
            df['date'] = pd.to_datetime(df['date'])
            df = df.sort_values('date')
            
            print("\n📊 DataFrame Columns after sorting and renaming:")
            print(df.columns.tolist())
            
            # Calculate technical indicators using finta
            try:
                print("\n🔄 Calculating indicators...")
                
                # Store a copy of the DataFrame for display purposes
                display_df = df.copy()
                
                # 1. EMA - Exponential Moving Average (12 days)
                print("Calculating EMA-12...")
                ema12 = TA.EMA(df, 12)
                display_df['EMA_12'] = ema12
                
                # 2. EMA - Exponential Moving Average (26 days)
                print("Calculating EMA-26...")
                ema26 = TA.EMA(df, 26)
                display_df['EMA_26'] = ema26
                
                # 3. RSI - Relative Strength Index (14 periods)
                print("Calculating RSI...")
                rsi = TA.RSI(df)
                display_df['RSI_14'] = rsi
                
                # 4. ROC - Rate of Change (14 periods)
                print("Calculating ROC...")
                roc = TA.ROC(df, 14)
                display_df['ROC_14'] = roc
                
                # 5. MACD - Moving Average Convergence Divergence (12/26)
                print("Calculating MACD...")
                macd = TA.MACD(df)  # Using default 12/26 periods
                if isinstance(macd, pd.DataFrame):
                    print("MACD columns:", macd.columns.tolist())
                    display_df['MACD_12_26'] = macd['MACD']
                    display_df['MACD_SIGNAL_9'] = macd['SIGNAL']
                
                # 6. Bollinger Bands (20 periods, 2 standard deviations)
                print("Calculating Bollinger Bands...")
                bb = TA.BBANDS(df)
                if isinstance(bb, pd.DataFrame):
                    print("BB columns:", bb.columns.tolist())
                    display_df['BBANDS_UPPER_20_2'] = bb['BB_UPPER']
                    display_df['BBANDS_MIDDLE_20_2'] = bb['BB_MIDDLE']
                    display_df['BBANDS_LOWER_20_2'] = bb['BB_LOWER']
                
                # 7. Stochastic Oscillator
                print("Calculating Stochastic Oscillator...")
                try:
                    stoch = TA.STOCH(df)
                    print("STOCH type:", type(stoch))
                    if isinstance(stoch, pd.DataFrame):
                        print("STOCH columns:", stoch.columns.tolist())
                    display_df['STOCH_K'] = stoch['K'] if isinstance(stoch, pd.DataFrame) else stoch
                    display_df['STOCH_D'] = stoch['D'] if isinstance(stoch, pd.DataFrame) else None
                except Exception as stoch_error:
                    print(f"Stochastic calculation error: {str(stoch_error)}")
                
                # 8. Williams %R
                print("Calculating Williams %R...")
                williams = TA.WILLIAMS(df)
                display_df['WILLIAMS_R'] = williams
                
                # Capitalize the basic price columns for display
                display_df.rename(columns={
                    'date': 'DATE',
                    'open': 'OPEN',
                    'high': 'HIGH',
                    'low': 'LOW',
                    'close': 'CLOSE',
                    'volume': 'VOLUME',
                    'adj close': 'ADJ_CLOSE',
                    'dividends': 'DIVIDENDS',
                    'stock splits': 'STOCK_SPLITS',
                    'bb_upper': 'BBANDS_UPPER_20_2',
                    'bb_middle': 'BBANDS_MIDDLE_20_2',
                    'bb_lower': 'BBANDS_LOWER_20_2'
                }, inplace=True)
                
                print("\n📊 Final DataFrame Head:")
                print(display_df.head())
                print("\n📊 Final DataFrame Columns:")
                print(display_df.columns.tolist())
                
                # Display headers and data
                # Clear existing content first
                prices_sheet["B7:Z1000"].clear_contents()
                
                # Write headers in row 7
                prices_sheet["B7"].value = display_df.columns.tolist()
                
                # Write data starting from row 8
                prices_sheet["B8"].value = display_df.values.tolist()
                
                # Format as table including headers
                data_range = prices_sheet["B7"].resize(len(display_df) + 1, len(display_df.columns))
                prices_sheet.tables.add(data_range)
                
                print("\n✅ Successfully added technical indicators:")
                print("📊 Added: EMA_12, EMA_26, RSI_14, ROC_14, MACD_12_26, MACD_SIGNAL_9, BB_UPPER/MIDDLE/LOWER, STOCH_K/D, WILLIAMS_R")
            except Exception as e:
                print(f"\n❌ Error calculating indicators: {str(e)}")
                print("Error type:", type(e))
                print("Error args:", e.args)
                print("Current DataFrame columns:", df.columns.tolist())
                prices_sheet["B8"].value = f"Error calculating technical indicators: {str(e)}"
        else:
            error_msg = data.get("error") if isinstance(data, dict) else "Invalid response format"
            prices_sheet["B8"].value = f"N/A - {error_msg}"
    else:
        prices_sheet["B8"].value = "N/A - Service temporarily unavailable"
    
    print("✓ ENDING create_technicals ✓")

@script
def get_technical_analysis_from_gemini(book: xw.Book):
    """Get technical analysis from Gemini API based on price data and technical indicators."""
    print("▶ STARTING get_technical_analysis_from_gemini ◀")
    
    # Get the TA sheet
    ta_sheet = book.sheets["TA"]
    
    # Read model name and API key from TA sheet
    model_name = ta_sheet["B3"].value
    api_key = ta_sheet["B4"].value
    
    if not all([model_name, api_key]):
        ta_sheet["B8"].value = "Please enter model name (B3) and API key (B4)"
        return
    
    # First get the price data and technical indicators
    print("📊 Getting price data and technical indicators...")
    create_technicals(book)
    
    # Get the PRICES sheet data
    prices_sheet = book.sheets["PRICES"]
    data_range = prices_sheet["B7"].expand()
    data = data_range.options(pd.DataFrame, index=False).value
    
    # Get the latest data point
    latest_data = data.iloc[-1]
    
    # Create a prompt for Gemini
    prompt = f"""
    [SYSTEM INSTRUCTIONS]
    Generate a technical analysis report in clean markdown format. Follow these guidelines:

    1. Use proper markdown headers with single # for main title and ## for sections
    2. Use bold (**) for metric names and values
    3. Keep paragraphs clean without extra markdown symbols
    4. Use simple bullet points with single hyphens
    5. No special formatting for analysis sections
    6. Keep the format clean and consistent
    7. No investment advice or recommendations
    8. Focus on technical analysis only
    9. For each technical indicator, provide a detailed 100-150 word analysis that covers:
       - Current indicator values and what they suggest
       - Recent trends and patterns
       - Potential technical implications
       - Relationship with other relevant indicators
       - Historical context when relevant

    Generate the report with this structure:

    # Technical Analysis - [DATE]

    ## Price Action Overview
    [Two paragraphs analyzing price action and market behavior]

    ## Current Session Data
    **Close**: **{latest_data['CLOSE']}** | **Open**: **{latest_data['OPEN']}** | **High**: **{latest_data['HIGH']}** | **Low**: **{latest_data['LOW']}** | **Volume**: **{latest_data['VOLUME']}**

    ## Technical Indicators

    ### 1. Exponential Moving Averages (EMA)
    - **EMA_12**: **{latest_data['EMA_12']}**
    - **EMA_26**: **{latest_data['EMA_26']}**

    Analysis: [Provide 100-150 word analysis of EMA relationships, trends, crossovers, and potential support/resistance levels. Include both short-term and medium-term perspectives.]

    ### 2. Relative Strength Index (RSI)
    - **RSI_14**: **{latest_data['RSI_14']}**

    Analysis: [Provide 100-150 word analysis of RSI, including current momentum, overbought/oversold conditions, potential divergences, and trend strength implications.]

    ### 3. Rate of Change (ROC)
    - **ROC_14**: **{latest_data['ROC_14']}**

    Analysis: [Provide 100-150 word analysis of ROC, examining momentum strength, trend confirmation, potential reversals, and relationship with price action.]

    ### 4. MACD
    - **MACD_12_26**: **{latest_data['MACD_12_26']}**
    - **MACD_SIGNAL_9**: **{latest_data['MACD_SIGNAL_9']}**

    Analysis: [Provide 100-150 word analysis of MACD, including signal line crossovers, histogram patterns, divergences, and trend strength confirmation.]

    ### 5. Bollinger Bands
    - **Upper Band (20,2)**: **{latest_data['BBANDS_UPPER_20_2']}**
    - **Middle Band (20,2)**: **{latest_data['BBANDS_MIDDLE_20_2']}**
    - **Lower Band (20,2)**: **{latest_data['BBANDS_LOWER_20_2']}**

    Analysis: [Provide 100-150 word analysis of Bollinger Bands, including price position relative to bands, bandwidth trends, potential mean reversion, and volatility patterns.]

    ### 6. Stochastic Oscillator
    - **STOCH_K**: **{latest_data['STOCH_K']}**
    - **STOCH_D**: **{latest_data['STOCH_D']}**

    Analysis: [Provide 100-150 word analysis of Stochastic Oscillator, examining overbought/oversold conditions, crossovers, divergences, and momentum confirmation.]

    ### 7. Williams %R
    - **WILLIAMS_R**: **{latest_data['WILLIAMS_R']}**

    Analysis: [Provide 100-150 word analysis of Williams %R, including current market position, potential reversals, confirmation of trends, and relationship with price momentum.]
    """
    
    # Prepare API payload
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.7,
            "maxOutputTokens": 7500
        }
    }
    
    # Make API call
    api_url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
    
    response = requests.post(
        api_url,
        headers={"Content-Type": "application/json"},
        json=payload
    )
    
    if response.status_code == 200:
        response_json = response.json()
        if 'candidates' in response_json:
            analysis = response_json['candidates'][0]['content']['parts'][0]['text']
            
            # Insert our note at the beginning of the analysis with a proper heading
            disclaimer_note = """## Important Disclaimer

*The content in this report is for xlwings Lite demo purposes only and is not investment research, investment analysis, or financial advice. This is a technical demonstration of how to use xlwings Lite to send data to an LLM/AI via a web API call, receive a markdown-formatted response, convert it to PDF, and display the downloadable link in a cell.*

"""
            
            # Remove any existing note if it was included in the response
            if analysis.startswith("*") or analysis.startswith("#"):
                # Find the first technical analysis header
                header_start = analysis.find("# Technical Analysis")
                if header_start != -1:
                    analysis = analysis[header_start:]
            
            # Clean up any potential markdown formatting issues
            analysis = analysis.replace("**Close**:", "**Close**:").replace("**Open**:", "**Open**:")
            analysis = analysis.replace("*Analysis:*", "Analysis:")
            
            # Combine the note with the analysis and ensure proper markdown
            final_markdown = f"{disclaimer_note}{analysis}"
            
            # Convert markdown to PDF using the API
            try:
                pdf_api_url = "https://mdpdf.tigzig.com/text-input"
                pdf_payload = {
                    "text": final_markdown
                }
                
                print("\n📄 Converting markdown to PDF...")
                print(f"Sending request to: {pdf_api_url}")
                
                # First try to get JSON response with URL
                pdf_response = requests.post(
                    pdf_api_url,
                    headers={
                        "Content-Type": "application/json",
                        "Accept": "application/json"
                    },
                    json=pdf_payload
                )
                
                print(f"Response status: {pdf_response.status_code}")
                print(f"Response headers: {dict(pdf_response.headers)}")
                
                if pdf_response.status_code == 200:
                    response_data = pdf_response.json()
                    pdf_url = response_data.get('pdf_url')
                    
                    if pdf_url:
                        # Write the label and URL
                        ta_sheet["B7"].value = "PDF URL:"
                        ta_sheet["B8"].value = pdf_url
                        
                        print("\n✅ Technical analysis URL saved!")
                        print(f"PDF URL: {pdf_url}")
                    else:
                        error_msg = "No PDF URL in response"
                        print(f"❌ {error_msg}")
                        ta_sheet["B7"].value = error_msg
                else:
                    error_msg = f"PDF conversion failed with status {pdf_response.status_code}"
                    print(f"❌ {error_msg}")
                    if pdf_response.text:
                        print(f"Error response: {pdf_response.text}")
                    ta_sheet["B7"].value = error_msg
            except Exception as e:
                error_msg = f"Error converting to PDF: {str(e)}"
                print(f"❌ {error_msg}")
                ta_sheet["B7"].value = error_msg
        else:
            ta_sheet["B7"].value = "Error: No response from Gemini API"
    else:
        ta_sheet["B7"].value = f"Error: API call failed with status {response.status_code}"
        print(f"❌ API call failed with status {response.status_code}")
        print("Error:", response.text)
    
    print("✓ ENDING get_technical_analysis_from_gemini ✓")
