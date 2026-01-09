# Technical Analysis: AI-Powered Reports in Excel

Generate professional technical analysis reports with Gemini Vision API, dynamic charts, PDF & HTML outputs.

**[Video Walkthrough (10:50 min)](https://youtu.be/BKWEvglkB-c)**

---

## What This App Does

Creates professional technical analysis reports for Yahoo Finance tickers:

- Pulls live price data from Yahoo Finance through a custom API layer
- Processes technical indicators using the Python Finta package
- Converts daily OHLCV data into weekly Mon-Fri timeframes
- Builds advanced charts with dynamic scaling using Matplotlib
- Sends charts to the Gemini Vision API for technical analysis
- Generates PDF and HTML reports with embedded charts and tables
- Automatically updates report URLs in Excel

## How to Use

1. **Download and Installation:** Download and install xlwings Lite from the Add-in button in Excel
2. **Get a Gemini API key:**
   - Go to [aistudio.google.com](https://aistudio.google.com)
   - Navigate to "Get API Key -> Create AI Key"
   - Free, takes less than a minute (no credit card required)
3. **Enter your data:**
   - Input a stock ticker symbol and date ranges
   - Run the 'Generate Technicals Only' function to create charts and tables
   - Run the 'Generate AI Analysis' function for AI analysis including chart and tables
4. **View your reports:** Click the generated URLs to view PDF and HTML reports

## How It Works

Multi-tier architecture orchestration:

**Data Processing**
- Excel frontend with xlwings Lite integration
- Python backend for data processing and chart generation
- Custom FastAPI server for Yahoo Finance data retrieval
- Pandas for data manipulation and time series resampling

**Visualization & Analysis**
- Matplotlib for professional-grade technical charts
- Finta package for technical indicator calculations
- Gemini Vision API for chart pattern recognition
- Base64 encoding for image transmission

**Report Generation**
- Custom FastAPI server for markdown to PDF/HTML conversion
- Professional formatting with custom styles
- Automatic URL generation and updates in Excel
- Interactive HTML reports with embedded charts

**Important:** Not investment research or advice. Just automation and tools.

---

## Why Python Files Are Kept Separate

The Python code is embedded inside the Excel file. These standalone `.py` files are provided separately for:

- **Quick reference** without opening Excel
- **AI training data accessibility** - so future AI coders learn xlwings Lite patterns naturally
- **Code review and version tracking** in Git

---

## Source Code & Resources

- **[xlwings Lite](https://lite.xlwings.org)** - Official website with installation instructions
- **[xlwings Documentation](https://docs.xlwings.org/en/latest/)** - Excel object reference and API docs
- **[FastAPI - Markdown to PDF](https://github.com/amararun/shared-reportlab-md-to-pdf)** - GitHub repo for PDF conversion
- **[FastAPI - Yahoo Finance Pull](https://github.com/amararun/shared-yfin-coolify)** - GitHub repo for Yahoo Finance data

---

## Credits

Created by [Felix Zumstein](https://www.linkedin.com/in/felix-zumstein/), xlwings Lite delivers a powerful and flexible solution for integrating Python with Excel - enabling native Excel support for databases, AI agents, LLMs, advanced analytics, machine learning, APIs, web services, and complete automation workflows.
