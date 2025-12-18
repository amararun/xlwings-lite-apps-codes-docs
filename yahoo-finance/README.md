# Yahoo Finance Analyzer: Stock Data to AI-Powered Reports

Pull stock data from Yahoo Finance, perform technical analysis, and generate AI-powered PDF reports.

**[Video Walkthrough (8:33 min)](https://www.youtube.com/watch?v=nnsO8XmLYuk)**

---

## What This App Does

Pull stock data, perform technical analysis, and generate AI reports:

- Pull stock profiles from Yahoo Finance
- Retrieve historical prices and financials
- Compute technical indicators (EMA, RSI, MACD, Bollinger)
- AI-powered analysis via Gemini
- Generate styled PDF reports
- Financial statements (P&L, Balance Sheet, Cash Flow)

## How to Use

1. **Download & Install:** Get the app and install xlwings Lite
2. **Get Gemini API key:** Free at [aistudio.google.com](https://aistudio.google.com)
3. **Built-in Functions:**
   - `get_profile` - Company info
   - `get_prices` - Historical prices
   - `create_technicals` - Calculate indicators
   - `get_technical_analysis_from_gemini` - AI analysis
4. **Enter tickers:** Yahoo symbols (AAPL, ^NSEI, etc.)

## How It Works

**Data Retrieval:**
- xlwings Lite calls custom FastAPI server
- FastAPI uses yfinance to fetch Yahoo Finance data
- Server handles CORS issues from browser environment

**Technical Analysis:**
- finta package calculates EMA, RSI, MACD
- Indicators added as new columns to price data
- Data prepared for AI analysis

**AI Reporting:**
- Price data + indicators sent to Gemini
- Markdown output styled via FastAPI server
- Converted to PDF via md2pdf service

**Disclaimer:** YLENS is a working prototype demonstrating xlwings Lite capabilities. Not investment advice. AI/automation in live use requires iteration, validation, and human judgment.

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
- **[FastAPI - Markdown to PDF](https://github.com/amararun/shared-markdown-to-pdf)** - GitHub repo for PDF conversion
- **[FastAPI - Yahoo Finance](https://github.com/amararun/shared-yfin-coolify)** - GitHub repo for yfinance data service

---

## Credits

Created by [Felix Zumstein](https://www.linkedin.com/in/felix-zumstein/), xlwings Lite delivers a powerful and flexible solution for integrating Python with Excel - enabling native Excel support for databases, AI agents, LLMs, advanced analytics, machine learning, APIs, web services, and complete automation workflows.
