# INTELISCAPE-X: Intelligent Web Scraping

Transform Excel into a versatile web scraping platform that extracts structured data from multiple web pages using AI-powered analysis.

**[Video Walkthrough (0:33 min)](https://www.youtube.com/watch?v=41ZX46DibV4)**

---

## What This App Does

This xlwings Lite app turns Excel into a web scraping platform:

- Extract structured data from multiple websites
- Process batches of URLs with custom column specifications
- Use AI to identify and extract data elements from web pages
- Apply custom filtering instructions to refine results
- Track processing status and errors automatically
- Monitor performance metrics and token usage
- Format results directly into Excel tables

## How to Use

1. **Install xlwings Lite** from the Add-in button in Excel
2. **Configure the MASTER sheet:**
   - Add your Jina and Gemini API keys (or set environment variables `JINA_API_KEY` and `GEMINI_API_KEY`)
   - Select your Gemini model (e.g., `gemini-2.5-flash`)
   - Set scraping parameters (request timeout, max retries, delays)
   - Configure LLM parameters (temperature, topP, output tokens, thinking budget)
3. **Define extraction columns in COLUMN_INPUTS:**
   - Specify column names and descriptions for the data you want to extract
   - Add custom filtering instructions if needed
4. **Add target URLs in URL_LIST:** Enter the URLs you want to scrape
5. **Run the scraper:** Execute `scrape_urls_from_list` from the xlwings tab
6. **Review results in four sheets:**
   - **DATA:** Extracted structured data with your defined columns
   - **ERROR_LOG:** Detailed error records with timestamps and types
   - **DASHBOARD:** Performance metrics, token usage, and processing statistics
   - **URL_LIST:** Updated with processing status and timestamps

## How It Works

The web scraper processes URLs through a three-stage pipeline:

**Stage 1: Configuration and Setup**
- xlwings Lite runs Python code directly in Excel via WebAssembly
- Reads configuration from MASTER sheet (API keys, model settings, processing parameters)
- Loads column definitions and custom filtering instructions from COLUMN_INPUTS
- Reads URLs from URL_LIST and filters out already-processed entries

**Stage 2: Data Extraction**
- Sends each URL to Jina API to fetch and render the webpage as markdown
- Passes rendered content to Gemini API with your column specifications
- Gemini extracts structured data matching your column definitions and filtering criteria
- Updates URL_LIST with status and timestamp after each URL

**Stage 3: Results and Reporting**
- Writes extracted data to DATA sheet as a formatted table
- Logs errors to ERROR_LOG sheet with timestamps, error types, and messages
- Generates DASHBOARD sheet with performance metrics

### Use Cases
- Product data collection from e-commerce sites
- Contact list building from business directories
- Price monitoring and competitor analysis
- Statistics aggregation from sports or finance websites

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
- **[Jina AI Web Scraping API](https://jina.ai/api-dashboard/reader)** - Dashboard for web page rendering
- **[Google Gemini API](https://ai.google.dev/gemini-api/docs/structured-output?lang=rest)** - AI-powered data extraction

---

## Credits

Created by [Felix Zumstein](https://www.linkedin.com/in/felix-zumstein/), xlwings Lite delivers a powerful and flexible solution for integrating Python with Excel - enabling native Excel support for databases, AI agents, LLMs, advanced analytics, machine learning, APIs, web services, and complete automation workflows.
