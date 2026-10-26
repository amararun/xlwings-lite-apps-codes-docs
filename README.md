# xlwings Lite Apps Collection

Excel tools built with [xlwings Lite](https://lite.xlwings.org), created by [Felix Zumstein](https://www.linkedin.com/in/felix-zumstein/). xlwings Lite runs Python inside Excel.

Every tool here was updated in October 2026 for the current xlwings Lite (add-in build 1.0.0.104, Python 3.14, pandas 3). Each folder has the workbook, its Python code as a plain file, and a short README. The same workbooks are on [tigzig.com/xlwings-starter](https://www.tigzig.com/xlwings-starter).

---

## Tools

| Folder | What it does |
|---|---|
| [web-scraper](./web-scraper/) | AI web scraper: describe the columns in plain English, get a table from a list of web pages |
| [llm-api-schema](./llm-api-schema/) | An LLM (Gemini or any OpenRouter model) detects a table's schema, then Python runs the EDA |
| [database-ml](./database-ml/) | Query a Neon Postgres database over HTTPS, then EDA and an XGBoost response model |
| [technical-analysis](./technical-analysis/) | Prices, indicators and charts for any Yahoo Finance symbol, and an AI-written report (PDF and HTML) |
| [yahoo-finance](./yahoo-finance/) | YLENS: company profile, financial statements, prices with indicators, and an AI-written report |
| [mf-portfolio](./mf-portfolio/) | Compares what a set of Indian mutual funds held in two months |
| [data-importer](./data-importer/) | Imports files from Google Drive, Dropbox, GitHub or any link into DuckDB, or downloads them as-is |
| [sp500-tracker](./sp500-tracker/) | Earnings dates and daily prices for hundreds of tickers, pulled in batches |
| [local-file-patterns](./local-file-patterns/) | Reads and writes files in a folder on your computer: zip, CSV, Parquet, HTML, PDF, PowerPoint |
| [wingman](./wingman/) | Examples of the =WINGMAN() function built into xlwings Lite |

Also in this repo:

- [ai_coder_instructions](./ai_coder_instructions/): the instructions file I give an AI coder before it writes xlwings Lite code. The latest copy is always at [tigzig.com/files/xlwings/AI_CODER_INSTRUCTIONS.md](https://www.tigzig.com/files/xlwings/AI_CODER_INSTRUCTIONS.md).
- [remote_modules](./remote_modules/): the statistics modules the Data Importer loads at run time.
- [cloudflare-workers](./cloudflare-workers/): the CORS proxy the Data Importer uses for Google Drive, Dropbox and GitHub Releases links.

---

## How to run a tool

1. Download the workbook from its folder.
2. Open it in Excel and start xlwings Lite (Home > Add-ins).
3. If the tool needs an API key, add it in xlwings Lite > Settings > Environment Variables, scope Local. The folder README says which one.
4. Click the buttons in order. Most workbooks have a HOW_IT_WORKS sheet that explains each step.

---

## Why the Python files are kept separately

The Python code is inside each Excel file; that is how xlwings Lite works. The `.py` file in each folder is a plain copy of that code, so you can read and search it without opening Excel, and so AI coders can find working xlwings Lite code. The workbook is the source; the `.py` file is extracted from it.

---

## Resources

- [xlwings Lite](https://lite.xlwings.org): official website and documentation
- [xlwings documentation](https://docs.xlwings.org/en/stable/)
- [tigzig.com/xlwings-starter](https://www.tigzig.com/xlwings-starter): practice workbooks and all the tools above

## Author

Built by [Amar Harolikar](https://www.linkedin.com/in/amarharolikar/). More open source tools for analytics, databases and automation at [tigzig.com](https://www.tigzig.com).
