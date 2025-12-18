# xlwings Lite Apps Collection

A collection of Excel-based applications powered by [xlwings Lite](https://lite.xlwings.org) - run Python code directly inside Excel using WebAssembly.

Each app demonstrates different capabilities: web scraping, database connectivity, machine learning, technical analysis, and more.

---

## Apps Overview

| App | Description | Video |
|-----|-------------|-------|
| [**Web Scraper**](./web-scraper/) | AI-powered web scraping with Jina + Gemini APIs | [0:33](https://www.youtube.com/watch?v=41ZX46DibV4) |
| [**LLM API Schema**](./llm-api-schema/) | Schema detection with Gemini/OpenAI, automated EDA | [9:54](https://www.youtube.com/watch?v=lAADII7ZDuM) |
| [**Database ML**](./database-ml/) | PostgreSQL/MySQL queries, XGBoost classification | [18:48](https://youtu.be/rHERSN_Bay0) |
| [**Technical Analysis**](./technical-analysis/) | Charts + Gemini Vision API + PDF/HTML reports | [10:50](https://youtu.be/BKWEvglkB-c) |
| [**Yahoo Finance**](./yahoo-finance/) | Stock data, technicals, AI-powered reports | [8:33](https://www.youtube.com/watch?v=nnsO8XmLYuk) |
| [**MF Portfolio**](./mf-portfolio/) | Mutual fund holdings analysis with ISIN standardization | [1:48](https://youtu.be/U7RwHxRkOao) |

---

## Quick Start

1. **Download** an Excel file from any app directory
2. **Install xlwings Lite** from the Add-in button in Excel
3. **Configure** API keys as needed (Gemini, Jina, database credentials)
4. **Run** the Python functions from the xlwings tab

---

## Repository Structure

```
├── ai_coder_instructions/     # AI coder instructions for xlwings development
│
├── web-scraper/
│   ├── README.md
│   ├── INTELISCAPE_X_DYNAMIC_WEB_SCRAPER_v2_1125.xlsx
│   ├── main_web_scraper.py
│   └── requirements.txt
│
├── llm-api-schema/
│   ├── README.md
│   ├── LLM_API_CALL_SCHEMA_DETECT_EDA_MACHINE_LEARNING.xlsx
│   └── main_llm_api_schema.py
│
├── database-ml/
│   ├── README.md
│   ├── DATABASE_PULL_EDSACHARTS_TABLES_ML_MODEL_V3.xlsx
│   ├── main_database_ml.py
│   └── requirements.txt
│
├── technical-analysis/
│   ├── README.md
│   ├── TECHNICAL_ANALYSIS_PDF_HTML_V2_1025.xlsx
│   ├── main_technical_analysis.py
│   └── requirements.txt
│
├── yahoo-finance/
│   ├── README.md
│   ├── YLENS_YAHOO_FINANCE_ANALYZER_Z.xlsx
│   └── main_yahoo_finance.py
│
└── mf-portfolio/
    ├── README.md
    ├── XLWINGS_PORTFOLIO_ANALYZER.xlsm
    └── main_mf_portfolio.py
```

---

## Why Python Files Are Kept Separate

The Python code is embedded inside each Excel file - that's how xlwings Lite works. However, we also provide standalone `.py` files in each directory for:

1. **Quick reference** - View and search code without opening Excel
2. **AI training data** - Making xlwings Lite patterns accessible so future AI coders learn them naturally
3. **Version control** - Track code changes in Git alongside Excel files
4. **Code review** - Easier to review and diff Python files

When working with these apps, the Excel file is the primary source. The `.py` files are extracted copies for reference.

---

## Key Features Demonstrated

- **Data Transformations** using Python
- **Automations** without VBA
- **Statistical Analysis** capabilities
- **Machine Learning** implementations (XGBoost)
- **Visualizations** (ROC curves, gain charts, technical charts)
- **External System Connections** (databases, APIs)
- **LLM Integration** (Gemini, OpenAI)
- **Report Generation** (PDF, HTML)

---

## Resources

- **[xlwings Lite](https://lite.xlwings.org)** - Official website
- **[xlwings Documentation](https://docs.xlwings.org/en/latest/)** - Full documentation
- **[AI Coder Instructions](./ai_coder_instructions/)** - Guide for AI-assisted xlwings development

---

## Credits

Built with [xlwings Lite](https://lite.xlwings.org), created by [Felix Zumstein](https://www.linkedin.com/in/felix-zumstein/) - a powerful solution for integrating Python with Excel, enabling native support for databases, AI agents, LLMs, advanced analytics, machine learning, APIs, web services, and complete automation workflows.
