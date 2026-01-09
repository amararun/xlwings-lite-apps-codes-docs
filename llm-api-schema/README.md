# AI Schema Detection: Automate Workflows with LLM-Powered Analysis

Connect Excel directly to LLM APIs for AI-powered schema detection, automated workflows, and exploratory data analysis.

**[Video Walkthrough (9:54 min)](https://www.youtube.com/watch?v=lAADII7ZDuM)**

---

## What This App Does

Connect Excel directly to LLM APIs for AI-powered analysis:

- Connect to **Gemini (2.0-Flash)** or **OpenAI (GPT-4o)**
- Automatic schema detection with structured JSON output
- Identify categorical and numerical variables
- Generate SQL-compatible data types
- AI-guided exploratory data analysis (EDA)
- Data visualizations based on detected schema

## How to Use

1. **Download & Install:** Get the app and install xlwings Lite
2. **Add API keys:**
   - **Gemini:** Free at [aistudio.google.com](https://aistudio.google.com)
   - **OpenAI:** At [platform.openai.com](https://platform.openai.com)
3. **Built-in Functions:**
   - `analyze_table_schema_gemini` - Detect column types with Gemini
   - `analyze_table_schema_openai` - Detect column types with OpenAI
   - `perform_eda` - EDA with visualizations

## How It Works

**LLM API Integration:**
- Samples your data table and sends to chosen LLM
- Crafted prompt instructs LLM to return structured JSON
- JSON response parsed and formatted for Excel display
- Best performance with Gemini's Flash model

**Automated Workflows:**
- Detected schema configures EDA operations
- Numeric columns: statistical analysis
- Categorical columns: frequency distributions
- Enables automation of database operations

### Practical Use Cases
- Automated workflows connected to API backends
- Web scraping with structured output
- Text classification and summarization
- Text-to-SQL with Excel frontend
- Data preparation for database uploads

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
- **[Mutual Fund Processor](https://mf.tigzig.com)** - Live app using schema detection
- **[TIGZIG Co-Analyst](https://rexdb.tigzig.com)** - Database connector with schema detection

---

## Credits

Created by [Felix Zumstein](https://www.linkedin.com/in/felix-zumstein/), xlwings Lite delivers a powerful and flexible solution for integrating Python with Excel - enabling native Excel support for databases, AI agents, LLMs, advanced analytics, machine learning, APIs, web services, and complete automation workflows.
