# Database & ML: Connect to Databases, Build ML Models

Connect to PostgreSQL/MySQL databases, run queries, and perform advanced analytics including XGBoost classification directly in Excel.

**[Video Walkthrough (18:48 min)](https://youtu.be/rHERSN_Bay0)**

---

## What This App Does

Connect to databases and perform advanced analytics directly in Excel:

- Connect to PostgreSQL/MySQL databases
- Explore tables, run custom SQL queries
- Statistical analysis and EDA
- Build ML models with **XGBoost**
- Decile tables, ROC curves, gains charts
- Model scoring and evaluation metrics

## How to Use

1. **Download & Install:** Get the app and install xlwings Lite
2. **Configure DB:** Update database credentials
3. **Built-in Functions:**
   - `list_tables` - List all tables
   - `get_custom_query` - Run SQL queries
   - `perform_eda` - Exploratory data analysis
   - `score_and_deciles` - XGBoost + decile tables

## How It Works

**Database Connectivity:**
- FastAPI server handles database connections
- API requests connect to PostgreSQL/MySQL
- Responses returned to Excel for processing

**ML Processing:**
- All processing occurs locally in xlwings
- XGBoost runs directly in browser environment
- Generates charts and tables for analysis

### Free Database Options
- **[neon.tech](https://neon.tech)** - Postgres, 500MB free
- **[Supabase](https://supabase.com)** - Postgres + auth, 500MB free
- **[Aiven](https://aiven.io)** - Postgres/MySQL, 5GB free

## XGBoost ML Pipeline

**Note:** This demonstrates ML in Excel - not a point-and-click app. For new datasets, work iteratively with an AI coder (Claude/ChatGPT) to adapt the script.

**Complete Pipeline Includes:**
- Data loading, feature engineering, train/test split
- XGBoost classifier training
- **Decile tables** - performance across score bands
- **ROC curves** and **Cumulative Gains charts**
- Gini coefficient, confusion matrix, F1-score
- Scoring output appended to source data

**Adaptation Workflow:**
1. Pass `main_database_ml.py` to AI coder
2. Specify your dataset and variables
3. Request modifications, iterate on results

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
- **[FastAPI Server Code](https://github.com/amararun/shared-fastapi-rex-db-coolify)** - Database connectivity implementation
- **[Deploy on Render](https://lnkd.in/g2A9h8f2)** - Deployment guide for FastAPI server

---

## Credits

Created by [Felix Zumstein](https://www.linkedin.com/in/felix-zumstein/), xlwings Lite delivers a powerful and flexible solution for integrating Python with Excel - enabling native Excel support for databases, AI agents, LLMs, advanced analytics, machine learning, APIs, web services, and complete automation workflows.
