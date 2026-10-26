# Database and ML in Excel

Query a Neon Postgres database straight from Excel over HTTPS, with no server in between: list tables, pull rows, run SQL. Then run EDA and an XGBoost response model on any Excel table: Gini, deciles with lift, a score per row, ROC and gain charts. A 10,000-row sample table ships with the workbook.

Tool page: https://www.tigzig.com/xlwings-api-db

## Files

- `DATABASE_PULL_EDSACHARTS_TABLES_ML_MODEL_V3_1225.xlsx`: the tool. The Python code is inside the workbook.
- `main_database_ml.py`: the same code as a plain file, to read or to give to an AI coder.
- `requirements.txt`: the packages xlwings Lite installs for this workbook.

## How to run

Open the workbook in Excel and start xlwings Lite (Home > Add-ins). Click the numbered buttons in order. The HOW_IT_WORKS sheet explains each step.

## Keys and setup

NEON_DATABASE_URL (your Neon connection string). Optional: the EDA and model work on the sample without a database.

Add keys in xlwings Lite > Settings > Environment Variables, scope Local. They stay in your browser and are not saved in the workbook.
