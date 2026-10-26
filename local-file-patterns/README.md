# Local File Access Patterns

Reads and writes files in a folder on your computer: downloads a zip, combines 3,006 CSV files into one, converts to Parquet, runs statistics with DuckDB, and writes HTML, PDF and PowerPoint reports and a Parquet file back to the folder. An optional step emails the reports.

Tool page: https://www.tigzig.com/xlwings-starter

## Files

- `XLWINGS_LITE_LOCAL_FILES_PATTERNS.xlsx`: the tool. The Python code is inside the workbook.
- `main_local_file_patterns.py`: the same code as a plain file, to read or to give to an AI coder.
- `requirements.txt`: the packages xlwings Lite installs for this workbook.

## How to run

Open the workbook in Excel and start xlwings Lite (Home > Add-ins). Click the numbered buttons in order. The HOW_IT_WORKS sheet explains each step.

## Keys and setup

Add a folder in xlwings Lite (Files > Local Folders) and type its path in MASTER!B5. BREVO_API_KEY only for the email step.

Add keys in xlwings Lite > Settings > Environment Variables, scope Local. They stay in your browser and are not saved in the workbook.
