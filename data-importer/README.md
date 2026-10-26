# Data Importer

Imports a file from a link (Google Drive, Dropbox, GitHub or any direct link) or from private storage with your own token, loads data files (Parquet, DuckDB, SQLite, CSV, TSV, pipe, TXT, JSON) into DuckDB and shows the tables, a column profile and sample rows. Any other file can be downloaded as-is, to /tmp or to a local folder.

Tool page: https://www.tigzig.com/xlwings-data-importer

## Files

- `XLWINGS_LITE_DATA_IMPORTER_V2_17JAN2025_C.xlsx`: the tool. The Python code is inside the workbook.
- `main_data_importer.py`: the same code as a plain file, to read or to give to an AI coder.
- `requirements.txt`: the packages xlwings Lite installs for this workbook.

## How to run

Open the workbook in Excel and start xlwings Lite (Home > Add-ins). Click the numbered buttons in order. The HOW_IT_WORKS sheet explains each step.

## Keys and setup

None for public links. Token access: DROPBOX.REFRESH_TOKEN, DROPBOX.APP_KEY, DROPBOX.APP_SECRET; GDRIVE.REFRESH_TOKEN, GDRIVE.CLIENT_ID, GDRIVE.CLIENT_SECRET; GITHUB.PAT.

Add keys in xlwings Lite > Settings > Environment Variables, scope Local. They stay in your browser and are not saved in the workbook.
