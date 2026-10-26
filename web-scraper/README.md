# INTELISCAPE-X: AI Web Scraper

Describe the columns you want in plain English. Jina Reader turns each web page into text and Gemini returns the rows as a table in Excel. Status per URL shows live; results are kept across runs.

Tool page: https://www.tigzig.com/web-scraper

## Files

- `INTELISCAPE_X_DYNAMIC_WEB_SCRAPER_v2_1125.xlsx`: the tool. The Python code is inside the workbook.
- `main_web_scraper.py`: the same code as a plain file, to read or to give to an AI coder.
- `requirements.txt`: the packages xlwings Lite installs for this workbook.

## How to run

Open the workbook in Excel and start xlwings Lite (Home > Add-ins). Click the numbered buttons in order. The HOW_IT_WORKS sheet explains each step.

## Keys and setup

GEMINI_API_KEY (free from aistudio.google.com). JINA_API_KEY is optional, for higher Jina limits.

Add keys in xlwings Lite > Settings > Environment Variables, scope Local. They stay in your browser and are not saved in the workbook.
