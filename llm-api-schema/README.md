# AI Workflows: LLM API Calls from Excel

An LLM (Gemini, or any model on OpenRouter) reads the first rows of an Excel table and returns the schema as JSON: categorical and numeric columns, a PostgreSQL type and a description per column. Python then runs the EDA: statistics, distributions, correlations and charts.

Tool page: https://www.tigzig.com/xlwings-llm-api

## Files

- `LLM_API_CALL_SCHEMA_DETECT_EDA_MACHINE_LEARNING.xlsx`: the tool. The Python code is inside the workbook.
- `main_llm_api_schema.py`: the same code as a plain file, to read or to give to an AI coder.
- `requirements.txt`: the packages xlwings Lite installs for this workbook.

## How to run

Open the workbook in Excel and start xlwings Lite (Home > Add-ins). Click the numbered buttons in order. The HOW_IT_WORKS sheet explains each step.

## Keys and setup

GEMINI_API_KEY or OPENROUTER_API_KEY.

Add keys in xlwings Lite > Settings > Environment Variables, scope Local. They stay in your browser and are not saved in the workbook.
