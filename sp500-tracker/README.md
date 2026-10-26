# S&P 500 Tracker

Pulls the next earnings date and daily adjusted close prices for a list of tickers (the S&P 500 list ships with it) from the TIGZIG Yahoo Finance API, in batches, writing each batch to Excel as it arrives.

Tool page: https://www.tigzig.com/post/sp500-excel-tracker-xlwings-yfin-aug2026

## Files

- `SP500_TRACKER_PULL.xlsx`: the tool. The Python code is inside the workbook.
- `main_sp500_tracker.py`: the same code as a plain file, to read or to give to an AI coder.
- `requirements.txt`: the packages xlwings Lite installs for this workbook.

## How to run

Open the workbook in Excel and start xlwings Lite (Home > Add-ins). Click the numbered buttons in order. The HOW_IT_WORKS sheet explains each step.

## Keys and setup

None.
