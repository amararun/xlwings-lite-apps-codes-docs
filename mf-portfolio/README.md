# MF Portfolio Analyzer: Holdings Change Analysis

Analyze changes in mutual fund holdings between two periods with automated data quality checks, ISIN name standardization, and human-in-the-loop correction workflow.

**[Video Walkthrough (1:48 min)](https://youtu.be/U7RwHxRkOao)**

---

## What This App Does

A pipeline for analyzing mutual fund portfolio holdings across two periods:

- **Smart file import:** Auto-detects CSV/TXT/pipe-delimited formats
- **2-period detection:** Auto-identifies latest two months
- Automatic ISIN name standardization and deduplication
- Data quality validation reports
- Human-in-the-loop correction workflow
- Summary analysis tables with holdings changes
- Automated Top 10 holdings charts

**Prerequisites:**
- Requires standardized text files - use [MF File Converter](https://app.tigzig.com/mf-files-ai) first if you have raw Excel files
- VBA macro included - after download: right-click -> Properties -> Unblock -> OK

## How to Use

1. **Prerequisite:** Convert raw MF files using [MF File Converter](https://app.tigzig.com/mf-files-ai)
2. **Download & Install:** Get the app and install xlwings Lite
3. **Import Data:** Click "Import Data" on Control Sheet
4. **Run Stage 1:** Review ISIN_Mapping sheet
5. **Manual Correction:** Fix grouping issues if needed
6. **Run Stage 2:** Generate final analysis
7. **View Results:** Check Summary_Analysis and Charts

## How It Works

A **3-step workflow** for data quality and accuracy:

**Smart Detection:**
- **File import:** Auto-detects CSV/TXT/pipe formats
- **2-period detection:** Identifies latest two months automatically

**ISIN Standardization:**
- Handles name variations (HDFC, HDFC Bank Ltd, etc.)
- Selects best name for each ISIN
- Name truncation to identify conflicts

**Validation Reports:**
- Namecut_Exceptions & Multiple_ISINs reports
- Guides human review for edge cases

**Human-in-the-Loop:** User edits ISIN_Mapping for accuracy. Stage 2 uses corrected mappings.

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

---

## Credits

Created by [Felix Zumstein](https://www.linkedin.com/in/felix-zumstein/), xlwings Lite delivers a powerful and flexible solution for integrating Python with Excel - enabling native Excel support for databases, AI agents, LLMs, advanced analytics, machine learning, APIs, web services, and complete automation workflows.
