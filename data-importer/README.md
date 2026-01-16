# xlwings Lite Data Importer

Python script for importing data files into xlwings Lite from various cloud sources.

## downloader-links-secure-token-stats-raw.py

Main script with all data import functionality. Embedded in the xlwings Lite Data Importer Excel workbook.

### Features

#### 2A: Shareable Link Access (import_via_sharelink)
- Import from any public/shareable URL
- Supports: DuckIt, GitHub, Google Drive, Dropbox
- File types: Parquet, DuckDB, SQLite, CSV, TSV, JSON
- Uses CORS proxy for cross-origin downloads
- Auto-detects file type and delimiter

#### 2B: Token Access (import_via_token)
- Import from private cloud storage
- Dropbox: OAuth refresh token
- Google Drive: OAuth refresh token
- GitHub: Personal Access Token (PAT)
- Direct API access (no proxy for Dropbox/GDrive)

#### 2C: Raw File Import (import_raw_sharelink, import_raw_token)
- Download any file type as-is
- PDF, images, ZIP, Excel, or data files
- No DuckDB conversion
- Test functions: test_last_image, test_last_pdf, test_last_zip

### Sheet Layout

**SHARE_LINK_ACCESS sheet:**
- B5: Source URL
- B7: Raw mode flag (1 = raw)
- B8: Output filename (for raw mode)
- D8: Output file path

**TOKEN_ACCESS sheet:**
- B5: Provider (Dropbox, Google Drive, GitHub)
- B6: File path/ID
- B7: Auth proxy URL (GitHub only)
- B9: Raw mode flag (1 = raw)
- B10: Output filename (for raw mode)
- D10: Output file path

### Environment Variables (for Token Access)

**Dropbox:**
- DROPBOX.APP_KEY
- DROPBOX.APP_SECRET
- DROPBOX.REFRESH_TOKEN

**Google Drive:**
- GDRIVE.CLIENT_ID
- GDRIVE.CLIENT_SECRET
- GDRIVE.REFRESH_TOKEN

**GitHub:**
- GITHUB.PAT

### Functions

| Function | Description |
|----------|-------------|
| import_via_sharelink | Import via shareable URL |
| import_via_token | Import via OAuth/PAT |
| import_raw_sharelink | Raw file download via URL |
| import_raw_token | Raw file download via OAuth/PAT |
| run_stats | Auto-detect and run statistics |
| test_last_image | Verify last imported image |
| test_last_pdf | Verify last imported PDF |
| test_last_zip | Verify last imported ZIP |
