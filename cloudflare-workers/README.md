# Cloudflare Workers

This directory contains Cloudflare Worker scripts used by xlwings Lite apps.

## github-proxy-worker.js

**Deployed at:** https://github-proxy.tigzig.com

**Purpose:** CORS proxy for browser-based file downloads. Browser apps (like xlwings Lite running in Pyodide) cannot directly fetch files from certain services due to CORS restrictions.

### Supported Services

| Service | Why Proxy Needed |
|---------|------------------|
| GitHub Releases | Redirects to Azure blob which blocks CORS |
| Google Drive | No CORS headers on direct downloads |
| Dropbox | No CORS on shared links |

### Usage

```
https://github-proxy.tigzig.com/?url=<ENCODED_URL>
```

### Examples

```javascript
// GitHub Release
?url=https://github.com/user/repo/releases/download/v1.0/file.parquet

// Google Drive
?url=https://drive.google.com/uc?export=download&id=FILE_ID

// Dropbox
?url=https://www.dropbox.com/s/abc123/file.ext?dl=1
```

### Features

- **Streaming:** No file size limit (streams directly, no buffering)
- **Security:** Domain whitelist (only GitHub, Google Drive, Dropbox allowed)
- **CORS:** Adds proper headers for browser compatibility
- **Pass-through:** No data stored or logged

### Tested

Successfully tested with files up to 535MB.
