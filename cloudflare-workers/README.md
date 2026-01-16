# Cloudflare Workers

This directory contains the Cloudflare Worker script used by xlwings Lite Data Importer.

## github-proxy-auth-worker.js

**Deployed at:** https://github-proxy-auth.tigzig.com

**Purpose:** Unified multi-cloud CORS proxy supporting shareable links AND token-based private repo access. Used by both Shareable Link Access (2A) and Token Access (2B) modes.

### Why a Proxy?

Browser apps (like xlwings Lite running in Pyodide) cannot directly fetch files from certain services due to CORS restrictions:

| Service | Why Proxy Needed |
|---------|------------------|
| GitHub Releases | Redirects to Azure blob which blocks CORS |
| Google Drive | No CORS headers on direct downloads |
| Dropbox | No CORS on shared links |

### Supported Services

| Service | Public URLs | Private (Token) |
|---------|-------------|-----------------|
| GitHub Releases | Yes | Yes (PAT) |
| GitHub Raw Content | Yes | Yes (PAT) |
| Google Drive | Yes | - |
| Dropbox | Yes | - |

### Usage

```
https://github-proxy-auth.tigzig.com/?url=<ENCODED_URL>
```

For private GitHub repos, include Authorization header:
```
Authorization: token ghp_xxxxx
```

### Features

- **Streaming:** No file size limit (streams directly, no buffering)
- **Security:** Domain whitelist (only GitHub, Google Drive, Dropbox allowed)
- **CORS:** Adds proper headers for browser compatibility
- **Pass-through:** No data stored or logged
- **Private GitHub repos:** Forwards PAT for authentication
- **GitHub API integration:** Uses GitHub API for private release asset downloads
- **Google Drive (new domains):** Supports new 2024-2025 drive.usercontent.google.com domain

### Examples

```javascript
// GitHub Public Release
?url=https://github.com/user/repo/releases/download/v1.0/file.parquet

// GitHub Private Release (with Authorization header)
?url=https://github.com/user/private-repo/releases/download/v1.0/file.parquet
// Header: Authorization: token ghp_xxxxx

// Google Drive (old domain)
?url=https://drive.google.com/uc?export=download&id=FILE_ID

// Google Drive (new domain)
?url=https://drive.usercontent.google.com/download?id=FILE_ID&export=download&confirm=t

// Dropbox
?url=https://www.dropbox.com/s/abc123/file.ext?dl=1
```

---

## Deployment

Deploy using Cloudflare Workers dashboard or Wrangler CLI.

### Free Cloudflare Account

Works with free Cloudflare account - uses free workers.dev subdomain. No custom domain required.
