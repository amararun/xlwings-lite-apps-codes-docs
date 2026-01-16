/**
 * Cloudflare Worker: Unified Multi-Cloud File Proxy
 *
 * Purpose: Proxy file downloads to bypass browser CORS restrictions.
 * Supports both shareable links AND token-based private repo access.
 *
 * Supported Services:
 * - GitHub Releases (public and private repos)
 * - GitHub Raw Content (public and private repos)
 * - Google Drive (old and new 2024-2025 domains)
 * - Dropbox (shareable links and API)
 *
 * Features:
 * - Pure pass-through (no data stored/logged)
 * - Streams large files (no size limit due to streaming)
 * - Domain whitelist for security
 * - Adds CORS headers for browser compatibility
 * - Forwards Authorization header for private GitHub repos
 * - GitHub API-based download for private release assets
 *
 * Deployed at: https://github-proxy-auth.tigzig.com
 * Usage: https://github-proxy-auth.tigzig.com/?url=<ENCODED_URL>
 *
 * For private repos, include Authorization header:
 *   Authorization: token ghp_xxxxx
 *
 * Examples:
 * - GitHub Public: ?url=https://github.com/user/repo/releases/download/tag/file.ext
 * - GitHub Private: Same URL + Authorization header
 * - Google Drive (old): ?url=https://drive.google.com/uc?export=download&id=FILE_ID
 * - Google Drive (new): ?url=https://drive.usercontent.google.com/download?id=FILE_ID&export=download&confirm=t
 * - Dropbox: ?url=https://www.dropbox.com/s/abc123/file.ext?dl=1
 */

// Allowed domains whitelist (unified - supports all shareable + token access)
const ALLOWED_DOMAINS = [
  // GitHub
  "github.com",
  "api.github.com",
  "githubusercontent.com",
  "raw.githubusercontent.com",
  "github.io",
  "ghcr.io",
  // Google Drive (OLD and NEW domains)
  "drive.google.com",
  "drive.usercontent.google.com",  // NEW 2024-2025 domain for downloads
  "docs.google.com",
  "googleusercontent.com",
  "googleapis.com",
  // Dropbox
  "dropbox.com",
  "www.dropbox.com",
  "dl.dropboxusercontent.com",
  "dropboxusercontent.com",
  "content.dropboxapi.com"
];

addEventListener("fetch", event => {
  event.respondWith(handleRequest(event.request));
});

function isAllowedUrl(urlString) {
  try {
    const url = new URL(urlString);
    const hostname = url.hostname.toLowerCase();
    return ALLOWED_DOMAINS.some(domain =>
      hostname === domain || hostname.endsWith("." + domain)
    );
  } catch {
    return false;
  }
}

function getServiceName(urlString) {
  try {
    const url = new URL(urlString);
    const hostname = url.hostname.toLowerCase();
    if (hostname.includes("github") || hostname.includes("githubusercontent")) return "GitHub";
    if (hostname.includes("google") || hostname.includes("drive.google")) return "Google Drive";
    if (hostname.includes("dropbox")) return "Dropbox";
    return "Unknown";
  } catch {
    return "Unknown";
  }
}

/**
 * Convert GitHub release browser URL to API URL for private repos.
 * Browser URL: https://github.com/owner/repo/releases/download/tag/filename
 * API URL: https://api.github.com/repos/owner/repo/releases/tags/tag (to get asset ID)
 * Then: https://api.github.com/repos/owner/repo/releases/assets/{asset_id}
 *
 * Returns: { needsApiLookup: true, owner, repo, tag, filename } or { needsApiLookup: false }
 */
function parseGitHubReleaseUrl(urlString) {
  try {
    const url = new URL(urlString);
    if (!url.hostname.includes("github.com")) {
      return { needsApiLookup: false };
    }

    // Match: /owner/repo/releases/download/tag/filename
    const match = url.pathname.match(/^\/([^\/]+)\/([^\/]+)\/releases\/download\/([^\/]+)\/(.+)$/);
    if (match) {
      return {
        needsApiLookup: true,
        owner: match[1],
        repo: match[2],
        tag: match[3],
        filename: match[4]
      };
    }
    return { needsApiLookup: false };
  } catch {
    return { needsApiLookup: false };
  }
}

/**
 * For private repos, we need to:
 * 1. Get the release by tag to find the asset ID
 * 2. Download the asset using the asset ID with Accept: application/octet-stream
 */
async function downloadGitHubReleaseAsset(owner, repo, tag, filename, authHeader) {
  // Step 1: Get release info to find asset ID
  const releaseUrl = `https://api.github.com/repos/${owner}/${repo}/releases/tags/${tag}`;

  const releaseResponse = await fetch(releaseUrl, {
    headers: {
      "Authorization": authHeader,
      "Accept": "application/vnd.github+json",
      "User-Agent": "CloudflareWorker-GitHubProxy/2.0"
    }
  });

  if (!releaseResponse.ok) {
    return {
      ok: false,
      status: releaseResponse.status,
      error: `Failed to get release info: ${releaseResponse.status} ${releaseResponse.statusText}`
    };
  }

  const releaseData = await releaseResponse.json();

  // Find the asset by filename
  const asset = releaseData.assets.find(a => a.name === filename);
  if (!asset) {
    return {
      ok: false,
      status: 404,
      error: `Asset '${filename}' not found in release '${tag}'`
    };
  }

  // Step 2: Download the asset using API URL with octet-stream Accept header
  const assetResponse = await fetch(asset.url, {
    headers: {
      "Authorization": authHeader,
      "Accept": "application/octet-stream",
      "User-Agent": "CloudflareWorker-GitHubProxy/2.0"
    },
    redirect: "follow"
  });

  if (!assetResponse.ok) {
    return {
      ok: false,
      status: assetResponse.status,
      error: `Failed to download asset: ${assetResponse.status} ${assetResponse.statusText}`
    };
  }

  return {
    ok: true,
    response: assetResponse,
    filename: asset.name,
    size: asset.size
  };
}

async function handleRequest(request) {
  // Handle CORS preflight
  if (request.method === "OPTIONS") {
    return new Response(null, {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization, Accept",
        "Access-Control-Max-Age": "86400"
      }
    });
  }

  // Only allow GET requests
  if (request.method !== "GET") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*"
      }
    });
  }

  // Get target URL from query parameter
  const url = new URL(request.url);
  const targetUrl = url.searchParams.get("url");

  // Return usage info if no URL provided
  if (!targetUrl) {
    return new Response(JSON.stringify({
      name: "Unified Multi-Cloud File Proxy",
      version: "3.0",
      error: "Missing url parameter",
      usage: {
        github_public: "?url=https://github.com/user/repo/releases/download/tag/file.ext",
        github_private: "Same URL + Authorization header: 'token ghp_xxxxx'",
        gdrive_old: "?url=https://drive.google.com/uc?export=download&id=FILE_ID",
        gdrive_new: "?url=https://drive.usercontent.google.com/download?id=FILE_ID&export=download&confirm=t",
        dropbox: "?url=https://www.dropbox.com/s/abc123/file.ext?dl=1"
      },
      supported: ["GitHub Releases (public/private)", "Google Drive (old & new domains)", "Dropbox"]
    }), {
      status: 400,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*"
      }
    });
  }

  // Security: Only allow whitelisted domains
  if (!isAllowedUrl(targetUrl)) {
    return new Response(JSON.stringify({
      error: "Invalid URL - domain not allowed",
      allowed: ["github.com", "drive.google.com", "drive.usercontent.google.com", "dropbox.com"],
      hint: "Only GitHub, Google Drive, and Dropbox URLs are supported"
    }), {
      status: 400,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*"
      }
    });
  }

  const serviceName = getServiceName(targetUrl);
  const authHeader = request.headers.get("Authorization");

  try {
    // Special handling for GitHub release URLs with authentication (private repos)
    // For private repos, direct download URLs don't work - need to use GitHub API
    const releaseInfo = parseGitHubReleaseUrl(targetUrl);
    if (releaseInfo.needsApiLookup && authHeader && serviceName === "GitHub") {
      // Use GitHub API to download the asset (works for private repos)
      const result = await downloadGitHubReleaseAsset(
        releaseInfo.owner,
        releaseInfo.repo,
        releaseInfo.tag,
        releaseInfo.filename,
        authHeader
      );

      if (!result.ok) {
        return new Response(JSON.stringify({
          error: result.error,
          status: result.status,
          hint: "Make sure the PAT has 'repo' scope for private repositories"
        }), {
          status: result.status,
          headers: {
            "Content-Type": "application/json",
            "Access-Control-Allow-Origin": "*"
          }
        });
      }

      // Build response headers with CORS
      const headers = new Headers();
      headers.set("Access-Control-Allow-Origin", "*");
      headers.set("Access-Control-Expose-Headers", "Content-Length, Content-Type, X-Proxy-Service, X-Proxy-Auth, X-Proxy-Method");
      headers.set("Content-Type", result.response.headers.get("Content-Type") || "application/octet-stream");
      headers.set("X-Proxy-Service", "GitHub");
      headers.set("X-Proxy-Auth", "yes");
      headers.set("X-Proxy-Method", "api-asset");
      headers.set("Cache-Control", "no-store");

      if (result.size) {
        headers.set("Content-Length", String(result.size));
      }

      // Stream the response body directly
      return new Response(result.response.body, {
        status: 200,
        headers: headers
      });
    }

    // Standard handling for other URLs (public GitHub, Google Drive, Dropbox)
    const fetchHeaders = {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
      "Accept": "application/octet-stream, */*"
    };

    // Forward Authorization header if present (for private GitHub repos)
    if (authHeader && serviceName === "GitHub") {
      fetchHeaders["Authorization"] = authHeader;
    }

    // Forward Accept header if provided (for GitHub API octet-stream)
    const acceptHeader = request.headers.get("Accept");
    if (acceptHeader) {
      fetchHeaders["Accept"] = acceptHeader;
    }

    // Fetch from target (follows redirects)
    const response = await fetch(targetUrl, {
      method: "GET",
      redirect: "follow",
      headers: fetchHeaders
    });

    if (!response.ok) {
      return new Response(JSON.stringify({
        error: `Failed to fetch from ${serviceName}`,
        status: response.status,
        statusText: response.statusText,
        hint: response.status === 404 ? "File not found or private repo requires Authorization header" : undefined
      }), {
        status: response.status,
        headers: {
          "Content-Type": "application/json",
          "Access-Control-Allow-Origin": "*"
        }
      });
    }

    // Build response headers with CORS
    const headers = new Headers();
    headers.set("Access-Control-Allow-Origin", "*");
    headers.set("Access-Control-Expose-Headers", "Content-Length, Content-Type, X-Proxy-Service, X-Proxy-Auth");
    headers.set("Content-Type", response.headers.get("Content-Type") || "application/octet-stream");
    headers.set("X-Proxy-Service", serviceName);
    headers.set("X-Proxy-Auth", authHeader ? "yes" : "no");

    const contentLength = response.headers.get("Content-Length");
    if (contentLength) {
      headers.set("Content-Length", contentLength);
    }

    // Forward Content-Disposition if present
    const contentDisposition = response.headers.get("Content-Disposition");
    if (contentDisposition) {
      headers.set("Content-Disposition", contentDisposition);
    }

    // No caching to ensure fresh data
    headers.set("Cache-Control", "no-store");

    // Stream the response body directly (no buffering = no size limit)
    return new Response(response.body, {
      status: 200,
      headers: headers
    });

  } catch (error) {
    return new Response(JSON.stringify({
      error: "Proxy error",
      service: serviceName,
      message: error.message
    }), {
      status: 500,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*"
      }
    });
  }
}
