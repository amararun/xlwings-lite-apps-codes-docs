/**
 * Cloudflare Worker: Multi-Cloud File Proxy
 *
 * Purpose: Proxy file downloads to bypass browser CORS restrictions.
 * Browser apps cannot directly fetch from these services due to CORS/redirect issues.
 *
 * Supported Services:
 * - GitHub Releases (redirects to Azure blob - CORS blocked)
 * - Google Drive (no CORS headers on direct download)
 * - Dropbox (no CORS on shared links)
 *
 * Features:
 * - Pure pass-through (no data stored/logged)
 * - Streams large files (no size limit due to streaming)
 * - Domain whitelist for security
 * - Adds CORS headers for browser compatibility
 *
 * Deployed at: https://github-proxy.tigzig.com
 * Usage: https://github-proxy.tigzig.com/?url=<ENCODED_URL>
 *
 * Examples:
 * - GitHub: ?url=https://github.com/user/repo/releases/download/tag/file.ext
 * - Google Drive: ?url=https://drive.google.com/uc?export=download&id=FILE_ID
 * - Dropbox: ?url=https://www.dropbox.com/s/abc123/file.ext?dl=1
 *
 * Tested with files up to 535MB successfully.
 */

// Allowed domains whitelist
const ALLOWED_DOMAINS = [
  // GitHub
  "github.com",
  "githubusercontent.com",
  "raw.githubusercontent.com",
  // Google Drive
  "drive.google.com",
  "docs.google.com",
  "googleusercontent.com",
  // Dropbox
  "dropbox.com",
  "www.dropbox.com",
  "dl.dropboxusercontent.com",
  "dropboxusercontent.com"
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

async function handleRequest(request) {
  // Handle CORS preflight
  if (request.method === "OPTIONS") {
    return new Response(null, {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
        "Access-Control-Max-Age": "86400"
      }
    });
  }

  // Only allow GET requests
  if (request.method !== "GET") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { "Content-Type": "application/json" }
    });
  }

  // Get target URL from query parameter
  const url = new URL(request.url);
  const targetUrl = url.searchParams.get("url");

  // Return usage info if no URL provided
  if (!targetUrl) {
    return new Response(JSON.stringify({
      error: "Missing url parameter",
      usage: {
        github: "?url=https://github.com/user/repo/releases/download/tag/file.ext",
        gdrive: "?url=https://drive.google.com/uc?export=download&id=FILE_ID",
        dropbox: "?url=https://www.dropbox.com/s/abc123/file.ext?dl=1"
      },
      supported: ["GitHub Releases", "Google Drive", "Dropbox"]
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
      allowed: ["github.com", "drive.google.com", "dropbox.com"],
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

  try {
    // Fetch from target (follows redirects)
    const response = await fetch(targetUrl, {
      method: "GET",
      redirect: "follow",
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/octet-stream, */*"
      }
    });

    if (!response.ok) {
      return new Response(JSON.stringify({
        error: `Failed to fetch from ${serviceName}`,
        status: response.status,
        statusText: response.statusText
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
    headers.set("Access-Control-Expose-Headers", "Content-Length, Content-Type, X-Proxy-Service");
    headers.set("Content-Type", response.headers.get("Content-Type") || "application/octet-stream");
    headers.set("X-Proxy-Service", serviceName);

    const contentLength = response.headers.get("Content-Length");
    if (contentLength) {
      headers.set("Content-Length", contentLength);
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
