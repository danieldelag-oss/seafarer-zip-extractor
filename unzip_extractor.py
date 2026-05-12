"""
Seafarer ZIP Extractor (scoped walk + parallel, year-month filter fix)
=======================================================================
Polls OneDrive for .zip attachments under
/Email attachments/{YYYY-MM ...}/{DD}/{Sailor}/Original/, extracts allowed
file types in place, renames source zips to _processed_*.zip.

Strategy:
  - Walk only the N most recent year-month folders (default 2).
  - Skip non-date top-level folders (Invoices, General Documents, etc.).
  - Make sailor-level listings concurrently for speed.

Safety guards:
  - Hardcoded prefix check: /Email attachments only
  - Only processes zips whose parent folder ends with /Original
  - Caps extracted size per zip (default 500 MB)
  - Caps file count per zip (default 200)
  - Allow-listed extensions only
  - Skips/quarantines password-protected and corrupt zips
"""

import io
import logging
import os
import sys
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import quote

import msal
import requests

TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
TARGET_USER_UPN = os.environ["TARGET_USER_UPN"]

EMAIL_ATTACHMENTS_ROOT = "/Email attachments"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

MAX_YEAR_MONTHS_TO_SCAN = 2
PARALLEL_WORKERS = 5

MAX_EXTRACTED_SIZE_BYTES = 500 * 1024 * 1024
MAX_FILES_PER_ZIP = 200
ALLOWED_EXTENSIONS = {
    ".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".gif", ".bmp",
}
EXCLUDE_NAMES = {"Invoices", "General Documents"}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("unzip")


# ---------------------------------------------------------------------------
# Graph helpers with simple retry
# ---------------------------------------------------------------------------
def _request_with_retry(method, url, token, *, max_retries=3, **kw):
    headers = kw.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"
    for attempt in range(max_retries):
        r = requests.request(method, url, headers=headers, timeout=60, **kw)
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", "2"))
            log.warning("Throttled (429), retrying after %ds", wait)
            time.sleep(wait)
            continue
        if r.status_code in (500, 502, 503, 504):
            wait = 2 ** attempt
            log.warning("Graph %d, retrying in %ds (attempt %d)", r.status_code, wait, attempt + 1)
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r
    raise RuntimeError(f"Graph call failed after {max_retries} retries: {url}")


def graph_get(url, token):
    return _request_with_retry("GET", url, token).json()


def graph_get_bytes(url, token):
    return _request_with_retry("GET", url, token).content


def graph_put_bytes(url, token, content):
    _request_with_retry(
        "PUT", url, token,
        headers={"Content-Type": "application/octet-stream"},
        data=content,
    )


def graph_patch_json(url, token, body):
    _request_with_retry(
        "PATCH", url, token,
        headers={"Content-Type": "application/json"},
        json=body,
    )


def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Token acquisition failed: {result.get('error_description', result)}")
    return result["access_token"]


def get_drive_id(token):
    url = f"{GRAPH_BASE}/users/{quote(TARGET_USER_UPN)}/drive"
