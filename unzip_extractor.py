"""
Seafarer ZIP Extractor (scoped walk + parallel)
================================================
Polls OneDrive for .zip attachments under /Email attachments/{YYYY-MM ...}/{DD}/{Sailor}/Original/,
extracts allowed file types in place, renames source zips to _processed_*.zip.

Strategy:
  - Walk only the N most recent year-month folders (default 2).
  - Make sailor-level listings concurrently (5 at a time) for speed.
  - Use search NOT used — tenant blocks app-only search with 403.

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

# Performance tuning
MAX_YEAR_MONTHS_TO_SCAN = 2     # Scan only the N most recent year-month folders
PARALLEL_WORKERS = 5            # Concurrent Graph calls (Graph throttles ~10+)

# Safety limits
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
    return graph_get(url, token)["id"]


def list_folder_by_path(drive_id, path, token):
    encoded = quote(path)
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:{encoded}:/children?$top=200"
    items = []
    while url:
        data = graph_get(url, token)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def _list_safe(drive_id, path, token):
    """Like list_folder_by_path but returns [] on 404 instead of raising."""
    try:
        return list_folder_by_path(drive_id, path, token)
    except requests.HTTPError as e:
        if e.response is not None and e.response.status_code == 404:
            return []
        raise


# ---------------------------------------------------------------------------
# Scoped walk: last N year-months only, parallel sailor scans
# ---------------------------------------------------------------------------
def _scan_sailor_for_zips(drive_id, sailor_path, token):
    """Return list of (Original_path, zip_item) tuples for unprocessed zips in
    {sailor_path}/Original. Used as a worker in the thread pool."""
    original_path = f"{sailor_path}/Original"
    files = _list_safe(drive_id, original_path, token)
    out = []
    for f in files:
        if f.get("folder"):
            continue
        name = f["name"]
        if not name.lower().endswith(".zip"):
            continue
        if name.startswith("_processed_") or name.startswith("_REJECTED_"):
            continue
        out.append((original_path, f))
    return out


def find_zip_files(drive_id, token):
    results = []

    # Step 1: list year-month folders, take last N (most recent by name)
    ym_items = _list_safe(drive_id, EMAIL_ATTACHMENTS_ROOT, token)
    ym_folders = [
        it for it in ym_items
        if it.get("folder") and not it["name"].startswith(("_", "."))
    ]
    ym_folders.sort(key=lambda x: x["name"], reverse=True)
    ym_folders = ym_folders[:MAX_YEAR_MONTHS_TO_SCAN]
    log.info("Scanning %d year-month folder(s): %s",
             len(ym_folders), [it["name"] for it in ym_folders])

    # Step 2: walk year-month -> day -> sailor; collect sailor paths
    sailor_paths = []
    for ym in ym_folders:
        ym_path = f"{EMAIL_ATTACHMENTS_ROOT}/{ym['name']}"
        day_items = _list_safe(drive_id, ym_path, token)
        for d in day_items:
            if not d.get("folder") or d["name"].startswith(("_", ".")):
                continue
            d_path = f"{ym_path}/{d['name']}"
            sailor_items = _list_safe(drive_id, d_path, token)
            for s in sailor_items:
                if not s.get("folder"):
                    continue
                if s["name"].startswith(("_", ".")) or s["name"] in EXCLUDE_NAMES:
                    continue
                sailor_paths.append(f"{d_path}/{s['name']}")
    log.info("Total sailor folders to scan: %d", len(sailor_paths))

    # Step 3: scan all sailor/Original folders concurrently
    with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as ex:
        futures = {
            ex.submit(_scan_sailor_for_zips, drive_id, sp, token): sp
            for sp in sailor_paths
        }
        for fut in as_completed(futures):
            sp = futures[fut]
            try:
                results.extend(fut.result())
            except Exception:
                log.exception("Error scanning %s", sp)

    log.info("Found %d zip(s) to process", len(results))
    return results


# ---------------------------------------------------------------------------
# Zip validation + processing (unchanged from before)
# ---------------------------------------------------------------------------
def validate_zip(zip_bytes):
    try:
        zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    except zipfile.BadZipFile:
        return False, "corrupt"
    infos = [i for i in zf.infolist() if not i.is_dir()]
    for info in infos:
        if info.flag_bits & 0x1:
            return False, "password_protected"
    if len(infos) > MAX_FILES_PER_ZIP:
        return False, f"too_many_files_{len(infos)}"
    total = sum(i.file_size for i in infos)
    if total > MAX_EXTRACTED_SIZE_BYTES:
        return False, f"too_large_{total}"
    return True, None


def upload_to_folder(drive_id, folder_path, filename, content, token):
    if not folder_path.startswith(EMAIL_ATTACHMENTS_ROOT):
        raise RuntimeError(f"Refusing upload outside allowed root: {folder_path}")
    encoded = quote(f"{folder_path}/{filename}")
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:{encoded}:/content"
    graph_put_bytes(url, token, content)


def rename_item(drive_id, item_id, new_name, token):
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
    graph_patch_json(url, token, {"name": new_name})


def process_zip(drive_id, folder_path, zip_item, token):
    name = zip_item["name"]
    item_id = zip_item["id"]
    log.info("Processing %s/%s", folder_path, name)

    if not folder_path.startswith(EMAIL_ATTACHMENTS_ROOT):
        log.error("Refusing — outside allowed root: %s", folder_path)
        return

    download_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
    try:
        content = graph_get_bytes(download_url, token)
    except requests.HTTPError as e:
        log.error("Failed to download %s: %s", name, e)
        return

    ok, reason = validate_zip(content)
    if not ok:
        log.warning("Rejecting %s: %s", name, reason)
        rename_item(drive_id, item_id, f"_REJECTED_{reason}_{name}", token)
        return

    extracted = 0
    skipped = 0
    with zipfile.ZipFile(io.BytesIO(content)) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            entry_name = os.path.basename(info.filename)
            if not entry_name:
                continue
            ext = os.path.splitext(entry_name)[1].lower()
            if ext in ALLOWED_EXTENSIONS or ext == ".zip":
                log.info("  Extracting: %s", entry_name)
                upload_to_folder(drive_id, folder_path, entry_name, zf.read(info), token)
                extracted += 1
            else:
                log.info("  Skipping disallowed type: %s", entry_name)
                skipped += 1

    log.info("Extracted %d, skipped %d from %s", extracted, skipped, name)
    rename_item(drive_id, item_id, f"_processed_{name}", token)


def main():
    log.info("Seafarer ZIP Extractor (scoped walk + parallel) starting")
    if not EMAIL_ATTACHMENTS_ROOT.startswith("/Email attachments"):
        log.error("Path safety check failed — refusing to run.")
        return 2
    try:
        token = get_access_token()
        log.info("Got Graph token")
        drive_id = get_drive_id(token)
        log.info("Target drive id: %s", drive_id)

        zips = find_zip_files(drive_id, token)
        for folder_path, zip_item in zips:
            try:
                process_zip(drive_id, folder_path, zip_item, token)
            except Exception:
                log.exception("Error processing %s/%s",
                              folder_path, zip_item.get("name", "?"))

        log.info("Done.")
        return 0
    except Exception:
        log.exception("Fatal error")
        return 1


if __name__ == "__main__":
    sys.exit(main())
