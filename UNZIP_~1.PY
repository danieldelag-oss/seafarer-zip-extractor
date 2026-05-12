"""
Seafarer ZIP/RAR Extractor
===========================
Polls OneDrive for archive attachments (.zip, .rar, .zip.rar) under
/Email attachments/{YYYY-MM ...}/{DD}/{Sailor}/Original/, extracts contents
in place, renames the source archive to _processed_*.

Archive type is detected by MAGIC BYTES of the file content (not by
filename), so files renamed by email security tools (e.g. a real .zip
renamed to .zip.rar) are still handled correctly.

Safety guards:
  - Refuses to operate outside /Email attachments/
  - Only processes archives in folders ending with /Original
  - Caps extracted size (default 500 MB) and file count (default 200)
  - Allow-listed file types only
  - Skips password-protected and corrupt archives (renames to _REJECTED_*)
"""

import io
import logging
import os
import sys
import tempfile
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import quote

import msal
import rarfile
import requests

# rarfile uses an external system tool. The workflow installs 'unar'.
rarfile.UNRAR_TOOL = "unar"

TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
TARGET_USER_UPN = os.environ["TARGET_USER_UPN"]

EMAIL_ATTACHMENTS_ROOT = "/Email attachments"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

MAX_YEAR_MONTHS_TO_SCAN = 2
PARALLEL_WORKERS = 5

MAX_EXTRACTED_SIZE_BYTES = 500 * 1024 * 1024
MAX_FILES_PER_ARCHIVE = 200
ALLOWED_EXTENSIONS = {
    ".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".gif", ".bmp",
}
EXCLUDE_NAMES = {"Invoices", "General Documents"}

ZIP_MAGIC = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
RAR_MAGIC = (b"Rar!\x1a\x07\x00", b"Rar!\x1a\x07\x01\x00")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("unzip")


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
            log.warning("Graph %d, retrying in %ds", r.status_code, wait)
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
    try:
        return list_folder_by_path(drive_id, path, token)
    except requests.HTTPError as e:
        if e.response is not None and e.response.status_code == 404:
            return []
        raise


def is_archive_filename(name):
    n = name.lower()
    return n.endswith(".zip") or n.endswith(".rar")


def detect_archive_type(content):
    head = content[:16]
    if any(head.startswith(m) for m in ZIP_MAGIC):
        return "zip"
    if any(head.startswith(m) for m in RAR_MAGIC):
        return "rar"
    return None


def _scan_sailor_for_archives(drive_id, sailor_path, token):
    original_path = f"{sailor_path}/Original"
    files = _list_safe(drive_id, original_path, token)
    out = []
    for f in files:
        if f.get("folder"):
            continue
        name = f["name"]
        if not is_archive_filename(name):
            continue
        if name.startswith("_processed_") or name.startswith("_REJECTED_"):
            continue
        out.append((original_path, f))
    return out


def find_archives(drive_id, token):
    results = []

    ym_items = _list_safe(drive_id, EMAIL_ATTACHMENTS_ROOT, token)
    ym_folders = [
        it for it in ym_items
        if it.get("folder")
        and not it["name"].startswith(("_", "."))
        and it["name"] not in EXCLUDE_NAMES
        and it["name"][:1].isdigit()
    ]
    ym_folders.sort(key=lambda x: x["name"], reverse=True)
    ym_folders = ym_folders[:MAX_YEAR_MONTHS_TO_SCAN]
    log.info("Scanning %d year-month folder(s): %s",
             len(ym_folders), [it["name"] for it in ym_folders])

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

    with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as ex:
        futures = {
            ex.submit(_scan_sailor_for_archives, drive_id, sp, token): sp
            for sp in sailor_paths
        }
        for fut in as_completed(futures):
            sp = futures[fut]
            try:
                results.extend(fut.result())
            except Exception:
                log.exception("Error scanning %s", sp)

    log.info("Found %d archive(s) to process", len(results))
    return results


def _validate_zip(content):
    try:
        zf = zipfile.ZipFile(io.BytesIO(content))
    except zipfile.BadZipFile:
        return False, "corrupt"
    infos = [i for i in zf.infolist() if not i.is_dir()]
    for info in infos:
        if info.flag_bits & 0x1:
            return False, "password_protected"
    if len(infos) > MAX_FILES_PER_ARCHIVE:
        return False, f"too_many_files_{len(infos)}"
    total = sum(i.file_size for i in infos)
    if total > MAX_EXTRACTED_SIZE_BYTES:
        return False, f"too_large_{total}"
    return True, None


def _validate_rar(content):
    fd, tmp_path = tempfile.mkstemp(suffix=".rar")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(content)
        try:
            rf = rarfile.RarFile(tmp_path)
        except (rarfile.BadRarFile, rarfile.NotRarFile, rarfile.Error):
            return False, "corrupt"
        if rf.needs_password():
            return False, "password_protected"
        infos = [i for i in rf.infolist() if not i.is_dir()]
        if len(infos) > MAX_FILES_PER_ARCHIVE:
            return False, f"too_many_files_{len(infos)}"
        total = sum(i.file_size for i in infos)
        if total > MAX_EXTRACTED_SIZE_BYTES:
            return False, f"too_large_{total}"
        return True, None
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def validate_archive(content, archive_type):
    if archive_type == "zip":
        return _validate_zip(content)
    if archive_type == "rar":
        return _validate_rar(content)
    return False, "unrecognized_format"


def _iter_zip_entries(content):
    with zipfile.ZipFile(io.BytesIO(content)) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            yield info.filename, zf.read(info)


def _iter_rar_entries(content):
    fd, tmp_path = tempfile.mkstemp(suffix=".rar")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(content)
        with rarfile.RarFile(tmp_path) as rf:
            for info in rf.infolist():
                if info.is_dir():
                    continue
                yield info.filename, rf.read(info)
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def iter_archive_entries(content, archive_type):
    if archive_type == "zip":
        yield from _iter_zip_entries(content)
    elif archive_type == "rar":
        yield from _iter_rar_entries(content)


def upload_to_folder(drive_id, folder_path, filename, content, token):
    if not folder_path.startswith(EMAIL_ATTACHMENTS_ROOT):
        raise RuntimeError(f"Refusing upload outside allowed root: {folder_path}")
    encoded = quote(f"{folder_path}/{filename}")
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:{encoded}:/content"
    graph_put_bytes(url, token, content)


def rename_item(drive_id, item_id, new_name, token):
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
    graph_patch_json(url, token, {"name": new_name})


def process_archive(drive_id, folder_path, archive_item, token):
    name = archive_item["name"]
    item_id = archive_item["id"]
    log.info("Processing %s/%s", folder_path, name)

    if not folder_path.startswith(EMAIL_ATTACHMENTS_ROOT):
        log.error("Refusing - outside allowed root: %s", folder_path)
        return

    download_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
    try:
        content = graph_get_bytes(download_url, token)
    except requests.HTTPError as e:
        log.error("Failed to download %s: %s", name, e)
        return

    archive_type = detect_archive_type(content)
    if archive_type is None:
        log.warning("Rejecting %s: unrecognized archive format", name)
        rename_item(drive_id, item_id, f"_REJECTED_unknown_format_{name}", token)
        return
    log.info("  Detected archive type: %s", archive_type)

    ok, reason = validate_archive(content, archive_type)
    if not ok:
        log.warning("Rejecting %s: %s", name, reason)
        rename_item(drive_id, item_id, f"_REJECTED_{reason}_{name}", token)
        return

    extracted = 0
    skipped = 0
    try:
        for entry_filename, entry_bytes in iter_archive_entries(content, archive_type):
            entry_name = os.path.basename(entry_filename)
            if not entry_name:
                continue
            ext = os.path.splitext(entry_name)[1].lower()
            if ext in ALLOWED_EXTENSIONS or ext in (".zip", ".rar"):
                log.info("  Extracting: %s", entry_name)
                upload_to_folder(drive_id, folder_path, entry_name, entry_bytes, token)
                extracted += 1
            else:
                log.info("  Skipping disallowed type: %s", entry_name)
                skipped += 1
    except Exception:
        log.exception("Failed during extraction of %s", name)
        return

    log.info("Extracted %d, skipped %d from %s", extracted, skipped, name)
    rename_item(drive_id, item_id, f"_processed_{name}", token)


def main():
    log.info("Seafarer ZIP/RAR Extractor starting")
    if not EMAIL_ATTACHMENTS_ROOT.startswith("/Email attachments"):
        log.error("Path safety check failed - refusing to run.")
        return 2
    try:
        token = get_access_token()
        log.info("Got Graph token")
        drive_id = get_drive_id(token)
        log.info("Target drive id: %s", drive_id)

        archives = find_archives(drive_id, token)
        for folder_path, archive_item in archives:
            try:
                process_archive(drive_id, folder_path, archive_item, token)
            except Exception:
                log.exception("Error processing %s/%s",
                              folder_path, archive_item.get("name", "?"))

        log.info("Done.")
        return 0
    except Exception:
        log.exception("Fatal error")
        return 1


if __name__ == "__main__":
    sys.exit(main())
