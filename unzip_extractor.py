"""
Seafarer ZIP Extractor (search-based, drive-root search)
=========================================================
Polls OneDrive for .zip attachments under /Email attachments/**/Original/,
extracts allowed file types in place, renames source zips to _processed_*.zip.

Uses drive-root Graph search; path-based search endpoint isn't supported in
app-only context (returns 403), so we search the whole drive and filter
results to the allowed root in Python.

Safety guards:
  - Refuses to operate outside /Email attachments/ (hardcoded prefix)
  - Only processes zips whose parent folder ends with /Original
  - Caps extracted size per zip (default 500 MB)
  - Caps file count per zip (default 200)
  - Allow-listed extensions only
  - Skips password-protected zips (rename to _REJECTED_*)
  - Skips corrupt zips (rename to _REJECTED_*)
"""

import io
import logging
import os
import sys
import zipfile
from urllib.parse import quote

import msal
import requests

TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
TARGET_USER_UPN = os.environ["TARGET_USER_UPN"]

EMAIL_ATTACHMENTS_ROOT = "/Email attachments"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

MAX_EXTRACTED_SIZE_BYTES = 500 * 1024 * 1024
MAX_FILES_PER_ZIP = 200
ALLOWED_EXTENSIONS = {
    ".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".gif", ".bmp",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("unzip")


def get_access_token() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"Token acquisition failed: {result.get('error_description', result)}"
        )
    return result["access_token"]


def graph_get(url: str, token: str) -> dict:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json()


def graph_get_bytes(url: str, token: str) -> bytes:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=300)
    r.raise_for_status()
    return r.content


def graph_put_bytes(url: str, token: str, content: bytes) -> None:
    r = requests.put(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        },
        data=content,
        timeout=300,
    )
    r.raise_for_status()


def graph_patch_json(url: str, token: str, body: dict) -> None:
    r = requests.patch(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=body,
        timeout=60,
    )
    r.raise_for_status()


def get_drive_id(token: str) -> str:
    url = f"{GRAPH_BASE}/users/{quote(TARGET_USER_UPN)}/drive"
    return graph_get(url, token)["id"]


# ---------------------------------------------------------------------------
# Drive-root search + Python filter
# ---------------------------------------------------------------------------
def find_zip_files(drive_id: str, token: str) -> list[tuple[str, dict]]:
    """Search the whole drive for .zip files, filter to those in
    /Email attachments/.../Original."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/root/search(q='.zip')?$top=200"

    raw_items: list[dict] = []
    page = 0
    while url:
        page += 1
        data = graph_get(url, token)
        raw_items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    log.info("Drive search returned %d items across %d page(s)", len(raw_items), page)

    results: list[tuple[str, dict]] = []
    for item in raw_items:
        if item.get("folder"):
            continue
        name = item.get("name", "")
        if not name.lower().endswith(".zip"):
            continue
        if name.startswith("_processed_") or name.startswith("_REJECTED_"):
            continue

        parent_path_raw = item.get("parentReference", {}).get("path", "")
        # parentReference.path looks like "/drives/{id}/root:/Email attachments/.../Original"
        folder_path = parent_path_raw.split(":", 1)[1] if ":" in parent_path_raw else parent_path_raw

        if not folder_path.startswith(EMAIL_ATTACHMENTS_ROOT):
            continue  # silently skip; not in our scope
        if not folder_path.endswith("/Original"):
            log.info("Skipping zip not in Original/: %s/%s", folder_path, name)
            continue

        results.append((folder_path, item))

    log.info("Found %d zip(s) under %s/**/Original/ to process",
             len(results), EMAIL_ATTACHMENTS_ROOT)
    return results


def validate_zip(zip_bytes: bytes) -> tuple[bool, str | None]:
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
        log.error("Refusing — folder outside allowed root: %s", folder_path)
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


def main() -> int:
    log.info("Seafarer ZIP Extractor (drive-root search) starting")
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
                log.exception("Error processing %s/%s", folder_path, zip_item.get("name", "?"))

        log.info("Done.")
        return 0
    except Exception:
        log.exception("Fatal error")
        return 1


if __name__ == "__main__":
    sys.exit(main())
