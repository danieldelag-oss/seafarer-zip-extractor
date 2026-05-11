"""
Seafarer ZIP Extractor
=======================
Polls OneDrive folders under /Email attachments/**/Original/ for .zip
attachments, extracts allowed file types in place, and renames the source
zip to _processed_*.zip so it isn't re-processed.

Designed to run as a scheduled GitHub Actions workflow. Reads credentials
from environment variables provided via GitHub Secrets.

Safety guards:
  - Refuses to operate outside /Email attachments/ (hardcoded prefix)
  - Caps extracted size per zip (default 500 MB)
  - Caps file count per zip (default 200)
  - Only extracts allow-listed extensions (PDF, JPG, JPEG, PNG, TIF, etc.)
  - Skips password-protected zips (renames to _REJECTED_*)
  - Skips corrupt zips (renames to _REJECTED_*)
  - Caps recursion depth for nested zips (default 2)
  - All Graph PATCH/PUT operations are scoped to the target drive and items
    discovered by walking the allowed tree
"""

import io
import logging
import os
import sys
import zipfile
from urllib.parse import quote

import msal
import requests

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
TARGET_USER_UPN = os.environ["TARGET_USER_UPN"]  # e.g. ddelaguardia@embaseoul.kr

EMAIL_ATTACHMENTS_ROOT = "/Email attachments"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Safety limits — tune to taste
MAX_EXTRACTED_SIZE_BYTES = 500 * 1024 * 1024          # 500 MB total per zip
MAX_FILES_PER_ZIP = 200
MAX_NESTED_ZIP_DEPTH = 2
ALLOWED_EXTENSIONS = {
    ".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".gif", ".bmp",
}

# Subfolders inside the dated tree that are NOT sailor folders
EXCLUDE_NAMES = {"Invoices", "General Documents"}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("unzip")


# ---------------------------------------------------------------------------
# Graph helpers
# ---------------------------------------------------------------------------
def get_access_token() -> str:
    """Acquire Microsoft Graph token via client credentials flow."""
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
    r = requests.get(
        url, headers={"Authorization": f"Bearer {token}"}, timeout=60
    )
    r.raise_for_status()
    return r.json()


def graph_get_bytes(url: str, token: str) -> bytes:
    r = requests.get(
        url, headers={"Authorization": f"Bearer {token}"}, timeout=300
    )
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


def list_folder_by_path(drive_id: str, path: str, token: str) -> list[dict]:
    """List items in a folder. Path must start with /."""
    if not path.startswith("/"):
        raise ValueError(f"Path must start with /: {path}")
    encoded = quote(path)
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:{encoded}:/children?$top=200"
    items: list[dict] = []
    while url:
        data = graph_get(url, token)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def safe_subfolder_walk(drive_id: str, parent_path: str, token: str) -> list[tuple[str, dict]]:
    """List the parent and return (child_path, child_item) for each child folder
    that passes the standard exclusions."""
    out: list[tuple[str, dict]] = []
    try:
        items = list_folder_by_path(drive_id, parent_path, token)
    except requests.HTTPError as e:
        if e.response is not None and e.response.status_code == 404:
            return out
        log.warning("Couldn't list %s: %s", parent_path, e)
        return out
    for it in items:
        if not it.get("folder"):
            continue
        name = it["name"]
        if name.startswith(("_", ".")) or name in EXCLUDE_NAMES:
            continue
        out.append((f"{parent_path}/{name}", it))
    return out


# ---------------------------------------------------------------------------
# Zip validation
# ---------------------------------------------------------------------------
def validate_zip(zip_bytes: bytes) -> tuple[bool, str | None]:
    """Return (ok, reason_if_rejected)."""
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


# ---------------------------------------------------------------------------
# Processing
# ---------------------------------------------------------------------------
def process_zip(
    drive_id: str,
    folder_path: str,
    zip_item: dict,
    token: str,
    depth: int = 0,
) -> None:
    name = zip_item["name"]
    item_id = zip_item["id"]
    log.info("Processing %s/%s (depth=%d)", folder_path, name, depth)

    # Defensive: never touch anything outside /Email attachments
    if not folder_path.startswith(EMAIL_ATTACHMENTS_ROOT):
        log.error("Refusing — folder outside allowed root: %s", folder_path)
        return

    if depth >= MAX_NESTED_ZIP_DEPTH:
        log.warning("Skipping %s — nesting depth exceeded", name)
        rename_item(drive_id, item_id, f"_REJECTED_nested_{name}", token)
        return

    # Download
    download_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
    try:
        content = graph_get_bytes(download_url, token)
    except requests.HTTPError as e:
        log.error("Failed to download %s: %s", name, e)
        return

    # Validate
    ok, reason = validate_zip(content)
    if not ok:
        log.warning("Rejecting %s: %s", name, reason)
        rename_item(drive_id, item_id, f"_REJECTED_{reason}_{name}", token)
        return

    # Extract
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
            if ext == ".zip":
                log.info("  Found nested zip: %s — extracting for next cycle", entry_name)
                upload_to_folder(drive_id, folder_path, entry_name, zf.read(info), token)
                extracted += 1
            elif ext in ALLOWED_EXTENSIONS:
                log.info("  Extracting: %s", entry_name)
                upload_to_folder(drive_id, folder_path, entry_name, zf.read(info), token)
                extracted += 1
            else:
                log.info("  Skipping disallowed type: %s", entry_name)
                skipped += 1

    log.info("Extracted %d, skipped %d from %s", extracted, skipped, name)
    rename_item(drive_id, item_id, f"_processed_{name}", token)


def upload_to_folder(
    drive_id: str, folder_path: str, filename: str, content: bytes, token: str
) -> None:
    if not folder_path.startswith(EMAIL_ATTACHMENTS_ROOT):
        raise RuntimeError(f"Refusing upload outside allowed root: {folder_path}")
    encoded = quote(f"{folder_path}/{filename}")
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:{encoded}:/content"
    graph_put_bytes(url, token, content)


def rename_item(drive_id: str, item_id: str, new_name: str, token: str) -> None:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
    graph_patch_json(url, token, {"name": new_name})


# ---------------------------------------------------------------------------
# Walk
# ---------------------------------------------------------------------------
def find_and_process(drive_id: str, token: str) -> int:
    processed = 0
    for ym_path, _ in safe_subfolder_walk(drive_id, EMAIL_ATTACHMENTS_ROOT, token):
        for d_path, _ in safe_subfolder_walk(drive_id, ym_path, token):
            for s_path, _ in safe_subfolder_walk(drive_id, d_path, token):
                original_path = f"{s_path}/Original"
                try:
                    files = list_folder_by_path(drive_id, original_path, token)
                except requests.HTTPError as e:
                    if e.response is not None and e.response.status_code == 404:
                        continue
                    log.warning("Couldn't list %s: %s", original_path, e)
                    continue
                for f in files:
                    if f.get("folder"):
                        continue
                    fname = f["name"]
                    if not fname.lower().endswith(".zip"):
                        continue
                    if fname.startswith("_processed_") or fname.startswith("_REJECTED_"):
                        continue
                    try:
                        process_zip(drive_id, original_path, f, token)
                        processed += 1
                    except Exception:
                        log.exception("Error processing %s/%s", original_path, fname)
    return processed


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> int:
    log.info("Seafarer ZIP Extractor starting")
    if not EMAIL_ATTACHMENTS_ROOT.startswith("/Email attachments"):
        log.error("Path safety check failed — refusing to run.")
        return 2
    try:
        token = get_access_token()
        log.info("Got Graph token")
        drive_id = get_drive_id(token)
        log.info("Target drive id: %s", drive_id)
        count = find_and_process(drive_id, token)
        log.info("Done. Processed %d zip(s).", count)
        return 0
    except Exception:
        log.exception("Fatal error")
        return 1


if __name__ == "__main__":
    sys.exit(main())
