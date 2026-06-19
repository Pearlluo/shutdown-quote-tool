"""
operations_folders.py
Path: Operations/01 Acme Contracts/{ProjectFolder}/{JobFolder}/00 Commercial (Secured)/01 Client Quotes
"""

import os
import re
from typing import Any, Dict, List, Optional
from urllib.parse import quote as url_quote, unquote, urlparse, parse_qs

from graph_client import (
    GRAPH_BASE,
    graph_get,
    graph_put_bytes,
    get_operations_drive_id,
)

BASE_PATH         = "01 Acme Contracts"
COMMERCIAL_FOLDER = "00 Commercial (Secured)"
QUOTES_FOLDER     = "01 Client Quotes"
OPERATIONS_DRIVE  = "Operations"
# Last-resort "browse manually" link shown in the UI. Configured via env so the
# company URL isn't hard-coded; if unset the link simply doesn't appear.
IMS_OPS_URL       = os.getenv("IMS_OPS_URL", "")


# ── Low-level helpers ─────────────────────────────

def _children_all(drive_id: str, item_id: str, token: str) -> List[Dict]:
    """List ALL children (folders + files) of a drive item."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/children?$top=500"
    items = []
    while url:
        data = graph_get(url, token)
        items += data.get("value", [])
        url = data.get("@odata.nextLink")
    return items


def _children(drive_id: str, item_id: str, token: str) -> List[Dict]:
    """List only folder children."""
    return [i for i in _children_all(drive_id, item_id, token) if i.get("folder") is not None]


def _get_item_by_path(drive_id: str, rel_path: str, token: str) -> Optional[Dict[str, Any]]:
    """GET /drives/{drive_id}/root:/{rel_path} — returns None on any error."""
    encoded = url_quote(rel_path, safe="/")
    try:
        return graph_get(f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded}", token)
    except Exception:
        return None


def _drive_rel_path(sharepoint_url: str) -> str:
    """
    Extract path relative to the Operations drive root from a SharePoint URL.
    Handles both direct URLs and Forms URLs (path in ?id= parameter).

    Direct:  ".../sites/IMS/Operations/01 Acme Contracts/..." → "01 Acme Contracts/..."
    Forms:   ".../Operations/Forms/...aspx?id=%2F...Operations%2F01%20Acme..." → same
    """
    marker = f"/{OPERATIONS_DRIVE}/"

    # Try ?id= parameter first (Forms URL)
    parsed = urlparse(sharepoint_url)
    if parsed.query:
        qs = parse_qs(parsed.query)
        for id_val in qs.get("id", []):
            text = unquote(id_val)
            idx = text.find(marker)
            if idx != -1:
                return text[idx + len(marker):]

    # Direct URL
    text = unquote(sharepoint_url)
    idx = text.find(marker)
    if idx == -1:
        return ""
    # If the remaining path starts with "Forms/" it's a Forms base URL — skip
    rel = text[idx + len(marker):]
    if rel.startswith("Forms/"):
        return ""
    return rel


def _path_parent(rel_path: str) -> str:
    """Return the parent directory of a relative path."""
    parts = rel_path.rstrip("/").rsplit("/", 1)
    return parts[0] if len(parts) > 1 else ""


def _is_app_quote(item: Dict[str, Any], job_id: str) -> bool:
    """
    Return True if this DriveItem is an app-generated quote for job_id.
    Primary rule (matches the /api/job-quotes filter): a .xlsx/.xlsm whose name
    starts with 'QT-' and contains the job id. The legacy
    '{job_id}_Quote_{8-digit-date}.xlsx' pattern is still accepted so older
    quotes keep showing up.
    """
    if not item.get("file"):
        return False
    name = (item.get("name") or "").strip()
    if not name.lower().endswith((".xlsx", ".xlsm")):
        return False
    up = name.upper()
    if up.startswith("QT-") and job_id.upper() in up:
        return True
    # legacy naming
    return bool(re.match(rf'^{re.escape(job_id)}_Quote_\d{{8}}\.xlsx$', name, re.I))


# ── Quotes folder finder (3-URL priority) ─────────

def find_quotes_folder(token: str,
                       url_commercial: str = "",
                       url_job: str = "",
                       url_planning: str = "") -> Optional[Dict[str, Any]]:
    """
    Find the 01 Client Quotes folder using 3 URL candidates in priority order:

    1. url_commercial → {commercial_path}/01 Client Quotes
    2. url_job        → {job_path}/00 Commercial (Secured)/01 Client Quotes
    3. url_planning   → parent of planning folder → 00 Commercial (Secured)/01 Client Quotes

    Returns the folder DriveItem or None if all fail.
    """
    drive_id = get_operations_drive_id(token)
    candidates = []

    if url_commercial:
        rel = _drive_rel_path(url_commercial)
        if rel:
            candidates.append(f"{rel.rstrip('/')}/{QUOTES_FOLDER}")

    if url_job:
        rel = _drive_rel_path(url_job)
        if rel:
            candidates.append(f"{rel.rstrip('/')}/{COMMERCIAL_FOLDER}/{QUOTES_FOLDER}")

    if url_planning:
        rel = _drive_rel_path(url_planning)
        if rel:
            parent = _path_parent(rel)
            if parent:
                candidates.append(f"{parent}/{COMMERCIAL_FOLDER}/{QUOTES_FOLDER}")

    for path in candidates:
        item = _get_item_by_path(drive_id, path, token)
        if item:
            return item

    return None


def list_existing_quotes(job_id: str, token: str,
                         url_commercial: str = "",
                         url_job: str = "",
                         url_planning: str = "",
                         client_business_name: str = "") -> List[Dict[str, Any]]:
    """
    List app-generated quote files in 01 Client Quotes for a job.
    Uses find_quotes_folder (3-URL priority), falls back to IMS folder scan.
    Only returns files matching the app quote naming pattern.
    """
    drive_id = get_operations_drive_id(token)

    # Try the 3-URL approach first
    quotes_folder = find_quotes_folder(token, url_commercial, url_job, url_planning)

    # Fallback: full IMS traversal
    if not quotes_folder:
        quotes_folder = _find_quotes_via_ims(job_id, token, client_business_name)

    if not quotes_folder:
        return []

    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{quotes_folder['id']}/children?$top=500"
    data = graph_get(url, token)

    files = [
        {
            "drive_item_id": f["id"],
            "file_name":     f["name"],
            "web_url":       f.get("webUrl", ""),
            "modified":      f.get("lastModifiedDateTime", ""),
            "size":          f.get("size", 0),
        }
        for f in data.get("value", [])
        if f.get("file") and _is_app_quote(f, job_id)
    ]
    return sorted(files, key=lambda x: x["modified"], reverse=True)


def upload_quote_file(job_id: str, file_bytes: bytes, file_name: str, token: str,
                      url_commercial: str = "",
                      url_job: str = "",
                      url_planning: str = "",
                      client_business_name: str = "",
                      direct_folder_url: str = "") -> Dict[str, Any]:
    """Upload Excel to a quotes folder.

    Normal mode: locate 01 Client Quotes (3-URL priority → deep IMS traversal).
    Direct mode (direct_folder_url set): upload straight into that exact folder
    — used when 01 Client Quotes can't be found and the user picks a folder
    (e.g. the Ops/contract folder) to drop the file into as-is.
    """
    drive_id = get_operations_drive_id(token)

    if direct_folder_url:
        quotes_folder = None
        rel = _drive_rel_path(direct_folder_url)
        if rel:
            quotes_folder = _get_item_by_path(drive_id, rel, token)
        if not quotes_folder:
            raise RuntimeError("Selected save folder could not be resolved in SharePoint")
    else:
        quotes_folder = find_quotes_folder(token, url_commercial, url_job, url_planning)
        if not quotes_folder:
            quotes_folder = _find_quotes_via_ims(job_id, token, client_business_name)
        if not quotes_folder:
            raise RuntimeError(f"01 Client Quotes folder not found for job {job_id}")

    encoded_name = url_quote(file_name, safe="")
    upload_url = (
        f"{GRAPH_BASE}/drives/{drive_id}/items/"
        f"{quotes_folder['id']}:/{encoded_name}:/content"
    )
    result = graph_put_bytes(
        upload_url, token, file_bytes,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    return {
        "drive_item_id": result.get("id", ""),
        "file_name":     result.get("name", file_name),
        "web_url":       result.get("webUrl", ""),
    }


# ── IMS traversal fallback ────────────────────────

def _client_to_folder_name(name: str) -> str:
    """Map a client business name to its contract folder name.

    Folders are 'Cxxxx Name'. The business name may arrive in several formats:
        'C0001-Example Client'   /  'C0001 - Example Client'  /  'C0001  Example Client'
    All of these must resolve to 'C0001 Example Client'. So: drop the single
    separator (an optional hyphen with optional surrounding spaces) right after
    the contract number, then collapse any repeated whitespace to one space.
    A later hyphen in the name (e.g. 'C0001 PP-Site') is left untouched.
    """
    name = (name or "").strip()
    m = re.match(r'^(C\d{4})\s*-?\s*(.*)$', name)
    if m:
        name = f"{m.group(1)} {m.group(2)}"
    return re.sub(r'\s+', ' ', name).strip()


def _find_quotes_via_ims(job_id: str, token: str,
                          client_business_name: str = "") -> Optional[Dict[str, Any]]:
    """Full IMS traversal: find job folder then navigate to 01 Client Quotes."""
    drive_id = get_operations_drive_id(token)

    job_folder = _find_job_folder(job_id, token, drive_id, client_business_name)
    if not job_folder:
        return None

    children = _children(drive_id, job_folder["id"], token)
    commercial = next((f for f in children if f.get("name", "").startswith("00 Commercial")), None)
    if not commercial:
        return None

    children2 = _children(drive_id, commercial["id"], token)
    return next((f for f in children2 if f.get("name", "").startswith("01 Client Quotes")), None)


def _find_job_folder(job_id: str, token: str, drive_id: str,
                     client_business_name: str = "") -> Optional[Dict[str, Any]]:
    # Fast path via contract folder name
    if client_business_name:
        folder_name = _client_to_folder_name(client_business_name)
        pf = _get_item_by_path(drive_id, f"{BASE_PATH}/{folder_name}", token)
        if pf:
            for jf in _children(drive_id, pf["id"], token):
                if jf.get("name", "").upper().startswith(job_id.upper()):
                    return jf

    # Full scan fallback
    base = _get_item_by_path(drive_id, BASE_PATH, token)
    if not base:
        return None
    for pf in _children(drive_id, base["id"], token):
        for jf in _children(drive_id, pf["id"], token):
            if jf.get("name", "").upper().startswith(job_id.upper()):
                return jf
    return None


def get_job_folder_links(job_id: str, token: str,
                         client_business_name: str = "") -> Dict[str, Any]:
    """Return fresh webUrls for Job Root, Commercial, Planning folders."""
    drive_id = get_operations_drive_id(token)
    job_folder = _find_job_folder(job_id, token, drive_id, client_business_name)
    if not job_folder:
        return {"found": False}

    url_commercial = url_planning = ""
    for f in _children(drive_id, job_folder["id"], token):
        name = f.get("name", "")
        if name.startswith("00 Commercial"):
            url_commercial = f.get("webUrl", "")
        elif "plan" in name.lower():
            url_planning = f.get("webUrl", "")

    return {
        "found":           True,
        "job_folder_name": job_folder.get("name", ""),
        "url_job":         job_folder.get("webUrl", ""),
        "url_commercial":  url_commercial,
        "url_planning":    url_planning,
    }


def find_quotes_folder_any(token: str,
                           url_commercial: str = "",
                           url_job: str = "",
                           url_planning: str = "",
                           job_id: str = "",
                           client_business_name: str = "") -> Optional[Dict[str, Any]]:
    """Locate 01 Client Quotes by the 3-URL priority first; if every link is
    stale/dead, fall back to a full IMS traversal (find the contract folder by
    name → job folder → 00 Commercial → 01 Client Quotes). This keeps the quote
    lookup working even when the Commercial/Ops/Planning links have expired."""
    qf = find_quotes_folder(token, url_commercial, url_job, url_planning)
    if qf:
        return qf
    if job_id:
        return _find_quotes_via_ims(job_id, token, client_business_name)
    return None


def get_contract_folder_url(token: str, client_business_name: str) -> Dict[str, str]:
    """Return {name, url} of the contract folder under 01 Acme Contracts
    (e.g. 'C0001 Example Client') so the popup can drop the user there to filter
    manually when all links fail. Empty dict if it can't be resolved."""
    if not client_business_name:
        return {}
    drive_id = get_operations_drive_id(token)
    folder_name = _client_to_folder_name(client_business_name)
    pf = _get_item_by_path(drive_id, f"{BASE_PATH}/{folder_name}", token)
    if pf:
        return {"name": pf.get("name", folder_name), "url": pf.get("webUrl", "")}
    return {}
