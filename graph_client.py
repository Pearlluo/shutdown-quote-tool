"""
graph_client.py
Handles: token, graph_get/patch/post, site_id, drive_id
"""

import os
import requests
from typing import Any, Dict, Optional
from dotenv import load_dotenv

load_dotenv()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

TENANT_ID     = os.getenv("SHAREPOINT_TENANT_ID")
CLIENT_ID     = os.getenv("SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")
SP_HOST       = os.getenv("SHAREPOINT_HOST", "your-tenant.sharepoint.com")

# BMS site (JMS-Jobs, PPL-People, etc.)
BMS_SITE_NAME = os.getenv("SITE_NAME", "BMS")

# IMS site (Operations document library / file folders)
IMS_SITE_NAME = os.getenv("SITE_NAME1", "IMS")


# ── Token cache ───────────────────────────────────
_token_cache: Dict[str, str] = {}

def get_token() -> str:
    if _token_cache.get("token"):
        return _token_cache["token"]

    missing = [k for k, v in {
        "SHAREPOINT_TENANT_ID": TENANT_ID,
        "SHAREPOINT_CLIENT_ID": CLIENT_ID,
        "SHAREPOINT_CLIENT_SECRET": CLIENT_SECRET,
    }.items() if not v]
    if missing:
        raise RuntimeError(f"Missing .env values: {', '.join(missing)}")

    r = requests.post(
        TOKEN_URL_TEMPLATE.format(tenant_id=TENANT_ID),
        data={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=60,
    )
    r.raise_for_status()
    token = r.json()["access_token"]
    _token_cache["token"] = token
    return token

def clear_token_cache():
    _token_cache.clear()


# ── HTTP helpers ──────────────────────────────────
def _headers(token: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}

def graph_get(url: str, token: str) -> Dict[str, Any]:
    r = requests.get(url, headers=_headers(token), timeout=90)
    if not r.ok:
        raise RuntimeError(f"GET {url}\n{r.status_code}: {r.text}")
    return r.json()

def graph_patch(url: str, token: str, body: Dict[str, Any]) -> Dict[str, Any]:
    headers = _headers(token)
    headers["Content-Type"] = "application/json"
    r = requests.patch(url, headers=headers, json=body, timeout=60)
    if not r.ok:
        raise RuntimeError(f"PATCH {url}\n{r.status_code}: {r.text}")
    return r.json()

def graph_post(url: str, token: str, body: Dict[str, Any]) -> Dict[str, Any]:
    headers = _headers(token)
    headers["Content-Type"] = "application/json"
    r = requests.post(url, headers=headers, json=body, timeout=60)
    if not r.ok:
        raise RuntimeError(f"POST {url}\n{r.status_code}: {r.text}")
    return r.json()

def graph_put_bytes(url: str, token: str, data: bytes, content_type: str = "application/octet-stream") -> Dict[str, Any]:
    headers = {"Authorization": f"Bearer {token}", "Content-Type": content_type}
    r = requests.put(url, headers=headers, data=data, timeout=120)
    if not r.ok:
        raise RuntimeError(f"PUT {url}\n{r.status_code}: {r.text}")
    return r.json()


# ── Site IDs ──────────────────────────────────────
_site_id_cache: Dict[str, str] = {}

def get_site_id(site_name: str, token: str) -> str:
    if site_name in _site_id_cache:
        return _site_id_cache[site_name]
    url = f"{GRAPH_BASE}/sites/{SP_HOST}:/sites/{site_name}"
    data = graph_get(url, token)
    sid = data["id"]
    _site_id_cache[site_name] = sid
    return sid

def get_bms_site_id(token: str) -> str:
    return get_site_id(BMS_SITE_NAME, token)

def get_ims_site_id(token: str) -> str:
    return get_site_id(IMS_SITE_NAME, token)


# ── Drive IDs ─────────────────────────────────────
_drive_id_cache: Dict[str, str] = {}

def get_drive_id(site_id: str, drive_name: str, token: str) -> str:
    cache_key = f"{site_id}:{drive_name}"
    if cache_key in _drive_id_cache:
        return _drive_id_cache[cache_key]

    url = f"{GRAPH_BASE}/sites/{site_id}/drives"
    data = graph_get(url, token)
    for d in data.get("value", []):
        if d.get("name") == drive_name:
            _drive_id_cache[cache_key] = d["id"]
            return d["id"]

    # fallback: return default drive
    url2 = f"{GRAPH_BASE}/sites/{site_id}/drive"
    d2 = graph_get(url2, token)
    _drive_id_cache[cache_key] = d2["id"]
    return d2["id"]

def get_operations_drive_id(token: str) -> str:
    """IMS site → Operations document library drive"""
    ims_site_id = get_ims_site_id(token)
    return get_drive_id(ims_site_id, "Operations", token)


# ── List helpers ──────────────────────────────────
def get_list_id(site_id: str, list_name: str, token: str) -> str:
    url = f"{GRAPH_BASE}/sites/{site_id}/lists?$select=id,name,displayName"
    data = graph_get(url, token)
    for item in data.get("value", []):
        if item.get("name") == list_name or item.get("displayName") == list_name:
            return item["id"]
    available = [f"{x.get('displayName')} ({x.get('name')})" for x in data.get("value", [])]
    raise RuntimeError(f"List not found: {list_name}\nAvailable:\n- " + "\n- ".join(available))

def get_list_items(site_id: str, list_name: str, token: str) -> list:
    list_id = get_list_id(site_id, list_name, token)
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?$expand=fields&$top=5000"
    items = []
    while url:
        data = graph_get(url, token)
        for item in data.get("value", []):
            fields = item.get("fields", {})
            fields["_item_id"] = item.get("id")
            items.append(fields)
        url = data.get("@odata.nextLink")
    return items
