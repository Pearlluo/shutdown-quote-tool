"""
blob_index.py
Manages quote file index as a local JSON file.

File: static/quotes.json
Structure:
{
  "JOB-00001": [
    {
      "drive_item_id": "01ABC...",
      "file_name": "QT-JOB-00001_Quote_20260101.xlsx",
      "client_business_name": "C0001-Example Client",
      "job_title": "Example Job Title",
      "uploaded_at": "2026-01-01T10:30:00+00:00",
      "sharepoint_url": "https://...",
      "payload": { ... }
    }
  ]
}
"""

from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

import storage

INDEX_NAME = "quotes.json"   # blob name / local filename


# ── Read / Write ──────────────────────────────────

def load_index() -> Dict[str, List[Dict[str, Any]]]:
    """Load the full quote index (Azure Blob or local file). Returns {} if missing."""
    return storage.get_json(INDEX_NAME, {}) or {}


def save_index(index: Dict[str, List[Dict[str, Any]]]) -> None:
    """Persist the full quote index (Azure Blob or local file)."""
    storage.put_json(INDEX_NAME, index)


# ── Quote record helpers ──────────────────────────

def get_quotes_for_job(job_id: str) -> List[Dict[str, Any]]:
    """Get all quote records for a job_id."""
    return load_index().get(job_id, [])


def add_quote_record(
    job_id: str,
    drive_item_id: str,
    file_name: str,
    sharepoint_url: str,
    client_business_name: str = "",
    job_title: str = "",
    payload: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Add a new quote record for a job. Returns the new record."""
    index = load_index()

    record = {
        "drive_item_id":      drive_item_id,
        "file_name":          file_name,
        "client_business_name": client_business_name,
        "job_title":          job_title,
        "uploaded_at":        datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "sharepoint_url":     sharepoint_url,
        "payload":            payload or {},
    }

    records = index.setdefault(job_id, [])

    # Replace existing record with same drive_item_id (don't duplicate)
    for i, existing in enumerate(records):
        if existing.get("drive_item_id") == drive_item_id:
            records[i] = record
            save_index(index)
            return record

    records.append(record)
    save_index(index)
    return record


def get_record_by_drive_item_id(drive_item_id: str) -> Optional[Dict[str, Any]]:
    """Find a quote record by DriveItem ID across all jobs."""
    for records in load_index().values():
        for r in records:
            if r.get("drive_item_id") == drive_item_id:
                return r
    return None


def verify_drive_item(drive_item_id: str, token: str, drive_id: str) -> Optional[Dict[str, Any]]:
    """Verify a DriveItem still exists in SharePoint. Returns item info or None."""
    from graph_client import GRAPH_BASE, graph_get
    try:
        return graph_get(f"{GRAPH_BASE}/drives/{drive_id}/items/{drive_item_id}", token)
    except Exception:
        return None


def resolve_quotes_for_job(job_id: str, token: str, drive_id: str) -> List[Dict[str, Any]]:
    """
    Resolve quote records for a job by verifying each DriveItem still exists.
    Returns list of records with 'exists': True/False.
    """
    resolved = []
    for r in get_quotes_for_job(job_id):
        item = verify_drive_item(r["drive_item_id"], token, drive_id)
        if item:
            resolved.append({
                **r,
                "exists":   True,
                "web_url":  item.get("webUrl", r.get("sharepoint_url", "")),
                "modified": item.get("lastModifiedDateTime", ""),
            })
        else:
            resolved.append({**r, "exists": False})
    return resolved
