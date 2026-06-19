import os
import re
import json
import argparse
from datetime import datetime, date
from urllib.parse import quote
from typing import Any, Dict, List, Optional
from pathlib import Path

import requests
from dotenv import load_dotenv
from flask import Flask, render_template, request, jsonify

import storage   # Azure Blob (or local-file fallback) persistence layer

QUOTE_DATA_NAME = "quote_data.json"   # blob name / local filename


# ============================================================
# Config
# ============================================================

load_dotenv()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")
CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")

SHAREPOINT_HOST = os.getenv("SHAREPOINT_HOST", "your-tenant.sharepoint.com")
SITE_NAME = os.getenv("SITE_NAME", "BMS")

LIST_NAMES = {
    "jobs": os.getenv("LIST_ID_JMS_JOBS", "JMS-Jobs"),
    "rates": os.getenv("LIST_ID_JMS_RATES", "JMS-Rates"),
    "positions": os.getenv("LIST_ID_PPL_POSITIONS", "PPL-Positions"),
    "projects": os.getenv("LIST_ID_JMS_PROJECTS", "JMS-Projects"),
    "people": os.getenv("LIST_ID_PPL_PEOPLE", "PPL-People"),
    "clients": os.getenv("LIST_ID_JMS_CLIENTS", "JMS-Clients"),
}

CLIENT_QUOTES_FOLDER_NAME = "01 Client Quotes"

OPEN_STATUSES = {
    "",
    "In Progress",
    "Quote Sent",
    "Requested",
    "Approved",
    "Repeat Order",
    "Job Lead",
    "UNSURE",
}

CLOSED_STATUSES = {
    "Cancelled",
    "Complete- INVOICE",
    "COMPLETE/CLOSED",
    "Unsuccesful",
    "Unsuccessful",
}

SHIFT_CODES = {
    "FIP0": {"hours": 0, "type": "Day"},
    "FOA3": {"hours": 3, "type": "Day"},
    "DS12.5": {"hours": 12.5, "type": "Day"},
    "DS1": {"hours": 1, "type": "Day"},
    "DAYRATE": {"hours": 1, "type": "Day"},
    "NS12.5": {"hours": 12.5, "type": "Night"},
    "NS1": {"hours": 1, "type": "Night"},
    "FOA0": {"hours": 0, "type": "Day"},
    "DIA12": {"hours": 12, "type": "Day"},
    "DOA12": {"hours": 12, "type": "Day"},
}


# ============================================================
# Helpers
# ============================================================

def safe_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def safe_num(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except Exception:
        return None


def safe_date(value: Any) -> Optional[str]:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    text = str(value)
    return text[:10] if text else None


def get_first(fields: Dict[str, Any], keys: List[str], default: Any = None) -> Any:
    for key in keys:
        if key in fields and fields[key] not in (None, ""):
            return fields[key]
    return default


def lookup_value(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, dict):
        return safe_str(
            value.get("LookupValue")
            or value.get("lookupValue")
            or value.get("Title")
            or value.get("title")
            or value.get("Name")
            or value.get("name")
            or value.get("Email")
            or value.get("email")
            or ""
        )

    if isinstance(value, list):
        values = [lookup_value(v) for v in value]
        return ", ".join([v for v in values if v])

    return safe_str(value)


def extract_url(value: Any) -> Optional[str]:
    if not value:
        return None

    text = str(value)
    match = re.search(r"https://[^<>\s]+", text)

    if match:
        return match.group(0).strip()

    return None


def append_child_folder_url(parent_url: Optional[str], folder_name: str) -> Optional[str]:
    if not parent_url:
        return None

    return f"{parent_url.rstrip('/')}/{quote(folder_name)}"


def infer_client_from_folder(url: Optional[str]) -> str:
    if not url:
        return ""

    text = str(url).replace("%20", " ")
    match = re.search(r"C\d{4}\s+([^/]+)", text)

    if not match:
        return ""

    client = match.group(1).strip()
    client = re.sub(r"\s+OLD\s+\(Archived\)", "", client, flags=re.I)
    client = re.sub(r"\s+\(Archived\)", "", client, flags=re.I)

    return client.strip()


def is_open_status(status: str) -> bool:
    status = safe_str(status)

    if status in CLOSED_STATUSES:
        return False

    if status in OPEN_STATUSES:
        return True

    return True


# ============================================================
# Graph API
# ============================================================

def get_access_token() -> str:
    required = {
        "SHAREPOINT_TENANT_ID": TENANT_ID,
        "SHAREPOINT_CLIENT_ID": CLIENT_ID,
        "SHAREPOINT_CLIENT_SECRET": CLIENT_SECRET,
    }

    missing = [key for key, value in required.items() if not value]

    if missing:
        raise RuntimeError(f"Missing .env values: {', '.join(missing)}")

    token_url = TOKEN_URL_TEMPLATE.format(tenant_id=TENANT_ID)

    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    response = requests.post(token_url, data=data, timeout=60)
    response.raise_for_status()

    return response.json()["access_token"]


def graph_get(url: str, token: str) -> Dict[str, Any]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    response = requests.get(url, headers=headers, timeout=90)

    if not response.ok:
        raise RuntimeError(
            f"Graph request failed:\n"
            f"URL: {url}\n"
            f"Status: {response.status_code}\n"
            f"Response: {response.text}"
        )

    return response.json()


def get_site_id(token: str) -> str:
    url = f"{GRAPH_BASE}/sites/{SHAREPOINT_HOST}:/sites/{SITE_NAME}"
    data = graph_get(url, token)
    return data["id"]


def get_list_id_by_name(site_id: str, list_name: str, token: str) -> str:
    url = f"{GRAPH_BASE}/sites/{site_id}/lists?$select=id,name,displayName"
    data = graph_get(url, token)

    for item in data.get("value", []):
        if item.get("name") == list_name or item.get("displayName") == list_name:
            return item["id"]

    available = [
        f"{x.get('displayName')} ({x.get('name')})"
        for x in data.get("value", [])
    ]

    raise RuntimeError(
        f"Cannot find SharePoint list: {list_name}\n"
        f"Available lists:\n- " + "\n- ".join(available)
    )


def get_list_items(site_id: str, list_name: str, token: str) -> List[Dict[str, Any]]:
    list_id = get_list_id_by_name(site_id, list_name, token)

    url = (
        f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
        f"?$expand=fields"
        f"&$top=5000"
    )

    items = []

    while url:
        data = graph_get(url, token)

        for item in data.get("value", []):
            fields = item.get("fields", {})
            fields["_item_id"] = item.get("id")
            items.append(fields)

        url = data.get("@odata.nextLink")

    return items


# ============================================================
# Transform SharePoint Lists
# ============================================================

def transform_jobs(items: List[Dict[str, Any]], active_only: bool = False) -> List[Dict[str, Any]]:
    jobs = []

    for f in items:
        job_id = safe_str(get_first(f, ["JobID", "Job_x0020_ID", "Title"]))
        title = safe_str(get_first(f, ["Title", "JobTitle", "Job_x0020_Title"]))

        if not job_id:
            continue

        status = lookup_value(get_first(
            f,
            ["JobStatus", "JobStatusValue", "JobStatus0", "Status"]
        ))

        active_raw = get_first(f, ["Active"], True)
        active = str(active_raw).upper() not in {"FALSE", "0", "NO"}

        open_job = is_open_status(status)

        if active_only and not open_job:
            continue

        # Project is a Lookup field — real field name is ProjectLookupId (numeric ID)
        project_lookup_id = safe_str(f.get("ProjectLookupId", ""))
        project = lookup_value(get_first(f, ["Project", "ProjectLookup", "ProjectId"]))
        site = lookup_value(get_first(f, ["WorkLocation", "WorkLocation0", "Site", "WorkLocationLookupId"]))

        ops_raw = get_first(f, ["OpsFolder", "OpsFolderEdit", "JobFolder"])
        com_raw = get_first(f, ["ComFolderContribute", "ComFolder", "CommercialFolder"])
        planning_raw = get_first(f, ["PlanningFolder", "PlanningFolderContribute"])
        client_raw = get_first(f, ["ClientFolder", "ClientFolderContribute"])

        url_job = extract_url(ops_raw)
        url_commercial = extract_url(com_raw)
        url_planning = extract_url(planning_raw)
        url_client_shared = extract_url(client_raw)
        url_client_quotes = append_child_folder_url(url_commercial, CLIENT_QUOTES_FOLDER_NAME)

        # Extract full contract folder name (e.g. "C0001 Example Client") from IMS URLs
        contract_folder = ""
        for _u in [url_job, url_commercial]:
            if _u:
                import re as _re
                from urllib.parse import unquote as _unquote
                _m = _re.search(r'01[_ ]Acme[_ ]Contracts[/\\%]([^/%?]+)', _unquote(str(_u)))
                if _m:
                    contract_folder = _m.group(1).strip()
                    break

        client = lookup_value(get_first(f, ["Client", "ClientName", "Client0"]))

        if not client:
            client = infer_client_from_folder(url_job or url_commercial)

        # OpsFolderContribute and ComFolderContribute — person names for manager lookup
        ops_contribute_raw = get_first(f, ["OpsFolderContribute", "OpsContribute"])
        com_contribute_raw = get_first(f, ["ComFolderContribute0", "ComFolderContribute", "ComContribute"])

        ops_contribute = lookup_value(ops_contribute_raw) if ops_contribute_raw else ""
        com_contribute = lookup_value(com_contribute_raw) if com_contribute_raw else ""

        jobs.append({
            "id": job_id,
            "job_id": job_id,
            "title": title,
            "display_name": f"{job_id} - {title}".strip(" -"),
            "status": status,
            "active": active,
            "open": open_job,
            "client": client,
            "project": project,
            "project_lookup_id": project_lookup_id,
            "site": site,
            "ops_folder_contribute": ops_contribute,
            "com_folder_contribute": com_contribute,
            "client_contact": safe_str(get_first(f, ["ClientContact", "Client_x0020_Contact"])),
            "client_po": safe_str(get_first(f, ["ClientPurchaseOrder", "ClientPO", "Client_x0020_PO"])),
            "start_date": safe_date(get_first(f, ["StartDate", "Start_x0020_Date"])),
            "end_date": safe_date(get_first(f, ["FinDate", "EndDate", "FinishDate"])),
            "quote_value": safe_num(get_first(f, ["QuoteValue", "Quote_x0020_Value"])),
            "po_value": safe_num(get_first(f, ["POValue", "PO_x0020_Value"])),
            "opms_id": safe_str(get_first(f, ["OPMSID", "OPMS_x0020_ID"])),
            "comments": safe_str(get_first(f, ["Comments"])),
            "url_job": url_job,
            "url_commercial": url_commercial,
            "url_planning": url_planning,
            "url_client_shared": url_client_shared,
            "contract_folder": contract_folder,
            "folders": {
                "job": url_job,
                "commercial": url_commercial,
                "client_quotes": url_client_quotes,
                "planning": url_planning,
                "client_shared": url_client_shared,
            },
            "folder_status": {
                "has_job_folder": bool(url_job),
                "has_commercial_folder": bool(url_commercial),
                "client_quotes_assumed": bool(url_client_quotes),
                "needs_manual_folder_selection": not bool(url_commercial),
            },
            "raw_item_id": safe_str(f.get("_item_id")),
        })

    jobs.sort(key=lambda x: (not x["open"], x["job_id"]))

    return jobs


def transform_rates(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rates = []

    for f in items:
        # Project & Position are Lookup columns — Graph returns only the numeric
        # <Field>LookupId, not the text. Keep the IDs so build_quote_data() can
        # resolve them to JMS-Projects / PPL-Positions names (see resolve loop).
        project = lookup_value(get_first(f, ["Project", "ProjectLookup", "ProjectId"]))
        position = lookup_value(get_first(f, ["Position", "PositionLookup", "PositionId"]))
        project_lookup_id = safe_str(f.get("ProjectLookupId", ""))
        position_lookup_id = safe_str(f.get("PositionLookupId", ""))
        title = safe_str(get_first(f, ["Title"]))

        if not project and not position and not title and not project_lookup_id and not position_lookup_id:
            continue

        rates.append({
            "project": project,
            "position": position,
            "project_lookup_id": project_lookup_id,
            "position_lookup_id": position_lookup_id,
            "title": title,
            "day_shift_rate": safe_num(get_first(f, ["DayShift", "Day_x0020_Shift", "DSRate", "DS_x0020_Rate"])),
            "night_shift_rate": safe_num(get_first(f, ["NightShift", "Night_x0020_Shift", "NSRate", "NS_x0020_Rate"])),
            "raw_item_id": safe_str(f.get("_item_id")),
        })

    return rates


def transform_positions(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    positions = []
    seen = set()

    for f in items:
        title = safe_str(get_first(f, ["Title", "Position", "Role"]))

        if not title:
            continue

        key = title.lower()

        if key in seen:
            continue

        seen.add(key)

        active_raw = get_first(f, ["Active"], True)
        active = str(active_raw).upper() not in {"FALSE", "0", "NO"}

        positions.append({
            "title": title,
            "name": title,
            "position": title,
            "type": "P&E" if title.startswith("Z.") else "Labour",
            "discipline": lookup_value(get_first(f, ["Discipline", "Trade", "Category"])),
            "active": active,
            "raw_item_id": safe_str(f.get("_item_id")),
        })

    positions.sort(key=lambda x: x["title"].lower())

    return positions


def transform_projects(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    projects = []

    for f in items:
        atitle = safe_str(get_first(f, ["ATitle"]))
        title = safe_str(get_first(f, ["Title"]))
        # Use ATitle if available, otherwise fall back to Title
        project_key = atitle or title
        # Client is a Lookup field — store LookupId for ID-based matching
        client_lookup_id = safe_str(f.get("ClientLookupId", ""))
        client = lookup_value(get_first(f, ["Client", "ClientId", "ClientLookupId"]))
        lead = lookup_value(get_first(f, ["AcmeProjectLead", "AcmeProjectLeadId"]))

        com_raw = get_first(f, ["ComFolderContribute", "ComFolder", "CommercialFolder"])
        ops_raw = get_first(f, ["OpsFolder", "OpsFolderEdit"])

        url_commercial = extract_url(com_raw)
        url_ops = extract_url(ops_raw)

        # OpsFolderContribute and ComFolderContribute at project level
        ops_contribute_raw = get_first(f, ["OpsFolderContribute", "OpsContribute"])
        com_contribute_raw = get_first(f, ["ComFolderContribute0", "ComFolderContribute", "ComContribute"])

        ops_contribute = lookup_value(ops_contribute_raw) if ops_contribute_raw else ""
        com_contribute = lookup_value(com_contribute_raw) if com_contribute_raw else ""

        projects.append({
            "project": project_key,
            "atitle": atitle,
            "title": title,
            "client": client,
            "client_lookup_id": client_lookup_id,
            "raw_item_id": safe_str(f.get("_item_id")),
            "project_lead": lead,
            "ops_folder_contribute": ops_contribute,
            "com_folder_contribute": com_contribute,
            "commercial_folder": url_commercial,
            "ops_folder": url_ops,
            "client_quotes_folder": append_child_folder_url(url_commercial, CLIENT_QUOTES_FOLDER_NAME),
            "raw_item_id": safe_str(f.get("_item_id")),
        })

    return projects


def transform_people(items: List[Dict[str, Any]], positions_map: Dict[str, str] = None) -> List[Dict[str, Any]]:
    people = []

    for f in items:
        title = safe_str(get_first(f, ["Title"]))
        first_name = safe_str(get_first(f, ["FirstName", "First_x0020_Name"]))
        last_name = safe_str(get_first(f, ["LastName", "Last_x0020_Name"]))

        full_name = " ".join([x for x in [first_name, last_name] if x]).strip()

        if not full_name:
            full_name = title

        email = safe_str(get_first(f, ["AccessControl", "Email", "EMail", "WorkEmail"]))
        mobile = safe_str(get_first(f, ["Mobile", "Phone", "ContactNumber"]))

        # Position is a Lookup field — resolve via ID → PPL-Positions.Title
        position_lookup_id = safe_str(f.get("PositionLookupId", ""))
        if positions_map and position_lookup_id:
            position = positions_map.get(position_lookup_id, "")
        else:
            position = lookup_value(get_first(f, ["Position", "PositionId"]))

        if not full_name and not email:
            continue

        people.append({
            "name": full_name,
            "title": title,
            "first_name": first_name,
            "last_name": last_name,
            "email": email,
            "mobile": mobile,
            "position": position,
            "raw_item_id": safe_str(f.get("_item_id")),
        })

    people.sort(key=lambda x: x["name"].lower())

    return people


def transform_clients(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    clients = []

    for f in items:
        name = safe_str(get_first(f, ["Title", "ClientName", "Name"]))

        if not name:
            continue

        clients.append({
            "name": name,
            "address": safe_str(get_first(f, [
                "Address",
                "SiteAddress",
                "ClientAddress",
                "Site_x0020_Address",
                "Client_x0020_Address",
            ])),
            "raw_item_id": safe_str(f.get("_item_id")),
        })

    return clients


# ============================================================
# Build Quote Data
# ============================================================

def default_template() -> Dict[str, Any]:
    # Fallback template used only when none is found in storage. Personal
    # details are intentionally blank — the real default is configured in Blob.
    return {
        "mgr_name": "",
        "mgr_role": "Operations Manager",
        "mgr_phone": "",
        "mgr_email": "",
        "multiplier": 1.35,
        "shifts": [
            {"desc": "FIP0", "hours": "0"},
            {"desc": "FOA3", "hours": "3"},
            {"desc": "DS12.5", "hours": "12.5"},
            {"desc": "DS1", "hours": "1"},
            {"desc": "DAYRATE", "hours": "1"},
            {"desc": "NS12.5", "hours": "12.5"},
            {"desc": "NS1", "hours": "1"},
            {"desc": "FOA0", "hours": "0"},
            {"desc": "DIA12", "hours": "12"},
            {"desc": "DOA12", "hours": "12"},
        ],
        "new_roles": [
            {
                "role": "Dog walker",
                "required": 2,
                "comments": "",
                "ds_rate": 10,
                "ns_rate": 20,
                "emp_rate": 0,
            },
            {
                "role": "Smooth talker",
                "required": 1,
                "comments": "",
                "ds_rate": 100,
                "ns_rate": 200,
                "emp_rate": 0,
            },
        ],
        "staffing": [
            {
                "role": "Rigger - Intermediate",
                "required": 1,
                "comments": "",
                "ds_rate": 100,
                "ns_rate": 180,
                "emp_rate": 0,
            },
            {
                "role": "Tablet Device",
                "required": 1,
                "comments": "",
                "ds_rate": 0,
                "ns_rate": 0,
                "emp_rate": 0,
            },
        ],
    }


def split_names(name_str: str) -> List[str]:
    """Split a comma-separated name string into a list of stripped names, excluding archived entries."""
    if not name_str:
        return []
    parts = [p.strip() for p in name_str.split(",") if p.strip()]
    return [p for p in parts if "ARCHIVED" not in p.upper()]


def resolve_managers(job: Dict[str, Any], project_map: Dict[str, Any], people_map: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Given a job, resolve the Requesting Manager candidates:
    1. Collect names from job.ops_folder_contribute + job.com_folder_contribute
    2. VLOOKUP job.project → JMS-Projects → collect project.ops_folder_contribute + project.com_folder_contribute
    3. Merge and deduplicate all names
    4. For each name, look up PPL-People to get Position, Mobile, WorkEmail
    Returns a list of manager dicts ready for the frontend dropdown.
    """
    name_set: List[str] = []

    # Step 1: from the job itself
    for name in split_names(job.get("ops_folder_contribute", "")):
        if name not in name_set:
            name_set.append(name)
    for name in split_names(job.get("com_folder_contribute", "")):
        if name not in name_set:
            name_set.append(name)

    # Step 2: from the matching project
    project_key = safe_str(job.get("project", "")).lower()
    if project_key:
        proj = project_map.get(project_key)
        if proj:
            for name in split_names(proj.get("ops_folder_contribute", "")):
                if name not in name_set:
                    name_set.append(name)
            for name in split_names(proj.get("com_folder_contribute", "")):
                if name not in name_set:
                    name_set.append(name)

    # Step 3: look up each name in PPL-People
    managers = []
    for name in name_set:
        person = people_map.get(name.lower())
        managers.append({
            "name": name,
            "position": person.get("position", "") if person else "",
            "mobile": person.get("mobile", "") if person else "",
            "email": person.get("email", "") if person else "",
        })

    return managers


def build_quote_data(active_only: bool = False) -> Dict[str, Any]:
    token = get_access_token()
    site_id = get_site_id(token)

    print(f"Connected to SharePoint site: {SITE_NAME}")
    print(f"Site ID: {site_id}")
    print("\nPulling SharePoint lists...")

    raw_jobs = get_list_items(site_id, LIST_NAMES["jobs"], token)
    print(f"  {LIST_NAMES['jobs']}: {len(raw_jobs)} items")

    raw_rates = get_list_items(site_id, LIST_NAMES["rates"], token)
    print(f"  {LIST_NAMES['rates']}: {len(raw_rates)} items")

    raw_positions = get_list_items(site_id, LIST_NAMES["positions"], token)
    print(f"  {LIST_NAMES['positions']}: {len(raw_positions)} items")

    raw_projects = get_list_items(site_id, LIST_NAMES["projects"], token)
    print(f"  {LIST_NAMES['projects']}: {len(raw_projects)} items")

    raw_people = get_list_items(site_id, LIST_NAMES["people"], token)
    print(f"  {LIST_NAMES['people']}: {len(raw_people)} items")

    jobs = transform_jobs(raw_jobs, active_only=active_only)
    rates = transform_rates(raw_rates)
    positions = transform_positions(raw_positions)
    projects = transform_projects(raw_projects)

    # Build positions_map: PPL-Positions ID → Title (e.g. "2" → "Operations Manager")
    positions_map: Dict[str, str] = {}
    for item in raw_positions:
        f = item if "_item_id" in item else item
        pid = safe_str(f.get("_item_id") or f.get("id", ""))
        ptitle = safe_str(f.get("Title", ""))
        if pid and ptitle:
            positions_map[pid] = ptitle

    people = transform_people(raw_people, positions_map=positions_map)

    try:
        raw_clients = get_list_items(site_id, LIST_NAMES["clients"], token)
        print(f"  {LIST_NAMES['clients']}: {len(raw_clients)} items")
        clients = transform_clients(raw_clients)
    except Exception as e:
        print(f"  Warning: could not pull {LIST_NAMES['clients']}: {e}")
        clients = []

    # project_map: keyed by raw_item_id (matches job.project_lookup_id)
    # also keyed by atitle/title text as fallback
    project_map = {}
    for p in projects:
        if p.get("raw_item_id"):
            project_map[p["raw_item_id"]] = p          # ID match: "82" → project
        if p.get("project"):
            project_map[p["project"].lower()] = p      # ATitle text fallback
        if p.get("title") and p["title"].lower() not in project_map:
            project_map[p["title"].lower()] = p        # Title text fallback

    # client_map: keyed by raw_item_id (matches project.client_lookup_id)
    # also keyed by name text as fallback
    client_map = {}
    for c in clients:
        if c.get("raw_item_id"):
            client_map[c["raw_item_id"]] = c           # ID match: "73" → client
        if c.get("name"):
            client_map[c["name"].lower()] = c          # Title text fallback

    # people_map keyed by full name (lower) for manager resolution
    people_map = {
        p["name"].lower(): p
        for p in people
        if p.get("name")
    }

    # Resolve JMS-Rates Lookup IDs to text:
    #   ProjectLookupId  -> JMS-Projects (use ATitle, e.g. "C0071-Plus Pumps")
    #   PositionLookupId -> PPL-Positions Title (matches the Staffing dropdown)
    # so the frontend can VLOOKUP rates by Client Business Name + Position.
    for rt in rates:
        if not rt.get("position") and rt.get("position_lookup_id"):
            rt["position"] = positions_map.get(rt["position_lookup_id"], "")
        if not rt.get("project") and rt.get("project_lookup_id"):
            proj = project_map.get(rt["project_lookup_id"])
            if proj:
                rt["project"] = proj.get("atitle") or proj.get("project", "")

    for job in jobs:
        # Resolve project via ID first, then text fallback
        proj = None
        if job.get("project_lookup_id"):
            proj = project_map.get(job["project_lookup_id"])
        if not proj and job.get("project"):
            proj = project_map.get(job["project"].lower())

        # CLIENT BUSINESS NAME = JMS-Projects.ATitle
        if proj:
            job["client_business_name"] = proj.get("atitle") or proj.get("project", "")
        else:
            job["client_business_name"] = job.get("project", "")

        # Resolve client via ID first, then text fallback
        client_info = None
        if proj and proj.get("client_lookup_id"):
            client_info = client_map.get(proj["client_lookup_id"])
        if not client_info and proj and proj.get("client"):
            client_info = client_map.get(proj["client"].lower())
        if not client_info and job.get("client"):
            client_info = client_map.get(job["client"].lower())

        # SITE ADDRESS from JMS-Clients
        job["client_address"] = client_info.get("address", "") if client_info else ""

        # Also update client name on job if missing
        if not job.get("client") and proj:
            job["client"] = proj.get("client", "")

        # Resolve Requesting Manager candidates
        job["requesting_managers"] = resolve_managers(job, project_map, people_map)

    generated = datetime.now().isoformat(timespec="seconds")

    output = {
        "generated_at": generated,
        "generated": generated,
        "source": "SharePoint via App.py",
        "count": len(jobs),
        "open_count": sum(1 for j in jobs if j.get("open")),

        "site": {
            "host": SHAREPOINT_HOST,
            "site_name": SITE_NAME,
            "site_id": site_id,
        },

        "settings": {
            "client_quotes_folder_name": CLIENT_QUOTES_FOLDER_NAME,
            "active_only": active_only,
        },

        "lists": LIST_NAMES,
        "shift_codes": SHIFT_CODES,
        "template": default_template(),

        "jobs": jobs,
        "rates": rates,
        "positions": positions,
        "projects": projects,
        "people": people,
        "clients": clients,

        "summary": {
            "jobs": len(jobs),
            "open_jobs": sum(1 for j in jobs if j.get("open")),
            "rates": len(rates),
            "positions": len(positions),
            "projects": len(projects),
            "people": len(people),
            "clients": len(clients),
        },
    }

    return output


# ============================================================
# Flask App
# ============================================================

flask_app = Flask(__name__)
# WSGI entry point — gunicorn/Azure look for "app" by default (gunicorn app:app)
app = flask_app

# Always re-read templates from disk. Flask's mtime-based auto-reload is
# unreliable on OneDrive-synced folders (the file's mtime doesn't always update
# right away), so disabling the compiled-template cache guarantees edits to
# html3.html show up on the next request without a server restart.
flask_app.config["TEMPLATES_AUTO_RELOAD"] = True
flask_app.jinja_env.auto_reload = True
flask_app.jinja_env.cache_size = 0


@flask_app.after_request
def _gzip_response(resp):
    """Gzip large text/JSON responses (the index HTML carries the job cache).
    Cuts ~150KB+ HTML to ~20-30KB on the wire — much faster page load on Azure."""
    try:
        import gzip
        accept = request.headers.get("Accept-Encoding", "")
        ctype = (resp.content_type or "")
        if ("gzip" in accept.lower()
                and resp.status_code == 200
                and not resp.direct_passthrough
                and "Content-Encoding" not in resp.headers
                and ("text/html" in ctype or "application/json" in ctype
                     or "text/css" in ctype or "javascript" in ctype)):
            raw = resp.get_data()
            if len(raw) > 1024:
                packed = gzip.compress(raw, 6)
                resp.set_data(packed)
                resp.headers["Content-Encoding"] = "gzip"
                resp.headers["Content-Length"] = str(len(packed))
                resp.headers.add("Vary", "Accept-Encoding")
    except Exception:
        pass
    return resp

# Data directory holds quote_data.json / quotes.json.
# On Azure App Service set QUOTE_DATA_DIR=/home/data so the data survives
# restarts and re-deploys (the wwwroot folder is replaced on each deploy).
# Locally it defaults to the repo's ../static folder (matches blob_index.py).
_DEFAULT_DATA_DIR = Path(__file__).resolve().parent.parent / "static"
STATIC_DIR = Path(os.getenv("QUOTE_DATA_DIR") or _DEFAULT_DATA_DIR)
DRAFTS_DIR = Path(os.getenv("QUOTE_DRAFTS_DIR") or (Path(__file__).resolve().parent.parent / "drafts"))
OUTPUT_DIR = Path(os.getenv("QUOTE_OUTPUT_DIR") or (Path(__file__).resolve().parent.parent / "output"))

for _d in (STATIC_DIR, DRAFTS_DIR, OUTPUT_DIR):
    try:
        _d.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass  # read-only filesystem (e.g. some hosting); ignore


@flask_app.route("/")
def index():
    # Inject the cached job data at render time (kept OUT of the template /
    # git repo). Reads the local cache only — never calls SharePoint on page
    # load. Build/refresh the cache via /api/refresh-data or `python app.py`.
    embedded = storage.get_text(QUOTE_DATA_NAME, "{}", ttl=300)
    try:
        json.loads(embedded)  # validate it's JSON
    except Exception:
        embedded = "{}"
    import brands
    return render_template(
        "html3.html",
        embedded_data=embedded,
        brands_cfg=json.dumps(brands.public_config()),
    )


@flask_app.route("/brand-logo/<key>", methods=["GET"])
def brand_logo(key):
    """Serve a brand logo image from blob storage (cached in-process)."""
    import brands
    from flask import Response, abort
    b = brands.get_brand(key)
    data = storage.get_bytes(b["logo_blob"], ttl=600)
    if not data:
        abort(404)
    ctype = "image/png" if b["logo_blob"].lower().endswith(".png") else "image/jpeg"
    return Response(data, mimetype=ctype,
                    headers={"Cache-Control": "public, max-age=3600"})


@flask_app.route("/api/quote-data", methods=["GET"])
def api_quote_data():
    try:
        active_only = request.args.get("active_only", "false").lower() == "true"

        data = build_quote_data(active_only=active_only)

        # Persist to Blob (or local file) and refresh the index page cache
        storage.put_json(QUOTE_DATA_NAME, data)
        storage.invalidate(QUOTE_DATA_NAME)

        return jsonify({
            "success": True,
            "message": "Quote data loaded from SharePoint",
            "data": data,
            "summary": data.get("summary", {}),
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


@flask_app.route("/api/refresh-data", methods=["GET", "POST"])
def api_refresh_data():
    return api_quote_data()


@flask_app.route("/api/save-draft", methods=["POST"])
def api_save_draft():
    try:
        draft = request.get_json(force=True)

        job_id = (
            draft.get("data_input", {}).get("job_id")
            or draft.get("job", {}).get("job_id")
            or draft.get("job", {}).get("id")
            or "unknown_job"
        )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{job_id}_quote_draft_{timestamp}.json"

        draft_path = DRAFTS_DIR / filename
        draft_path.write_text(
            json.dumps(draft, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        return jsonify({
            "success": True,
            "message": "Draft saved",
            "draft_file": str(draft_path),
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


@flask_app.route("/api/load-draft/<filename>", methods=["GET"])
def api_load_draft(filename):
    draft_path = DRAFTS_DIR / filename

    if not draft_path.exists():
        return jsonify({
            "success": False,
            "message": "Draft not found",
        }), 404

    draft = json.loads(draft_path.read_text(encoding="utf-8"))

    return jsonify({
        "success": True,
        "draft": draft,
    })


@flask_app.route("/api/generate-quote", methods=["POST"])
def api_generate_quote():
    try:
        draft = request.get_json(force=True)

        job_id = (
            draft.get("data_input", {}).get("job_id")
            or draft.get("job", {}).get("job_id")
            or draft.get("job", {}).get("id")
            or "unknown_job"
        )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{job_id}_quote_payload_{timestamp}.json"

        output_path = OUTPUT_DIR / filename
        output_path.write_text(
            json.dumps(draft, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        return jsonify({
            "success": True,
            "message": "Local quote payload saved. Excel generation is next step.",
            "output_file": str(output_path),
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


@flask_app.route("/api/job-managers/<job_id>", methods=["GET"])
def api_job_managers(job_id: str):
    """
    Returns Requesting Manager candidates for a given job_id.
    Reads from the cached quote_data.json if available, otherwise pulls live.

    Response:
    {
      "success": true,
      "job_id": "LH-00001",
      "client_business_name": "Example Client",
      "client_address": "123 Example St, Somewhere WA 6000",
      "managers": [
        {
          "name": "Jane Doe",
          "position": "Operations Manager",
          "mobile": "61400000000",
          "email": "manager@example.com"
        },
        ...
      ]
    }
    """
    try:
        # Try cached data first for speed
        data = storage.get_json(QUOTE_DATA_NAME, None, ttl=300)
        if data is None:
            data = build_quote_data()

        jobs = data.get("jobs", [])
        job = next((j for j in jobs if j.get("job_id", "").upper() == job_id.upper()), None)

        if not job:
            return jsonify({
                "success": False,
                "message": f"Job '{job_id}' not found",
            }), 404

        return jsonify({
            "success": True,
            "job_id": job.get("job_id"),
            "client_business_name": job.get("client_business_name", job.get("project", "")),
            "client_address": job.get("client_address", ""),
            "managers": job.get("requesting_managers", []),
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


# ============================================================
# Quote File Routes (IMS folders + Blob + Excel)
# ============================================================

@flask_app.route("/api/job-folders/<job_id>", methods=["GET"])
def api_job_folders(job_id: str):
    """
    Returns OpsFolder / ComFolder / PlanningFolder URLs for a job.
    Primary source: cached JMS-Jobs data.
    Fallback: IMS/Operations traversal for any URL that is missing.
    Response: { success, url_job, url_commercial, url_planning, source }
    """
    try:
        url_job = url_commercial = url_planning = client_business_name = ""

        cached = storage.get_json(QUOTE_DATA_NAME, {}, ttl=300)
        if cached:
            job = next(
                (j for j in cached.get("jobs", [])
                 if j.get("job_id", "").upper() == job_id.upper()),
                None,
            )
            if job:
                url_job              = job.get("url_job", "")
                url_commercial       = job.get("url_commercial", "")
                url_planning         = job.get("url_planning", "")
                client_business_name = job.get("contract_folder", "") or job.get("client_business_name", "")

        # All three present — return immediately, no IMS call needed
        if url_job and url_commercial and url_planning:
            return jsonify({
                "success": True, "source": "cache",
                "url_job": url_job, "url_commercial": url_commercial,
                "url_planning": url_planning,
            })

        # At least one is missing — try IMS traversal
        from graph_client import get_token
        from operations_folders import get_job_folder_links

        token = get_token()
        ims   = get_job_folder_links(job_id, token, client_business_name)

        if ims.get("found"):
            return jsonify({
                "success": True, "source": "ims",
                "url_job":        url_job        or ims.get("url_job", ""),
                "url_commercial": url_commercial or ims.get("url_commercial", ""),
                "url_planning":   url_planning   or ims.get("url_planning", ""),
            })

        # IMS also couldn't find it — return whatever we have
        return jsonify({
            "success": True, "source": "partial",
            "url_job": url_job, "url_commercial": url_commercial,
            "url_planning": url_planning,
        })

    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@flask_app.route("/api/job-quotes/<job_id>", methods=["GET"])
def api_job_quotes(job_id: str):
    """
    Get existing quotes for a job.
    1. Check Blob index by DriveItem ID
    2. Verify each still exists in SharePoint
    3. Fallback: scan 01 Quotes folder directly
    Returns: { success, job_id, quotes: [...], has_quotes: bool }
    """
    try:
        from graph_client import get_token, get_operations_drive_id
        from blob_index import resolve_quotes_for_job, get_quotes_for_job
        from operations_folders import list_existing_quotes, IMS_OPS_URL

        token    = get_token()
        drive_id = get_operations_drive_id(token)

        # Load all 3 folder URLs + contract_folder from cache
        url_commercial = url_job = url_planning = client_business_name = ""
        cached = storage.get_json(QUOTE_DATA_NAME, {}, ttl=300)
        if cached:
            try:
                job = next(
                    (j for j in cached.get("jobs", [])
                     if j.get("job_id", "").upper() == job_id.upper()),
                    None,
                )
                if job:
                    url_commercial       = job.get("url_commercial", "")
                    url_job              = job.get("url_job", "")
                    url_planning         = job.get("url_planning", "")
                    client_business_name = job.get("contract_folder", "") or job.get("client_business_name", "")
            except Exception:
                pass

        # Try blob index first (fastest — DriveItem ID direct lookup)
        blob_records = get_quotes_for_job(job_id)
        if blob_records:
            resolved = resolve_quotes_for_job(job_id, token, drive_id)
            live = [r for r in resolved if r.get("exists")]
            # Deduplicate: keep one record per drive_item_id, then per file_name
            seen_ids = set(); seen_names = set(); deduped = []
            for r in sorted(live, key=lambda x: x.get("uploaded_at",""), reverse=True):
                did = r.get("drive_item_id",""); fn = r.get("file_name","")
                if did and did in seen_ids: continue
                if fn and fn in seen_names: continue
                if did: seen_ids.add(did)
                if fn:  seen_names.add(fn)
                deduped.append(r)
            if deduped:
                return jsonify({"success": True, "job_id": job_id,
                                "quotes": deduped, "has_quotes": True,
                                "source": "blob", "ims_ops_url": IMS_OPS_URL})

        # Scan SharePoint using 3-URL priority order, then deep IMS traversal
        from operations_folders import find_quotes_folder_any
        quotes_folder_url  = ""
        quotes_folder_name = ""
        sp_files = []
        try:
            qf = find_quotes_folder_any(token, url_commercial, url_job, url_planning,
                                        job_id=job_id,
                                        client_business_name=client_business_name)
            if qf:
                quotes_folder_url  = qf.get("webUrl", "")
                quotes_folder_name = qf.get("name", "01 Client Quotes")
                # List files from the found folder (reuse the drive_id above)
                ch = graph_get(
                    f"{GRAPH_BASE}/drives/{drive_id}/items/{qf['id']}/children?$top=500",
                    token
                )

                def _is_quote_file(f, jid):
                    name = f.get("name", "").strip()

                    if not f.get("file"):
                        return False

                    if not name.lower().endswith((".xlsx", ".xlsm")):
                        return False

                    # Online Quote Tool generated files must start with QT-
                    if not name.upper().startswith("QT-"):
                        return False

                    # Must belong to selected Job
                    return jid.upper() in name.upper()
                sp_files = sorted([
                    {"drive_item_id": f["id"], "file_name": f["name"],
                     "web_url": f.get("webUrl",""),
                     "modified": f.get("lastModifiedDateTime",""),
                     "size": f.get("size",0)}
                    for f in ch.get("value",[])
                    if _is_quote_file(f, job_id)
                ], key=lambda x: x["modified"], reverse=True)
        except Exception:
            pass

        # When nothing was found, give the popup a targeted landing folder
        # (the contract folder, e.g. "C0001 Example Client") instead of the bare
        # IMS/Operations root, so the user can filter from the right place.
        contract_url = contract_name = ""
        if not sp_files:
            try:
                from operations_folders import get_contract_folder_url
                cf = get_contract_folder_url(token, client_business_name)
                contract_url  = cf.get("url", "")
                contract_name = cf.get("name", "")
            except Exception:
                pass

        return jsonify({
            "success":             True,
            "job_id":              job_id,
            "quotes":              sp_files,
            "has_quotes":          bool(sp_files),
            "source":              "sharepoint",
            "ims_ops_url":         IMS_OPS_URL,
            "quotes_folder_url":   quotes_folder_url,
            "quotes_folder_name":  quotes_folder_name,
            "contract_folder_url":  contract_url,
            "contract_folder_name": contract_name,
        })

    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@flask_app.route("/api/load-quote/<drive_item_id>", methods=["GET"])
def api_load_quote(drive_item_id: str):
    """
    Load a previously generated quote's form payload by DriveItem ID.
    Returns the stored payload so the frontend can pre-fill the form.
    """
    try:
        from blob_index import get_record_by_drive_item_id
        record = get_record_by_drive_item_id(drive_item_id)
        if not record or not record.get("payload"):
            return jsonify({"success": False, "message": "No saved form data for this quote"}), 404
        return jsonify({"success": True, "payload": record["payload"],
                        "file_name": record.get("file_name", ""),
                        "uploaded_at": record.get("uploaded_at", "")})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@flask_app.route("/api/generate-quote-excel", methods=["POST"])
def api_generate_quote_excel():
    """
    Generate Excel from form payload, upload to SharePoint 01 Quotes,
    save DriveItem ID to Blob index.
    Body: { job_id, client_business_name, job_title, data_input, shift_types,
            new_roles, staff_roles, manning, manning_dates, multiplier,
            other_costs, summary }
    """
    try:
        from graph_client import get_token
        from excel_generator import generate_quote_excel, make_file_name
        from operations_folders import upload_quote_file
        from blob_index import add_quote_record

        payload = request.get_json(force=True)
        job_id  = payload.get("job_id") or payload.get("data_input", {}).get("job_id", "unknown")
        client_business_name = payload.get("client_business_name", "")
        job_title = payload.get("job_title", payload.get("data_input", {}).get("proj_title", ""))

        # Generate Excel bytes
        excel_bytes = generate_quote_excel(payload)
        # Use custom filename if provided by frontend, otherwise auto-generate
        file_name = safe_str(payload.get("file_name", "")).strip() or make_file_name(job_id)

        if not file_name.upper().startswith("QT-"):
            file_name = f"QT-{file_name}"

        if not file_name.lower().endswith(".xlsx"):
            file_name += ".xlsx"

        # Upload to SharePoint.
        #  - save_direct: user explicitly picked a folder → drop the file straight
        #    in there (used when 01 Client Quotes can't be found).
        #  - otherwise: auto-locate 01 Client Quotes via the 3-URL priority search
        #    (Commercial → Ops → Planning) then deep IMS traversal.
        token           = get_token()
        save_folder_url = payload.get("save_folder_url", "")
        save_direct     = bool(payload.get("save_direct")) and bool(save_folder_url)
        if save_direct:
            result = upload_quote_file(
                job_id, excel_bytes, file_name, token,
                direct_folder_url=save_folder_url,
                client_business_name=client_business_name,
            )
        else:
            result = upload_quote_file(
                job_id, excel_bytes, file_name, token,
                url_commercial=payload.get("url_commercial", ""),
                url_job=payload.get("url_job", ""),
                url_planning=payload.get("url_planning", ""),
                client_business_name=client_business_name,
            )

        # Save to Blob index (including full payload for future reload)
        record = add_quote_record(
            job_id=job_id,
            drive_item_id=result["drive_item_id"],
            file_name=result["file_name"],
            sharepoint_url=result["web_url"],
            client_business_name=client_business_name,
            job_title=job_title,
            payload=payload,
        )

        return jsonify({
            "success": True,
            "message": "Quote generated and uploaded",
            "file_name": result["file_name"],
            "web_url": result["web_url"],
            "drive_item_id": result["drive_item_id"],
            "record": record,
        })

    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@flask_app.route("/api/download-quote-excel", methods=["POST"])
def api_download_quote_excel():
    """
    Generate Excel and return as download (no SharePoint upload).
    Useful for preview before uploading.
    """
    try:
        from excel_generator import generate_quote_excel, make_file_name
        from flask import Response

        payload     = request.get_json(force=True)
        job_id      = payload.get("job_id") or payload.get("data_input", {}).get("job_id", "unknown")
        excel_bytes = generate_quote_excel(payload)
        file_name   = safe_str(payload.get("file_name", "")).strip() or make_file_name(job_id)
        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"

        return Response(
            excel_bytes,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{file_name}"'},
        )

    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@flask_app.route("/api/download-quote-pdf", methods=["POST"])
def api_download_quote_pdf():
    """Generate a formal client-facing Quote PDF and return as download."""
    try:
        from pdf_generator import generate_quote_pdf, make_pdf_file_name
        from flask import Response

        payload   = request.get_json(force=True)
        job_id    = payload.get("job_id") or payload.get("data_input", {}).get("job_id", "unknown")
        pdf_bytes = generate_quote_pdf(payload)

        file_name = safe_str(payload.get("file_name", "")).strip()
        if file_name:
            if file_name.lower().endswith((".xlsx", ".xlsm")):
                file_name = file_name.rsplit(".", 1)[0]
            if not file_name.lower().endswith(".pdf"):
                file_name += ".pdf"
        else:
            file_name = make_pdf_file_name(job_id)

        return Response(
            pdf_bytes,
            mimetype="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{file_name}"'},
        )
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@flask_app.route("/api/download-rr-schedule", methods=["POST"])
def api_download_rr_schedule():
    """Generate the internal RR (Resource & Rates) Schedule PDF — roster + worker
    pay rates only, with all client-facing pricing/multiplier removed."""
    try:
        from pdf_generator import generate_rr_schedule_pdf, make_rr_pdf_file_name
        from flask import Response

        payload   = request.get_json(force=True)
        job_id    = payload.get("job_id") or payload.get("data_input", {}).get("job_id", "unknown")
        pdf_bytes = generate_rr_schedule_pdf(payload)

        file_name = safe_str(payload.get("file_name", "")).strip()
        if file_name:
            if file_name.lower().endswith((".xlsx", ".xlsm")):
                file_name = file_name.rsplit(".", 1)[0]
            # Re-badge a quote file name as an RR Schedule so the two downloads
            # don't clash (QT-..._Quote → RR-..._Schedule).
            base = file_name
            for pfx in ("QT-", "RR-"):
                if base.upper().startswith(pfx):
                    base = base[len(pfx):]
            base = base.replace("_Quote", "").replace("_quote", "")
            file_name = f"RR-{base}_Schedule"
            if not file_name.lower().endswith(".pdf"):
                file_name += ".pdf"
        else:
            file_name = make_rr_pdf_file_name(job_id)

        return Response(
            pdf_bytes,
            mimetype="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{file_name}"'},
        )
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


def run_local_server():
    # Local dev only. In production Azure runs gunicorn against `app` (see startup.txt).
    host  = os.getenv("HOST", "127.0.0.1")
    port  = int(os.getenv("PORT", "5000"))
    debug = os.getenv("FLASK_DEBUG", "1") == "1"
    flask_app.run(host=host, port=port, debug=debug)


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Acme Online Quote Tool SharePoint Data Extractor"
    )

    parser.add_argument(
        "--output",
        default="static/quote_data.json",
        help="Output JSON file path",
    )

    parser.add_argument(
        "--active-only",
        action="store_true",
        help="Only include open jobs",
    )

    parser.add_argument(
        "--serve",
        action="store_true",
        help="Run local Flask server for testing the online form",
    )

    args = parser.parse_args()

    if args.serve:
        run_local_server()
        return

    data = build_quote_data(active_only=args.active_only)

    # Persist to the configured backend (Azure Blob or local file)
    storage.put_json(QUOTE_DATA_NAME, data)
    storage.invalidate(QUOTE_DATA_NAME)

    print("\nDone.")
    print(f"Saved job cache to: {storage.backend()}")

    print("\nSummary:")
    for key, value in data["summary"].items():
        print(f"  {key}: {value}")


if __name__ == "__main__":
    main()