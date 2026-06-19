"""
brands.py
Central brand definitions for the quote tool.

Used by:
  - app.py            → injects a public copy into the page so the UI can build
                        the brand picker, and serves logos via /brand-logo/<key>
  - pdf_generator.py  → header banner colours / title / logo / footer
  - excel_generator.py→ title row colours / text / logo

Colours are stored as 6-digit hex WITHOUT a leading '#'. Helpers add it where a
given library needs it (ReportLab wants '#RRGGBB', openpyxl wants 'RRGGBB').

Logo binaries live in Blob storage under brand/… (see upload_brand.py) and are
fetched at render time via storage.get_bytes(). This keeps the app filesystem
ephemeral-safe on Azure App Service — same model as the JSON data in storage.py.
"""
from typing import Any, Dict

DEFAULT_BRAND = "acme"

BRANDS: Dict[str, Dict[str, Any]] = {
    "acme": {
        "key":         "acme",
        "ui_name":     "Acme Group",            # sidebar label
        "doc_name":    "ACME GROUP",             # PDF banner / Excel title
        "tagline":     "Shutdown — Quotation",
        "excel_title": "ACME GROUP — Shutdown Price Calculator & Quote",
        "footer_note": "Acme Group standard terms & conditions.",
        "footer_strip":"Acme Group  -  Shutdown Quotation",
        "banner_bg":   "FFFFFF",   # white banner
        "banner_fg":   "1A1A1A",   # dark text
        "accent":      "2E8B8B",   # UI accent
        "xl_title_fg": "1A6060",   # Excel title text (dark teal)
        "xl_title_bg": "FFFFFF",   # Excel title fill (white)
        "logo_blob":   "brand/acme-logo.jpg",
        # white banner → logo sits directly, no chip needed
        "logo_on_chip": False,
    },
    "northwind": {
        "key":         "northwind",
        "ui_name":     "Northwind Workforce",
        "doc_name":    "NORTHWIND WORKFORCE",
        "tagline":     "Shutdown — Quotation",
        "excel_title": "NORTHWIND WORKFORCE — Shutdown Price Calculator & Quote",
        "footer_note": "Northwind Workforce standard terms & conditions.",
        "footer_strip":"Northwind Workforce  -  Shutdown Quotation",
        "banner_bg":   "FFFFFF",   # white banner
        "banner_fg":   "1A1A1A",   # dark text
        "accent":      "6FBF1B",   # green
        "xl_title_fg": "2E6B0F",   # Excel title text (dark green)
        "xl_title_bg": "FFFFFF",   # Excel title fill (white)
        "logo_blob":   "brand/northwind-logo.png",
        # transparent logo on a white banner → no chip needed
        "logo_on_chip": False,
    },
}


def get_brand(key: Any) -> Dict[str, Any]:
    """Return the brand config for `key`, falling back to the default brand."""
    k = str(key or "").strip().lower()
    return BRANDS.get(k, BRANDS[DEFAULT_BRAND])


def brand_key(payload: Dict[str, Any]) -> str:
    """Pull the selected brand key out of an export payload (default-safe)."""
    k = str((payload or {}).get("brand") or "").strip().lower()
    return k if k in BRANDS else DEFAULT_BRAND


def public_config() -> Dict[str, Any]:
    """The subset of brand info safe to embed in the page for the UI picker."""
    return {
        "default": DEFAULT_BRAND,
        "brands": [
            {
                "key":     b["key"],
                "ui_name": b["ui_name"],
                "accent":  "#" + b["accent"],
                "logo_url": f"/brand-logo/{b['key']}",
            }
            for b in BRANDS.values()
        ],
    }
