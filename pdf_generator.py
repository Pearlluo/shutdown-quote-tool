"""
pdf_generator.py
Generates a clean, formal client-facing Labour Hire Quote PDF.

Uses the same cost engine as excel_generator (_calc_role_cost) so the numbers
always match the Excel output. Designed as a proper document (header, project
info, scope, labour breakdown, pricing summary) — NOT a print of the web form.
"""

import io
from datetime import datetime
from typing import Any, Dict, List

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether, Image,
)

from excel_generator import _calc_role_cost, _build_shift_map, _num
from brands import get_brand, brand_key

# ── Brand palette ──────────────────────────────────────────────
TEAL    = colors.HexColor("#1A6060")
TEAL_MID = colors.HexColor("#2E8B8B")
NAVY    = colors.HexColor("#17375E")
ORANGE  = colors.HexColor("#C0681A")
PURPLE  = colors.HexColor("#7A4EA0")
GREY_BG = colors.HexColor("#F0F2F5")
GREY_LN = colors.HexColor("#C8D0DC")
WK_HDR  = colors.HexColor("#1A5060")
WK_TINT = colors.HexColor("#FFF4E8")
PH_TINT = colors.HexColor("#F3ECFA")
GREEN_LT = colors.HexColor("#E8F8F0")
BLUE_LT  = colors.HexColor("#E8E8FF")
GREY_LT  = colors.HexColor("#F8F9FB")
YELLOW  = colors.HexColor("#FFE699")

GST_RATE = 0.10


def _money(v) -> str:
    try:
        return f"${float(v):,.2f}"
    except (TypeError, ValueError):
        return "$0.00"


def _s(v) -> str:
    """Sanitize text for the built-in Helvetica font (latin-1 only).
    Replaces unsupported glyphs so they don't render as black boxes."""
    s = "" if v is None else str(v)
    repl = {"—": "-", "–": "-", "‘": "'", "’": "'",
            "“": '"', "”": '"', "×": "x", "…": "..."}
    for k, val in repl.items():
        s = s.replace(k, val)
    return s.encode("latin-1", "replace").decode("latin-1").replace("?", " ").strip() or "-"


def _styles():
    ss = getSampleStyleSheet()
    out = {
        "title":   ParagraphStyle("t", parent=ss["Normal"], fontName="Helvetica-Bold",
                                  fontSize=17, textColor=colors.white, leading=20),
        "sub":     ParagraphStyle("s", parent=ss["Normal"], fontName="Helvetica",
                                  fontSize=8.5, textColor=colors.white, leading=11),
        "band":    ParagraphStyle("b", parent=ss["Normal"], fontName="Helvetica-Bold",
                                  fontSize=10.5, textColor=colors.white, leading=13),
        "klabel":  ParagraphStyle("kl", parent=ss["Normal"], fontName="Helvetica-Bold",
                                  fontSize=7.5, textColor=colors.HexColor("#4A5568"), leading=10),
        "kval":    ParagraphStyle("kv", parent=ss["Normal"], fontName="Helvetica",
                                  fontSize=9, textColor=colors.HexColor("#1A1A1A"), leading=11),
        "body":    ParagraphStyle("bd", parent=ss["Normal"], fontName="Helvetica",
                                  fontSize=8.5, leading=11),
        "note":    ParagraphStyle("nt", parent=ss["Normal"], fontName="Helvetica-Oblique",
                                  fontSize=7.5, textColor=colors.HexColor("#666666"), leading=10),
    }
    return out


def _brand_logo(brand: Dict[str, Any], max_h: float, max_w: float):
    """Return (Image flowable, width_pt, height_pt) for the brand logo fetched
    from blob, scaled to fit within max_h × max_w. Returns (None, 0, 0) if the
    logo can't be loaded (PDF then falls back to text-only banner)."""
    try:
        import storage
        from PIL import Image as PILImage
        data = storage.get_bytes(brand["logo_blob"])
        if not data:
            return None, 0, 0
        iw, ih = PILImage.open(io.BytesIO(data)).size
        scale = min(max_h / ih, max_w / iw)
        w, h = iw * scale, ih * scale
        return Image(io.BytesIO(data), width=w, height=h), w, h
    except Exception:
        return None, 0, 0


def _band(text: str, width: float, st, bg=ORANGE):
    """Full-width coloured section banner."""
    t = Table([[Paragraph(text, st["band"])]], colWidths=[width])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), bg),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t


def _project_info(di: Dict, st, width: float):
    rows = [
        ("Job #", di.get("job_id", ""),            "Start Date", di.get("start_date", "")),
        ("Status", di.get("status", ""),           "End Date", di.get("end_date", "")),
        ("Client Job No.", di.get("client_job_no", ""), "Quotation Validity", di.get("validity", "")),
        ("Project / Job Title", di.get("proj_title", ""), "Requesting Manager", di.get("mgr_name", "")),
        ("Client", di.get("client_biz", ""),       "Manager Role", di.get("mgr_role", "")),
        ("Site", di.get("site", ""),               "Manager Contact", di.get("mgr_phone", "")),
        ("Site Address", di.get("site_addr", ""),  "Manager Email", di.get("mgr_email", "")),
        ("Client POC Email", di.get("client_email", ""), "Client Rep", di.get("client_rep", "")),
    ]
    data = []
    for l1, v1, l2, v2 in rows:
        data.append([
            Paragraph(l1, st["klabel"]), Paragraph(_s(v1), st["kval"]),
            Paragraph(l2, st["klabel"]), Paragraph(_s(v2), st["kval"]),
        ])
    c = width / 2
    t = Table(data, colWidths=[c * 0.33, c * 0.67, c * 0.33, c * 0.67])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.4, GREY_LN),
        ("BACKGROUND", (0, 0), (0, -1), GREY_BG),
        ("BACKGROUND", (2, 0), (2, -1), GREY_BG),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t


def _shift_table(shift_types, width):
    """Shift Types: Code | Hrs/Shift | Type."""
    data = [["Code", "Hrs / Shift", "Type"]]
    for s in shift_types:
        data.append([_s(s.get("desc", "")),
                     f"{float(s.get('hrs', 0) or 0):g}",
                     _s(s.get("type", "DS"))])
    t = Table(data, colWidths=[width * 0.45, width * 0.3, width * 0.25], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), TEAL),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.4, GREY_LN),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]
    for i in range(1, len(data)):
        if data[i][2].upper() == "NS":
            style.append(("BACKGROUND", (0, i), (-1, i), BLUE_LT))
        else:
            style.append(("BACKGROUND", (0, i), (-1, i), GREEN_LT))
    t.setStyle(TableStyle(style))
    return t


def _grouped_header(hdr_row0, hdr_row1, base_style, group_spans):
    """Apply two-row grouped header styling. group_spans: list of
    (c1, c2, bg) for the top banner cells; row1 cells coloured to match."""
    style = list(base_style)
    style += [
        ("BACKGROUND", (0, 0), (-1, 1), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 1), colors.white),
        ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 1), 6.5),
        ("ALIGN", (0, 0), (-1, 1), "CENTER"),
        ("VALIGN", (0, 0), (-1, 1), "MIDDLE"),
    ]
    for c1, c2, bg in group_spans:
        style.append(("SPAN", (c1, 0), (c2, 0)))
        style.append(("BACKGROUND", (c1, 0), (c2, 0), bg))
        style.append(("BACKGROUND", (c1, 1), (c2, 1), bg))
    return style


def _rates_table(all_roles, width):
    """Shutdown rate card: a single flat DS and NS $/hr per role."""
    hdr = ["Role", "Qty", "DS $/hr", "NS $/hr"]
    data = [hdr]
    for r in all_roles:
        data.append([
            _s(r.get("role", "")), str(r.get("qty", 1)),
            _money(_num(r, "dsRate", "dsStd", "ds_rate")),
            _money(_num(r, "nsRate", "nsStd", "ns_rate")),
        ])
    w = width
    col = [w*0.46, w*0.10, w*0.22, w*0.22]
    t = Table(data, colWidths=col, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, GREY_LN),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]))
    return t


def _manning_grid(all_roles, manning, shift_map, dates, multiplier, width,
                  include_client=True):
    """Full roster grid: roles x dates, colour-coded; trailing Client $ column
    (omitted when include_client is False, e.g. the RR Schedule)."""
    role_w   = width * 0.16
    client_w = width * 0.085 if include_client else 0
    day_w = (width - role_w - client_w) / max(len(dates), 1)
    fs = 5 if day_w < 24 else 6

    hdr = ["Role"]
    for d in dates:
        try:
            o = datetime.strptime(d, "%Y-%m-%d")
            lbl = f"{o.strftime('%a')}\n{o.day}/{o.month}"
        except Exception:
            lbl = d
        hdr.append(lbl)
    if include_client:
        hdr.append("Client $")
    data = [hdr]

    cell_styles = []  # (col, row, bg)
    for ri, r in enumerate(all_roles, start=1):
        rid = str(r.get("id", ""))
        row = [_s(r.get("role", "")) + f"  x{r.get('qty',1)}"]
        rm = manning.get(rid, {})
        for ci, d in enumerate(dates, start=1):
            v = rm.get(d, "OFF")
            row.append("-" if v in ("OFF", "—", "") else _s(v))
            if v not in ("OFF", "—", ""):
                if str(v).startswith("NS"):
                    cell_styles.append((ci, ri, BLUE_LT))
                else:
                    cell_styles.append((ci, ri, GREEN_LT))
        if include_client:
            c = _calc_role_cost(r, manning, shift_map, dates, multiplier)
            row.append(_money(c["client_charge"]))
        data.append(row)

    col = [role_w] + [day_w]*len(dates) + ([client_w] if include_client else [])
    t = Table(data, colWidths=col, repeatRows=1)
    style = [
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), fs),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), fs),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 0), (0, 0), TEAL),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, GREY_LN),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#F0FAFA")),
        ("TEXTCOLOR", (0, 1), (0, -1), TEAL),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
    ]
    # date header colour (uniform — no weekday/weekend/PH distinction)
    for ci in range(1, len(dates) + 1):
        style.append(("BACKGROUND", (ci, 0), (ci, 0), TEAL_MID))
    if include_client:   # trailing Client $ column header + right-align
        style.append(("BACKGROUND", (-1, 0), (-1, 0), NAVY))
        style.append(("ALIGN", (-1, 1), (-1, -1), "RIGHT"))
    for ci, ri, bg in cell_styles:
        style.append(("BACKGROUND", (ci, ri), (ci, ri), bg))
    t.setStyle(TableStyle(style))
    return t


def _alloc_table(all_roles, manning, shift_map, dates, multiplier, st, width,
                 include_client=True):
    """Per-role breakdown: DS / NS shifts, hours, rate and cost. Trailing Total $ /
    Client $ columns are omitted when include_client is False (RR Schedule)."""
    tail_hdr = ["Total $", "Client $"] if include_client else []
    hdr = ["Role", "Qty",
           "DS Shifts", "DS hrs", "DS $/hr", "DS Cost",
           "NS Shifts", "NS hrs", "NS $/hr", "NS Cost"] + tail_hdr
    data = [hdr]
    tot_base = tot_client = 0.0
    for r in all_roles:
        c = _calc_role_cost(r, manning, shift_map, dates, multiplier)
        tot_base += c["total"]; tot_client += c["client_charge"]
        f = lambda x: f"{x:.1f}"
        tail = [_money(c["total"]), _money(c["client_charge"])] if include_client else []
        data.append([
            _s(r.get("role", "")) + f"  x{r.get('qty',1)}", str(r.get("qty", 1)),
            str(c["ds_shifts"]), f(c["ds_hrs"]), _money(c["ds_rate"]), _money(c["ds_cost"]),
            str(c["ns_shifts"]), f(c["ns_hrs"]), _money(c["ns_rate"]), _money(c["ns_cost"]),
        ] + tail)
    n_cols = len(hdr)
    tot_tail = [_money(tot_base), _money(tot_client)] if include_client else []
    data.append(["TOTALS"] + [""]*9 + tot_tail)

    w = width
    fixed = w*0.20 + w*0.05
    small = (w - fixed) / (n_cols - 2)
    col = [w*0.20, w*0.05] + [small]*(n_cols - 2)
    t = Table(data, colWidths=col, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 6.5),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 6.3),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("ALIGN", (0, 1), (0, -1), "LEFT"),
        ("ALIGN", (3, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, GREY_LN),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND", (0, -1), (-1, -1), TEAL),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.white),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("SPAN", (0, -1), (9, -1)),
    ]
    t.setStyle(TableStyle(style))
    return t


def _pricing_summary(payload, all_roles, manning, shift_map, dates, multiplier, st, width):
    agg = {"ds": 0.0, "ns": 0.0, "base": 0.0, "client": 0.0}
    for r in all_roles:
        c = _calc_role_cost(r, manning, shift_map, dates, multiplier)
        agg["ds"]   += c["ds_cost"]; agg["ns"] += c["ns_cost"]
        agg["base"] += c["total"]; agg["client"] += c["client_charge"]

    oc = payload.get("other_costs", [])
    if isinstance(oc, list):
        oc_items = [(str(o.get("label", "") or "Other Cost"), float(o.get("amount", 0) or 0)) for o in oc]
    elif isinstance(oc, dict):
        oc_items = [("Plant & Equipment", float(oc.get("plant", 0) or 0)),
                    ("Contractors", float(oc.get("contractors", 0) or 0)),
                    ("Materials / Other", float(oc.get("materials", 0) or 0))]
    else:
        oc_items = []
    other_total = sum(a for _, a in oc_items)
    markup = agg["client"] - agg["base"]
    grand = agg["client"] + other_total

    rows = [
        ("Day Shift Cost", agg["ds"], 0),
        ("Night Shift Cost", agg["ns"], 0),
        ("Total Labour (base)", agg["base"], 1),
        ("Role Markup (Labour)", markup, 0),
        ("Labour Cost (Client)", agg["client"], 1),
    ]
    for lbl, amt in oc_items:
        rows.append((lbl, amt, 0))
    rows.append(("Total Other Costs", other_total, 1))
    gst = grand * GST_RATE
    rows.append(("Project Total (ex GST)", grand, 1))
    rows.append(("GST (10%)", gst, 0))
    rows.append(("PROJECT TOTAL (inc GST)", grand + gst, 2))

    data = [[lbl, _money(amt)] for lbl, amt, _ in rows]
    t = Table(data, colWidths=[width * 0.62, width * 0.38])
    style = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.3, GREY_LN),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]
    for i, (_, _, kind) in enumerate(rows):
        if kind == 1:      # subtotal
            style += [("BACKGROUND", (0, i), (-1, i), GREY_BG),
                      ("FONTNAME", (0, i), (-1, i), "Helvetica-Bold")]
        elif kind == 2:    # grand total
            style += [("BACKGROUND", (0, i), (-1, i), TEAL),
                      ("TEXTCOLOR", (0, i), (-1, i), YELLOW),
                      ("FONTNAME", (0, i), (-1, i), "Helvetica-Bold"),
                      ("FONTSIZE", (0, i), (-1, i), 11),
                      ("TOPPADDING", (0, i), (-1, i), 7),
                      ("BOTTOMPADDING", (0, i), (-1, i), 7)]
    t.setStyle(TableStyle(style))
    return t


def generate_quote_pdf(payload: Dict[str, Any], rr_schedule: bool = False) -> bytes:
    """Build the client-facing Quote PDF. When rr_schedule=True, produce the
    internal RR (Resource & Rates) Schedule instead: same layout but with the
    client-facing pricing stripped out — no multiplier in the header, no Client $
    column on the manning grid, no Total/Client $ columns on the allocation
    breakdown, and no pricing summary. Only the workers' pay rates remain."""
    st = _styles()
    di = payload.get("data_input", {})
    new_roles = payload.get("new_roles", [])
    staff_roles = payload.get("staff_roles", [])
    all_roles = new_roles + staff_roles
    manning = payload.get("manning", {})
    dates = payload.get("manning_dates", [])
    multiplier = float(payload.get("multiplier", 1.35) or 1.35)
    shift_map = _build_shift_map(payload.get("shift_types", []))

    buf = io.BytesIO()
    page = landscape(A4)
    PW = page[0] - 24 * mm   # usable width
    doc = SimpleDocTemplate(
        buf, pagesize=page,
        leftMargin=12 * mm, rightMargin=12 * mm,
        topMargin=10 * mm, bottomMargin=12 * mm,
        title=f"Quote {di.get('job_id','')}",
    )

    elems: List = []

    # ── Header band (brand-aware: colours, title, logo) ──
    brand = get_brand(brand_key(payload))
    bg = colors.HexColor("#" + brand["banner_bg"])
    fg = colors.HexColor("#" + brand["banner_fg"])
    st["title"].textColor = fg
    st["sub"].textColor = fg

    generated = datetime.now().strftime("%d/%m/%Y %H:%M")
    logo, logo_w, _ = _brand_logo(brand, max_h=42, max_w=150)
    chip = bool(brand.get("logo_on_chip")) and logo is not None
    col0_w = (logo_w + 28) if logo is not None else 0
    inner_w = PW - col0_w

    inner = Table([
        [Paragraph(brand["doc_name"], st["title"]),
         Paragraph(brand["tagline"].upper(), st["sub"])],
        [Paragraph(f"Job: {di.get('job_id','')}", st["sub"]),
         Paragraph(f"Generated: {generated}" if rr_schedule
                   else f"Generated: {generated}&nbsp;&nbsp;|&nbsp;&nbsp;Multiplier: ×{multiplier:.2f}",
                   st["sub"])],
    ], colWidths=[inner_w * 0.5, inner_w * 0.5])
    inner.setStyle(TableStyle([
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("SPAN", (1, 0), (1, 0)),
    ]))

    if logo is not None:
        head = Table([[logo, inner]], colWidths=[col0_w, inner_w])
    else:
        head = Table([[inner]], colWidths=[PW])
    hs = [
        ("BACKGROUND", (0, 0), (-1, -1), bg),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (0, 0), 12),
        ("RIGHTPADDING", (-1, 0), (-1, 0), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]
    if logo is not None:
        hs += [("RIGHTPADDING", (0, 0), (0, 0), 8),
               ("LEFTPADDING", (1, 0), (1, 0), 6)]
        if chip:   # white panel behind a light-background logo on a dark band
            hs += [("BACKGROUND", (0, 0), (0, 0), colors.white)]
    head.setStyle(TableStyle(hs))
    elems += [head, Spacer(1, 7)]

    # ── Project information ──
    elems += [_band("PROJECT INFORMATION", PW, st), Spacer(1, 3)]
    elems += [_project_info(di, st, PW)]

    scope = di.get("scope", "")
    if scope:
        elems += [Spacer(1, 5), _band("SCOPE OF WORKS", PW, st), Spacer(1, 3),
                  Paragraph(_s(scope).replace("\n", "<br/>"), st["body"])]
    add_labour = di.get("add_labour", "")
    if add_labour:
        elems += [Spacer(1, 5), _band("ADDITIONAL LABOUR / EXCLUSIONS", PW, st), Spacer(1, 3),
                  Paragraph(_s(add_labour).replace("\n", "<br/>"), st["body"])]

    # ── Labour rates & shift types ──
    shift_types = payload.get("shift_types", [])
    if all_roles:
        elems += [Spacer(1, 7), _band("LABOUR RATES & SHIFT TYPES", PW, st), Spacer(1, 3)]
        if shift_types:
            elems += [_shift_table(shift_types, PW * 0.42), Spacer(1, 5)]
        elems += [_rates_table(all_roles, PW)]

    # ── Manning schedule (full roster) ──
    # A long roster won't fit the page width, so split it into multiple grids:
    # ≤35 days = one table; longer = evenly-sized chunks of ≤35 days each.
    if all_roles and dates:
        elems += [Spacer(1, 8), _band("MANNING SCHEDULE", PW, st), Spacer(1, 3)]
        MAX_DAYS = 32
        n_chunks = (len(dates) + MAX_DAYS - 1) // MAX_DAYS
        size     = (len(dates) + n_chunks - 1) // n_chunks
        chunks   = [dates[i:i + size] for i in range(0, len(dates), size)]
        for idx, chunk in enumerate(chunks):
            if idx > 0:
                elems += [Spacer(1, 5),
                          Paragraph(f"Manning Schedule (continued — part {idx + 1} of {len(chunks)})",
                                    st["klabel"])]
            elems += [_manning_grid(all_roles, manning, shift_map, chunk, multiplier, PW,
                                    include_client=not rr_schedule)]

    # ── Labour allocation details ──
    if all_roles and dates:
        elems += [Spacer(1, 8), _band("LABOUR ALLOCATION DETAILS (Cost Breakdown per Role)", PW, st),
                  Spacer(1, 3),
                  _alloc_table(all_roles, manning, shift_map, dates, multiplier, st, PW,
                               include_client=not rr_schedule)]

    # ── Pricing summary ── (client-facing pricing — omitted on the RR Schedule)
    if not rr_schedule:
        sw = PW * 0.55
        elems += [Spacer(1, 9),
                  KeepTogether([_band("PRICING SUMMARY", sw, st), Spacer(1, 3),
                                _pricing_summary(payload, all_roles, manning, shift_map,
                                                 dates, multiplier, st, sw)])]

    _note = _s(brand["footer_note"]).replace("&", "&amp;")
    elems += [Spacer(1, 10),
              Paragraph("This quotation is valid for the period stated above and is subject to "
                        f"{_note}", st["note"])]

    def _footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#888888"))
        canvas.drawString(12 * mm, 6 * mm,
                          f"{_s(brand['footer_strip'])}  -  Job {_s(di.get('job_id',''))}")
        canvas.drawRightString(page[0] - 12 * mm, 6 * mm, f"Page {doc_.page}")
        canvas.restoreState()

    doc.build(elems, onFirstPage=_footer, onLaterPages=_footer)
    buf.seek(0)
    return buf.read()


def generate_rr_schedule_pdf(payload: Dict[str, Any]) -> bytes:
    """RR (Resource & Rates) Schedule — internal roster + pay-rate document with
    all client-facing pricing stripped out (see generate_quote_pdf rr_schedule)."""
    return generate_quote_pdf(payload, rr_schedule=True)


def make_pdf_file_name(job_id: str) -> str:
    date_str = datetime.now().strftime("%Y%m%d")
    return f"QT-{job_id}_Quote_{date_str}.pdf"


def make_rr_pdf_file_name(job_id: str) -> str:
    date_str = datetime.now().strftime("%Y%m%d")
    return f"RR-{job_id}_Schedule_{date_str}.pdf"
