"""
wa_holidays.py
Western Australia (WA) gazetted public holidays.

Source: https://www.wa.gov.au/service/employment/workplace-arrangements/public-holidays-western-australia
Dates for 2025 & 2026 are officially published. 2027 King's Birthday and any
weekend-substitute days are best-estimate and should be confirmed/updated once
formally gazetted.

To keep the front-end (html3.html WA_PUBLIC_HOLIDAYS const) and this module in
sync, update BOTH when adding a new year.
"""

from typing import Set

# ISO date string (YYYY-MM-DD) -> holiday name
WA_PUBLIC_HOLIDAYS = {
    # ── 2025 ───────────────────────────────────────
    "2025-01-01": "New Year's Day",
    "2025-01-27": "Australia Day (observed)",
    "2025-03-03": "Labour Day",
    "2025-04-18": "Good Friday",
    "2025-04-21": "Easter Monday",
    "2025-04-25": "Anzac Day",
    "2025-06-02": "Western Australia Day",
    "2025-09-29": "King's Birthday",
    "2025-12-25": "Christmas Day",
    "2025-12-26": "Boxing Day",

    # ── 2026 ───────────────────────────────────────
    "2026-01-01": "New Year's Day",
    "2026-01-26": "Australia Day",
    "2026-03-02": "Labour Day",
    "2026-04-03": "Good Friday",
    "2026-04-06": "Easter Monday",
    "2026-04-25": "Anzac Day",
    "2026-06-01": "Western Australia Day",
    "2026-09-28": "King's Birthday",
    "2026-12-25": "Christmas Day",
    "2026-12-28": "Boxing Day (observed)",

    # ── 2027 (provisional) ─────────────────────────
    "2027-01-01": "New Year's Day",
    "2027-01-26": "Australia Day",
    "2027-03-01": "Labour Day",
    "2027-03-26": "Good Friday",
    "2027-03-29": "Easter Monday",
    "2027-04-26": "Anzac Day (observed)",
    "2027-06-07": "Western Australia Day",
    "2027-09-27": "King's Birthday",
    "2027-12-27": "Christmas Day (observed)",
    "2027-12-28": "Boxing Day (observed)",
}


def is_wa_public_holiday(date_str: str) -> bool:
    """Return True if the ISO date string (YYYY-MM-DD) is a WA public holiday."""
    return date_str in WA_PUBLIC_HOLIDAYS


def wa_holiday_name(date_str: str) -> str:
    """Return the holiday name, or '' if the date is not a WA public holiday."""
    return WA_PUBLIC_HOLIDAYS.get(date_str, "")


def wa_holiday_dates() -> Set[str]:
    """Return the set of all WA public holiday ISO date strings."""
    return set(WA_PUBLIC_HOLIDAYS.keys())
