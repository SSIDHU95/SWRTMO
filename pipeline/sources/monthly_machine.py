"""Adapter for the `monthly_machine` table.

Live sync: monthly_machine_sync.gs. One row, re-pushed daily, covering the
"Monthly Machine" workbook -- per-category, per-period machine progress
tabs (date-wise progress columns discovered dynamically per period block).
Confirmed against the real live payload (not just the docstring this
module used to carry): `months` is a DICT keyed by month label (e.g.
"OCT-26", "SEP-26" -- her own Nov tab is literally named "NOV 26" with a
space instead of a hyphen like every other month, which is almost
certainly why the live sync hadn't picked up November yet as of
17-Nov-2026: _parse_month_key() below tolerates that variant anyway, so
this adapter won't silently drop November once the sync does carry it,
but the naming slip on the sheet itself is worth fixing at the source).
    {
      "months": {
        "OCT-26": {"categories": [
            {"category": "SQRS", "title": "SQRS MACHINE PROGRESS FOR THE
             MONTH OF OCTOBER", "periods": [
                {"periodLabel": "Date wise machine progress from
                 01.10.26 to 15.10.2026", "extraFieldNames": [...],
                 "totals": {"daily": {...}, "extra": {...}},
                 "rows": [
                    {"divn": "UBL", "machine": "SQRS 37 & 38",
                     "section": "GDG-HG (Sr.DEN.C)",
                     "daily": {"1": 0.26, "2": 0.26, "3": "Depot work",
                               ...},
                     "extra": {"Scope upto 30th SEP-26 in kms": 3.285,
                               "Progress achieved upto 31st OCT-26": 3.073,
                               "Balance of work upto 31st OCT-26": 0.212}},
                    ...
                 ]},
                {... second half of the month ...}
            ]},
            ...
        ]},
        ...
      }
    }
`daily`'s values are EITHER a plain number (actual km/unit progress that
day) OR a free-text remark string when no numeric progress was recorded
("Depot work", "Block not permitted by operating department.", etc).

IMPORTANT: a category's own label is NOT a reliable type classifier --
confirmed against the live "T-28" category, which actually mixes real
T-28 (PCCM), PCTM and UNI machines in one block (it even carries an
embedded second section-header row, "PCTM & UNIMAT MACHINE PROGRESS...",
with a blank machine/daily/extra -- skipped below like any other
blank-machine row). extract_monthly_machine_by_machine() below classifies
every ROW by its OWN machine name instead, which is also simply the
safer general rule.

Two extraction functions, added/changed 18-Nov-2026 per Shruti: "there is
one monthly machine progress sheet also... connected with the monthly
progress sheet" (this table had been synced for later use but read by no
caller until now):
  - extract_monthly_machine(payload) -- unchanged in spirit, still a thin
    pass-through for the Live Data Sources tab's raw viewer, just
    corrected to the real dict shape (it used to claim/assume a list).
  - extract_monthly_machine_by_machine(payload) -- NEW: reshapes the
    sheet into a per-machine day-by-day log, feeding the Network Map
    popup's "last team's progress" and the machine detail drawer's
    "last 15 days" breakdown (see aggregate.py).
"""
import datetime
import re

import pipeline_helpers as ph
import supabase_client

TABLE = "monthly_machine"

_MONTH_NUM = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def _parse_month_key(key):
    """'OCT-26' / 'NOV 26' / 'nov-2026' -> (year, month) as (2026, 11),
    or None if unparseable. Accepts a hyphen, a space, or nothing between
    the month name and the year, and a 2- or 4-digit year, so a sheet
    tab's own naming slip (her Nov tab is "NOV 26", not "NOV-26" like
    every other month) doesn't also break this adapter."""
    m = re.match(r'\s*([A-Za-z]+)[\s\-]?(\d{2,4})\s*$', str(key or ''))
    if not m:
        return None
    mon = _MONTH_NUM.get(m.group(1).strip().lower()[:3])
    if not mon:
        return None
    yr = int(m.group(2))
    yr = 2000 + yr if yr < 100 else yr
    return yr, mon


def extract_monthly_machine(payload):
    """-> the payload's own `months` dict, unchanged (a thin, honest
    pass-through -- see cumulative_failure.py's docstring for why this
    project doesn't invent extraction logic for a source no dashboard tab
    reads yet). This is what feeds payload.rawSources.monthlyMachine for
    the Live Data Sources tab's own generic viewer (index.html's
    ldsSectionMonthlyMachine) -- kept exactly pass-through, unlike
    extract_monthly_machine_by_machine() below, so that raw/debug view
    keeps showing the sheet exactly as Supabase has it.

    Corrected 18-Nov-2026: this used to claim (and callers used to
    assume) `months` was a LIST; the real live payload is a DICT keyed
    by month label ("OCT-26", etc). Returns {} if the payload is empty."""
    months = (payload or {}).get('months')
    return months if isinstance(months, dict) else {}


def extract_monthly_machine_by_machine(payload):
    """-> {machine_key_aliased(machine): [{"date": "YYYY-MM-DD",
    "value": <number-or-remark-string>}, ...]}, one entry per calendar
    day this sheet actually has a value for, sorted oldest-first, across
    EVERY month currently in the payload (not just the latest one) --
    callers slice however many recent days they need. Added 18-Nov-2026
    for the Network Map popup's "last team's progress" and the machine
    detail drawer's "last 15 days" day-by-day breakdown (see
    aggregate.py) -- a separate function from extract_monthly_machine()
    above (which stays a plain pass-through for the Live Data Sources
    tab) since this one re-shapes the data rather than just forwarding it.

    Keyed by machine_key_aliased(norm_machine(...)), the SAME join key
    swr_machine_details.py's home_div fix and the age-profile join use,
    so a BRM/PBR, SBCM/FRM, DUO/WST naming mismatch between this sheet
    and Cumulative Progress never silently drops a machine's daily log
    (same bug class as the 18-Nov-2026 PBR/home-division fix) -- even
    though in practice this sheet's own machine names already use the
    canonical prefix (confirmed: its BRM/SBCM/DUO categories all say
    "BRM"/"SBCM"/"DUO", not "PBR"/"FRM"/"WST"), aliasing the key is a
    cheap, free safety net against the same class of mismatch recurring
    here too.

    Returns {} if the payload is empty or carries no usable months."""
    months = (payload or {}).get('months')
    if not isinstance(months, dict):
        return {}
    out = {}
    for month_key, month_val in months.items():
        parsed = _parse_month_key(month_key)
        if not parsed:
            continue
        year, mon = parsed
        for cat in (month_val or {}).get('categories') or []:
            for period in cat.get('periods') or []:
                for row in period.get('rows') or []:
                    name = (row.get('machine') or '').strip()
                    if not name or name.lower().startswith('total progress'):
                        continue
                    key = ph.machine_key_aliased(ph.norm_machine(ph.canon_machine_alias(name)))
                    if not key:
                        continue
                    log = out.setdefault(key, {})
                    for day_str, val in (row.get('daily') or {}).items():
                        if val is None or val == '' or val == '-':
                            continue
                        try:
                            d = datetime.date(year, mon, int(day_str))
                        except (TypeError, ValueError):
                            continue
                        log[d.isoformat()] = val
    return {k: [dict(date=dt, value=v) for dt, v in sorted(entries.items())]
            for k, entries in out.items()}
