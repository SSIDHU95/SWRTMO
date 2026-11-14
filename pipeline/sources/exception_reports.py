"""Adapter for the `exception_reports` table.

Live sync: exception_sync.gs. It finds whichever date-tab (named
dd.mm.yyyy) in its source Google Sheet is the most recent one that
actually has data, and pushes ONE new Supabase row per day shaped:
    {
      "asOfDate": "17.09.2026",
      "reportTitle": "...",
      "rowCount": 42,
      "rows": [
        {"slNo": 1, "category": "Under repair", "machine": "BCM-01",
         "division": "MYS", "location": "...", "remarks": "...",
         "underSince": <date-or-string>, "tentative": <date-or-string>,
         "daysUnder": 7},
        ...
      ]
    }

Unlike swr_machine_details/unit_cost (one row = the current truth, old
rows superseded), this sync pushes a NEW row every day without deleting
earlier ones - so the table itself accumulates one day's snapshot per
run, which is exactly the multi-day history run_pipeline.py's
extract_exceptions(base) used to build by reading one workbook per day
out of the local 'Exception Sheet/' folder. This adapter therefore calls
fetch_all_payloads() (every row seen so far), not fetch_latest_payload(),
and folds them into the same {day_key: [row, ...]} shape.

NOTE on underSince/tentative: exception_sync.gs's parseTab() reads those
two columns as raw cell values rather than formatting them first, so a
real Date-typed cell comes through JSON as an ISO-8601 timestamp, not the
'dd.mm.yyyy' string run_pipeline.py always produced. pipeline_helpers.
coerce_date_ddmmyyyy() normalizes either shape to 'dd.mm.yyyy' so the
merged payload looks the same regardless of which path produced it.
"""
import datetime
import re

import pipeline_helpers as ph
import supabase_client

TABLE = "exception_reports"


def fetch_all():
    return supabase_client.fetch_all_payloads(TABLE)


def _day_key(as_of_date):
    """'17.09.2026' (or '2026-09-17', or an ISO timestamp) -> '17-Sep',
    matching run_pipeline.py's dt.strftime('%d-%b').lstrip('0') exactly."""
    s = str(as_of_date or "").strip()
    m = re.match(r'^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})', s)
    if m:
        day, mon, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m2 = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
        if not m2:
            return None
        year, mon, day = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
    try:
        dt = datetime.date(year, mon, day)
    except ValueError:
        return None
    return dt.strftime('%d-%b').lstrip('0')


def extract_exceptions(all_rows):
    """[(payload, updated_at), ...] from fetch_all() -> {day_key: [row, ...]}
    with the same per-row keys run_pipeline.py's extract_exceptions(base)
    produces: Category, Machine, Division, Location, Remarks, UnderSince,
    Tentative, DaysUnder. If the sync ever runs more than once for the
    same asOfDate (a re-run after a fix, say), the later push wins, since
    fetch_all_payloads() returns oldest-first and each day_key is simply
    overwritten by whichever payload is processed later in the list."""
    ex = {}
    for payload, _updated_at in all_rows:
        key = _day_key((payload or {}).get("asOfDate"))
        if not key:
            continue
        out_rows = []
        for r in (payload or {}).get("rows") or []:
            machine = r.get("machine")
            if not machine:
                continue
            days_under = r.get("daysUnder")
            out_rows.append(dict(
                Category=r.get("category"),
                Machine=str(machine).strip(),
                Division=r.get("division"),
                Location=r.get("location"),
                Remarks=r.get("remarks"),
                UnderSince=ph.coerce_date_ddmmyyyy(r.get("underSince")),
                Tentative=ph.coerce_date_ddmmyyyy(r.get("tentative")),
                DaysUnder=days_under if isinstance(days_under, (int, float)) else None,
            ))
        ex[key] = out_rows
    return ex
