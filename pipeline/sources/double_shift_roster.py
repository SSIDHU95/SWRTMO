"""Adapter for the `double_shift_roster` table.

Live sync: double_shift_sync.gs. One row, pushed daily, re-reading and
re-pushing the ENTIRE history every time (the source is one long-running
workbook of repeating date blocks, not a rolling tab-per-day) - so unlike
exception_reports, fetch_latest_payload() is enough here; there's no
accumulation across rows to fold. Payload shape:
    {
      "blockCount": 111, "totalRecords": 1012, "asOfDate": "29/30.10.26",
      "blocks": [
        {"date": "11/12.07.26", "rows": [
            {"slNo": 1, "division": "MYS", "srDen": "...", "section": "...",
             "machine": "BCM-01", "shift": "DAY",
             "maintenanceStarted": "09:15"|9.3, "readiness": ..., "arrived": ...,
             "blockTimings": "...", "progress": "...", "staff": "..."},
            ...
        ]},
        ...
      ]
    }
division/srDen/section/machine are already carry-forwarded onto blank
NIGHT rows server-side (double_shift_sync.gs's `carry` object), matching
the local-file pipeline's own forward-fill pass - so this adapter does
NOT need to re-do that fill.

Replaces run_pipeline.py's extract_double_shift(base), which returns a
flat list of shift records (one per blocks[*].rows[*] entry) each
carrying: date, sl, div, srden, section, machine, shift, started_raw,
ready_raw, arrived_raw, started_hr, ready_hr, arrived_hr, blockTimings,
progress, staff, durationHours, over12h.
"""
import datetime
import re

import supabase_client

TABLE = "double_shift_roster"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def to_hours(v):
    """Same contract as run_pipeline.py's to_hours(): a real time-of-day
    becomes a float hour count, a hand-typed decimal (9.3) is read as
    9h18m (not 9.3 decimal hours - matches the source sheet's own typing
    convention, confirmed in the local-file pipeline), an 'HH:MM'-style
    string parses directly, '-'/blank is None."""
    if v is None or v in ('', '-'):
        return None
    if isinstance(v, (int, float)):
        h = int(v)
        frac = round((v - h) * 100)
        return h + frac / 60
    if isinstance(v, str):
        m = re.match(r'^(\d{1,2})[.:](\d{1,2})$', v.strip())
        if m:
            return int(m.group(1)) + int(m.group(2)) / 60
        try:
            return float(v)
        except Exception:
            return None
    return None


def _display_time(v):
    """The live sync already formats a real Excel time cell to 'HH:mm'
    (formatTimeCell() in double_shift_sync.gs) and passes a hand-typed
    decimal straight through unchanged - so unlike run_pipeline.py's own
    display_time() (which has to do that formatting itself from a raw
    openpyxl cell), this just needs a str() for a decimal value so the
    *_raw display fields are always text, matching the local version's
    return type."""
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return str(v)
    return str(v)


def _duration_hours(started_hr, arrived_hr):
    if started_hr is None or arrived_hr is None:
        return None
    d = arrived_hr - started_hr
    if d < 0:
        d += 24
    return round(d, 2)


def extract_double_shift(payload):
    """-> [record, ...], flattening every block's rows into run_pipeline.
    py's extract_double_shift(base) record shape. A row whose shift text
    isn't exactly DAY/NIGHT after upper-casing is skipped - the live sync
    already normalizes common typing slips (Day, NiGHT, etc.) to upper
    case before pushing, same as the local-file version's own filter."""
    out = []
    for block in (payload or {}).get("blocks") or []:
        date = block.get("date")
        for r in block.get("rows") or []:
            shift = str(r.get("shift") or "").strip().upper()
            if shift not in ("DAY", "NIGHT"):
                continue
            started = r.get("maintenanceStarted")
            ready = r.get("readiness")
            arrived = r.get("arrived")
            started_hr = to_hours(started)
            ready_hr = to_hours(ready)
            arrived_hr = to_hours(arrived)
            duration = _duration_hours(started_hr, arrived_hr)
            out.append(dict(
                date=date, sl=r.get("slNo"), div=r.get("division"), srden=r.get("srDen"),
                section=r.get("section"), machine=r.get("machine"), shift=shift,
                started_raw=_display_time(started), ready_raw=_display_time(ready),
                arrived_raw=_display_time(arrived),
                started_hr=started_hr, ready_hr=ready_hr, arrived_hr=arrived_hr,
                blockTimings=r.get("blockTimings"),
                progress=(str(r["progress"]) if r.get("progress") is not None else None),
                staff=(str(r["staff"]) if r.get("staff") is not None else None),
                durationHours=duration,
                over12h=bool(duration is not None and duration > 12),
            ))
    return out
