"""Adapter for the `shortfall_reasons` table.

Live sync: shortfall_reasons_sync.gs. One row, re-pushed daily, covering a
DIFFERENT workbook than its name might suggest - "AVAILABILITY" (division-
wise machine availability + two fleet-utilization snapshots) and
"UTILISATION" (a flat per-machine-type table with its own Shortfall
remarks column) and "Machine wise short fall" (down-days, tamping vs
non-tamping, grouped by machine type). This is NOT the Daily Progress
shortfall-category log that run_pipeline.py's extract_shortfall_records
(base) reads (Task #53's still-open gap, ~15,000 Remarks rows classified
into 19 categories) - confirmed by the .gs file's own header comment
("NOT wired into any dashboard tab today"), and by this source's sheet
names simply not matching anything in input-data-specification.md. Synced
for later use, same as cumulative_failure/hq_summary/monthly_machine/
daily_positions. Payload shape:
    {
      "divisionAvailability": {"period", "headers", "rows", "total"},
      "fleetUtilizationSept": {"asOf", "headers", "rows"},
      "fleetUtilizationOct": {"asOf", "headers", "rows"},
      "machineWiseShortfall": {"tamping": {...}|None, "nonTamping": {...}|None}
    }

Kept as a thin, honest pass-through - see cumulative_failure.py's
docstring for why this project doesn't invent extraction logic for a
source no dashboard tab reads yet.
"""
import supabase_client

TABLE = "shortfall_reasons"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def extract_shortfall_reasons(payload):
    """-> {divisionAvailability, fleetUtilizationSept, fleetUtilizationOct,
    machineWiseShortfall}, unchanged from the live payload. Returns an
    all-empty shape if the payload is missing."""
    p = payload or {}
    return dict(
        divisionAvailability=p.get('divisionAvailability') or {},
        fleetUtilizationSept=p.get('fleetUtilizationSept') or {},
        fleetUtilizationOct=p.get('fleetUtilizationOct') or {},
        machineWiseShortfall=p.get('machineWiseShortfall') or {},
    )
