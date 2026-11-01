"""Adapter for the `hq_summary` table.

Live sync: hq_summary_sync.gs. One row, re-pushed daily, covering the
"HQ Summary Items" workbook's most recent filled date-tab - a fleet-by-
type holding/progress summary (fleetRows: machine-type rows plus their
own subtotal/total rows) and a set of per-machine-type progress blocks.
NOT referenced in input-data-specification.md and not read by any
extract_* function in run_pipeline.py - synced for later use, same as
cumulative_failure/monthly_machine/shortfall_reasons/daily_positions.
Payload shape:
    {
      "asOfDate": "...",
      "fleetRows": [{"type": "machine"|"total"|"subtotal", ...}, ...],
      "progressBlocks": [{"machineType": "...", "rows": [...]}, ...]
    }

Kept as a thin, honest pass-through - see cumulative_failure.py's
docstring for why this project doesn't invent extraction logic for a
source no dashboard tab reads yet.
"""
import supabase_client

TABLE = "hq_summary"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def extract_hq_summary(payload):
    """-> {asOfDate, fleetRows, progressBlocks}, unchanged from the live
    payload (each already a clean list of dicts - no local-file precedent
    exists to port against). Returns an all-empty shape if the payload is
    missing."""
    p = payload or {}
    return dict(asOfDate=p.get('asOfDate'), fleetRows=p.get('fleetRows') or [],
                progressBlocks=p.get('progressBlocks') or [])
