"""Adapter for the `monthly_machine` table.

Live sync: monthly_machine_sync.gs. One row, re-pushed daily, covering the
"Monthly Machine" workbook - per-category, per-period machine progress
tabs (date-wise progress columns discovered dynamically per period block).
NOT referenced in input-data-specification.md and not read by any
extract_* function in run_pipeline.py (double_shift_roster now lives in
its own "Double shift" folder - see sources/double_shift_roster.py's
docstring on the 07-Oct-2026 folder split) - synced for later use, same as
cumulative_failure/hq_summary/shortfall_reasons/daily_positions. Payload
shape:
    { "months": [{"category": "...", "periods": [...]}, ...] }

Kept as a thin, honest pass-through - see cumulative_failure.py's
docstring for why this project doesn't invent extraction logic for a
source no dashboard tab reads yet.
"""
import supabase_client

TABLE = "monthly_machine"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def extract_monthly_machine(payload):
    """-> the payload's own `months` list, unchanged. Returns [] if the
    payload is empty/missing."""
    return (payload or {}).get('months') or []
