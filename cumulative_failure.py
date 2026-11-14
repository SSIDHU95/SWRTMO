"""Adapter for the `cumulative_failure` table.

Live sync: cumulative_failure_sync.gs. One row, re-pushed daily, covering
the "CUMULATIVE FAILURES FOR TMC RECORD" workbook. NOT referenced in
input-data-specification.md and not read by any extract_* function in
run_pipeline.py - synced for later use, same as several other sources
this round (hq_summary, monthly_machine, shortfall_reasons,
daily_positions). Payload shape:
    { "years": { "<FY label>": {...per-year parse...}, ... } }

Kept as a thin, honest pass-through rather than inventing a shape no
dashboard tab asks for yet: fetch_payload() does the one REST call, and
extract_cumulative_failure() hands back the `years` dict unchanged so a
future consumer (or an integrity test asserting the table is reachable
and non-empty) has a single, documented entry point rather than reaching
into supabase_client directly.
"""
import supabase_client

TABLE = "cumulative_failure"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def extract_cumulative_failure(payload):
    """-> the payload's own `years` dict, unchanged. Returns {} if the
    payload is empty/missing, matching this pipeline's usual degrade-to-
    empty convention for an optional source."""
    return (payload or {}).get('years') or {}
