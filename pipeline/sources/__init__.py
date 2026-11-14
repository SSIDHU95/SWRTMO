"""Supabase-backed source adapters - one module per live-sync table.

Every adapter in this package follows the same shape, so aggregate.py can
drive all 16 the same way and the integrity tests can feed each one a
canned payload without touching the network:

    fetch_payload()            -> (payload: dict, updated_at: str)
        The one REST call against this adapter's table (via
        supabase_client.fetch_latest_payload/fetch_all_payloads). Always a
        thin wrapper - no reshaping happens here.

    extract_*(payload) -> ...
        Pure functions that reshape the raw Supabase payload into exactly
        what the corresponding extract_*() in claude/pipeline/run_pipeline.py
        used to return from local files. These never touch the network, so
        tests/ can call them directly against fixture payloads.

Keeping "fetch" and "extract" separate (rather than one function doing
both, as run_pipeline.py's local-file version does) is deliberate here:
it means every adapter's reshaping logic is unit-testable without a live
Supabase connection, and a single adapter can fetch once and derive
several differently-shaped outputs from that one payload (e.g.
swr_machine_details needs both a home_div map and an age_profile list out
of the same row) without a second network round trip.
"""
