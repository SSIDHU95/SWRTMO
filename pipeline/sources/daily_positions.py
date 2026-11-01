"""Adapter for the `daily_positions` table.

Live sync: daily_progress_sync.gs. One row, re-pushed daily, giving every
machine's "Stabled At" position as of the most recent filled date-tab in
the Daily Progress sheet - built for the SWR Network Map tab (plotting
machines on the map), not for anything run_pipeline.py's run() reads
today. Payload shape:
    {
      "asOfDate": "dd.mm.yyyy",
      "machineCount": 87,
      "machines": [{"machine": "BCM-01", "div": "MYS", "stabledAt": "..."
          or None}, ...]
    }

NOTE: despite the sheet's name, this is NOT the Daily Progress shortfall-
category log (Task #53's still-open gap) - that free-text Remarks/Category
data lives in a different part of the same Google Sheet and has no
Supabase sync at all. This table only carries each machine's current
stabled-at location.
"""
import supabase_client

TABLE = "daily_positions"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def extract_positions(payload):
    """-> {machineKeyAliased: {"machine", "div", "stabledAt"}, ...}, keyed
    the same way the rest of this pipeline cross-references machines
    (pipeline_helpers.machine_key_aliased), so a future Network Map
    adapter consumer can join this straight onto progress.machines."""
    import pipeline_helpers as ph
    out = {}
    for m in (payload or {}).get('machines') or []:
        name = m.get('machine')
        if not name:
            continue
        out[ph.machine_key_aliased(ph.norm_machine(name))] = dict(
            machine=str(name).strip(), div=m.get('div'), stabledAt=m.get('stabledAt'),
        )
    return dict(asOfDate=(payload or {}).get('asOfDate'), positions=out)
