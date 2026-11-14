"""Adapter for the `daily_progress_srden` table.

Live sync: daily_progress_srden_sync.gs. Shruti explicitly chose this
narrower scope (14-Nov-2026): just the Sr.DEN (jurisdiction) field out of
the Daily Progress sheet, not the full shortfall-category/reason-for-loss
log from that same sheet -- that remains Task #53, a separate, still-open,
declined-for-now piece of work (extract_shortfall_records() in
aggregate.py still unconditionally returns [] and is NOT touched by this
adapter).

This is independent of daily_positions (also synced from the same Daily
Progress Google Sheet, but reading its "Stabled At" column for the
Network Map, not Sr.DEN) -- two separate syncs, two separate tables, same
source sheet.

One row, re-pushed daily, giving every machine's Sr.DEN jurisdiction as
of the most recent filled date-tab in the Daily Progress sheet. Payload
shape:
    {
      "asOfDate": "dd.mm.yyyy",
      "machines": [
        {"machine": "SQRS - 51 & 52", "srDen": "OL Sr.DEN.W.", "division": "UBL"},
        ...
      ]
    }

"srDen" is the sheet's own raw banner text, forward-filled per
daily_progress_srden_sync.gs's own docstring -- deliberately NOT
canonicalized by the Apps Script (that would duplicate interpretation
logic in two languages); this adapter canonicalizes it into the exact
"<DIV> Sr.DEN/<code>" key format every other part of this pipeline
already expects (pipeline_helpers.canon_srden, ported verbatim from
run_pipeline.py -- the same key shape progress.machines[].srden,
machinesForSrden(), and the dashboard's own Sr.DEN jump buttons all key
on). "srDen" can legitimately be null for a machine that is the very
first row of its division section in the source sheet with no Sr.DEN
banner stated at all that day (confirmed directly against a real
uploaded file, 07/06/05-Nov-2026 tabs: SQRS 7&8 in the SBC section) --
not a parsing bug, an honest reflection of a genuine gap in that day's
source data.

The Konkan Railway section (a different railway's machines, not SWR) is
already excluded entirely by daily_progress_srden_sync.gs itself before
this ever reaches Supabase, so this adapter never has to filter it out.
"""
import pipeline_helpers as ph
import supabase_client

TABLE = "daily_progress_srden"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def extract_srden(payload):
    """-> {"asOfDate": ..., "srden": {machineKeyAliased: canonicalSrDenKey}}.

    Keyed the same way the rest of this pipeline cross-references
    machines (pipeline_helpers.machine_key_aliased(pipeline_helpers.norm_
    machine(...))), same convention sources/daily_positions.py already
    uses for this exact same Google Sheet. A machine whose raw Sr.DEN text
    doesn't canonicalize to one of the 6 known codes (N/S/E/W/C/HQ) -- or
    has none at all that day -- is simply left out of the dict rather than
    guessed, so a lookup miss degrades to "no live Sr.DEN for this
    machine" rather than a wrong one.

    NOTE on machine_key_aliased collisions: machine_key_aliased() drops
    parenthetical suffixes before keying, so the project's two distinct
    physical T-28 duo-tamping units ("T-28 (908 A&B)" and "T-28
    (403&404)") key identically here and only the later one in sheet order
    survives in this dict -- the same accepted, pre-existing simplification
    sources/daily_positions.py already has for "Stabled At" on these same
    two machines, not a new regression introduced by this adapter.
    """
    out = {}
    for m in (payload or {}).get('machines') or []:
        name = m.get('machine')
        if not name:
            continue
        div = m.get('division')
        key = ph.canon_srden(div, m.get('srDen'))
        if not key:
            continue
        out[ph.machine_key_aliased(ph.norm_machine(name))] = key
    return dict(asOfDate=(payload or {}).get('asOfDate'), srden=out)
