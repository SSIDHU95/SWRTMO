"""Adapter for the `daily_positions` table.

Live sync: daily_progress_sync.gs. One row, re-pushed daily, giving every
machine's "Stabled At" position AND (as of 14-Nov-2026) its full set of
that day's progress figures, as of the most recent filled date-tab in the
Daily Progress sheet - built for the SWR Network Map tab (plotting
machines on the map, and -- since 14-Nov-2026 -- showing each machine's
last working details in its map-pin drawer). Payload shape:
    {
      "asOfDate": "dd.mm.yyyy",
      "machineCount": 87,
      "machines": [{
          "machine": "BCM-01", "div": "MYS", "stabledAt": "..." or None,
          # Added 14-Nov-2026 -- see DETAIL_FIELDS below and
          # daily_progress_sync.gs's own OFFSETS comment for exactly
          # which sheet column each one is. Present whenever the sync
          # could locate the sheet's Stabled At column for that
          # machine's division that day (true in every real run checked
          # so far); otherwise simply absent from that machine's dict,
          # never a fabricated null.
          "blockSection": "...", "majorSection": "...",
          "locationFrom": 180.35, "locationTo": 180.75,
          "blockAvailedFrom": 8.45, "blockAvailedTo": 10.3,
          "durationMinutes": 105.0, "effBlockMinutes": 70.0,
          "totalBlockGross": 105.0, "totalBlockEffective": 70.0,
          "stipulatedPerEffHr": 0.8, "stipulatedTotal": 0.933,
          "actualPerSpell": 0.4, "actualTotal": 0.7,
          "pctActualVsStipulated": 0.75,
          "cumBlockAvailedTotal": 39.2, "cumBlockAvailedEffective": 16.5,
          "cumProgressForMonth": 14.68, "propTargetTillDate": 20.162,
          "monthlyTarget": 46.528, "remarks": "..."
      }, ...]
    }

NOTE: `remarks` here is the sheet's own free-text note for the day (e.g.
"Block not permitted due to heavy traffic...") -- this is NOT the
separate 19-category shortfall classification (Task #53's still-open
gap); that categorisation work remains undone, this is just the raw
sentence straight from the sheet's own Remarks column.
"""
import supabase_client

TABLE = "daily_positions"

# Every per-machine "last working details" field daily_progress_sync.gs
# can push, beyond machine/div/stabledAt -- kept as an explicit list
# (rather than "pass everything through") so this adapter stays in
# control of its own output shape even if the .gs script is later
# extended with a field this pipeline doesn't know about yet.
DETAIL_FIELDS = [
    'blockSection', 'majorSection', 'locationFrom', 'locationTo',
    'blockAvailedFrom', 'blockAvailedTo', 'durationMinutes', 'effBlockMinutes',
    'totalBlockGross', 'totalBlockEffective', 'stipulatedPerEffHr', 'stipulatedTotal',
    'actualPerSpell', 'actualTotal', 'pctActualVsStipulated',
    'cumBlockAvailedTotal', 'cumBlockAvailedEffective', 'cumProgressForMonth',
    'propTargetTillDate', 'monthlyTarget', 'remarks',
]


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def extract_positions(payload):
    """-> {key: {"machine", "div", "stabledAt", ...DETAIL_FIELDS present}}.

    Fixed 14-Nov-2026: this used to key on
    pipeline_helpers.machine_key_aliased(norm_machine(name)) -- the
    TYPE+NUMBER key the rest of this pipeline uses to cross-reference a
    machine against progress/exception data by its type code (e.g. both
    "BRM 027" and "BRM-027" collapse to the same "BRM27" key, which is
    exactly the point there). That key deliberately also strips out
    parenthetical suffixes (machine_key()'s own `re.sub(r'\\(.*?\\)', ...)`)
    -- harmless for a normal single machine, but this sheet has two
    genuinely separate T-28 units distinguished ONLY by that suffix
    ("T-28 (403&404)" and "T-28(908 A&B)"), so both collapsed onto the
    identical "T28" key and the second one silently overwrote the first
    in this dict -- one real machine quietly vanishing from the
    Network Map's own machine count every run (confirmed directly: of
    the 87 real machine rows in a live daily_positions payload, this
    function returned only 86 distinct keys). Neither of this function's
    two actual callers (aggregate.py's `rawSources.dailyPositions` and
    network_map.py's live map build) does the type-keyed join the old
    docstring anticipated -- both just iterate the dict's own values --
    so there is no reason to collapse distinct machine names together
    here at all. Now keyed on the plain normalized name (whitespace/dash
    normalized, but nothing else stripped), which can never collide two
    genuinely different machine labels the way the type+number key
    could."""
    import pipeline_helpers as ph
    out = {}
    for m in (payload or {}).get('machines') or []:
        name = m.get('machine')
        if not name:
            continue
        entry = dict(machine=str(name).strip(), div=m.get('div'), stabledAt=m.get('stabledAt'))
        for field in DETAIL_FIELDS:
            if field in m:
                entry[field] = m[field]
        out[ph.norm_machine(name)] = entry
    return dict(asOfDate=(payload or {}).get('asOfDate'), positions=out)
