"""Adapter for the `sidings_rest_house` table.

Live sync: sidings_rest_house_sync.gs. One row, pushed daily, covering the
4 tabs run_pipeline.py's extract_sidings_rest_houses(base) reads (Master
Summary, Construction, CE TM Sr DEN detail, TM Sanctioned Siding Status)
PLUS 7 supplementary tabs nothing reads yet (synced per the user's
request, 31-Oct-2026 - kept in the raw payload but not surfaced by this
adapter, same as the rest of this pipeline leaving an unused source alone
rather than guessing a use for it). Payload shape:
    {
      "masterSummary": {"asOf": "...", "note": "...",
          "siding": {"byDivision": [...], "swr": {...}, "target": {...}},
          "restHouse": {"byDivision": [...], "swr": {...}, "target": {...}}},
      "construction": [{"division", "section", "sectionLength", "stations",
          "stationRequired", "sidingAvailable", "sidingSanctioned",
          "sidingProposed", "rhAvailable", "rhSanctioned", "rhProposed"}, ...],
      "ceTmSrDenDetail": [{"division", "asOf", "srden", "section", "station",
          "km", "interSidingDistance", "sidingStatus", "rhStatus", "remark"}, ...],
      "sanctionedSidingStatus": [{"division", "asOf", "srden", "srdenKey",
          "station", "esp", "sip", "crs", "agencyAvailable",
          "trackWorksCompleted", "sntWorksCompleted", "remark", "stage"}, ...],
      "supplementary": {"UBL": [...], "SBC": [...], ...}   # not used here
    }

Confirmed against the real live payload: masterSummary.siding/restHouse's
byDivision rows, and sanctionedSidingStatus's rows, already carry the exact
same field names run_pipeline.py's own div_block()/parseSanctionedSiding
Status-equivalent code produces (the .gs script mirrors that extraction
logic directly) - so those two pass through close to unchanged. The two
things the live sync does NOT do that the local-file version does are:
(1) normalize_infra_status() on sidingStatus/rhStatus (the live payload's
values are the sheet's raw text, e.g. 'Available and Sanctioned' already
clean but '????'/'Not feasible'-vs-'Not Feasible' casing not yet
collapsed), (2) classify_infra_responsibility() on the free-text remark,
and (3) the srdenKey canon_srden_word() computation for CE TM Sr DEN detail
rows (sanctionedSidingStatus already computes its own srdenKey server-side,
confirmed identical to canon_srden_word()'s own logic) - all three applied
here, plus the srdenTable/responsibility rollups run_pipeline.py's run()
derives from the station rows afterward.
"""
import re
from collections import Counter, defaultdict

import pipeline_helpers as ph
import supabase_client

TABLE = "sidings_rest_house"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


_STATUS_WORD = {'available': 'Available', 'sanctioned': 'Sanctioned', 'proposed': 'Proposed'}


def _status_of(avail, sanc, prop):
    if avail:
        return 'Available'
    if sanc:
        return 'Sanctioned'
    if prop:
        return 'Proposed'
    return None


def stations_from_srden_detailed(payload):
    """One station row per 'SrDEN Detailed' entry that has a siding or a
    rest-house status, in the stationRows shape the website already uses."""
    from sources import infra_map  # local import: infra_map imports network_map
    as_of = None
    for b in (payload.get('supplementary') or {}).get('SrDEN Detailed') or []:
        m = re.search(r'(\d{1,2}\.\d{1,2}\.\d{4})', b.get('title') or '')
        if m:
            as_of = m.group(1)
    rows = []
    for r in infra_map._parse_srden_detailed(payload, {}):
        siding = _status_of(r['sidingAvailTM'] or r['sidingAvailRGM'], r['sidingSancTM'] or r['sidingSancRGM'],
                            r['sidingPropTM'] or r['sidingPropRGM'])
        rh = _status_of(r['rhAvail'], r['rhSanc'], r['rhProp'])
        if not (siding or rh):
            continue
        rows.append(dict(
            division=r['division'], asOf=as_of, srden=r['srdenRaw'], srdenKey=r['srdenKey'],
            section=r['section'], station=r['station'], km=r['km'],
            interSidingDistance=r['interSidingDistance'],
            sidingStatus=siding, rhStatus=rh,
            sidingType=('TM + RGM' if (r['sidingAvailTM'] or r['sidingSancTM'] or r['sidingPropTM'])
                        and (r['sidingAvailRGM'] or r['sidingSancRGM'] or r['sidingPropRGM'])
                        else 'RGM' if (r['sidingAvailRGM'] or r['sidingSancRGM'] or r['sidingPropRGM'])
                        else 'TM' if siding else None),
            remark=r['remark'], responsibility=ph.classify_infra_responsibility(r['remark']),
        ))
    return rows


def extract_sidings_rest_houses(payload):
    """-> the same dict shape run_pipeline.py's extract_sidings_rest_houses
    (base) returns (asOf, summary, note, construction, stationRows,
    srdenTable, responsibility, sanctionedSidingStatus, sourceFile), or
    None if the live payload itself is empty/missing (mirrors the
    local-file version's own "workbook not found" -> None degrade path)."""
    if not payload:
        return None

    ms = payload.get('masterSummary') or {}
    summary = dict(sidings=ms.get('siding'), restHouses=ms.get('restHouse'))
    note = ms.get('note')
    as_of_overall = ms.get('asOf')

    construction = payload.get('construction') or []

    station_rows = []
    ce_rows = payload.get('ceTmSrDenDetail') or []
    if not ce_rows:
        # 28-Nov-2026: the old "CE TM Sr DEN detail" tab no longer syncs
        # any rows (the workbook moved to "SrDEN Detailed"), which left
        # Stations Covered = 0, Pending - CN Units = 0, an empty Sr.DEN
        # table and "No data" in the Pending-responsibility chart. Build
        # the same station rows from the live 'SrDEN Detailed' tab instead
        # (the one the Infra map already reads), keeping siding and rest
        # house status separate, exactly as the map does.
        station_rows = stations_from_srden_detailed(payload)
    for row in ce_rows:
        division = row.get('division')
        srden_raw = row.get('srden')
        remark = row.get('remark')
        station_rows.append(dict(
            division=division, asOf=row.get('asOf'), srden=srden_raw,
            srdenKey=ph.canon_srden_word(division, srden_raw),
            section=row.get('section'), station=row.get('station'), km=row.get('km'),
            interSidingDistance=row.get('interSidingDistance'),
            sidingStatus=ph.normalize_infra_status(row.get('sidingStatus')),
            rhStatus=ph.normalize_infra_status(row.get('rhStatus')),
            remark=remark, responsibility=ph.classify_infra_responsibility(remark),
        ))

    # ---- Sr.DEN jurisdiction rollup (same shape as run_pipeline.py's run()) ----
    groups = {}
    for row in station_rows:
        key = row['srdenKey']
        if not key:
            continue
        g = groups.get(key)
        if g is None:
            g = groups[key] = dict(key=key, division=row['division'], srden=row['srden'],
                                    stations=0, sidingStatus=Counter(), rhStatus=Counter(),
                                    responsibility=Counter())
        g['stations'] += 1
        if row['sidingStatus']:
            g['sidingStatus'][row['sidingStatus']] += 1
        if row['rhStatus']:
            g['rhStatus'][row['rhStatus']] += 1
        if row['responsibility']:
            g['responsibility'][row['responsibility']] += 1
    srden_table = [
        dict(key=g['key'], division=g['division'], srden=g['srden'], stations=g['stations'],
             sidingStatus=dict(g['sidingStatus']), rhStatus=dict(g['rhStatus']),
             responsibility=dict(g['responsibility']))
        for g in groups.values()
    ]

    resp_overall = Counter()
    resp_by_div = defaultdict(Counter)
    for row in station_rows:
        if row['responsibility']:
            resp_overall[row['responsibility']] += 1
            if row['division']:
                resp_by_div[row['division']][row['responsibility']] += 1
    responsibility = dict(
        overall=dict(resp_overall),
        byDivision={d: dict(c) for d, c in resp_by_div.items()},
    )

    # ---- TM Sanctioned Siding Status rollup - rows already match
    # run_pipeline.py's sanc_rows shape field-for-field (srdenKey included),
    # so only the byDivision/stageOverall/stageByDivision aggregation needs
    # redoing here. ----
    sanc_rows = payload.get('sanctionedSidingStatus') or []
    sanctioned_status = None
    if sanc_rows:
        stage_overall = Counter(r['stage'] for r in sanc_rows)
        stage_by_div = defaultdict(Counter)
        div_counts = defaultdict(lambda: dict(total=0, espDone=0, sipDone=0, crsDone=0,
                                                agencyAvailable=0, trackWorksCompleted=0,
                                                sntWorksCompleted=0, completed=0))
        for r in sanc_rows:
            stage_by_div[r['division']][r['stage']] += 1
            dc = div_counts[r['division']]
            dc['total'] += 1
            if r.get('esp'):
                dc['espDone'] += 1
            if r.get('sip'):
                dc['sipDone'] += 1
            if r.get('crs'):
                dc['crsDone'] += 1
            if r.get('agencyAvailable'):
                dc['agencyAvailable'] += 1
            if r.get('trackWorksCompleted'):
                dc['trackWorksCompleted'] += 1
            if r.get('sntWorksCompleted'):
                dc['sntWorksCompleted'] += 1
            if r['stage'] == 'Completed':
                dc['completed'] += 1
        sanctioned_status = dict(
            asOf=max((r.get('asOf') for r in sanc_rows if r.get('asOf')), default=None),
            stations=sanc_rows,
            stageOverall=dict(stage_overall),
            stageByDivision={d: dict(c) for d, c in stage_by_div.items()},
            byDivision={d: dict(c) for d, c in div_counts.items()},
        )

    return dict(
        asOf=as_of_overall, summary=summary, note=note, construction=construction,
        stationRows=station_rows, srdenTable=srden_table, responsibility=responsibility,
        sanctionedSidingStatus=sanctioned_status,
        sourceFile="live sync (sidings_rest_house)",
    )
