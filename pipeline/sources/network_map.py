"""Builds the live SWR Network Map tab's `networkMap` payload.

Added 13-Nov-2026 per Shruti's request to "fix the live map refresh
according to the machine stabling data from daily progress" -- until
this module existed, the map's station/track/machine data
(`DATA.networkMap` / the template's `NMDATA`) had never once been
produced by any live pipeline (confirmed by grep across run_pipeline.py,
aggregate.py's earlier rounds, and every sources/*.py module); the
website's Network Map tab only ever showed a frozen, one-off snapshot
baked directly into the HTML (see claude/network-map-tab-concept.md and
claude/supabase-live-deployment-notes.md's 01-Nov-2026 entry). This
module is what finally produces a real, refreshing payload.

It combines three things:

  1. Station + track GEOMETRY -- live from the `rdps_points` Supabase
     table, via sources/rdps_points.py's fetch_rows()/extract_points().
     Grouped into one station entry per `stationCode` and one track
     polyline per (division, `route`) -- the same grouping concept the
     retired one-off build script
     (/home/claude/mapwork/build/rebuild_data.py) used on the old local
     sbc.json/mys.json/ubl.json files, just re-pointed at the live
     table's own column names (`route` where that script's files said
     `sheet`, etc.). This is what makes the map self-healing if Shruti's
     RDPS survey corrections are ever re-synced (`rdps_sync.gs`, run by
     hand after a fresh correction pass) -- a future re-sync needs no
     code change here, just a fresh pipeline run.

  2. A small, hand-researched DEPOTS / APPROX_STATIONS supplement,
     below. These codes have no RDPS survey point at all and never will
     (3 depot/workshop facilities outside the surveyed network, plus a
     handful of station codes the RDPS extract doesn't carry) -- every
     one was individually researched and confidence-tagged by Shruti
     over several Oct-2026 rounds (full trail in
     claude/network-map-tab-concept.md). Committed here durably, word-
     for-word from /home/claude/mapwork/depots.json and
     /home/claude/mapwork/approx_stations.json's CURRENT contents (not
     the larger historical list quoted in early rounds of that doc --
     several of those early approximate codes, e.g. CLR/BLRR/SBH/DKN/
     BLGA/LD/KIGL, were later superseded by genuine RDPS-surveyed points
     once later correction passes picked them up for real, and are
     correctly gone from the current files; verified directly against
     /home/claude/mapwork/build/map_stations.json, the last actual
     build's output, before trusting this as the current truth) -- this
     is exactly the durability gap network-map-tab-concept.md's own
     "Next steps" section flagged ("Commit the RDPS-to-station lookup
     and the ... web-researched approximate-station entries somewhere
     durable (currently session scratch files)").

  3. Current machine PLACEMENT -- live from the `daily_positions`
     Supabase table, via sources/daily_positions.py's
     fetch_payload()/extract_positions(). Each machine's `stabledAt`
     text is matched to a station using the exact same logic
     rebuild_data.py used for the old one-off snapshot: try the whole
     code, then (if it contains '/') each slash-separated part.

HONEST DEGRADATION (see claude/supabase-live-deployment-notes.md's
13-Nov-2026 entry for the full explanation, written for Shruti in plain
terms): the OLD baked machine entries carried rich per-machine work-
detail fields (block, line, corridorTiming, km/time ranges, stipulated
vs actual km, duty hours, loss reasons, work-done narrative, Sr.DEN,
working/under-overhaul status, IOH/POH type) sourced from a one-off
Daily Progress PDF/Excel parse that had no live Supabase sync at all on
13-Nov-2026. Per-machine entries built here on that date carried only
machine/type/typeName/cat/div/stabledAt/station/posConfidence/source --
nothing invented, everything else simply absent.

UPDATED 14-Nov-2026: daily_progress_sync.gs now reads the sheet's full
per-machine row, not just Stabled At (see that script's own OFFSETS
comment, and sources/daily_positions.py's DETAIL_FIELDS), so the
`daily_positions` table -- and therefore this module's per-machine
entries -- genuinely DO carry the day's real working detail now: which
section/block the machine worked (or "At <station>" if it was idle/
stabled rather than working), how much block time was granted vs
actually used (gross/effective minutes), stipulated vs actual progress
for the day, the cumulative-for-month figures, and the sheet's own
free-text remark. These are copied straight through from
daily_positions.extract_positions()'s DETAIL_FIELDS onto each placed/
unmatched entry below -- unchanged field names, so the template reads
them directly off each machine object. A field is simply absent from a
machine's entry on any day the sync could not locate that division's
Stabled At column (the one case daily_progress_sync.gs itself already
treats as a degrade-gracefully fallback) -- never a fabricated null.

Still genuinely absent, and NOT part of this update: corridorTiming
(multi-spell timing breakdown -- only the day's first spell/row is
read, see daily_progress_sync.gs's own file-header comment for why
that's sufficient for the TOTAL figures), Sr.DEN (that's
daily_progress_srden.py's own separate table, not joined onto these
entries here), and the categorised (not free-text) shortfall reason --
Task #53, still open.

mtype/catMeta/divColor/divName/typeGroup/kpi ARE computed here. As of
the 13-Nov-2026 morning round (before this task), `index.html`'s own
NMDATA is baked with real sample data too (a stopgap fix, since no
pipeline had ever produced `networkMap` before now) -- but `NM.setData()`
still does a WHOLESALE replace of NMDATA once a real payload arrives
(see the template-side fix in the same dated changelog entry), so this
payload must be fully self-contained regardless, matching the old baked
nmdata.json's top-level shape exactly, or the replace would silently
drop the taxonomy metadata the page's rendering code depends on.
catMeta/divColor/divName/typeGroup are static reference tables (copied
from the last baked build, which itself mirrors aggregate.py's own
TAMPING_INFO/UNI_MERGE) that essentially never change; kpi is a genuine
rollup of this run's own real data, computed fresh each time, not
carried over from anywhere.

Per the separation every other adapter in this package uses (see
sources/__init__.py's own docstring), fetching is left to the caller
(aggregate.py) -- this module's own build_network_map() is a pure
function over already-fetched rdps_points rows + a daily_positions
payload, so it is fully unit-testable with synthetic data and never
touches the network itself. See tests/test_network_map.py.
"""
import re
from collections import Counter

import pipeline_helpers as ph

# Relative imports (this module lives inside the `sources` package
# itself, alongside both of these) -- avoids depending on `sources/`
# being separately on sys.path the way the swr_pipeline/ directory is
# for every adapter's `import supabase_client`/`import pipeline_helpers`.
from . import rdps_points
from . import daily_positions

# ---------------------------------------------------------------------
# Static reference tables -- copied verbatim from the last baked
# nmdata.json build's own mtype/catMeta/divColor/divName/typeGroup
# (/home/claude/mapwork/build/nmdata.json), which itself mirrors
# aggregate.py's TAMPING_INFO / pipeline_helpers.UNI_MERGE. These
# essentially never change (a new machine TYPE prefix showing up in
# Daily Progress is rare and would need a TAMPING_INFO update anyway,
# same as every other tab), so they are not re-derived from scratch
# every run -- only `mtype` (which *types* actually appear today) and
# `kpi` (a genuine rollup of today's real data) are computed fresh.
# ---------------------------------------------------------------------
DIV_COLOR = {'UBL': '#0284C7', 'SBC': '#DB2777', 'MYS': '#059669', 'SHARED': '#B45309'}
DIV_NAME = {
    'UBL': 'Hubballi (Hubli) Division', 'SBC': 'Bengaluru (SBC) Division',
    'MYS': 'Mysuru (MYS) Division', 'SHARED': 'Depots / Workshops (CPOH, ZBD, STMRD)',
}
CAT_META = {
    'tamping': {'label': 'Tamping', 'color': '#059669'},
    'nontamping': {'label': 'Non-Tamping', 'color': '#6366F1'},
    'other': {'label': 'Other / Unclassified', 'color': '#64748B'},
}
TYPE_GROUP = dict(ph.UNI_MERGE)  # {'PCTM': 'UNI', 'PCT': 'UNI'}

# ---------------------------------------------------------------------
# Hand-curated supplement -- committed durably, 13-Nov-2026. Copied
# verbatim (same keys/values) from /home/claude/mapwork/depots.json and
# /home/claude/mapwork/approx_stations.json's current contents. Do not
# derive these from any live sync -- they never will have one; see the
# module docstring above and each entry's own `source` note for why.
# ---------------------------------------------------------------------
DEPOTS = {
    "CPOH/RYP": {
        "lat": 16.571, "lon": 80.5688,
        "label": "CPOH Track Machine Workshop, Rayanapadu",
        "source": ("Not in RDPS survey (outside SWR zone -- the POH/periodic-overhaul "
                   "workshop is a shared facility near Vijayawada, Andhra Pradesh). Shruti "
                   "identified the facility from Google results (\"CPOH Track Machine\", "
                   "Hyderabad-Vijayawada Hwy, Guntupalli, Andhra Pradesh 521241); coordinate "
                   "is Rayanapadu town centre from Wikipedia (16°34'16\"N 80°34'08\"E), "
                   "town-level precision, not the exact depot gate."),
    },
    "STMRD/UBL": {
        "lat": 15.333422, "lon": 75.17565,
        "label": "STMRD -- Satellite Maintenance Depot, Hubballi",
        "source": ("Shruti identified this as the satellite maintenance depot at Hubli. "
                   "Co-located with Hubballi Jn station; reusing UBL's RDPS-surveyed position "
                   "(sheet UBL-MRJ DN, km 465.0)."),
    },
    "ZBD/YPR": {
        "lat": 13.026389, "lon": 77.550278,
        "label": "Track Machine Care Center, Yeshwantpur",
        "source": ("Not in RDPS survey (a depot/yard facility, not a surveyed track-chainage "
                   "point). Shruti supplied the location via a Google Maps listing (\"Track "
                   "Machine Care Center\", Railway Engine Service Yard (YPR), Yeshwanthpur, "
                   "Bengaluru) whose Wikimapia description gives coordinates 13°1'35\"N "
                   "77°33'1\"E, converted to decimal here."),
    },
}

# Codes confirmed = exact coordinate from a verified non-RDPS source;
# approximate = nearest confidently-identified place, always with a
# source note (same 3-tier confidence convention as RDPS-surveyed
# stations -- see each entry's own 'confidence').
APPROX_STATIONS = {
    "NAN": {
        "lat": 12.12277, "lon": 76.683522,
        "label": "Nanjangud area (reported as \"NAN\")", "confidence": "approximate",
        "source": ("Not in the RDPS survey extract, and \"NAN\" does not match any IR station "
                   "code found -- the Mysuru-area station on this line (Mysore-Chamarajanagar "
                   "branch, matching this machine's \"NAN-KKW\" block) is actually coded NTW "
                   "(Nanjangud Town). Using NTW's coordinate from Wikipedia as the nearest "
                   "identifiable match for the reported \"NAN\"."),
    },
    "HPT": {
        "lat": 15.2865, "lon": 76.3867,
        "label": "Hosapete Jn / Hospet (HPT)", "confidence": "confirmed",
        "source": ("Not in the RDPS survey extract for this corridor (RDPS sheet \"HPT-VC-KTY\" "
                   "starts a few km further on, at TBDM). HPT is a real, well-known IR junction "
                   "station -- Hosapete (Hospet), UBL division. Coordinate from Wikipedia "
                   "(15°17'11\"N 76°23'12\"E), station-level precision."),
    },
    "SGT": {
        "lat": 12.99712, "lon": 77.7614,
        "label": "Whitefield Satellite Goods Terminal (SGT)", "confidence": "approximate",
        "source": ("Not in the RDPS survey extract, and not a public IRCTC passenger station "
                   "code -- user identified it directly (19-Oct-2026) via a Google Maps "
                   "listing: \"Whitefield Satellite Goods Terminal\", Sadaramangala, "
                   "Krishnarajapuram, Bengaluru, Karnataka 560066, SBC division (consistent "
                   "with machine UTV 08's own SBC-DMM block and with the neighbouring UTV 003 "
                   "logged \"At KJM\" -- Krishnarajapuram -- a few km away in the same sheet). "
                   "Coordinate is the Whitefield goods/container terminal facility location "
                   "(12.99712N 77.7614E) from ShipsGo's INWFT port/terminal listing -- the "
                   "nearest verifiable fix for this specific facility found via web search; no "
                   "station-platform-precise source was available, so treat as an area marker, "
                   "not an exact pin."),
    },
    "UBL": {
        "lat": 15.35, "lon": 75.1491,
        "label": "Hubballi Jn (division HQ)", "confidence": "confirmed",
        "source": ("Wikipedia (Hubballi Junction railway station) -- not in the corrected RDPS "
                   "data (the division HQ junction is not tagged with its own station_code row "
                   "in Shruti's cleaned sheets), added as a confirmed station-level coordinate "
                   "so UNI 8277 (stabled \"At UBL\") maps correctly."),
    },
}

def _machine_type(machine):
    """-> the machine-type prefix, via pipeline_helpers.mtype_of() -- the
    same function the rest of this dashboard (progress/shortfall
    rollups) already uses to classify a machine, rather than
    reimplementing the old rebuild_data.py's own narrower
    `re.match(r'[A-Z]+', ...)`. Keeps this map's type classification
    identical to every other tab's."""
    return ph.mtype_of(machine)


def _type_name(typecode):
    """Short display name for a type code -- identical to every type
    code except 'T', which the source sheet and the rest of this
    dashboard display as 'T-28' (the 19-Oct-2026 short-form-names
    round's one deliberate exception, preserved here)."""
    return 'T-28' if typecode == 'T' else typecode


def _cat_of(typecode):
    """-> 'tamping' or 'nontamping', via pipeline_helpers.TAMPING_INFO --
    the dashboard's own definitive Tamping/Non-Tamping classification
    (user-confirmed 28-Sep-2026), same fallback ('Non-Tamping') aggregate.py
    itself uses for a type code not in that table."""
    label = ph.TAMPING_INFO.get(typecode, ('Non-Tamping', 'Unclassified'))[0]
    return 'tamping' if label == 'Tamping' else 'nontamping'


def _compute_kpi(stations_out, unmatched, as_of_date):
    """A genuine rollup of THIS run's own real data -- never carried over
    from a prior run or invented. Mirrors the old baked nmdata.json's
    kpi shape (total/placed/unplaced/types/tamping/nontamping/
    divCounts/confCounts/asOf) so the template's renderKpis()/
    renderTypeFilter() need no change to read it."""
    placed_entries = [m for st in stations_out.values() for m in st['machines']]
    all_entries = placed_entries + unmatched
    tamping = sum(1 for m in all_entries if m.get('cat') == 'tamping')
    return dict(
        total=len(all_entries),
        placed=len(placed_entries),
        unplaced=len(unmatched),
        types=len({m['type'] for m in all_entries}),
        tamping=tamping,
        nontamping=len(all_entries) - tamping,
        divCounts=dict(Counter(m['div'] for m in all_entries if m.get('div'))),
        # confidence tiers only ever apply to a PLACED machine's actual
        # position -- an unmatched machine has no position to be
        # confident/unconfident about.
        confCounts=dict(Counter(m['posConfidence'] for m in placed_entries if m.get('posConfidence'))),
        asOf=as_of_date,
    )


def _build_mtype(all_entries):
    """-> {typecode: {name, cat}}, one entry per machine type actually
    present in today's data (not a fixed universe) -- mirrors the old
    gen_map_tab.py's own MTYPE dict, which was also built from the types
    actually seen that day."""
    out = {}
    for m in all_entries:
        out.setdefault(m['type'], dict(name=m['typeName'], cat=m['cat']))
    return out


def build_stations_from_rdps(rdps_by_div):
    """rdps_by_div: {division: [{route, km, distanceM, lat, lon,
    stationCode, stationName, reviewNotes}, ...]}, i.e. the shape
    rdps_points.extract_points() returns.

    -> {code: {lat, lon, code, div, confidence: 'surveyed',
    kind: 'rdps_station', depotLabel: None, source}}, one entry per
    distinct stationCode actually present in the live data.

    Where more than one row shares a stationCode (a station surveyed
    from more than one adjoining route/chainage direction), the FIRST
    occurrence's lat/lon is used rather than averaged -- those rows are
    independent fixes of the same physical point, not noisy duplicates
    of one one; averaging them could blend two genuinely different
    chainage readings into a position neither one actually is. This
    mirrors the old rebuild_data.py's own behaviour (it also kept
    whichever single fix happened to land last in iteration order for a
    repeated code).
    """
    stations = {}
    for div, pts in (rdps_by_div or {}).items():
        for p in pts or []:
            code = (p.get('stationCode') or '').strip()
            if not code or p.get('lat') is None or p.get('lon') is None:
                continue
            if code in stations:
                continue
            stations[code] = dict(
                lat=p['lat'], lon=p['lon'], code=code, div=div,
                confidence='surveyed', kind='rdps_station', depotLabel=None,
                source=(f'RDPS {div} div, route "{p.get("route")}", km {p.get("km")}: '
                        f'"{p.get("stationName") or code}"'),
            )
    return stations


# ---------------------------------------------------------------------
# Fixed 14-Nov-2026: adjacent RDPS survey points are only ever a few
# metres apart -- far denser than a schematic web map needs. This was
# harmless while the map read a baked local file, but now that
# rdps_points.fetch_rows() pages past Supabase's 1000-row cap (see
# supabase_client.py's 14-Nov-2026 fix), build_tracks_from_rdps() below
# was handed the FULL ~430,370-point survey set for the first time ever
# -- and was embedding every single one of those points verbatim as
# track-line coordinates, ballooning `networkMap` alone to ~14.5 MB
# (confirmed: 430,370 points, one [lat, lon] pair each). That made the
# daily GitHub Actions push to Supabase's `dashboard_payload` table
# fail outright with a Postgres `57014 canceling statement due to
# statement timeout`, since a single REST insert that large has to be
# parsed and TOASTed by Postgres in one go.
#
# A web map rendering a route as a line has no use for metre-level
# chainage resolution -- the line's visual shape barely changes
# decimated down to a few hundred points per route -- so each route is
# now capped at MAX_POINTS_PER_ROUTE via uniform decimation (see
# _decimate() below), always keeping the first and last point so a
# route's endpoints (and any station snapped near them) are never
# lost. This caps the worst case at (number of routes) x
# MAX_POINTS_PER_ROUTE points (53 routes x 400 = 21,200 today) instead
# of growing with the survey's own raw point count, bringing the live
# payload back down to roughly 1 MB -- comfortably under whatever
# statement-timeout ceiling Supabase's Postgres instance enforces.
# ---------------------------------------------------------------------
MAX_POINTS_PER_ROUTE = 400


def _decimate(rows_sorted, max_points=MAX_POINTS_PER_ROUTE):
    """-> a subset of rows_sorted with at most max_points entries,
    picked at a uniform stride, always keeping the first and last row
    (so a route's two endpoints -- the ones nearest a named station --
    are never dropped). A no-op if there are already <= max_points
    rows."""
    n = len(rows_sorted)
    if n <= max_points:
        return rows_sorted
    step = (n - 1) / (max_points - 1)
    picked = [rows_sorted[round(i * step)] for i in range(max_points - 1)]
    picked.append(rows_sorted[-1])
    return picked


def build_tracks_from_rdps(rdps_by_div):
    """-> {div: [{"sheet": route, "pts": [[lat, lon], ...]}, ...]},
    points sorted by km (then distanceM as a tiebreaker) ascending
    within each route, then decimated to at most MAX_POINTS_PER_ROUTE
    points per route (see the module docstring's 14-Nov-2026 entry for
    why). Mirrors rebuild_data.py's map_tracks.json grouping (by
    {div}:{sheet}), re-keyed one level deeper to match nmdata.json's own
    `tracksByDiv` shape, and grouped by this table's `route` column
    where the old local sbc.json/mys.json/ubl.json files used `sheet`
    -- same concept, live table's own column name.

    Does NOT re-run the anchor-guided outlier-rejection pass the one-off
    geometry-cleaning scripts (clean_tracks.py through clean_tracks4.py,
    documented in claude/network-map-tab-concept.md) used to fix stray/
    duplicate GPS survey passes in the old local extract files -- the
    live `rdps_points` table already holds Shruti's own corrected survey
    data (rdps_sync.gs is run by hand only once a fresh correction pass
    is ready, never on a raw/unreviewed extract), so this function
    trusts the table's points as given, sorted by chainage (then
    decimated purely for payload size, not for data quality). If a
    future survey correction ever reintroduces stray points, re-applying
    that same cleaning method to this live path -- not redoing it from
    scratch -- would be the fix.
    """
    out = {}
    for div, pts in (rdps_by_div or {}).items():
        by_route = {}
        for p in pts or []:
            route = p.get('route')
            if not route or p.get('lat') is None or p.get('lon') is None:
                continue
            by_route.setdefault(route, []).append(p)
        sheets = []
        for route, rows in sorted(by_route.items()):
            rows_sorted = sorted(
                rows, key=lambda r: (r.get('km') if r.get('km') is not None else 0,
                                      r.get('distanceM') if r.get('distanceM') is not None else 0))
            rows_thinned = _decimate(rows_sorted)
            sheets.append(dict(sheet=route, pts=[[r['lat'], r['lon']] for r in rows_thinned]))
        out[div] = sheets
    return out


def merge_supplement_stations(stations):
    """Overlays DEPOTS/APPROX_STATIONS on top of the live RDPS-derived
    `stations` dict, exactly as rebuild_data.py's own map_stations loop
    order did (RDPS stations first, depots next, approx stations last --
    so the hand-curated supplement always wins on a code collision,
    though none is expected in practice: these codes have no RDPS point
    at all)."""
    out = dict(stations)
    for code, info in DEPOTS.items():
        conf = 'confirmed' if code in ('STMRD/UBL', 'ZBD/YPR') else 'approximate'
        out[code] = dict(lat=info['lat'], lon=info['lon'], code=code, div='SHARED',
                          confidence=conf, kind='depot', depotLabel=info['label'],
                          source=info['source'])
    for code, info in APPROX_STATIONS.items():
        # div=None here, same as the old approx_stations.json convention --
        # filled in per-machine below from whichever machine is placed here,
        # since these codes aren't tied to one division the way a depot or
        # an RDPS-surveyed point is.
        out[code] = dict(lat=info['lat'], lon=info['lon'], code=code, div=None,
                          confidence=info['confidence'], kind='web_geocoded',
                          depotLabel=info['label'], source=info['source'])
    return out


def match_machine_to_station(stabled_at, stations):
    """-> (found_code, found_info) or (None, None). Exact code match
    first, then (if the code contains '/') each slash-separated part --
    identical matching order to rebuild_data.py's own logic."""
    if not stabled_at:
        return None, None
    code = str(stabled_at).strip()
    if not code:
        return None, None
    candidates = [code] + (code.split('/') if '/' in code else [])
    for c in candidates:
        c = c.strip()
        if c and c in stations:
            return c, stations[c]
    return None, None


def _is_header_leakage(machine, stabled_at):
    """The live `daily_positions` payload has been observed to carry the
    source sheet's own header row through as if it were a machine
    ("Machine" / "STABLED AT") -- a pre-existing quirk of
    daily_positions.extract_positions(), not something this module can
    fix at the source (every other consumer of that adapter would need
    the same guard). Filtered out here so it never pollutes this map's
    `unmatched` list as a fake machine."""
    m = str(machine or '').strip().lower()
    s = str(stabled_at or '').strip().lower()
    return m == 'machine' or s == 'stabled at'


def build_network_map(rdps_rows, positions_payload):
    """Pure function: combines already-fetched `rdps_points` rows (as
    rdps_points.fetch_rows() returns them) and an already-fetched
    `daily_positions` payload (as daily_positions.fetch_payload()
    returns `payload`, the first element of its tuple) into the live
    `networkMap` shape:

        {stations: {<code>: {lat, lon, code, div, confidence, kind,
                              depotLabel, machines: [...]}, ...},
         tracksByDiv: {<div>: [{"sheet", "pts"}, ...], ...},
         unmatched: [...],
         asOfDate: "<dd.mm.yyyy or None>"}

    Each per-machine entry under a station's `machines` list, and each
    entry in `unmatched`, carries {machine, type, typeName, cat, div,
    stabledAt, station?, posConfidence?, source?, reason?} plus -- as of
    14-Nov-2026 -- whichever of daily_positions.DETAIL_FIELDS that
    machine's row actually had (blockSection, majorSection, block/
    duration/progress figures, remarks, etc. -- see this module's
    "HONEST DEGRADATION" docstring section above for the exact list and
    what is still genuinely absent).

    Also returns mtype/catMeta/divColor/divName/typeGroup/kpi, making
    this payload fully self-contained -- see the module docstring for
    why (the live/null-starting index.html build has nothing baked to
    fall back on for this metadata).
    """
    rdps_by_div = rdps_points.extract_points(rdps_rows)
    stations_geo = build_stations_from_rdps(rdps_by_div)
    stations_geo = merge_supplement_stations(stations_geo)
    tracks_by_div = build_tracks_from_rdps(rdps_by_div)

    positions = daily_positions.extract_positions(positions_payload)

    stations_out = {}
    unmatched = []
    for p in (positions.get('positions') or {}).values():
        machine = (p.get('machine') or '').strip()
        stabled_at = p.get('stabledAt')
        div = p.get('div')
        if not machine or _is_header_leakage(machine, stabled_at):
            continue
        mtype = _machine_type(machine)
        entry = dict(machine=machine, type=mtype, typeName=_type_name(mtype), cat=_cat_of(mtype),
                     div=div, stabledAt=stabled_at)
        # Added 14-Nov-2026: the day's real working-detail fields, when
        # the live sync could read them (see this module's "HONEST
        # DEGRADATION" docstring entry above) -- copied straight through
        # rather than re-derived, so a future daily_positions.py field
        # addition needs no change here.
        for field in daily_positions.DETAIL_FIELDS:
            if field in p:
                entry[field] = p[field]
        found_code, found_info = match_machine_to_station(stabled_at, stations_geo)
        if found_info:
            entry['station'] = found_code
            entry['posConfidence'] = found_info['confidence']
            entry['source'] = found_info.get('source')
            st = stations_out.setdefault(found_code, dict(
                lat=found_info['lat'], lon=found_info['lon'], code=found_code,
                div=(found_info.get('div') or div), confidence=found_info['confidence'],
                kind=found_info['kind'], depotLabel=found_info.get('depotLabel'),
                machines=[]))
            st['machines'].append(entry)
        else:
            entry['reason'] = 'no_code' if not stabled_at else 'station_not_found'
            unmatched.append(entry)

    all_entries = [m for st in stations_out.values() for m in st['machines']] + unmatched
    as_of_date = positions.get('asOfDate')

    return dict(
        stations=stations_out,
        tracksByDiv=tracks_by_div,
        unmatched=unmatched,
        asOfDate=as_of_date,
        mtype=_build_mtype(all_entries),
        catMeta=CAT_META,
        divColor=DIV_COLOR,
        divName=DIV_NAME,
        typeGroup=TYPE_GROUP,
        kpi=_compute_kpi(stations_out, unmatched, as_of_date),
    )
