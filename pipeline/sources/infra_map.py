"""Builds the Infra tab's new Sidings & Rest House map + Sr.DEN breakdown,
per Shruti's 26-Nov-2026 request (see claude/project-status.md section 3).

Combines THREE things:

1. The live `sidings_rest_house` Supabase payload's `supplementary` key --
   7 extra tabs synced since 31-Oct-2026 that sidings_rest_house.py's own
   extract_sidings_rest_houses() doesn't read. Two of them are what this
   module actually needs:
     - 'SrDEN Detailed': one row per station (165 stations across all 3
       divisions, as 3 per-division {rows, title, total} blocks). THIS
       CARRIES A LIVE 'Remark' FIELD per station -- confirmed by direct
       REST fetch 26-Nov-2026 (e.g. station E UBL: "RH- Not Feasible";
       NVU: "TM RH has been proposed by the Div in IRPSM"). This was
       mistakenly believed absent from the live sync in an earlier pass
       at this feature; it is not -- only the OTHER sheet below lacks it.
     - 'UBL Sr.DEN' / 'SBC Sr.DEN' / 'MYS Sr. DEN': the per-division
       Sr.DEN/Major-Section breakdown, each as 3 sub-blocks (TM siding /
       RGM siding / Rest House), sharing one 'Sl No' numbering AND one
       'Major Section' value across all 3 sub-blocks for the same
       division (confirmed directly against live data: Sl No 9's "Major
       Section" is identical, 'UBL-GDG', in both the TM and RGM blocks).
       This lets the 3 sub-blocks be merged back into one row per
       section. These sheets do NOT carry their own Remarks column live
       (confirmed: no row has a 'Remark' key) -- but every station named
       in one of these sections also appears in 'SrDEN Detailed' above,
       so this module attaches each station's live remark from there
       instead of needing that column live at all.

2. Live `rdps_points` rows (already fetched once by aggregate.py's
   extract_network_map() for the Network Map tab -- passed in here
   pre-grouped as `rdps_by_div`, exactly as rdps_points.extract_points()
   returns it, so this module adds no extra network round trip).

3. A static supplement file (pipeline/data/infra_supplement.json),
   genuinely necessary for 4 things with no live equivalent:
     - The Master Summary siding/restHouse KPI block. The live Apps
       Script's own computation of this (payload['masterSummary']) is
       confirmed WRONG for the restHouse half: every row is shifted down
       by one against the real sheet (its "swr" total is actually
       Mysuru's own row; the real SWR total was never read at all).
       Confirmed directly: live swr={available:17,...} vs. the real
       sheet value {available:63,...}. Until the Apps Script itself is
       fixed (not possible from this session -- it runs on Shruti's
       Google account), this module trusts the static supplement's
       from-the-workbook figures for this one KPI block instead of the
       live (wrong) masterSummary.
     - Strikethrough formatting (11→9 stations marking a withdrawn/"not
       feasible" proposal as of the 26-Nov-2026 workbook) -- cell
       formatting isn't carried by the live sync's getValues() call at
       all, live or not.
     - 'TM Sanctioned RH Status' and 'Status of Sanctioned Depot' -- two
       sheets added to the workbook after the live sync's sheet list was
       last set up; not live yet.
     - The RB Targets circular's figures (a signed PDF, issued ~once a
       year, same "can't live-sync a PDF" reasoning as
       pipeline/data/cpoh_ryp_planning.json).

   See pipeline/data/infra_supplement.json's own "_comment" key for
   exactly how Shruti replaces this file the next time she sends an
   updated workbook or a new RB circular -- no code change needed here.

HONEST DEGRADATION: if `payload` is falsy, returns None (same convention
as sidings_rest_house.extract_sidings_rest_houses). If `supplement` is
None (file missing/unreadable), the function still returns a result built
entirely from live data -- the KPI block's siding/restHouse figures fall
back to the live (confirmed-buggy for restHouse) masterSummary rather than
crashing, strikeouts/depot/RB-target sections are simply omitted, and this
is logged by the caller, not raised here (this module is a pure function,
same convention as network_map.build_network_map).
"""
import re
from collections import Counter, defaultdict

import pipeline_helpers as ph
from sources import network_map

DIV_LABEL = {'UBL': 'Hubballi', 'SBC': 'Bengaluru', 'MYS': 'Mysuru'}

_SRDEN_DETAILED_DIV_WORD = [('hubballi', 'UBL'), ('bengaluru', 'SBC'), ('mysuru', 'MYS')]


def _div_from_title(title):
    t = (title or '').lower()
    for word, div in _SRDEN_DETAILED_DIV_WORD:
        if word in t:
            return div
    return None


def _split_station_list(s):
    """'DWR, MGD, KBI, LWR' / '-' / None -> ['DWR','MGD','KBI','LWR'] / []."""
    if not s:
        return []
    out = []
    for tok in re.split(r'[,\n]+', str(s)):
        tok = tok.strip()
        if tok and tok != '-':
            out.append(tok)
    return out


def _strip_km_suffix(code):
    """'SBC (355)' -> 'SBC', 'KPN (251)' -> 'KPN'. Station codes in
    'SrDEN Detailed' occasionally carry a stray km-chainage annotation
    (confirmed: SBC (355), KPN (251), SKPT (220) in the live Bengaluru
    block) -- this is a workbook quirk, not a different station."""
    m = re.match(r'^(.*?)\s*\(\s*\d+\s*\)\s*$', str(code or '').strip())
    return m.group(1).strip() if m else str(code or '').strip()


# Manual code aliases for a handful of stations whose 'SrDEN Detailed' label
# doesn't match any RDPS/APPROX_STATIONS code directly -- confirmed by
# Shruti herself rather than guessed (e.g. 27-Nov-2026: "E UBL is BMAP").
# Add more here as she identifies them; this stays a short, explicit list
# rather than a fuzzy-matching heuristic, so every entry is traceable to
# something she actually said.
_STATION_CODE_ALIASES = {
    'E UBL': 'BMAP',
}


def _resolve_geo(code, stations_geo):
    """Exact match, then km-suffix-stripped, then no-space variant of
    each, then a known manual alias. Returns the geo dict or None (an
    honest, undecorated gap -- never guessed). Mirrors the matching used
    to build this feature's private preview (26-Nov-2026), extended with
    the km-suffix case found directly in the live data."""
    candidates = [code, code.replace(' ', '')]
    stripped = _strip_km_suffix(code)
    if stripped != code:
        candidates += [stripped, stripped.replace(' ', '')]
    alias = _STATION_CODE_ALIASES.get(code) or _STATION_CODE_ALIASES.get(stripped)
    if alias:
        candidates += [alias, alias.replace(' ', '')]
    for c in candidates:
        if c in stations_geo:
            return stations_geo[c]
    return None


def _build_stations_geo(rdps_by_div):
    """Live RDPS first; network_map.APPROX_STATIONS fills in ONLY the
    codes RDPS genuinely has no point for. Deliberately NOT using
    network_map.merge_supplement_stations() here -- that function always
    overwrites with APPROX_STATIONS regardless of whether RDPS already
    has the code, which was a safe assumption when it was written ("none
    is expected in practice") but is no longer true now that Shruti's
    own 26-Nov-2026 RDPS corrections added live points for codes (e.g.
    HPT) that also happen to have an old hardcoded APPROX_STATIONS entry
    -- the live, freshly-surveyed point should win."""
    stations = network_map.build_stations_from_rdps(rdps_by_div)
    for code, info in network_map.APPROX_STATIONS.items():
        if code not in stations:
            stations[code] = dict(lat=info['lat'], lon=info['lon'], code=code, div=None,
                                   confidence=info['confidence'], source=info['source'])
    return stations


def _parse_srden_detailed(payload, stations_geo):
    """-> list of per-station dicts (one per physical siding/RH entry),
    decoding the live generic-block-parser's key-naming scheme for this
    sheet (reverse-engineered 26-Nov-2026, confirmed against the raw
    workbook's column groups):
        TM / RGM           -> that type's siding is AVAILABLE
        TM (2) / RGM (2)   -> that type's siding is SANCTIONED
        TM (3) / RGM (3)   -> that type's siding is PROPOSED
        Available/Sanctioned/Proposed (bare) -> the REST HOUSE's own status
    Each key's value, when present, is just that status word repeated
    (confirmed: every non-null value across all these keys is the
    group's own label text) -- so presence alone is the real signal.

    Also recovers one genuine data row the live block-parser mis-filed
    under its 'total' key for the Hubballi sub-block specifically (that
    sub-block's sheet has no real trailing Total row, unlike the other
    two, so the parser's last-row-is-the-total assumption misfires;
    detected here by an empty 'Label', which every genuine Total row
    always has a real division name in)."""
    blocks = payload.get('supplementary', {}).get('SrDEN Detailed') or []
    out = []
    for block in blocks:
        div = _div_from_title(block.get('title'))
        rows = list(block.get('rows') or [])
        tot = block.get('total') or {}
        if tot.get('Station') and not (tot.get('Label') or '').strip():
            rows.append(tot)  # stray data row, not a real total (see docstring)
        section = None
        for r in rows:
            section = r.get('Section (Km)') or section
            code = str(r.get('Station') or '').strip()
            if not code:
                continue
            geo = _resolve_geo(code, stations_geo)
            out.append(dict(
                division=div, srdenRaw=r.get('Sr.DEN'),
                srdenKey=ph.canon_srden_word(div, r.get('Sr.DEN')),
                section=section, station=_strip_km_suffix(code), km=r.get('KM'),
                interSidingDistance=r.get('Inter Siding Distance'),
                sidingAvailTM=bool(r.get('TM')), sidingSancTM=bool(r.get('TM (2)')),
                sidingPropTM=bool(r.get('TM (3)')),
                sidingAvailRGM=bool(r.get('RGM')), sidingSancRGM=bool(r.get('RGM (2)')),
                sidingPropRGM=bool(r.get('RGM (3)')),
                rhAvail=bool(r.get('Available')), rhSanc=bool(r.get('Sanctioned')),
                rhProp=bool(r.get('Proposed')),
                remark=r.get('Remark'),
                lat=geo['lat'] if geo else None, lon=geo['lon'] if geo else None,
            ))
    return out


def _parse_srden_breakdown(payload, remark_by_station, strikeouts_by_div):
    """-> {div: [section dict, ...]}, merging each division's 3 live
    sub-blocks (TM siding / RGM siding / Rest House) back into one row
    per (Sl No, Major Section), and attaching each listed station's live
    remark (from 'SrDEN Detailed', via remark_by_station) plus a
    struck=True flag for any station in that division's static
    strikeout list."""
    SHEET_BY_DIV = {'UBL': 'UBL Sr.DEN', 'SBC': 'SBC Sr.DEN', 'MYS': 'MYS Sr. DEN'}
    out = {}
    for div, sheet_key in SHEET_BY_DIV.items():
        blocks = payload.get('supplementary', {}).get(sheet_key) or []
        by_slno = {}
        for block in blocks:
            title = (block.get('title') or '').lower()
            if 'rgm' in title:
                kind = 'rgm'
            elif 'rh' in title or 'rest house' in title:
                kind = 'rh'
            else:
                kind = 'tm'
            req_key = 'RH Required' if kind == 'rh' else 'Siding Required'
            for r in block.get('rows') or []:
                sl = r.get('Sl No')
                if not isinstance(sl, (int, float)):
                    continue
                sl = int(sl)
                row = by_slno.setdefault(sl, dict(
                    slNo=sl, division=div, srden=r.get('Sr.DEN'), majorSection=r.get('Major Section'),
                    sectionLength=r.get('Section Length'), stations=r.get('No. of stations'),
                    tm=None, rgm=None, rh=None,
                ))
                avail = _split_station_list(r.get('Available (2)'))
                sanc = _split_station_list(r.get('Sanctioned (2)'))
                prop = _split_station_list(r.get('Proposed (2)'))
                remarks, struck = [], []
                for st in avail + sanc + prop:
                    rm = remark_by_station.get(st)
                    if rm and rm not in remarks:
                        remarks.append(rm)
                    if st in strikeouts_by_div.get(div, ()):
                        struck.append(st)
                row[kind] = dict(
                    required=r.get(req_key), available=r.get('Available'), availableStations=avail,
                    sanctioned=r.get('Sanctioned'), sanctionedStations=sanc,
                    proposed=r.get('Proposed'), proposedStations=prop,
                    remark='; '.join(remarks) or None, struckStations=struck,
                )
        out[div] = sorted(by_slno.values(), key=lambda r: r['slNo'])
    return out


_STAGE_PATTERNS = [
    (re.compile(r'complet', re.I), 'Completed'),
    (re.compile(r'tender|estimate', re.I), 'Estimate / Tender Stage'),
    (re.compile(r'not\s*start', re.I), 'Not Started / Unknown'),
]


def _rh_work_stage(status_text):
    s = str(status_text or '')
    for pat, label in _STAGE_PATTERNS:
        if pat.search(s):
            return label
    return 'Work In Progress' if s.strip() else 'Not Started / Unknown'


def _feature_status(avail, sanc, prop):
    """One of 'available'/'sanctioned'/'proposed'/None for a single
    feature (siding, or rest house), given that feature's own
    avail/sanctioned/proposed flags. Available beats Sanctioned beats
    Proposed when more than one happens to be set for the same feature
    (shouldn't normally happen, but a real sheet occasionally does mark
    more than one stage) -- this is a priority WITHIN one feature, not a
    blend ACROSS features the way the station-level status used to be
    computed (see the 27-Nov-2026 fix below for why that distinction
    matters)."""
    if avail:
        return 'available'
    if sanc:
        return 'sanctioned'
    if prop:
        return 'proposed'
    return None


def build_infra_map(payload, rdps_by_div, supplement):
    if not payload:
        return None
    supplement = supplement or {}
    stations_geo = _build_stations_geo(rdps_by_div)
    station_rows = _parse_srden_detailed(payload, stations_geo)
    remark_by_station = {r['station']: r['remark'] for r in station_rows if r.get('remark')}
    strikeouts_by_div = {d: set(v) for d, v in (supplement.get('strikeouts') or {}).items()}
    for r in station_rows:
        r['struck'] = r['station'] in strikeouts_by_div.get(r['division'], ())
        # Fixed 27-Nov-2026 per Shruti: "sanctioned doesn't mean it is
        # available... the proposed ones mean they are not yet
        # sanctioned also." A station's SIDING status and REST HOUSE
        # status are now computed and exposed SEPARATELY, each from only
        # its own flags (TM+RGM combined for siding; rhAvail/Sanc/Prop
        # for rest house) -- never blended with each other the way the
        # frontend's old single combined per-station status used to.
        # struck is deliberately NOT read here any more: per Shruti's
        # 27-Nov-2026 "for the time being, don't show struck out on the
        # map... if it is sanctioned, show as sanctioned" -- a struck
        # station still gets its normal available/sanctioned/proposed
        # status from these flags; `struck` stays a separate field the
        # frontend only surfaces inside that station's own drawer.
        r['sidingStatus'] = _feature_status(
            r['sidingAvailTM'] or r['sidingAvailRGM'],
            r['sidingSancTM'] or r['sidingSancRGM'],
            r['sidingPropTM'] or r['sidingPropRGM'],
        )
        r['rhStatus'] = _feature_status(r['rhAvail'], r['rhSanc'], r['rhProp'])
        r['hasSiding'] = r['sidingStatus'] is not None
        r['hasRH'] = r['rhStatus'] is not None

    srden_breakdown = _parse_srden_breakdown(payload, remark_by_station, strikeouts_by_div)

    map_stations, unmapped = [], []
    for r in station_rows:
        # Fixed 27-Nov-2026 per Shruti ("I see the proposed ones you have
        # shown as zero, which is incorrect"): this gate used to omit the
        # *Prop flags entirely, so a station whose ONLY entry was
        # "Proposed" (no Available/Sanctioned flag at all, for siding OR
        # rest house) was silently dropped before ever reaching the map
        # or the unmapped list -- not shown anywhere, not even counted.
        # Proposed is a real, distinct pipeline stage she explicitly
        # wants to see and filter by, so it now counts as "has an entry"
        # too, same as Available/Sanctioned always did.
        has_entry = r['hasSiding'] or r['hasRH']
        if not has_entry:
            continue
        (map_stations if (r['lat'] is not None) else unmapped).append(r)

    # ---- Division-level sanctioned-execution rollup ----
    sanc_siding = payload.get('sanctionedSidingStatus') or []
    rh_sanc_rows = supplement.get('rhSanctionedStatus') or []
    division_rollup = {}
    for div in ('UBL', 'SBC', 'MYS'):
        siding_rows = [r for r in sanc_siding if r.get('division') == div]
        siding_stage = Counter(r.get('stage') or 'Not Started / Unknown' for r in siding_rows)
        rh_rows = [r for r in rh_sanc_rows if r.get('division') == div]
        rh_stage = Counter(_rh_work_stage(r.get('workStatus')) for r in rh_rows)
        division_rollup[div] = dict(
            sidingExecution=dict(total=len(siding_rows), byStage=dict(siding_stage)),
            rhExecution=dict(total=len(rh_rows), byStage=dict(rh_stage)),
        )

    ms = supplement.get('masterSummaryCorrected') or payload.get('masterSummary') or {}
    rb = supplement.get('rbTargets') or {}

    depots = []
    for d in (supplement.get('sanctionedDepots') or []):
        geo = _resolve_geo(d.get('station'), stations_geo)
        depots.append(dict(d, lat=geo['lat'] if geo else None, lon=geo['lon'] if geo else None))

    # ---- Track polylines, reused directly from the Network Map tab's own
    # builder (same live rdps_by_div input is already passed in here) --
    # added 26-Nov-2026 per Shruti's feature-parity request so the Infra
    # map shows the actual route alignment, not just bare station dots.
    tracks_by_div = network_map.build_tracks_from_rdps(rdps_by_div)

    return dict(
        asOf=ms.get('asOf') or payload.get('masterSummary', {}).get('asOf'),
        kpi=dict(
            stationsCovered=(ms.get('siding') or {}).get('swr', {}).get('stations'),
            siding=(ms.get('siding') or {}).get('swr'),
            restHouse=(ms.get('restHouse') or {}).get('swr'),
            sidingTargetRB=rb.get('stablingSiding'),
            satDepotRB=rb.get('satDepot'),
        ),
        mapStations=map_stations,
        unmappedStations=unmapped,
        depots=depots,
        srdenBreakdown=srden_breakdown,
        divisionRollup=division_rollup,
        divLabel=DIV_LABEL,
        tracksByDiv=tracks_by_div,
        divColor=network_map.DIV_COLOR,
    )
