"""
SWR Track Machine Dashboard -- live aggregation pipeline.

Supabase-backed replacement for claude/pipeline/run_pipeline.py's run().
Reads all 16 live-synced Supabase tables (via the sources/*.py adapter
package) instead of local Excel/Word files, applies the EXACT SAME
aggregation/rollup/reconciliation logic run_pipeline.py's
build_base_payload()/run() used -- ported verbatim wherever it has no
filesystem dependency -- and writes the same dashboard_payload.json shape
dashboard_template2.html already expects.

Two sources genuinely have NO live-sync equivalent and are gracefully
degraded here, exactly matching run_pipeline.py's own WARN-and-degrade
pattern for a missing optional local file:

  - Daily Progress shortfall-category log (run_pipeline.py's
    extract_shortfall_records). This is Task #53's confirmed-still-open
    gap -- none of the 16 live tables cover it (daily_positions and
    shortfall_reasons were both individually checked and are different
    datasets -- see their own adapter docstrings). shortfall_records=[]
    here; every shortfall-dependent payload section (shortfall.*,
    srdenBundle, the Reason-for-Shortfall tab) degrades to empty/zero,
    same as run_pipeline.py would if its local Daily Progress folder were
    empty.

  - IOH/POH current-year letter + CPOH/RYP forward-planning letter
    (run_pipeline.py's extract_ioh_poh_letter / extract_cpoh_ryp_planning).
    Both read a Word (.docx) letter with no Supabase table behind it.
    ioh_poh_letter=None, cpoh_planning=None here; Fleet Health's
    currentYear/cpohPlanning sections degrade to empty, same as
    run_pipeline.py's own "WARN: no .docx files found" path. (The
    FORECAST part of Fleet Health -- FY2027-28..2029-30 -- is NOT
    affected; that comes from ioh_poh_planning, which IS live-synced.)

The other 6 live tables (cumulative_failure, daily_positions, hq_summary,
monthly_machine, rdps_points, shortfall_reasons) have no consumption
anywhere in run_pipeline.py's run() either -- they were synced "for later
use" (see each adapter's own docstring). This rebuild still fetches and
reshapes all of them (so all 16 sources really are live in this pipeline,
per the user's explicit "full rebuild of all 16 sources" instruction),
and attaches them under the payload's own `rawSources` key rather than
inventing dashboard consumption for data no tab reads yet.

Usage:
    python3 aggregate.py <OutputDir> [--as-of YYYY-MM-DD]

Produces <OutputDir>/dashboard_payload.json. Requires SUPABASE_URL and
SUPABASE_ANON_KEY in the environment (the same two env vars every
adapter's supabase_client call already uses).
"""
import sys, os, re, json, datetime, difflib
from collections import defaultdict, Counter

import pipeline_helpers as ph
import supabase_client
from sources import (
    swr_machine_details, cumulative_progress, daily_positions, double_shift_roster,
    exception_reports, hq_summary, indent_position, ioh_poh_planning, monthly_machine,
    rdps_points, shortfall_reasons, sidings_rest_house, tenders_contracts, unit_cost,
    action_plan_53, cumulative_failure,
)

MONTH_ABBR = ph.MONTH_ABBR
MONTH_NUM = {'Apr': 4, 'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8, 'Sep': 9}
ZONE_LETTER = {'CENTRAL': 'C', 'EAST': 'E', 'WEST': 'W', 'NORTH': 'N', 'SOUTH': 'S', 'HQ': 'HQ'}

# Sourced from SWR's own Cumulative Progress workbooks' official
# Tamping/Non-Tamping grand-total banner rows, SUPERSEDED by the
# definitive, user-supplied classification of 28-Sep-2026 -- copied
# verbatim from run_pipeline.py's TAMPING_INFO. Do not revert without the
# user's explicit sign-off again (see run_pipeline.py's own comment for
# the full history).
TAMPING_INFO = {
    'CSM': ('Tamping', 'Continuous Action Tamping Machine (user-confirmed 28-Sep-2026 definitive list)'),
    'DUO': ('Tamping', 'Duomatic - dual-unit tamping machine (user-confirmed 28-Sep-2026 definitive list)'),
    'UNI': ('Tamping', 'Unimat / Universal tamping machine (points & crossings + plain track) (user-confirmed 28-Sep-2026 definitive list)'),
    'PCTM': ('Tamping', 'Points & Crossing Tamping Machine - same family as Unimat, merged with UNI'),
    'PCT': ('Tamping', 'Points & Crossing Tamping Machine - same family as Unimat, merged with UNI'),
    'MPT': ('Tamping', 'Multi-Purpose Tamper (user-confirmed 28-Sep-2026 definitive list)'),
    'DTE': ('Tamping', 'Tamping machine (user-confirmed 28-Sep-2026 definitive list)'),
    'BCM': ('Non-Tamping', 'Ballast Cleaning Machine (user-corrected 28-Sep-2026, definitive list)'),
    'FRM': ('Non-Tamping', 'Formation Rehabilitation Machine (user-corrected 28-Sep-2026, definitive list - was wrongly Tamping)'),
    'SRGM': ('Non-Tamping', 'Shoulder Ballast Cleaning/Regulating Machine (user-corrected 28-Sep-2026, definitive list - was wrongly Tamping)'),
    'RGM': ('Non-Tamping', 'Rail Grinding Machine (user-corrected 28-Sep-2026, definitive list - was wrongly Tamping)'),
    'SQRS': ('Non-Tamping', 'Quick Relay System (user-corrected 28-Sep-2026, definitive list - was wrongly Tamping)'),
    'T': ('Non-Tamping', 'T-28 Track machine (user-corrected 28-Sep-2026, definitive list - was wrongly Tamping)'),
    'PBR': ('Non-Tamping', 'Points/Plain-track Ballast Regulator (user-confirmed 28-Sep-2026 definitive list)'),
    'MDU': ('Non-Tamping', 'Maintenance / utility unit (user-confirmed 28-Sep-2026 definitive list)'),
    'RBMV': ('Non-Tamping', 'Rail-Borne Maintenance Vehicle / RBMB (user-confirmed 28-Sep-2026 definitive list)'),
    'UTV': ('Non-Tamping', 'Utility Track Vehicle (user-confirmed 28-Sep-2026 definitive list)'),
    'DGS': ('Non-Tamping', 'Dynamic Track Stabilizer (DGS/DTS) (user-confirmed 28-Sep-2026 definitive list)'),
    'BRM': ('Non-Tamping', 'Ballast Regulating Machine'),
    'WST': ('Non-Tamping', 'Special track machine (under commissioning)'),
    'RMV': ('Non-Tamping', 'RMV - maintained by the Track Machine Department but OPERATED by the divisions; '
                            'unique to SWR (not present on any other railway); the Railway Board issues no '
                            'progress target against this type, so it never appears in the Cumulative Progress '
                            'sheet, but it IS part of the SWR fleet (user-confirmed 21-Oct-2026) - added to the '
                            'fleet roster via RMV_WST_NO_TARGET_FLEET below, with NoProgressTarget=True.'),
}

# RMV/WST "no progress target" fleet additions (user-confirmed 21-Oct-2026)
# -- copied verbatim from run_pipeline.py. Pure/hardcoded, no filesystem
# dependency, so it ports unchanged.
RMV_WST_NO_TARGET_FLEET = [
    dict(machine='RMV-0332', div='MYS'),
    dict(machine='RMV-0305', div='MYS'),
    dict(machine='RMV-0288', div='MYS'),
    dict(machine='RMV-0295', div='MYS'),
    dict(machine='RMV-0296', div='MYS'),
    dict(machine='RMV-0330', div='MYS'),
    dict(machine='RMV-0329', div='MYS'),
    dict(machine='RMV-0314', div='UBL'),
    dict(machine='RMV-0315', div='UBL'),
    dict(machine='RMV-0333', div='SBC'),
    dict(machine='RMV-0304', div='SBC'),
    dict(machine='WST 01', div='SBC', status='Under Commissioning',
         note='Under Commissioning since 22-Feb-2024 (per Exception Sheet) - not yet operational, '
              'so no progress figures exist yet either.'),
]


def log(*a):
    print(*a, file=sys.stderr)


def build_rmv_wst_records(home_div, bare_type_div, norm, start_slno):
    """Builds synthetic cum_records-shaped entries for RMV_WST_NO_TARGET_FLEET,
    laid out identically to a real extract_cum_progress() record (same keys)
    so they flow through every downstream aggregation (fleetSize, div/mtype
    rollups, statusCounts, progress.machines) exactly like real data, just
    with every progress/target field left None. Copied verbatim from
    run_pipeline.py."""
    out = []
    for i, entry in enumerate(RMV_WST_NO_TARGET_FLEET):
        machine = entry['machine']
        mtype_match = re.match(r'^([A-Za-z]+)', machine)
        mtype = mtype_match.group(1).upper() if mtype_match else 'UNKNOWN'
        div = entry['div']
        status = entry.get('status', 'Active')
        rec = {
            'SlNo': start_slno + i, 'Machine': machine, 'MachineType': mtype,
            'CurrentDIV': div, 'HomeDivision': home_div.get(norm(machine)) or div,
            'Status': status,
            'Unit': None, 'AnnualTarget': None, 'TargetPerMonth': None,
            'ProportionateTarget': None, 'ActualProg': None, 'PctProg': None,
            'NoProgressTarget': True,
            'NoProgressTargetNote': entry.get(
                'note',
                'No progress target is issued by the Railway Board for this machine type (RMV) - '
                'maintained by the Track Machine Department but operated by the divisions; unique to '
                'SWR. Part of the SWR fleet, but excluded from Target-vs-Progress figures by design.'),
        }
        for abbr in MONTH_ABBR:
            rec[abbr] = None
        out.append(rec)
    log(f"RMV/WST no-progress-target fleet: {len(out)} machines added "
        f"({sum(1 for e in RMV_WST_NO_TARGET_FLEET if e['machine'].startswith('RMV'))} RMV, "
        f"{sum(1 for e in RMV_WST_NO_TARGET_FLEET if e['machine'].startswith('WST'))} WST).")
    return out


# ---------------------------------------------------------------------------
# Small pure helpers used by build_base_payload()/run() with no filesystem
# dependency -- copied verbatim from run_pipeline.py.
# ---------------------------------------------------------------------------
def _num_or_none(v):
    return v if isinstance(v, (int, float)) else None


def _cum_as_of_date(cum_records, fy_start_year):
    """Derives the real 'as of' date the Cumulative Progress sheet's own
    ProportionateTarget column is computed through, from content rather
    than any filename/title text. Copied verbatim from run_pipeline.py."""
    month_lengths = {4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30,
                      12: 31, 1: 31, 2: 29 if (fy_start_year + 1) % 4 == 0 and
                      ((fy_start_year + 1) % 100 != 0 or (fy_start_year + 1) % 400 == 0) else 28, 3: 31}
    month_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
    samples = []
    for r in cum_records:
        tm = r.get('TargetPerMonth')
        pt = r.get('ProportionateTarget')
        if isinstance(tm, (int, float)) and tm > 0 and isinstance(pt, (int, float)) and pt > 0:
            samples.append(pt / tm)
    if not samples:
        return None
    samples.sort()
    elapsed_months = samples[len(samples) // 2]
    full_months = max(0, min(11, int(elapsed_months)))
    frac = max(0.0, elapsed_months - full_months)
    cur_month = month_order[full_months]
    cur_year = fy_start_year if full_months <= 8 else fy_start_year + 1
    day_in_month = max(1, min(month_lengths[cur_month], round(frac * month_lengths[cur_month]) or 1))
    try:
        return datetime.date(cur_year, cur_month, day_in_month)
    except ValueError:
        return None


def _secondary_dict(sec):
    """Builds the client-facing `secondary` sub-object (BCM's Turnout/T/O
    line). Copied verbatim from run_pipeline.py."""
    if not sec:
        return None
    pct = (sec['ActualProg'] / sec['ProportionateTarget']) if sec.get('ProportionateTarget') else 0
    out = dict(unit=sec['Unit'], annualTarget=sec.get('AnnualTarget'), targetMonth=sec.get('TargetPerMonth'),
               targetToDate=sec.get('ProportionateTarget'), actualToDate=sec.get('ActualProg'), pct=round(pct, 4))
    for abbr in MONTH_ABBR:
        out[abbr.lower()] = sec.get(abbr)
    return out


def canon_ds_srden(div, zone):
    if not div or not zone:
        return None
    zl = ZONE_LETTER.get(str(zone).strip().upper())
    if not zl:
        return None
    return f"{div} Sr.DEN/{zl}"


_IOHPOH_MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'april': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'sept': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def parse_iohpoh_month_sort(s):
    """Best-effort chronological sort key for the letter's free-text month
    column. Copied verbatim from run_pipeline.py (kept here even though the
    letter itself is not live-synced, since ioh_poh_planning's own forecast
    events reuse this same sort key)."""
    if not s:
        return (9999, 99)
    low = str(s).lower()
    m = re.search(r'([a-z]{3,5})[.\-\s]*(\d{2,4})', low)
    if not m:
        return (9999, 99)
    mon = _IOHPOH_MONTH_MAP.get(m.group(1)[:4]) or _IOHPOH_MONTH_MAP.get(m.group(1)[:3])
    if mon is None:
        return (9999, 99)
    yr = int(m.group(2))
    if yr < 100:
        yr += 2000
    return (yr, mon)


def classify_iohpoh_status(remark):
    """Copied verbatim from run_pipeline.py -- retained for forward
    compatibility if a future live source ever carries this free-text
    remark; not currently exercised since ioh_poh_letter is always None
    here (see module docstring)."""
    s = str(remark or '').strip()
    if not s or s == '-':
        return 'Scheduled'
    low = s.lower()
    if 'completed' in low:
        return 'Completed'
    if 'yet to release' in low:
        return 'Yet to Release from Division'
    if 'overdue' in low:
        return 'Overdue for Release'
    if re.search(r'under\s*(ioh|poh)\s*attention|under\s*ioh|under\s*poh', low):
        return 'Ongoing / Under Attention'
    if 'due for' in low:
        return 'Due for Release'
    return s


def extract_overdue_months(remark):
    m = re.search(r'overdue\s*by\s*(\d+(?:\.\d+)?)\s*month', str(remark or ''), re.I)
    return float(m.group(1)) if m else None


# =====================================================================
# Supabase-backed extract_* wrappers.
#
# Each one carries the EXACT same name and call signature as its
# filesystem-backed counterpart in run_pipeline.py's run() (including the
# unused `base` parameter, kept purely so run()'s own call sites below
# needed zero changes when ported) -- only its body differs: a Supabase
# fetch instead of an openpyxl/docx read. `base` is accepted and ignored
# everywhere except as a readability marker of "this used to be a folder
# path, now it's just a placeholder".
#
# A payload a module's fetch_payload() might be asked for twice in one
# run() (e.g. extract_home_div/extract_age_profile both need
# swr_machine_details; extract_ioh_poh/extract_ioh_poh_duration_norms
# both need ioh_poh_planning) is cached here so each table is fetched
# over the network exactly once per run, not once per extract_* call.
# =====================================================================
_PAYLOAD_CACHE = {}


def _cached_payload(module):
    name = module.TABLE
    if name not in _PAYLOAD_CACHE:
        payload, updated_at = module.fetch_payload()
        _PAYLOAD_CACHE[name] = payload
        log(f"Fetched '{name}' (updated_at={updated_at})")
    return _PAYLOAD_CACHE[name]


def reset_payload_cache():
    """Tests / repeated runs in one process call this between runs so a
    stale payload from an earlier call never leaks into a later one."""
    _PAYLOAD_CACHE.clear()


def extract_home_div(base=None):
    payload = _cached_payload(swr_machine_details)
    return swr_machine_details.extract_home_div(payload)


def extract_age_profile(base=None):
    payload = _cached_payload(swr_machine_details)
    return swr_machine_details.extract_age_profile(payload)


def extract_cum_progress(base, home_div, bare_type_div, norm):
    payload = _cached_payload(cumulative_progress)
    return cumulative_progress.extract_cum_progress(payload, home_div, bare_type_div, norm)


def extract_prior_year_by_type(base=None):
    payload = _cached_payload(cumulative_progress)
    return cumulative_progress.extract_prior_year_by_type(payload)


def extract_shortfall_records(base=None):
    """Task #53's confirmed-still-open gap -- no live table covers the
    Daily Progress shortfall-category log (see module docstring). Returns
    [] so every downstream shortfall.*/srdenBundle/Reason-for-Shortfall
    section degrades to empty, exactly like run_pipeline.py's own
    behaviour when its local 'Daily Progress' folder is empty."""
    log("WARN: no live source for the Daily Progress shortfall-category log "
        "(Task #53, confirmed still open) - shortfall_records will be empty; "
        "shortfall.*/srdenBundle/Reason-for-Shortfall degrade to zero.")
    return []


def extract_exceptions(base=None):
    all_rows = exception_reports.fetch_all()
    return exception_reports.extract_exceptions(all_rows)


def extract_ioh_poh(base=None):
    payload = _cached_payload(ioh_poh_planning)
    return ioh_poh_planning.extract_ioh_poh(payload)


def extract_ioh_poh_duration_norms(base=None):
    payload = _cached_payload(ioh_poh_planning)
    return ioh_poh_planning.extract_ioh_poh_duration_norms(payload)


def extract_ioh_poh_letter(base, as_of):
    """No live table backs the current-year IOH/POH Dy.CE/TM office letter
    (a Word .docx with no Supabase equivalent) - degrades to None, same as
    run_pipeline.py's own 'WARN: no .docx files found' path. Fleet
    Health's futureYears (FY2027-28..2029-30, from ioh_poh_planning, which
    IS live) is unaffected; only currentYear degrades."""
    log("WARN: no live source for the IOH/POH current-year office letter "
        "(Word .docx, no Supabase table) - ioh_poh_letter will be None; "
        "Fleet Health's currentYear section degrades to empty.")
    return None


def extract_cpoh_ryp_planning(base=None):
    """No live table backs the CPOH/RYP forward POH-planning letter
    (also a Word .docx) - degrades to None, same reasoning as
    extract_ioh_poh_letter above."""
    log("WARN: no live source for the CPOH/RYP forward-planning letter "
        "(Word .docx, no Supabase table) - cpoh_planning will be None.")
    return None


def extract_double_shift(base=None):
    payload = _cached_payload(double_shift_roster)
    return double_shift_roster.extract_double_shift(payload)


def extract_action_plan_53(base=None):
    payload = _cached_payload(action_plan_53)
    return action_plan_53.extract_action_plan_53(payload)


def extract_sidings_rest_houses(base=None):
    payload = _cached_payload(sidings_rest_house)
    return sidings_rest_house.extract_sidings_rest_houses(payload)


def extract_tenders_contracts(base, as_of):
    payload = _cached_payload(tenders_contracts)
    return tenders_contracts.extract_tenders_contracts(payload, as_of)


def extract_indent_position(base, as_of):
    payload = _cached_payload(indent_position)
    return indent_position.extract_indent_position(payload, as_of)


def extract_unit_cost(base=None):
    payload = _cached_payload(unit_cost)
    return unit_cost.extract_unit_cost(payload)


# ---------------------------------------------------------------------
# The 6 "for later use" live sources (see module docstring) -- fetched
# and reshaped so all 16 tables really are live in this rebuild, but not
# consumed by any existing dashboard tab. Collected into one dict and
# attached to the payload under `rawSources`, exactly mirroring each
# adapter's own "kept as a thin, honest pass-through" framing.
# ---------------------------------------------------------------------
def extract_raw_sources():
    out = {}
    try:
        out['cumulativeFailure'] = cumulative_failure.extract_cumulative_failure(
            _cached_payload(cumulative_failure))
    except Exception as e:
        log(f"WARN: cumulative_failure fetch/extract failed ({e}) - omitted from rawSources.")
    try:
        out['dailyPositions'] = daily_positions.extract_positions(_cached_payload(daily_positions))
    except Exception as e:
        log(f"WARN: daily_positions fetch/extract failed ({e}) - omitted from rawSources.")
    try:
        out['hqSummary'] = hq_summary.extract_hq_summary(_cached_payload(hq_summary))
    except Exception as e:
        log(f"WARN: hq_summary fetch/extract failed ({e}) - omitted from rawSources.")
    try:
        out['monthlyMachine'] = monthly_machine.extract_monthly_machine(_cached_payload(monthly_machine))
    except Exception as e:
        log(f"WARN: monthly_machine fetch/extract failed ({e}) - omitted from rawSources.")
    try:
        out['rdpsPoints'] = rdps_points.extract_points(rdps_points.fetch_rows())
    except Exception as e:
        log(f"WARN: rdps_points fetch/extract failed ({e}) - omitted from rawSources.")
    try:
        out['shortfallReasonsRaw'] = shortfall_reasons.extract_shortfall_reasons(
            _cached_payload(shortfall_reasons))
    except Exception as e:
        log(f"WARN: shortfall_reasons fetch/extract failed ({e}) - omitted from rawSources.")
    log(f"rawSources (for-later-use tables): {len(out)}/6 fetched successfully.")
    return out


# =====================================================================
# AGGREGATION: base payload skeleton -- copied VERBATIM from
# run_pipeline.py's build_base_payload(). It has no filesystem dependency
# at all (it only ever touches the already-extracted cum_records/
# shortfall_records/exception_data/age_profile/as_of arguments), so it
# ports unchanged.
# =====================================================================
def build_base_payload(cum_records, shortfall_records, exception_data, age_profile, as_of):
    fy_year = as_of.year if as_of.month >= 4 else as_of.year - 1
    cum_as_of = _cum_as_of_date(cum_records, fy_year)
    if cum_as_of is not None:
        if cum_as_of.month == 9 and cum_as_of.day >= 28:
            cum_period_label = f"Apr–Sep {cum_as_of.year}"
        else:
            cum_period_label = f"Apr – {cum_as_of.strftime('%-d %b')} {cum_as_of.year}"
    else:
        cum_period_label = f"{MONTH_ABBR[0]}–{MONTH_ABBR[-1]} {as_of.year}"
    meta = dict(
        asOf=as_of.strftime('%d-%b-%Y'),
        fy=f"FY {fy_year}-{str(fy_year + 1)[-2:]}",
        periodCovered=cum_period_label,
        cumAsOfDate=(cum_as_of.strftime('%Y-%m-%d') if cum_as_of else None),
        fleetSize=len(cum_records),
        statusCounts=dict(Counter(r['Status'] for r in cum_records)),
    )

    def rollup(records, keyfn):
        by = defaultdict(lambda: dict(count=0, annualTarget=0.0, targetToDate=0.0, actualToDate=0.0))
        for r in records:
            k = keyfn(r)
            d = by[k]
            d['count'] += 1
            d['annualTarget'] += (r.get('AnnualTarget') or 0)
            d['targetToDate'] += (r.get('ProportionateTarget') or 0)
            d['actualToDate'] += (r.get('ActualProg') or 0)
        out = []
        for k, d in sorted(by.items(), key=lambda x: str(x[0])):
            pct = (d['actualToDate'] / d['targetToDate']) if d['targetToDate'] else 0
            out.append(dict(count=d['count'], annualTarget=round(d['annualTarget'], 1),
                             targetToDate=round(d['targetToDate'], 1), actualToDate=round(d['actualToDate'], 1),
                             pct=round(pct, 4), **{'div' if keyfn is div_key else 'type': k}))
        return out

    def div_key(r):
        return r['HomeDivision']

    div_rollup = rollup(cum_records, div_key)
    mtype_rollup = rollup(cum_records, lambda r: r['MachineType'])

    secondary_cum_records = []
    for r in cum_records:
        sec = r.get('Secondary')
        if sec:
            secondary_cum_records.append({
                'MachineType': f"{r['MachineType']} ({sec['Unit']})",
                'AnnualTarget': sec.get('AnnualTarget'),
                'ProportionateTarget': sec.get('ProportionateTarget'),
                'ActualProg': sec.get('ActualProg'),
                '_unit': sec['Unit'],
            })
    if secondary_cum_records:
        sec_rollup = rollup(secondary_cum_records, lambda r: r['MachineType'])
        unit_by_pseudo_type = {r['MachineType']: r['_unit'] for r in secondary_cum_records}
        for row in sec_rollup:
            row['unit'] = unit_by_pseudo_type.get(row['type'])
            row['isSecondaryMetric'] = True
        mtype_rollup = mtype_rollup + sec_rollup
        log(f"Cumulative Progress: {len(secondary_cum_records)} secondary-unit records rolled up into "
            f"{len(sec_rollup)} pseudo-type row(s): {', '.join(r['type'] for r in sec_rollup)}")

    machines_base = []
    for r in cum_records:
        d = r['CurrentDIV'] if r['CurrentDIV'] in ('MYS', 'SBC', 'UBL') else r['HomeDivision']
        m = dict(type=r['MachineType'], machine=r['Machine'], div=d, status=r['Status'], unit=r['Unit'],
                  annualTarget=r['AnnualTarget'], targetMonth=r['TargetPerMonth'])
        for abbr in MONTH_ABBR:
            m[abbr.lower()] = r.get(abbr)
        m['targetToDate'] = r['ProportionateTarget']
        m['actualToDate'] = r['ActualProg']
        m['pct'] = r['PctProg']
        m['secondary'] = _secondary_dict(r.get('Secondary'))
        machines_base.append(m)

    progress = dict(machines=machines_base, divRollup=div_rollup, mtypeRollup=mtype_rollup)

    cat_counts = Counter(r['Category'] for r in shortfall_records if r.get('Category') and r.get('Machine'))
    by_month_cat = defaultdict(Counter)
    by_div_cat = defaultdict(Counter)
    by_mtype_cat = defaultdict(Counter)
    machine_counts = Counter()

    def norm_div_for_shortfall(v):
        if v in ('MYS', 'SBC', 'UBL'):
            return v
        if not v:
            return 'Unknown'
        s = str(v).strip().upper()
        if 'KONKAN' in s or s == 'KRCL':
            return 'Deputed-KRCL'
        if 'NWR' in s or 'NORTH WESTERN' in s:
            return 'Deputed-NWR'
        if s in ('CR', 'C.RLY') or 'CENTRAL RLY' in s or 'CENTRAL RAILWAY' in s or 'C.RLY' in s:
            return 'Deputed-C.Rly'
        return 'Unknown'

    for r in shortfall_records:
        cat = r.get('Category')
        if not cat or not r.get('Machine'):
            continue
        by_month_cat[r.get('Month') or 'Unknown'][cat] += 1
        by_div_cat[norm_div_for_shortfall(r.get('Division'))][cat] += 1
        if r.get('Machine'):
            mt = ph.disp_type(ph.mtype_of(r['Machine']))
            by_mtype_cat[mt][cat] += 1
            if cat != 'Productive Work - Full Progress':
                machine_counts[r['Machine']] += 1
    top_machines = machine_counts.most_common(15)

    shortfall = dict(
        catCounts=dict(cat_counts),
        byMonthCat={k: dict(v) for k, v in by_month_cat.items()},
        byDivCat={k: dict(v) for k, v in by_div_cat.items()},
        byMtypeCat={k: dict(v) for k, v in by_mtype_cat.items()},
        topMachines=[list(x) for x in top_machines],
        totalEntries=sum(cat_counts.values()),
    )

    days_sorted = sorted(exception_data.keys(),
                          key=lambda k: datetime.datetime.strptime(k + f"-{as_of.year}", '%d-%b-%Y'))
    trend = [dict(day=d, count=len(exception_data[d])) for d in days_sorted]
    latest_day = days_sorted[-1] if days_sorted else None
    latest_rows = exception_data.get(latest_day, []) if latest_day else []
    latest_cat_counts = Counter(ph.classify_exc_category(r['Category'])[1] for r in latest_rows if r.get('Category'))

    def _fmt_day(d):
        return datetime.datetime.strptime(d + f"-{as_of.year}", '%d-%b-%Y').strftime('%d-%b-%Y') if d else None
    latest_day_label = _fmt_day(latest_day)
    trend_range_label = (f"{days_sorted[0]}–{_fmt_day(days_sorted[-1])}" if len(days_sorted) > 1
                          else (latest_day_label or ''))

    exceptions = dict(trend=trend, latest=latest_rows, latestCatCounts=dict(latest_cat_counts),
                       latestDay=latest_day, latestDayLabel=latest_day_label, trendRangeLabel=trend_range_label)

    def age_bucket(age):
        if age is None:
            return None
        if age <= 10:
            return '0-10 yrs'
        if age <= 20:
            return '11-20 yrs'
        if age <= 30:
            return '21-30 yrs'
        return '30+ yrs'

    age_buckets = Counter(age_bucket(a.get('AgeYears')) for a in age_profile if age_bucket(a.get('AgeYears')))
    fleet = dict(ageBuckets=dict(age_buckets))

    return dict(meta=meta, progress=progress, shortfall=shortfall, exceptions=exceptions, fleet=fleet), latest_day


def run(outdir, as_of=None):
    """Supabase-backed port of run_pipeline.py's run(base, outdir). Every
    line below this point that isn't the extraction block or the one
    `iohPohSchedule.futureYears.source` line is copied VERBATIM from
    run_pipeline.py's run() -- it only ever operates on the
    already-extracted record lists/dicts above, which are shape-identical
    regardless of whether they came from a local file or a live Supabase
    table, so the aggregation/rollup/reconciliation logic needed no
    changes at all.

    `base` is kept as a local dummy below (always None) purely so every
    extract_*(base) call site reads identically to its run_pipeline.py
    counterpart -- a deliberate port-fidelity choice, not a leftover.
    """
    reset_payload_cache()
    os.makedirs(outdir, exist_ok=True)
    as_of = as_of or datetime.date.today()
    base = None  # no local folder in this Supabase-backed pipeline -- every extract_*(base) call below ignores it

    home_div, bare_type_div, norm = extract_home_div(base)
    cum_records = extract_cum_progress(base, home_div, bare_type_div, norm)
    # RMV/WST "no progress target" fleet addition - user-confirmed 21-Oct-2026.
    cum_records = cum_records + build_rmv_wst_records(
        home_div, bare_type_div, norm,
        start_slno=max((r['SlNo'] for r in cum_records), default=0) + 1)
    age_profile = extract_age_profile(base)
    shortfall_records = extract_shortfall_records(base)
    exception_data = extract_exceptions(base)
    ioh_poh_records = extract_ioh_poh(base)
    ioh_poh_norms = extract_ioh_poh_duration_norms(base)
    ioh_poh_letter = extract_ioh_poh_letter(base, as_of)
    cpoh_planning = extract_cpoh_ryp_planning(base)
    ds_records = extract_double_shift(base)
    action_plan_53 = extract_action_plan_53(base)
    infra = extract_sidings_rest_houses(base)
    tenders_contracts = extract_tenders_contracts(base, as_of)
    indent_position = extract_indent_position(base, as_of)
    unit_cost = extract_unit_cost(base)
    raw_sources = extract_raw_sources()

    # Data-quality fix for ioh_poh_records' Division field: backfill from
    # the Cumulative Progress home-division map wherever the source
    # sheet didn't give a clean MYS/SBC/UBL value.
    for r in ioh_poh_records:
        if not r.get('Division'):
            hd = home_div.get(norm(r['Machine']))
            r['Division'] = hd or r.get('DivisionRaw') or 'Unknown'

    p, latest_ex_day = build_base_payload(cum_records, shortfall_records, exception_data, age_profile, as_of)
    p['actionPlan53'] = action_plan_53
    if infra:
        p['infra'] = infra
        log("Infra (Sidings/Rest House) added to payload.")
    else:
        log("WARN: 'infra' key omitted from payload (extraction returned nothing).")
    if tenders_contracts:
        p['tendersContracts'] = tenders_contracts
        log("Tenders & Contracts added to payload.")
    else:
        log("WARN: 'tendersContracts' key omitted from payload (extraction returned nothing).")
    if indent_position:
        p['indentPosition'] = indent_position
        log("Indent Position added to payload.")
    else:
        log("WARN: 'indentPosition' key omitted from payload (extraction returned nothing).")
    if unit_cost:
        p['unitCost'] = unit_cost
        log(f"Unit Cost added to payload ({len(unit_cost['years'])} year(s), latest {unit_cost['latestFy']}).")
    else:
        log("WARN: 'unitCost' key omitted from payload (extraction returned nothing).")
    if raw_sources:
        p['rawSources'] = raw_sources
        log(f"rawSources added to payload ({len(raw_sources)} for-later-use table(s)).")

    # ---- Sr.DEN rollups (build_srden_and_fixes.py equivalent) ----
    MONTH_ORDER = MONTH_NUM

    def parse_date(d):
        if not d:
            return None
        d = str(d).strip().split(' ')[0]
        try:
            return datetime.datetime.strptime(d, '%d.%m.%Y')
        except Exception:
            return None

    machine_latest = {}
    for r in shortfall_records:
        if not r.get('SrDENKey'):
            continue
        nm = ph.norm_machine(r['Machine'])
        if not nm:
            continue
        nm = ph.canon_machine_alias(nm)
        dt = parse_date(r['Date'])
        sortkey = (MONTH_ORDER.get(r['Month'], 0), dt or datetime.datetime.min)
        if nm not in machine_latest or sortkey > machine_latest[nm][0]:
            machine_latest[nm] = (sortkey, r['SrDENKey'])

    _key_groups = defaultdict(list)
    for nm, (sortkey, k) in machine_latest.items():
        _key_groups[ph.machine_key(nm)].append((nm, sortkey, k))
    _ambiguous_keys = {}
    for key, entries in _key_groups.items():
        if len(entries) < 2:
            continue
        if len({k for _, _, k in entries}) == 1:
            keep_nm = max(entries, key=lambda e: e[1])[0]
            for nm, _, _ in entries:
                if nm != keep_nm:
                    del machine_latest[nm]
        else:
            _ambiguous_keys[key] = entries
    if _ambiguous_keys:
        log(f"WARN: {len(_ambiguous_keys)} machine spelling(s) map to the same machine_key "
            f"but were logged against DIFFERENT Sr.DEN jurisdictions at different times - "
            f"left unmerged, not auto-attributed to either: " +
            "; ".join(f"{key}: " + ", ".join(f"{nm}->{k}" for nm, _, k in entries)
                      for key, entries in _ambiguous_keys.items()))

    machine_to_srden = {k: v[1] for k, v in machine_latest.items()}

    def _unambiguous_fallback(keyfn):
        groups = defaultdict(set)
        for nm, k in machine_to_srden.items():
            groups[keyfn(nm)].add(k)
        return {key: next(iter(ks)) for key, ks in groups.items() if len(ks) == 1}

    machine_to_srden_alnum = _unambiguous_fallback(ph.alnum)
    machine_to_srden_key = _unambiguous_fallback(ph.machine_key_aliased)

    def lookup_srden(machine_name):
        nm = ph.canon_machine_alias(ph.norm_machine(machine_name))
        if nm in machine_to_srden:
            return machine_to_srden[nm]
        a = machine_to_srden_alnum.get(ph.alnum(machine_name))
        if a:
            return a
        return machine_to_srden_key.get(ph.machine_key_aliased(machine_name))

    latest_month = max((r['Month'] for r in shortfall_records if r.get('SrDENKey')),
                        key=lambda m: MONTH_ORDER.get(m, 0), default=None)
    srden_month_cat = defaultdict(Counter)
    srden_machines = defaultdict(set)
    for nm, k in machine_to_srden.items():
        srden_machines[k].add(nm)
    srden_month_machine_cat = defaultdict(lambda: defaultdict(Counter))
    for r in shortfall_records:
        k = r.get('SrDENKey')
        if not k:
            continue
        if r['Month'] == latest_month:
            srden_month_cat[k][r['Category']] += 1
            if r['Category'] != 'Productive Work - Full Progress' and r.get('Machine'):
                srden_month_machine_cat[k][ph.norm_machine(r['Machine'])][r['Category']] += 1

    srden_exceptions = defaultdict(list)
    if latest_ex_day:
        for row in exception_data.get(latest_ex_day, []):
            nm = ph.norm_machine(row['Machine'])
            k = machine_to_srden.get(nm)
            srden_exceptions[k or f"UNMAPPED|{row.get('Division')}"].append(row)

    # ---- progress_fixed: tamping classification + status correction + IOH/POH dates ----
    progress_fixed = []
    for r in cum_records:
        r = dict(r)
        mt = r.get('MachineType')
        info = TAMPING_INFO.get(mt, ('Non-Tamping', 'Unclassified'))
        r['MachineTypeDisplay'] = ph.disp_type(mt)
        r['TampingCategory'] = info[0]
        r['TampingNote'] = info[1]
        if r.get('Status') == 'Deputed-NWR':
            k = lookup_srden(r['Machine'])
            if k and k.startswith('UBL'):
                r['StatusCorrectedFrom'] = 'Deputed-NWR'
                r['Status'] = 'Active'
                r['CurrentDIV'] = 'UBL'
                r['StatusCorrected'] = True
                r['StatusCorrectionNote'] = ('Cumulative Progress sheet showed Deputed-NWR, but Daily Progress '
                                              'confirms the machine has returned and is working in UBL division '
                                              '- corrected.')
        progress_fixed.append(r)

    ioh_poh_live = {}
    if latest_ex_day:
        for row in exception_data.get(latest_ex_day, []):
            if row.get('Category') == 'M/c under IOH/POH':
                ioh_poh_live[ph.machine_key(row['Machine'])] = row
    for r in progress_fixed:
        if r.get('Status') in ('IOH', 'POH'):
            hit = ioh_poh_live.get(ph.machine_key(r['Machine']))
            if hit:
                r['IOHPOH_StartDate'] = hit.get('UnderSince')
                r['IOHPOH_ExpectedCompletion'] = hit.get('Tentative')
                r['IOHPOH_DaysUnder'] = hit.get('DaysUnder')
                r['IOHPOH_Location'] = hit.get('Location')

    # ---- Status correction pass 2: stale IOH/POH flags ----
    STALE_IOHPOH_WINDOW_DAYS = 60
    _ioh_poh_by_machine_early = {ph.machine_key(r['Machine']): r for r in ioh_poh_records}
    for r in progress_fixed:
        if r.get('Status') in ('IOH', 'POH') and ph.machine_key(r['Machine']) not in ioh_poh_live:
            planning_rec = _ioh_poh_by_machine_early.get(ph.machine_key(r['Machine']))
            if not planning_rec:
                continue
            key = 'LastPOH_Date' if r['Status'] == 'POH' else 'LastIOH_Date'
            completion_dt = ph.parse_ddmmyyyy(planning_rec.get(key))
            if completion_dt and 0 <= (as_of - completion_dt.date()).days <= STALE_IOHPOH_WINDOW_DAYS:
                old_status = r['Status']
                r['StatusCorrectedFrom'] = old_status
                r['Status'] = 'Active'
                r['CurrentDIV'] = planning_rec.get('Division') or r.get('CurrentDIV')
                r['StatusCorrected'] = True
                r['StatusCorrectionNote'] = (
                    f"Cumulative Progress Status showed {old_status}, but IOH/POH Planning sheet's "
                    f"Last {old_status} Date ({completion_dt.strftime('%d-%b-%Y')}) confirms it completed, "
                    f"and it is absent from today's Exception Sheet - corrected to Active.")
        progress_fixed_item = r  # no-op, keeps loop body non-trivial for readability

    # ---- IOH/POH reconciled: latest Exception Sheet is authoritative ----
    ioh_poh_by_machine = {ph.machine_key(r['Machine']): r for r in ioh_poh_records}
    cp_iohpoh = {ph.machine_key(r['Machine']): r for r in cum_records if r['Status'] in ('IOH', 'POH')}
    ex_iohpoh = {ph.machine_key(row['Machine']): row for row in exception_data.get(latest_ex_day, [])
                 if row.get('Category') == 'M/c under IOH/POH'} if latest_ex_day else {}

    def _last_completion(planning_rec, status):
        if not planning_rec:
            return None
        key = 'LastPOH_Date' if status == 'POH' else 'LastIOH_Date'
        return ph.parse_ddmmyyyy(planning_rec.get(key))

    iohpoh_reconciled = []
    iohpoh_stale = []
    iohpoh_cp_only_unconfirmed = []
    for k in sorted(set(cp_iohpoh) | set(ex_iohpoh)):
        cpr, exr = cp_iohpoh.get(k), ex_iohpoh.get(k)
        name = (cpr or {}).get('Machine') or (exr or {}).get('Machine')
        if exr:
            sources = ['Exception Sheet (latest snapshot)'] + (['Cumulative Progress (Status field)'] if cpr else [])
            iohpoh_reconciled.append(dict(
                Machine=name, Status=(cpr or {}).get('Status') or 'IOH/POH',
                Division=(cpr or {}).get('CurrentDIV') or exr.get('Division'),
                StartDate=exr.get('UnderSince'), ExpectedCompletion=exr.get('Tentative'),
                DaysUnder=exr.get('DaysUnder'), Sources=sources,
            ))
            continue
        planning_rec = ioh_poh_by_machine.get(k)
        completion_dt = _last_completion(planning_rec, cpr.get('Status'))
        if completion_dt and 0 <= (as_of - completion_dt.date()).days <= STALE_IOHPOH_WINDOW_DAYS:
            iohpoh_stale.append(dict(
                Machine=name, Status=cpr.get('Status'), Division=cpr.get('CurrentDIV'),
                CompletionDate=completion_dt.strftime('%d.%m.%Y'),
                Note=(f"Cumulative Progress Status still shows {cpr.get('Status')}, but IOH/POH Planning "
                      f"sheet's Last {cpr.get('Status')} Date is {completion_dt.strftime('%d-%b-%Y')} "
                      f"(on/before as-of {as_of.strftime('%d-%b-%Y')}) and it is absent from today's "
                      f"Exception Sheet - treated as completed, excluded from the live IOH/POH list."),
            ))
            continue
        iohpoh_cp_only_unconfirmed.append(dict(
            Machine=name, Status=cpr.get('Status'), Division=cpr.get('CurrentDIV'),
            Note=(f"Cumulative Progress Status shows {cpr.get('Status')}, but the machine is not in "
                  f"today's ({as_of.strftime('%d-%b-%Y')}) Exception Sheet and there is no IOH/POH "
                  f"Planning completion date confirming it finished. The Exception Sheet is authoritative "
                  f"for 'currently under IOH/POH', so this machine is excluded from that list/count until "
                  f"it either reappears in the Exception Sheet or shows a completion date."),
        ))
    iohpoh_reconciled.sort(key=lambda r: r['Machine'] or '')
    iohpoh_stale.sort(key=lambda r: r['Machine'] or '')
    iohpoh_cp_only_unconfirmed.sort(key=lambda r: r['Machine'] or '')

    # ---- Sr.DEN messages (longest-pending-first) ----
    def fmt_list(items, n=3):
        return ', '.join(items[:n]) + (f' +{len(items)-n} more' if len(items) > n else '')

    srden_bundle = {}
    for k in sorted(srden_month_cat.keys()):
        cats = srden_month_cat[k]
        top_cats = sorted(cats.items(), key=lambda x: -x[1])[:5]
        n_machines = len(srden_machines.get(k, []))
        exc_rows = srden_exceptions.get(k, [])
        exc_by_type_rows = defaultdict(list)
        for row in exc_rows:
            t, label = ph.classify_exc_category(row['Category'])
            exc_by_type_rows[label].append(row)
        for label in exc_by_type_rows:
            exc_by_type_rows[label].sort(key=lambda r: -(r.get('DaysUnder') or 0))
        label_order = sorted(exc_by_type_rows.keys(),
                              key=lambda lbl: -max((r.get('DaysUnder') or 0) for r in exc_by_type_rows[lbl]))
        exc_by_type = {lbl: [r['Machine'] for r in rows] for lbl, rows in exc_by_type_rows.items()}

        div, code = k.split(' Sr.DEN/')
        lines = [f"*Track Machine Update — {k} — {as_of.strftime('%d-%b-%Y')}*",
                 f"({n_machines} machines working in this jurisdiction)", ""]
        if exc_by_type_rows:
            lines.append("*Machines needing your attention today (longest-pending first):*")
            for label in label_order:
                rows = exc_by_type_rows[label]
                lines.append(f"- {label}:")
                for r in rows:
                    days = r.get('DaysUnder')
                    days_txt = f"{days}d" if isinstance(days, (int, float)) else "—"
                    where = r.get('Location') or r.get('Division') or ''
                    lines.append(f"    • {r['Machine']} — {days_txt} pending" + (f" ({where})" if where else ""))
            lines.append("")
        machine_cats = srden_month_machine_cat.get(k, {})
        machine_shortfalls = sorted(machine_cats.items(), key=lambda kv: -sum(kv[1].values()))
        open_exc_machines = {ph.norm_machine(r['Machine']) for r in exc_rows}
        message_shortfalls = [(m, c) for m, c in machine_shortfalls if m in open_exc_machines]
        if message_shortfalls:
            lines.append(f"*Machines with shortfall reasons this month ({latest_month}), open exceptions only:*")
            for machine, cat_counter in message_shortfalls:
                cat_txt = ', '.join(f"{cat} ({cnt}x)" for cat, cnt in sorted(cat_counter.items(), key=lambda x: -x[1]))
                lines.append(f"    • {machine} — {cat_txt}")
            lines.append("")
        asks = []
        if any('Line Clear' in l or 'Block' in l for l in exc_by_type):
            asks.append('please expedite line clear / block sanction for the machines flagged above')
        if any('Failure' in l or 'Repair' in l for l in exc_by_type):
            asks.append('please arrange local maintenance support for machines under repair')
        longest = None
        if exc_by_type_rows:
            all_rows = [r for rows in exc_by_type_rows.values() for r in rows]
            longest = max(all_rows, key=lambda r: (r.get('DaysUnder') or 0))
        if longest and (longest.get('DaysUnder') or 0) >= 20:
            asks.append(f"{longest['Machine']} has been pending {longest.get('DaysUnder')} days — please prioritise this one first")
        if not asks:
            asks.append('no critical blockers today — request continued block support to sustain progress')
        lines.append("*Request:* " + '; '.join(asks) + '.')
        lines.append("")
        lines.append("— Dy.CE/TM, SWR")

        sf_lines = [f"*Shortfall Flag — {k} — {latest_month} {as_of.year}*", ""]
        if message_shortfalls:
            sf_lines.append("*Machine-wise shortfall reasons this month, open exceptions only:*")
            for machine, cat_counter in message_shortfalls:
                cat_txt = ', '.join(f"{cat} ({cnt}x)" for cat, cnt in sorted(cat_counter.items(), key=lambda x: -x[1]))
                sf_lines.append(f"    • {machine} — {cat_txt}")
            sf_lines.append("")
            worst_machine, worst_cats = message_shortfalls[0]
            worst_cat_name = sorted(worst_cats.items(), key=lambda x: -x[1])[0][0]
            sf_lines.append(f"*Request:* please review {worst_machine} ({worst_cat_name}) and the other machines "
                             f"above with the concerned section engineers before the next planning cycle.")
        else:
            sf_lines.append("No open-exception machines in this jurisdiction have a shortfall entry logged this month.")
            sf_lines.append("")
            sf_lines.append("*Request:* none — keep up the current block-utilisation discipline.")
        sf_lines.append("")
        sf_lines.append("— Dy.CE/TM, SWR")

        srden_bundle[k] = dict(key=k, division=div, code=code, machineCount=n_machines,
                                machines=sorted(srden_machines.get(k, [])), topCategories=top_cats,
                                exceptionsByType={lbl: sorted(set(m)) for lbl, m in exc_by_type.items()},
                                exceptionRows={lbl: exc_by_type_rows[lbl] for lbl in label_order},
                                machineShortfalls=[[m, sorted(c.items(), key=lambda x: -x[1])] for m, c in machine_shortfalls],
                                longestPending=longest, message='\n'.join(lines),
                                shortfallMessage='\n'.join(sf_lines))

    # ---- Reason-for-Shortfall tab: Sr.DEN-wise aggregates, full period + last 7 days ----
    dated_sf = [(r, ph.parse_ddmmyyyy(r.get('Date'))) for r in shortfall_records if r.get('SrDENKey')]
    dated_sf = [(r, dt) for r, dt in dated_sf if dt]
    max_sf_date = max((dt for _, dt in dated_sf), default=None)

    srden_cat_full = defaultdict(Counter)
    srden_machine_cat_full = defaultdict(lambda: defaultdict(Counter))
    for r in shortfall_records:
        k = r.get('SrDENKey')
        if not k:
            continue
        srden_cat_full[k][r['Category']] += 1
        if r['Category'] != 'Productive Work - Full Progress' and r.get('Machine'):
            srden_machine_cat_full[k][ph.norm_machine(r['Machine'])][r['Category']] += 1

    last7_start = last7_end = None
    srden_cat_last7 = defaultdict(Counter)
    srden_machine_cat_last7 = defaultdict(lambda: defaultdict(Counter))
    if max_sf_date:
        last7_end = max_sf_date.date()
        last7_start = last7_end - datetime.timedelta(days=6)
        for r, dt in dated_sf:
            if last7_start <= dt.date() <= last7_end:
                k = r['SrDENKey']
                srden_cat_last7[k][r['Category']] += 1
                if r['Category'] != 'Productive Work - Full Progress' and r.get('Machine'):
                    srden_machine_cat_last7[k][ph.norm_machine(r['Machine'])][r['Category']] += 1

    shortfall_by_srden_cat = {k: dict(c) for k, c in srden_cat_full.items()}
    shortfall_by_srden_machine_cat = {
        k: {m: dict(c) for m, c in mc.items()} for k, mc in srden_machine_cat_full.items()
    }
    shortfall_last7 = dict(
        startDate=last7_start.strftime('%d-%b-%Y') if last7_start else None,
        endDate=last7_end.strftime('%d-%b-%Y') if last7_end else None,
        bySrdenCat={k: dict(c) for k, c in srden_cat_last7.items()},
        bySrdenMachineCat={k: {m: dict(c) for m, c in mc.items()} for k, mc in srden_machine_cat_last7.items()},
    )

    # ---- progress.machines rebuild (v2 step) ----
    new_machines = []
    for r in progress_fixed:
        nm = dict(
            type=r.get('MachineTypeDisplay') or r.get('MachineType'), origType=r.get('MachineType'),
            machine=r['Machine'],
            div=(r.get('CurrentDIV') if r.get('CurrentDIV') in ('MYS', 'SBC', 'UBL') else r.get('HomeDivision')),
            homeDiv=r.get('HomeDivision'), workingAt=r.get('CurrentDIV'), status=r.get('Status'),
            statusCorrected=bool(r.get('StatusCorrected')), statusNote=r.get('StatusCorrectionNote'),
            unit=r.get('Unit'), annualTarget=r.get('AnnualTarget'), targetMonth=r.get('TargetPerMonth'),
            targetToDate=r.get('ProportionateTarget'), actualToDate=r.get('ActualProg'), pct=r.get('PctProg'),
            tamping=r.get('TampingCategory'), tampingNote=r.get('TampingNote'),
            iohPohStart=r.get('IOHPOH_StartDate'), iohPohExpected=r.get('IOHPOH_ExpectedCompletion'),
            iohPohDays=None, iohPohDaysUnderSrc=r.get('IOHPOH_DaysUnder'), iohPohLocation=r.get('IOHPOH_Location'),
            srden=machine_to_srden.get(ph.norm_machine(r['Machine'])),
            secondary=_secondary_dict(r.get('Secondary')),
        )
        for abbr in MONTH_ABBR:
            nm[abbr.lower()] = r.get(abbr)
        new_machines.append(nm)
    p['progress']['machines'] = new_machines
    p['progress']['tampingRollup'] = dict(Counter(m['tamping'] for m in new_machines))

    # ---- Unit Cost enrichment ----
    if p.get('unitCost'):
        progress_idx = {ph.machine_key_aliased(ph.norm_machine(m['machine'])): m for m in new_machines}
        fy_years = p['unitCost']['years']
        fy_order = sorted(fy_years.keys())
        for i, fy_label in enumerate(fy_order):
            year_block = fy_years[fy_label]
            prev_block = fy_years[fy_order[i - 1]] if i > 0 else None
            prev_by_key = {m['machineKey']: m for m in prev_block['machines']} if prev_block else {}
            type_rollup = defaultdict(lambda: dict(count=0, totalDailyCost=0.0, totalUnitCost=0.0,
                                                     unitCostCount=0, totalInitialCost=0.0))
            for rec in year_block['machines']:
                pm = progress_idx.get(rec['machineKey'])
                rec['unit'] = ph.norm_unit(pm.get('unit')) if pm else None
                rec['type'] = ph.disp_type(pm.get('type') or ph.mtype_of(rec['machine'])) if pm else ph.disp_type(ph.mtype_of(rec['machine']))
                rec['div'] = pm.get('div') if pm else None
                rec['status'] = pm.get('status') if pm else None
                rec['matchedToFleet'] = pm is not None
                prev = prev_by_key.get(rec['machineKey'])
                if prev and isinstance(prev.get('unitCost'), (int, float)) and isinstance(rec.get('unitCost'), (int, float)) and prev['unitCost']:
                    rec['prevUnitCost'] = prev['unitCost']
                    rec['unitCostChangePct'] = round((rec['unitCost'] - prev['unitCost']) / prev['unitCost'] * 100, 1)
                if prev and isinstance(prev.get('dailyCost'), (int, float)) and isinstance(rec.get('dailyCost'), (int, float)) and prev['dailyCost']:
                    rec['prevDailyCost'] = prev['dailyCost']
                    rec['dailyCostChangePct'] = round((rec['dailyCost'] - prev['dailyCost']) / prev['dailyCost'] * 100, 1)
                t = rec['type'] or 'Other'
                tr = type_rollup[t]
                tr['count'] += 1
                if isinstance(rec.get('dailyCost'), (int, float)):
                    tr['totalDailyCost'] += rec['dailyCost']
                if isinstance(rec.get('unitCost'), (int, float)):
                    tr['totalUnitCost'] += rec['unitCost']
                    tr['unitCostCount'] += 1
                if isinstance(rec.get('initialCost'), (int, float)):
                    tr['totalInitialCost'] += rec['initialCost']
            by_type = []
            for t, tr in sorted(type_rollup.items()):
                by_type.append(dict(
                    type=t, count=tr['count'],
                    avgDailyCost=round(tr['totalDailyCost'] / tr['count'], 2) if tr['count'] else 0,
                    avgUnitCost=round(tr['totalUnitCost'] / tr['unitCostCount'], 2) if tr['unitCostCount'] else None,
                    totalInitialCost=round(tr['totalInitialCost'], 2),
                ))
            year_block['byType'] = by_type
            all_daily = [m['dailyCost'] for m in year_block['machines'] if isinstance(m.get('dailyCost'), (int, float))]
            all_unit = [m['unitCost'] for m in year_block['machines'] if isinstance(m.get('unitCost'), (int, float))]
            year_block['kpi'] = dict(
                machineCount=len(year_block['machines']),
                totalDailyCost=round(sum(all_daily), 2) if all_daily else 0,
                avgDailyCost=round(sum(all_daily) / len(all_daily), 2) if all_daily else 0,
                avgUnitCost=round(sum(all_unit) / len(all_unit), 2) if all_unit else 0,
                costliestMachine=(max(year_block['machines'], key=lambda m: m.get('unitCost') or 0)['machine']
                                   if year_block['machines'] else None),
                costliestUnitCost=(max((m.get('unitCost') or 0) for m in year_block['machines'])
                                    if year_block['machines'] else 0),
            )

    for row in p['exceptions']['latest']:
        bucket, clean_label = ph.classify_exc_category(row['Category'])
        row['Type'] = bucket
        row['CategoryLabel'] = clean_label
        row['SrDEN'] = machine_to_srden.get(ph.norm_machine(row['Machine']))
    p['exceptions']['iohPohReconciled'] = iohpoh_reconciled
    p['exceptions']['iohPohReconciledCount'] = len(iohpoh_reconciled)
    p['exceptions']['iohPohStale'] = iohpoh_stale
    p['exceptions']['iohPohStaleCount'] = len(iohpoh_stale)
    p['exceptions']['iohPohCpOnlyUnconfirmed'] = iohpoh_cp_only_unconfirmed
    p['exceptions']['iohPohCpOnlyUnconfirmedCount'] = len(iohpoh_cp_only_unconfirmed)
    commissioning_rows = [r for r in p['exceptions']['latest'] if r['Type'] == 'commissioning']
    longest_iohpoh = max([r for r in p['exceptions']['latest'] if r['Type'] == 'iohpoh'],
                          key=lambda r: r.get('DaysUnder') or 0, default=None)
    longest_commissioning = max(commissioning_rows, key=lambda r: r.get('DaysUnder') or 0, default=None)
    p['exceptions']['longestIohPoh'] = longest_iohpoh
    p['exceptions']['longestCommissioning'] = longest_commissioning

    p['shortfall']['srdenBundle'] = srden_bundle
    p['shortfall']['bySrdenCat'] = shortfall_by_srden_cat
    p['shortfall']['bySrdenMachineCat'] = shortfall_by_srden_machine_cat
    p['shortfall']['last7'] = shortfall_last7

    overdue = []
    for r in ioh_poh_records:
        nd_ioh, nd_poh = r.get('NextDueIOH_Months'), r.get('NextDuePOH_Months')
        is_overdue_ioh = isinstance(nd_ioh, (int, float)) and nd_ioh < 0
        is_overdue_poh = isinstance(nd_poh, (int, float)) and nd_poh < 0
        if not (is_overdue_ioh or is_overdue_poh):
            continue
        mt = re.match(r'[A-Za-z]+', r['Machine'].replace(' ', '').replace('-', ''))
        typecode = mt.group(0).upper() if mt else ''
        tamping = TAMPING_INFO.get(typecode, ('Non-Tamping', ''))[0]
        overdue.append(dict(machine=r['Machine'], division=r['Division'], type=ph.disp_type(typecode), tamping=tamping,
                             commission=r['Commission'],
                             overdueIOHMonths=round(nd_ioh, 1) if is_overdue_ioh else None,
                             overduePOHMonths=round(nd_poh, 1) if is_overdue_poh else None,
                             lastIOHDate=r['LastIOH_Date'], lastPOHDate=r['LastPOH_Date'], remark=r['Remark']))
    overdue.sort(key=lambda x: min(x['overdueIOHMonths'] or 0, x['overduePOHMonths'] or 0))
    p['fleet']['iohPohOverdue'] = overdue
    p['fleet']['iohPohOverdueCount'] = len(overdue)
    p['fleet']['iohPohTotalTracked'] = len(ioh_poh_records)

    for r in progress_fixed:
        old = r.get('StatusCorrectedFrom')
        if not old:
            continue
        p['meta']['statusCounts']['Active'] = p['meta']['statusCounts'].get('Active', 0) + 1
        p['meta']['statusCounts'][old] = p['meta']['statusCounts'].get(old, 0) - 1
        if p['meta']['statusCounts'][old] <= 0:
            p['meta']['statusCounts'].pop(old, None)
    p['meta']['iohPohReconciledCount'] = len(iohpoh_reconciled)
    p['meta']['iohPohOverdueCount'] = len(overdue)
    p['meta']['iohPohStaleCount'] = len(iohpoh_stale)

    # ---- v3: age enrichment + catMachine/catDivMachine drilldowns + merged mtypeRollup ----
    age_by_machine = {ph.norm_machine(a['Machine']): a for a in age_profile}
    for m in p['progress']['machines']:
        a = age_by_machine.get(ph.norm_machine(m['machine']))
        if a:
            m['age'] = a['AgeYears']
            m['ageBucket'] = ('0-10 yrs' if (a['AgeYears'] or 0) <= 10 else
                               '11-20 yrs' if (a['AgeYears'] or 0) <= 20 else
                               '21-30 yrs' if (a['AgeYears'] or 0) <= 30 else '30+ yrs') if a['AgeYears'] is not None else None
            m['commissionDate'] = a['CommissioningDate']

    cat_div_machine = defaultdict(lambda: defaultdict(Counter))
    cat_mtype_machine = defaultdict(lambda: defaultdict(Counter))
    cat_machine = defaultdict(Counter)
    for r in shortfall_records:
        cat, div, mach = r.get('Category'), r.get('Division'), r.get('Machine')
        if not cat or not mach:
            continue
        cat_machine[cat][mach] += 1
        if div in ('MYS', 'SBC', 'UBL'):
            cat_div_machine[cat][div][mach] += 1
        cat_mtype_machine[cat][ph.disp_type(ph.mtype_of(mach))][mach] += 1
    p['shortfall']['catMachine'] = {cat: sorted(ctr.items(), key=lambda x: -x[1])[:12] for cat, ctr in cat_machine.items()}
    p['shortfall']['catDivMachine'] = {
        cat: {div: sorted(ctr.items(), key=lambda x: -x[1])[:10] for div, ctr in divs.items()}
        for cat, divs in cat_div_machine.items()
    }
    p['shortfall']['catMtypeMachine'] = {
        cat: {mt: sorted(ctr.items(), key=lambda x: -x[1])[:15] for mt, ctr in types.items()}
        for cat, types in cat_mtype_machine.items()
    }

    by_type = defaultdict(lambda: dict(count=0, annualTarget=0.0, targetToDate=0.0, actualToDate=0.0))
    for m in p['progress']['machines']:
        t = m['type']
        d = by_type[t]
        d['count'] += 1
        d['annualTarget'] += (m.get('annualTarget') or 0)
        d['targetToDate'] += (m.get('targetToDate') or 0)
        d['actualToDate'] += (m.get('actualToDate') or 0)
    new_mtype_rollup = []
    for t, d in sorted(by_type.items()):
        pct = (d['actualToDate'] / d['targetToDate']) if d['targetToDate'] else 0
        new_mtype_rollup.append(dict(type=t, count=d['count'], annualTarget=round(d['annualTarget'], 1),
                                      targetToDate=round(d['targetToDate'], 1), actualToDate=round(d['actualToDate'], 1),
                                      pct=round(pct, 4)))
    by_type_secondary = defaultdict(lambda: dict(count=0, annualTarget=0.0, targetToDate=0.0, actualToDate=0.0, unit=None))
    for m in p['progress']['machines']:
        sec = m.get('secondary')
        if not sec:
            continue
        t = f"{m['type']} ({sec['unit']})"
        d = by_type_secondary[t]
        d['count'] += 1
        d['annualTarget'] += (sec.get('annualTarget') or 0)
        d['targetToDate'] += (sec.get('targetToDate') or 0)
        d['actualToDate'] += (sec.get('actualToDate') or 0)
        d['unit'] = sec.get('unit')
    for t, d in sorted(by_type_secondary.items()):
        pct = (d['actualToDate'] / d['targetToDate']) if d['targetToDate'] else 0
        new_mtype_rollup.append(dict(type=t, count=d['count'], annualTarget=round(d['annualTarget'], 1),
                                      targetToDate=round(d['targetToDate'], 1), actualToDate=round(d['actualToDate'], 1),
                                      pct=round(pct, 4), unit=d['unit'], isSecondaryMetric=True))
    if by_type_secondary:
        log(f"mtypeRollup (v3 rebuild): re-attached {len(by_type_secondary)} secondary-unit pseudo-type "
            f"row(s): {', '.join(sorted(by_type_secondary))}")
    p['progress']['mtypeRollup'] = new_mtype_rollup

    # ---- v4: Sr.DEN fuzzy attach + double-shift roster ----
    for m in p['progress']['machines']:
        if not m.get('srden'):
            hit = lookup_srden(m['machine'])
            if hit:
                m['srden'] = hit
    machine_idx = {ph.alnum(m['machine']): m for m in p['progress']['machines']}

    for r in ds_records:
        r['srdenKey'] = canon_ds_srden(r.get('div'), r.get('srden'))
        hit = machine_idx.get(ph.alnum(r.get('machine') or ''))
        r['machineType'] = hit['type'] if hit else None

    over12 = [r for r in ds_records if r.get('over12h')]
    by_machine = defaultdict(list)
    for r in ds_records:
        if r.get('machine'):
            by_machine[r['machine']].append(r)
    machine_rollup = []
    for mc, recs in by_machine.items():
        durations = [r['durationHours'] for r in recs if r.get('durationHours') is not None]
        machine_rollup.append(dict(machine=mc, days=len(set(r['date'] for r in recs)), shifts=len(recs),
                                    maxDurationHours=max(durations) if durations else None,
                                    over12hCount=sum(1 for r in recs if r.get('over12h')),
                                    div=next((r['div'] for r in recs if r.get('div')), None),
                                    srdenKey=next((r['srdenKey'] for r in recs if r.get('srdenKey')), None),
                                    machineType=next((r['machineType'] for r in recs if r.get('machineType')), None)))
    machine_rollup.sort(key=lambda x: -x['days'])

    srden_over12 = defaultdict(list)
    for r in over12:
        if r.get('srdenKey'):
            srden_over12[r['srdenKey']].append(r)
    srden_ds_messages = {}
    for key, recs in srden_over12.items():
        recs_sorted = sorted(recs, key=lambda r: -(r['durationHours'] or 0))
        lines = [f"*Double-Shift Alert — {key} — {as_of.strftime('%d-%b-%Y')}*",
                 f"({len(recs_sorted)} shift(s) exceeded 12 hours — staff fatigue risk)", ""]
        for r in recs_sorted[:15]:
            lines.append(f"- {r['machine']} | {r['shift']} shift on {r['date']} | on-duty {r['durationHours']}h "
                          f"(started {r.get('started_raw')}, back at siding {r.get('arrived_raw')})")
        lines += ["", "*Request:* please review shift planning/relief arrangements for the machines above.", "",
                  "— Dy.CE/TM, SWR"]
        srden_ds_messages[key] = "\n".join(lines)

    p['fleet']['doubleShift'] = dict(records=ds_records, machineRollup=machine_rollup,
                                      over12hRecords=sorted(over12, key=lambda r: -(r['durationHours'] or 0)),
                                      over12hCount=len(over12), totalShifts=len(ds_records),
                                      daysCovered=len(sorted(set(r['date'] for r in ds_records if r['date']))),
                                      srdenMessages=srden_ds_messages)

    # ---- YoY comparison ----
    prior_by_type, prior_total, prior_by_machine, prior_by_type_secondary, prior_by_machine_secondary = \
        extract_prior_year_by_type(base)
    if prior_by_type is not None:
        cur_by_type = defaultdict(float)
        cur_total = 0.0
        cur_by_type_secondary = defaultdict(float)
        cur_by_machine_secondary = {}
        for m in p['progress']['machines']:
            s = sum((m.get(k) if isinstance(m.get(k), (int, float)) else 0) for k in
                    ['apr', 'may', 'jun', 'jul', 'aug', 'sep'])
            cur_by_type[m['type']] += s
            cur_total += s
            sec = m.get('secondary')
            if sec:
                ss = sum((sec.get(k) if isinstance(sec.get(k), (int, float)) else 0) for k in
                         ['apr', 'may', 'jun', 'jul', 'aug', 'sep'])
                t = f"{m['type']} ({sec['unit']})"
                cur_by_type_secondary[t] += ss
                cur_by_machine_secondary[f"{ph.machine_key(m['machine'])}|{sec['unit']}"] = ss
        types = sorted(set(list(prior_by_type.keys()) + list(cur_by_type.keys())))
        yoy_by_type = []
        for t in types:
            pv, cv = round(prior_by_type.get(t, 0), 2), round(cur_by_type.get(t, 0), 2)
            delta = round(cv - pv, 2)
            pct = round((cv - pv) / pv * 100, 1) if pv else None
            yoy_by_type.append(dict(type=t, priorYear=pv, currentYear=cv, delta=delta, pctChange=pct))

        sec_types = sorted(set(list((prior_by_type_secondary or {}).keys()) + list(cur_by_type_secondary.keys())))
        for t in sec_types:
            pv = round((prior_by_type_secondary or {}).get(t, 0), 2)
            cv = round(cur_by_type_secondary.get(t, 0), 2)
            delta = round(cv - pv, 2)
            pct = round((cv - pv) / pv * 100, 1) if pv else None
            unit = t[t.rfind('(') + 1:-1] if '(' in t else None
            yoy_by_type.append(dict(type=t, unit=unit, isSecondaryMetric=True,
                                     priorYear=pv, currentYear=cv, delta=delta, pctChange=pct))

        yoy_by_machine = []
        if prior_by_machine is not None:
            for m in p['progress']['machines']:
                cv = sum((m.get(k) if isinstance(m.get(k), (int, float)) else 0) for k in
                         ['apr', 'may', 'jun', 'jul', 'aug', 'sep'])
                pv = prior_by_machine.get(ph.machine_key(m['machine']))
                pv = round(pv, 2) if pv is not None else 0.0
                cv = round(cv, 2)
                delta = round(cv - pv, 2)
                pct = round((cv - pv) / pv * 100, 1) if pv else None
                yoy_by_machine.append(dict(machine=m['machine'], type=m['type'], div=m['div'],
                                            priorYear=pv, currentYear=cv, delta=delta, pctChange=pct))
                sec = m.get('secondary')
                if sec:
                    key = f"{ph.machine_key(m['machine'])}|{sec['unit']}"
                    spv = (prior_by_machine_secondary or {}).get(key)
                    spv = round(spv, 2) if spv is not None else 0.0
                    scv = round(cur_by_machine_secondary.get(key, 0), 2)
                    sdelta = round(scv - spv, 2)
                    spct = round((scv - spv) / spv * 100, 1) if spv else None
                    yoy_by_machine.append(dict(machine=f"{m['machine']} ({sec['unit']})",
                                                type=f"{m['type']} ({sec['unit']})", unit=sec['unit'],
                                                isSecondaryMetric=True, baseMachine=m['machine'], div=m['div'],
                                                priorYear=spv, currentYear=scv, delta=sdelta, pctChange=spct))

        p['progress']['yoy'] = dict(
            label=f'Apr–Sep {as_of.year-1} vs Apr–Sep {as_of.year} (cumulative progress, same period)',
            priorTotal=round(prior_total, 2), currentTotal=round(cur_total, 2),
            totalDelta=round(cur_total - prior_total, 2),
            totalPctChange=round((cur_total - prior_total) / prior_total * 100, 1) if prior_total else None,
            byType=yoy_by_type, byMachine=yoy_by_machine)

    # ---- IOH/POH days-elapsed AND days-designated ----
    for m in p['progress']['machines']:
        if m.get('iohPohStart'):
            dt = ph.parse_ddmmyyyy(m['iohPohStart'])
            if dt:
                m['iohPohDaysElapsed'] = (datetime.datetime.combine(as_of, datetime.time()) - dt).days
                exp = ph.parse_ddmmyyyy(m.get('iohPohExpected'))
                if exp:
                    m['iohPohTotalSpan'] = (exp - dt).days
                    m['iohPohDays'] = m['iohPohTotalSpan']

    # ---- Fleet Health - IOH/POH Schedule ----
    def attach_srden(machine_name):
        return lookup_srden(machine_name)

    current_events = []
    if ioh_poh_letter:
        for ev in ioh_poh_letter['events']:
            ev = dict(ev)
            ev['srden'] = attach_srden(ev['machine'])
            ev['division'] = ev.get('division') or (home_div.get(norm(ev['machine'])))
            current_events.append(ev)
    current_priority = []
    if ioh_poh_letter:
        for ev in ioh_poh_letter['priorityEvents']:
            ev = dict(ev)
            ev['srden'] = attach_srden(ev['machine'])
            ev['division'] = ev.get('division') or (home_div.get(norm(ev['machine'])))
            current_priority.append(ev)
    current_events.extend(current_priority)

    ioh_priority = sorted(
        (dict(e) for e in current_events
         if e.get('type') == 'IOH' and e.get('status') in ('Yet to Release from Division', 'Ongoing / Under Attention')),
        key=lambda e: (e['monthSort'], e['machine'])
    )

    FUTURE_FY_LABELS = ['2027-2028', '2028-2029', '2029-2030']
    future_events = []
    for r in ioh_poh_records:
        sched = r.get('Schedule') or {}
        mt = re.match(r'[A-Za-z]+', r['Machine'].replace(' ', '').replace('-', ''))
        typecode = ph.disp_type(mt.group(0).upper()) if mt else 'OTHER'
        for fy in FUTURE_FY_LABELS:
            ioh_v, poh_v = (sched.get(fy) or [None, None])
            for kind, val in (('IOH', ioh_v), ('POH', poh_v)):
                if val in (None, '', '-'):
                    continue
                fy_a, fy_b = fy.split('-')
                future_events.append(dict(
                    machine=r['Machine'], machineKey=ph.machine_key(r['Machine']),
                    division=r.get('Division'), type=typecode, kind=kind,
                    fy=f"{fy_a}-{fy_b[-2:]}",
                    fyRaw=fy, month=str(val), monthSort=list(parse_iohpoh_month_sort(str(val))),
                    srden=attach_srden(r['Machine']),
                ))
    future_events.sort(key=lambda e: (e['fyRaw'], e['monthSort']))

    def status_counts(events, status_field='status'):
        return dict(Counter(e.get(status_field) or 'Scheduled' for e in events))

    def by_division(events):
        out = defaultdict(list)
        for e in events:
            out[e.get('division') or 'Unknown'].append(e)
        return dict(out)

    current_by_div = by_division(current_events)
    current_div_summary = []
    for d in ('SBC', 'UBL', 'MYS'):
        evs = current_by_div.get(d, [])
        current_div_summary.append(dict(
            division=d, total=len(evs),
            iohCount=sum(1 for e in evs if e.get('type') == 'IOH'),
            pohCount=sum(1 for e in evs if e.get('type') == 'POH'),
            statusCounts=status_counts(evs),
        ))
    unknown_div_events = current_by_div.get('Unknown', [])

    future_by_div = by_division(future_events)
    future_div_summary = []
    for d in ('SBC', 'UBL', 'MYS'):
        evs = future_by_div.get(d, [])
        future_div_summary.append(dict(
            division=d, total=len(evs),
            byFY={fy: sum(1 for e in evs if e['fyRaw'] == fy) for fy in FUTURE_FY_LABELS},
            iohCount=sum(1 for e in evs if e['kind'] == 'IOH'),
            pohCount=sum(1 for e in evs if e['kind'] == 'POH'),
        ))

    current_srden_summary = defaultdict(lambda: defaultdict(list))
    for e in current_events:
        d = e.get('division') or 'Unknown'
        s = e.get('srden') or f"{d} — Not currently mapped"
        current_srden_summary[d][s].append(e)
    current_srden_summary = {
        d: [dict(srden=s, events=evs, count=len(evs)) for s, evs in sorted(by_s.items())]
        for d, by_s in current_srden_summary.items()
    }

    status_overall = status_counts(current_events)
    currently_ongoing = [e for e in current_events if e.get('status') == 'Ongoing / Under Attention']

    completed_ioh_count = sum(1 for e in current_events if e.get('type') == 'IOH' and e.get('status') == 'Completed')
    completed_poh_count = sum(1 for e in current_events if e.get('type') == 'POH' and e.get('status') == 'Completed')
    completed_total = status_overall.get('Completed', 0)
    if completed_ioh_count + completed_poh_count != completed_total:
        log(f"WARN: completed IOH ({completed_ioh_count}) + completed POH ({completed_poh_count}) "
            f"!= completedCount ({completed_total}) - a current-year event has an unexpected type/status combo")

    p['fleet']['iohPohSchedule'] = dict(
        currentYear=dict(
            source=(ioh_poh_letter or {}).get('sourceFile'),
            letterDate=(ioh_poh_letter or {}).get('letterDate'),
            fy=(ioh_poh_letter or {}).get('fy'),
            events=current_events,
            unmappedDivisionEvents=unknown_div_events,
            priorityEvents=current_priority,
            iohPriorityEvents=ioh_priority,
            byDivision=current_div_summary,
            bySrden=current_srden_summary,
            statusCounts=status_overall,
            ongoing=currently_ongoing,
            ongoingCount=len(currently_ongoing),
            completedCount=completed_total,
            completedIOHCount=completed_ioh_count,
            completedPOHCount=completed_poh_count,
            yetToReleaseCount=status_overall.get('Yet to Release from Division', 0),
            scheduledCount=status_overall.get('Scheduled', 0),
            totalCount=len(current_events),
        ),
        futureYears=dict(
            source="live sync (ioh_poh_planning)",
            fyLabels=FUTURE_FY_LABELS,
            events=future_events,
            byDivision=future_div_summary,
            totalCount=len(future_events),
        ),
        norms=ioh_poh_norms,
    )

    if cpoh_planning:
        for grp in ('proposed2728', 'tentative2829', 'tentative2930'):
            for ev in cpoh_planning[grp]:
                ev['division'] = home_div.get(norm(ev['machine']))
                ev['srden'] = attach_srden(ev['machine'])
        p['fleet']['cpohPlanning'] = cpoh_planning
    else:
        p['fleet']['cpohPlanning'] = None

    log(f"Fleet Health IOH/POH schedule: {len(current_events)} current-year events "
        f"({ioh_poh_letter and ioh_poh_letter.get('sourceFile')}), {len(future_events)} "
        f"future-year events (27-30 workbook); completed {completed_total} "
        f"(IOH {completed_ioh_count} + POH {completed_poh_count}); {len(ioh_priority)} "
        f"IOH priority entries derived, {len(current_priority)} POH priority entries from the letter.")

    out_path = os.path.join(outdir, 'dashboard_payload.json')
    with open(out_path, 'w') as f:
        json.dump(p, f, default=str)
    log("Saved", out_path, "size", os.path.getsize(out_path))
    return out_path


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description="SWR Track Machine Dashboard - live Supabase aggregation pipeline.")
    parser.add_argument('outdir', help="Directory to write dashboard_payload.json into.")
    parser.add_argument('--as-of', dest='as_of', default=None,
                         help="Override the as-of date (YYYY-MM-DD). Defaults to today.")
    args = parser.parse_args()
    as_of_override = datetime.date.fromisoformat(args.as_of) if args.as_of else None
    run(args.outdir, as_of=as_of_override)
