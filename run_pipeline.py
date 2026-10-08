"""
SWR Track Machine Dashboard — full pipeline, raw Excel -> final payload JSON.

Consolidates the entire chain that was originally built interactively across
many scripts/inline snippets into one script a fresh session can run
end-to-end. Reads every source file with glob-based discovery (never a
hardcoded exact filename) so it keeps working as monthly/daily files get
replaced in the "Machine Monitoring" folder.

Usage:
    python3 run_pipeline.py <MachineMonitoringFolder> <OutputDir>

Produces <OutputDir>/dashboard_payload.json — the final payload to embed
into dashboard_template2.html at the __DATA_JSON__ placeholder.

IMPORTANT KNOWN LIMITATION (read before relying on this unattended):
The whole schema (Cumulative Progress columns, Daily Progress monthly
files, progress.machines.{apr..sep} fields, the dashboard's "Apr-Sep 2026"
framing) is built around FY2026-27's first half (Apr-Sep). This script
auto-discovers whichever of those 6 standard months' files are present, but
does NOT extend the schema to Oct/Nov/etc. Once Cumulative Progress moves
past September, this pipeline (and the HTML template it feeds) needs a
follow-up redesign to cover a rolling/full-year window - it will not do
that automatically.
"""
import sys, os, re, glob, json, datetime, difflib
from collections import defaultdict, Counter
import openpyxl

MONTH_ABBR = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep']


# ---------------------------------------------------------------------------
# Fallback shortfall-category classifier.
#
# Normally each Daily Progress workbook has a "Shortfall Category" column
# hand-injected via Excel formula (see the project's
# shortfall-category-taxonomy.md doc - 19 categories, built and approved with
# the user from ~15,000 raw Remarks entries). That column can go missing on a
# given month's file if it gets re-synced/re-exported from wherever the
# underlying report is maintained (confirmed to happen for Sep 2026 - Apr-Aug
# still carry it). Rather than silently produce zero shortfall/category data
# for that month, reconstruct the same 19-category scheme directly from the
# Remarks free text as a fallback, only when the real tag column is absent.
#
# This is a best-effort reimplementation of the documented priority-ordered
# keyword rules, not a byte-exact copy of the original Excel formula (that
# formula itself isn't preserved verbatim anywhere) - tested against Sep 2026
# and directionally sound (~38% "Productive Work - Full Progress" vs ~40%
# historical), but a few categories (notably "Material Loading/Stocking") are
# under-matched on compound remarks. Records classified this way are flagged
# with CategorySource="reconstructed" so this can be distinguished from the
# original tagged data if that distinction is ever needed downstream.
# ---------------------------------------------------------------------------
_OBSTRUCTION_KW = ('obstruction', 'boulder', 'caked ballast', 'old gate', 'concrete slab')
_WEATHER_KW = ('wet muck',)
_PRODUCTIVE_KW = ('tamped', 'tamping', 'deep screening', 'boxing', 'brooming', 'sweeping',
                   'panels laid', 'worked behind', 'pt.no', 'pt no', 'track lifting',
                   'sleeper renewal', 'rail renewal', 'welding', 'packing', 'survey',
                   'patrolling', 'ballast', 'clamped', 'dismantling', 'shifting done',
                   'longitudinal shifting', 'shoulder ballast cleaning')
_MATERIAL_KW = ('rails loaded', 'rails stocked', 'rail loaded', 'rail stocked',
                 'sleeper', 'muck loaded', 'muck unloaded', 'muck disposed')


def classify_shortfall(text):
    """Best-effort reconstruction of the 19-category shortfall taxonomy from
    raw Remarks text. Returns (category, is_reconstructed=True)."""
    if not text:
        return None
    t = str(text).lower()
    bc = 'block cancel' in t
    rain_hit = bool(re.search(r'\brain\b', t))

    if 'block not permit' in t or 'heavy traffic' in t or 'not permitted by controller' in t:
        return 'Block Not Permitted by Traffic'
    if bc and any(k in t for k in ('controller', 'late running of trains', 'insufficient block time', 'power block')):
        return 'Block Not Permitted by Traffic'
    if any(k in t for k in ('pwi not utilis', 'pwi not utiliz', 'refused by pwi', 'block not availed', 'block not plann', 'depot work', 'under preparation')):
        return 'Block/Machine Not Utilised (All Reasons)'
    if bc and 'pwi' in t:
        return 'Block/Machine Not Utilised (All Reasons)'
    if 'no sanctioned block' in t or 'not sanctioned' in t:
        return 'Block Not Sanctioned'
    has_weather_or_obs = rain_hit or any(k in t for k in _WEATHER_KW) or any(k in t for k in _OBSTRUCTION_KW)
    if ('limited scope' in t or re.search(r'\blsp\b', t)) and not has_weather_or_obs:
        return 'LSP: Less Progress'
    if any(k in t for k in ('crew not available', 'loco not available', 'crew & loco', 'loco & crew')):
        return 'Crew / Loco Not Available'
    if bc and any(k in t for k in ('loco failure', 'loco issue', 'loco not arranged', 'labourer', 'crew')):
        return 'Crew / Loco Not Available'
    if 'under poh' in t:
        return 'Under POH'
    if 'under ioh' in t:
        return 'Under IOH'
    if 'under repair' in t or 'breakdown' in t:
        return 'Breakdown / Under Repair'
    if 'hrs schedule' in t or 'hr schedule' in t or 'schedule maintenance' in t:
        return 'Scheduled Maintenance (periodic hrs service)'
    if 'under commission' in t:
        return 'Under Commissioning'
    if any(k in t for k in _OBSTRUCTION_KW):
        return 'Site Obstruction (boulders/debris/utilities)'
    if rain_hit or 'wet muck' in t:
        return 'Site/Weather Condition Unsuitable'
    if 'formation' in t:
        return 'Formation under BPC'
    if 'waiting for' in t and any(k in t for k in ('programme', 'instruction', 'further')):
        return 'Waiting for Programme/Instructions'
    if 'availed cl' in t or 'availed nh' in t or (('festival' in t) and re.search(r'\b(cl|nh)\b', t)):
        return 'Staff Leave/Holiday (festival, CL/NH)'
    if any(k in t for k in ('planned to move', 'shifted to', 'movement from', 'transit', 'on run to')):
        return 'Machine in Transit/Relocation'
    if any(k in t for k in _MATERIAL_KW):
        return 'Material Loading/Stocking (logistics support)'
    if any(k in t for k in _PRODUCTIVE_KW):
        return 'Productive Work - Full Progress'
    return 'Other/Uncategorized'
MONTH_NUM = {'Apr': 4, 'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8, 'Sep': 9}


def log(*a):
    print(*a, file=sys.stderr)


def norm_machine(m):
    if not m:
        return None
    s = re.sub(r'\s+', ' ', str(m)).strip().upper()
    return s.replace(' - ', '-').replace(' -', '-').replace('- ', '-')


def alnum(m):
    return re.sub(r'[^A-Z0-9]', '', str(m).upper())


def machine_key(m):
    """Canonical join key for matching the same machine across sheets that spell its
    number differently (leading zeros, parenthetical location suffixes, etc.) - e.g.
    'UTV-058', 'UTV-58', 'UTV 058' and 'UTV-003 (SAN)' must all key to the same value.
    Added 28-Sep-2026: plain alnum() was missing exactly this (UTV-058 vs UTV-58, and
    UTV-003 vs 'UTV-003 (SAN)'), which silently broke the IOH/POH Planning cross-check
    for the two machines the user flagged (UTV-58, UTV-003)."""
    s = re.sub(r'\(.*?\)', '', str(m).upper())  # drop parenthetical suffixes e.g. "(SAN)"
    s = re.sub(r'[^A-Z0-9]', '', s)
    mtp = re.match(r'([A-Z]+)(\d+)', s)
    if mtp:
        return mtp.group(1) + str(int(mtp.group(2)))  # strip leading zeros on the numeric part
    return s


# 07-Oct-2026: confirmed with the user. The Progress/Cumulative-Progress sheet and the
# Daily Progress shortfall log use DIFFERENT type-prefix abbreviations for the same
# machine type, not just different punctuation - "FRM" (Progress sheet) is "SBCM"
# (shortfall log), "PBR" is "BRM", "PCT" is "PCTM". norm_machine()/machine_key() can't
# bridge this (the letters themselves differ), so it has to be an explicit table.
MACHINE_PREFIX_ALIASES = {
    'SBCM': 'FRM',
    'BRM': 'PBR',
    'PCTM': 'PCT',
}


def machine_key_aliased(m):
    """machine_key() plus MACHINE_PREFIX_ALIASES - use this (not plain machine_key())
    anywhere a machine name might need to cross between the Progress sheet and the
    shortfall log, so e.g. 'SBCM-1889' and 'FRM 1889' key the same."""
    s = machine_key(m)
    mtp = re.match(r'([A-Z]+)(\d+)$', s)
    if mtp:
        return MACHINE_PREFIX_ALIASES.get(mtp.group(1), mtp.group(1)) + mtp.group(2)
    return s


# 07-Oct-2026: confirmed with the user there are exactly 2 distinct physical T-28 duo
# tamping units, each logged under multiple inconsistent spellings across the Progress
# sheet and the shortfall log - "T-28 (908 A&B)" and "T-28 (403&404)". machine_key()
# can't disambiguate them safely (it drops parenthetical content entirely, which is
# where T-28's actual distinguishing number lives, unlike a location suffix like
# "(SAN)"), so these are explicit aliases to the one canonical spelling for each unit,
# applied right after norm_machine() - everywhere a T-28 variant is seen it becomes
# literally the same string before any other matching logic runs.
MACHINE_NAME_ALIASES = {
    'T-28(908A &B)': 'T-28 (908 A&B)',
    'T-28(908 A&B)': 'T-28 (908 A&B)',
    'T-28 (908A&B)': 'T-28 (908 A&B)',
    'T-28 (03&04)': 'T-28 (403&404)',
}


def canon_machine_alias(nm):
    """Apply MACHINE_NAME_ALIASES to an already-norm_machine()'d name."""
    return MACHINE_NAME_ALIASES.get(nm, nm)


def mtype_of(name):
    m = re.match(r'[A-Za-z]+', str(name).strip())
    return m.group(0).upper() if m else 'OTHER'


UNI_MERGE = {'PCTM': 'UNI', 'PCT': 'UNI'}


def disp_type(t):
    return UNI_MERGE.get(t, t)


# Sourced from SWR's own Cumulative Progress workbooks' official Tamping/Non-Tamping
# grand-total banner rows (see project doc dashboard-build-summary.md for detail) -
# NOT from web research.
#
# DEFINITIVE, USER-SUPPLIED CLASSIFICATION (28-Sep-2026) - supersedes all prior
# SWR-workbook-banner-based and IRICEN-handbook-based classification below. The user
# gave this list verbatim and explicitly stated the workbook-banner-derived Tamping
# count ("47 tamping") was wrong:
#   Tamping:      MPT, Duomatic, Unimat, DTE, CSM
#   Non-Tamping:  UTV, MDU, BCM, PBR, SQRS, FRM, DGS, RBMV(RBMB), T(28), RGM, SRGM(SRG)
# Do not revert to the old SWR-banner classification (which had FRM/SQRS/T/RGM/SRGM as
# Tamping) without the user's explicit sign-off again.
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
}

EXC_TYPE_MAP = {
    'Under repair': ('failure', 'Failure / Under Repair'),
    "Machine's not availing block for less than 5 Days": ('block', 'Not Availing Block (<5 days)'),
    "Machine's not availing block for more than 5 Days": ('block', 'Not Availing Block (>5 days)'),
    'Line clear not permitted': ('block', 'Line Clear Not Permitted'),
    'GSU/CN': ('block', 'GSU/Construction Block'),
    'Machine Under Movement': ('transit', 'Under Movement/Transit'),
    'Machine Under maintenance': ('maintenance', 'Under Maintenance'),
    'M/c under IOH/POH': ('iohpoh', 'Under IOH/POH'),
    'M/c under commissioning': ('commissioning', 'Under Commissioning'),
}

# Fuzzy, pattern-based version of EXC_TYPE_MAP (added 28-Sep-2026 per the user's
# "the open exceptions sheet is not clear, please re-categorise it properly"
# feedback). The exact-string lookup above silently dropped any exception sheet
# row whose raw Category text didn't match byte-for-byte into a raw, unreadable
# 'other' bucket -- e.g. 'GSU/CN/K- RIDE' (the sheet's real text) never matched
# the map's 'GSU/CN' key, so it showed up verbatim in every chart and list
# instead of as 'GSU / Construction Block'. This is what actually feeds the
# Overview and Exceptions-tab "Open Exceptions by Category" charts and the
# per-machine Category field, so every row now gets a clean, readable label
# even when the sheet's exact wording drifts.
EXC_TYPE_RULES = [
    (re.compile(r'under\s*repair|break\s*down|breakdown', re.I), ('failure', 'Failure / Under Repair')),
    (re.compile(r'not\s*availing\s*block.*less\s*than|less\s*than\s*5\s*days?.*block', re.I), ('block', 'Not Availing Block (<5 days)')),
    (re.compile(r'not\s*availing\s*block.*more\s*than|more\s*than\s*5\s*days?.*block', re.I), ('block', 'Not Availing Block (>5 days)')),
    (re.compile(r'line\s*clear', re.I), ('block', 'Line Clear Not Permitted')),
    (re.compile(r'gsu|cn/k|construction\s*block', re.I), ('block', 'GSU / Construction Block')),
    (re.compile(r'movement|transit', re.I), ('transit', 'Under Movement / Transit')),
    (re.compile(r'under\s*maintenance', re.I), ('maintenance', 'Under Maintenance')),
    (re.compile(r'ioh\s*/?\s*poh', re.I), ('iohpoh', 'Under IOH/POH')),
    (re.compile(r'under\s*commissioning|commissioning', re.I), ('commissioning', 'Under Commissioning')),
]


def classify_exc_category(raw):
    """Raw Exception-Sheet Category text -> (bucket_key, clean_label). Falls
    back to the raw text (trimmed) only when nothing recognisable matches, so
    a genuinely new category still surfaces rather than being hidden."""
    s = str(raw or '').strip()
    for pat, (bucket, label) in EXC_TYPE_RULES:
        if pat.search(s):
            return bucket, label
    return 'other', (s or 'Uncategorised')

ZONE_LETTER = {'CENTRAL': 'C', 'EAST': 'E', 'WEST': 'W', 'NORTH': 'N', 'SOUTH': 'S', 'HQ': 'HQ'}


def canon_srden(div, raw):
    """Division + raw free-text Sr.DEN label -> canonical 'DIV Sr.DEN/X' key, or None."""
    if div not in ('MYS', 'SBC', 'UBL') or not raw:
        return None
    m = re.search(r'DEN[./]?\s*/?\s*([A-Za-z]+)', str(raw), re.I)
    if not m:
        return None
    code = m.group(1).upper()
    if code not in ('N', 'S', 'E', 'W', 'C', 'HQ'):
        return None
    return f"{div} Sr.DEN/{code}"


def canon_ds_srden(div, zone):
    if not div or not zone:
        return None
    zl = ZONE_LETTER.get(str(zone).strip().upper())
    if not zl:
        return None
    return f"{div} Sr.DEN/{zl}"


def norm_srden_text(v):
    if not v:
        return None
    s = re.sub(r'\s+', ' ', str(v)).strip()
    s = s.replace('OL ', '').replace('Sr.DEN.', 'Sr.DEN/').replace('Sr.DEN. ', 'Sr.DEN/')
    s = s.replace(' /', '/').replace('/ ', '/')
    s = re.sub(r'Sr\.DEN\s*/\s*', 'Sr.DEN/', s)
    return s.rstrip('.').strip()


def parse_ddmmyyyy(d):
    if not d:
        return None
    d = str(d).strip().split(' ')[0]
    for fmt in ('%d.%m.%Y', '%d-%m-%Y', '%d/%m/%Y'):
        try:
            return datetime.datetime.strptime(d, fmt)
        except Exception:
            pass
    return None


def one_file(pattern, required=True):
    hits = sorted(glob.glob(pattern))
    hits = [h for h in hits if not os.path.basename(h).startswith('~$')]  # skip Excel lock files
    if not hits:
        if required:
            raise FileNotFoundError(f"No file matching {pattern}")
        return None
    # newest by mtime wins if more than one candidate
    return max(hits, key=os.path.getmtime)


# =====================================================================
# STEP 1: Fleet allocation + age profile (SWR Machine Details)
# =====================================================================
def extract_home_div(base):
    path = one_file(f"{base}/SWR Machine Details/*Original Allocation*.xlsx")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb['Sheet2'] if 'Sheet2' in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))

    def norm(name):
        if name is None:
            return None
        s = str(name).upper()
        s = re.sub(r'\(.*?\)', '', s)
        return re.sub(r'[^A-Z0-9]', '', s)

    home_div, bare_type_div = {}, {}
    for r in rows[1:]:
        for col_idx, div in [(1, 'MYS'), (2, 'SBC'), (3, 'UBL')]:
            if col_idx < len(r) and r[col_idx]:
                val = str(r[col_idx]).strip()
                home_div[norm(val)] = div
                if re.match(r'^[A-Za-z]+$', val):
                    bare_type_div[val.upper()] = div
    return home_div, bare_type_div, norm


def extract_age_profile(base):
    path = one_file(f"{base}/SWR Machine Details/*Comissioning*.xlsx", required=False) or \
        one_file(f"{base}/SWR Machine Details/*Commissioning*.xlsx", required=False)
    if not path:
        log("WARN: Machine Commissioning workbook not found - age_profile will be empty")
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    out = []
    for r in range(1, ws.max_row + 1):
        sno = ws.cell(row=r, column=1).value
        mtype = ws.cell(row=r, column=2).value
        machine = ws.cell(row=r, column=3).value
        cdate = ws.cell(row=r, column=4).value
        age_txt = ws.cell(row=r, column=5).value
        if not isinstance(sno, (int, float)) or not machine:
            continue
        agem = re.search(r'(\d+)', str(age_txt)) if age_txt else None
        age = int(agem.group(1)) if agem else None
        out.append(dict(MachineType=mtype, Machine=str(machine).strip(),
                         CommissioningDate=cdate, AgeYears=age))
    return out


# Canonical unit labels, keyed by a whitespace-stripped/uppercased lookup -
# the Cumulative Progress sheet's Unit column is hand-typed every update
# cycle and has shown up as 'Km'/'km', 'T/O'/'t/o', etc. Normalizing here
# means unitForType()/breakdownBy() on the dashboard side (which group
# purely by the literal unit string) merge same-unit machines correctly
# instead of splitting them into a spurious extra group over a casing slip.
UNIT_CANON = {'KM': 'Km', 'T/O': 'T/O', 'R/S': 'R/S', 'UNIT': 'Unit'}


def _norm_unit(v):
    if v is None:
        return v
    s = str(v).strip()
    return UNIT_CANON.get(s.upper(), s)


def _secondary_dict(sec):
    """Builds the client-facing `secondary` sub-object (BCM's Turnout/T/O
    line) from a raw `Secondary` dict attached by extract_cum_progress.
    Shared by the two places progress.machines records get built (the
    initial machines_base pass and the later tamping/status-correction/
    IOH-POH "v2 rebuild" pass) so both stay consistent."""
    if not sec:
        return None
    pct = (sec['ActualProg'] / sec['ProportionateTarget']) if sec.get('ProportionateTarget') else 0
    out = dict(unit=sec['Unit'], annualTarget=sec.get('AnnualTarget'), targetMonth=sec.get('TargetPerMonth'),
               targetToDate=sec.get('ProportionateTarget'), actualToDate=sec.get('ActualProg'), pct=round(pct, 4))
    for abbr in MONTH_ABBR:
        out[abbr.lower()] = sec.get(abbr)
    return out


def _num_or_none(v):
    """Keeps a real number as-is; anything else (blank, or a hand-typed
    annotation like 'New M/c', 'IOH', 'POH' sitting in a monthly-progress
    cell in place of a numeric value) becomes None rather than leaking a
    stray string into a field every downstream consumer assumes is numeric
    (confirmed against DUO-57434, whose April cell reads 'New M/c' since it
    was commissioned partway through the month)."""
    return v if isinstance(v, (int, float)) else None


def _cum_row_fields(r):
    """Pulls the common Unit/AnnualTarget/TargetPerMonth/month-values/
    ProportionateTarget/ActualProg/PctProg fields out of one Cumulative
    Progress row, shared by both a primary (Sl.No-bearing) row and a
    secondary continuation row (see below) since they're laid out
    identically from column D (Unit) onward."""
    unit = _norm_unit(r[3] if len(r) > 3 else None)
    annual_target = r[4] if len(r) > 4 else None
    target_month = r[5] if len(r) > 5 else None
    month_vals = {MONTH_ABBR[i]: _num_or_none(r[6 + i] if len(r) > 6 + i else None) for i in range(6)}
    prop_target = r[18] if len(r) > 18 else None
    actual_prog = r[19] if len(r) > 19 else None
    pct_prog = r[20] if len(r) > 20 else None
    out = {
        'Unit': unit, 'AnnualTarget': annual_target, 'TargetPerMonth': target_month,
        'ProportionateTarget': prop_target, 'ActualProg': actual_prog, 'PctProg': pct_prog,
    }
    out.update(month_vals)
    return out


# =====================================================================
# STEP 2: Cumulative Progress (current FY)
# =====================================================================
def extract_cum_progress(base, home_div, bare_type_div, norm):
    path = one_file(f"{base}/Cumulative Progress/FY 2026-2027/*.xlsx")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    # Latest month tab is normally the last sheet in the workbook.
    ws = wb[wb.sheetnames[-1]]
    rows = list(ws.iter_rows(values_only=True))

    records = []
    skip_labels = {'total (km)', 'total (t/o)', 'total', 'gt- cumulative', 'gt-cumulative'}
    secondary_count = 0
    for idx in range(3, len(rows)):
        r = rows[idx]
        slno = r[0] if len(r) > 0 else None
        machine = r[1] if len(r) > 1 else None
        if machine is None:
            continue
        label = str(slno).strip().lower() if slno is not None else ''
        if label in skip_labels or (isinstance(machine, str) and machine.lower().startswith('note')):
            continue
        if not isinstance(slno, (int, float)):
            continue
        div_cur = r[2] if len(r) > 2 else None
        fields = _cum_row_fields(r)
        mtype_match = re.match(r'^([A-Za-z]+)', str(machine).strip())
        mtype = mtype_match.group(1).upper() if mtype_match else 'UNKNOWN'
        hd = home_div.get(norm(machine)) or bare_type_div.get(mtype)
        if div_cur in ('MYS', 'SBC', 'UBL'):
            status = 'Active'
        elif div_cur in ('IOH', 'POH'):
            status = div_cur
        else:
            status = f'Deputed-{div_cur}'
        rec = {
            'SlNo': slno, 'Machine': str(machine).strip(), 'MachineType': mtype,
            'CurrentDIV': div_cur,
            'HomeDivision': hd if hd else (div_cur if div_cur in ('MYS', 'SBC', 'UBL') else 'UNKNOWN'),
            'Status': status,
        }
        rec.update(fields)
        # Some machine types (so far: BCM only) carry a SECOND target/progress
        # line immediately below their primary row, measured in a different
        # unit (e.g. BCM is tracked in Km *and* in Turnouts/T/O, since a BCM
        # also tamps turnouts as part of its duty) -- that continuation row has
        # no Sl.No/Machine/DIV of its own (blank), just a Unit value onward, so
        # it would otherwise be silently skipped by the `isinstance(slno, ...)`
        # check above. Detected and attached to the preceding machine's record
        # as `Secondary` (confirmed against the user's screenshot of the sheet:
        # BCM-351's 'T/O' row sits directly under its 'Km' row, 10 rows total
        # for 7 BCM x 2 units each, with its own 'Total (T/O)' subtotal row).
        nxt = rows[idx + 1] if idx + 1 < len(rows) else None
        if nxt is not None:
            na, nb, nc, nd = (nxt[0] if len(nxt) > 0 else None), (nxt[1] if len(nxt) > 1 else None), \
                (nxt[2] if len(nxt) > 2 else None), (nxt[3] if len(nxt) > 3 else None)
            if na is None and nb is None and nc is None and nd:
                rec['Secondary'] = _cum_row_fields(nxt)
                secondary_count += 1
        records.append(rec)
    log(f"Cumulative Progress: {len(records)} machine records from {os.path.basename(path)} / sheet {ws.title}"
        + (f" ({secondary_count} with a secondary-unit continuation row)" if secondary_count else ""))
    return records


# =====================================================================
# STEP 3: Daily Progress -> shortfall records (per-day category log)
# =====================================================================
class ColMerges:
    def __init__(self, ws):
        import bisect
        self.vert = {}
        self.horiz_rows = {}
        for mc in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = mc.bounds
            if min_col == max_col:
                self.vert.setdefault(min_col, []).append((min_row, max_row))
            else:
                for c in range(min_col, max_col + 1):
                    self.horiz_rows.setdefault(c, set()).update(range(min_row, max_row + 1))
        for c in self.vert:
            self.vert[c].sort()
        self._bisect = bisect

    def vertical_merge_for(self, row, col):
        lst = self.vert.get(col)
        if not lst:
            return None
        idx = self._bisect.bisect_right(lst, (row, float('inf'))) - 1
        if idx >= 0:
            mn, mx = lst[idx]
            if mn <= row <= mx:
                return (mn, mx)
        return None

    def is_horizontal_banner(self, row, col):
        s = self.horiz_rows.get(col)
        return bool(s and row in s)


def anchor_value(cm, rows_vals, row1, col1):
    m = cm.vertical_merge_for(row1, col1)
    r = m[0] if m else row1
    if r - 1 < 0 or r - 1 >= len(rows_vals):
        return None
    row = rows_vals[r - 1]
    return row[col1 - 1] if col1 - 1 < len(row) else None


def find_daily_progress_files(base):
    """Map month abbr -> filepath, matching whatever files currently exist (case/space tolerant).

    More than one file can match the same month (e.g. a stale copy left over from an
    earlier staging alongside a freshly re-staged one with slightly different spacing) -
    when that happens, the newest-by-mtime file wins rather than whichever the directory
    listing happens to return last."""
    d = f"{base}/Daily Progress"
    candidates = defaultdict(list)
    if not os.path.isdir(d):
        return {}
    for fn in os.listdir(d):
        if not fn.lower().endswith('.xlsx') or fn.startswith('~$'):
            continue
        low = re.sub(r'\s+', ' ', fn.strip().lower())
        for abbr in MONTH_ABBR:
            full = {'Apr': 'apr', 'May': 'may', 'Jun': 'jun', 'Jul': 'jul', 'Aug': 'aug', 'Sep': 'sep'}[abbr]
            if re.search(rf'\b{full}\w*\s*26\b', low) or re.search(rf'\b{full}\w*\s*2026\b', low):
                candidates[abbr].append(os.path.join(d, fn))
    out = {}
    for abbr, paths in candidates.items():
        if len(paths) > 1:
            log(f"WARN: multiple Daily Progress files matched {abbr}: {paths} - using the newest by mtime")
        out[abbr] = max(paths, key=os.path.getmtime)
    return out


def extract_shortfall_records(base):
    files = find_daily_progress_files(base)
    log("Daily Progress files found:", {k: os.path.basename(v) for k, v in files.items()})
    records = []
    for month, path in files.items():
        if month == 'Jun':
            continue  # June sheet uses the simple flat layout - handled separately below
        wb_vals = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sheet_name = 'Progress ' if 'Progress ' in wb_vals.sheetnames else wb_vals.sheetnames[0]
        ws_vals = wb_vals[sheet_name]
        max_col = ws_vals.max_column
        rows_vals = list(ws_vals.iter_rows(max_col=max_col, values_only=True))
        wb_vals.close()

        header_positions = []
        for i, row in enumerate(rows_vals):
            for j, v in enumerate(row):
                if isinstance(v, str) and v.strip().lower() == 'remarks':
                    header_positions.append((i, j))
        header_positions.sort()
        n = len(rows_vals)

        wb = openpyxl.load_workbook(path, read_only=False)
        ws = wb[sheet_name]
        cm = ColMerges(ws)
        wb.close()

        for idx, (hr0, hc0) in enumerate(header_positions):
            next_header_row0 = header_positions[idx + 1][0] if idx + 1 < len(header_positions) else n
            hdr_row1 = hr0 + 1
            remarks_col1 = hc0 + 1
            tag_col1 = hc0 + 2
            data_end1 = next_header_row0

            hdr_merge = cm.vertical_merge_for(hdr_row1, remarks_col1)
            header_bottom1 = hdr_merge[1] if hdr_merge else hdr_row1
            data_start1 = header_bottom1 + 1

            r = data_start1
            while r <= data_end1:
                if cm.is_horizontal_banner(r, remarks_col1):
                    r += 1
                    continue
                date_v = anchor_value(cm, rows_vals, r, 1)
                div_v = anchor_value(cm, rows_vals, r, 2)
                srden_v = anchor_value(cm, rows_vals, r, 3)
                mach_v = anchor_value(cm, rows_vals, r, 4)
                remarks_v = anchor_value(cm, rows_vals, r, remarks_col1)
                cat_v = anchor_value(cm, rows_vals, r, tag_col1)
                m = cm.vertical_merge_for(r, tag_col1) or cm.vertical_merge_for(r, remarks_col1)
                next_r = (m[1] + 1) if m else (r + 1)
                cat_source = 'tagged'
                if cat_v in (None, '') and remarks_v not in (None, ''):
                    # This month's file is missing the hand-injected "Shortfall Category"
                    # column (e.g. re-synced from source without it) - reconstruct from
                    # Remarks text instead of silently dropping the month's data.
                    cat_v = classify_shortfall(remarks_v)
                    cat_source = 'reconstructed'
                if cat_v not in (None, ''):
                    records.append(dict(Month=month, Date=str(date_v) if date_v else None,
                                         Division=div_v, SrDEN=norm_srden_text(srden_v), Machine=mach_v,
                                         Category=cat_v, Remarks=remarks_v, CategorySource=cat_source))
                r = next_r

    if 'Jun' in files:
        path = files['Jun']
        wb_vals = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sheet_name = 'Progress ' if 'Progress ' in wb_vals.sheetnames else wb_vals.sheetnames[0]
        ws_vals = wb_vals[sheet_name]
        rows_vals = list(ws_vals.iter_rows(max_col=12, values_only=True))
        wb_vals.close()
        for row in rows_vals[1:]:
            date_v, div_v, srden_v, mach_v = row[0], row[1], row[2], row[3]
            remarks_v = row[8] if len(row) > 8 else None
            cat_v = row[9] if len(row) > 9 else None
            cat_source = 'tagged'
            if cat_v in (None, '') and remarks_v not in (None, ''):
                cat_v = classify_shortfall(remarks_v)
                cat_source = 'reconstructed'
            if cat_v not in (None, ''):
                records.append(dict(Month='Jun', Date=str(date_v) if date_v else None,
                                     Division=div_v, SrDEN=norm_srden_text(srden_v), Machine=mach_v,
                                     Category=cat_v, Remarks=remarks_v, CategorySource=cat_source))

    for rec in records:
        rec['SrDENKey'] = canon_srden(rec.get('Division'), rec.get('SrDEN'))
    log("Shortfall/category records:", len(records))
    return records


# =====================================================================
# STEP 4: Exception Sheet (one workbook per day, all kept as a dict-by-day)
# =====================================================================
def extract_exceptions(base):
    files = sorted(glob.glob(f"{base}/Exception Sheet/*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    ex = {}
    for path in files:
        fn = os.path.basename(path)
        m = re.search(r'(\d{1,2})[.\-](\d{1,2})[.\-](\d{4})', fn)
        if not m:
            continue
        day, mon, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            dt = datetime.date(year, mon, day)
        except ValueError:
            continue
        key = dt.strftime('%d-%b').lstrip('0')
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet_name = 'Exception' if 'Exception' in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        rows = []
        cur_cat = None
        r = 3
        while r <= ws.max_row:
            desc = ws.cell(row=r, column=2).value
            nos = ws.cell(row=r, column=3).value
            machine = ws.cell(row=r, column=4).value
            div = ws.cell(row=r, column=5).value
            loc = ws.cell(row=r, column=6).value
            remarks = ws.cell(row=r, column=7).value
            since = ws.cell(row=r, column=9).value
            tentative = ws.cell(row=r, column=10).value
            days_under = ws.cell(row=r, column=11).value
            if desc is not None and str(desc).strip():
                cur_cat = str(desc).strip()
            if machine:
                rows.append(dict(
                    Category=cur_cat, Machine=str(machine).strip(), Division=div, Location=loc,
                    Remarks=remarks,
                    UnderSince=(since.strftime('%d.%m.%Y') if isinstance(since, (datetime.date, datetime.datetime)) else since),
                    Tentative=(tentative.strftime('%d.%m.%Y') if isinstance(tentative, (datetime.date, datetime.datetime)) else tentative),
                    DaysUnder=(days_under if isinstance(days_under, (int, float)) else None),
                ))
            r += 1
        ex[key] = rows
    log("Exception Sheet days:", sorted(ex.keys()), {k: len(v) for k, v in ex.items()})
    return ex


# =====================================================================
# STEP 5: IOH/POH Planning — PRIMARY SOURCE for Fleet Health (revised
# 01-Oct-2026). See the big comment block right above
# extract_ioh_poh_letter() further down for the full design.
# =====================================================================
def extract_ioh_poh(base):
    path = one_file(f"{base}/IOH POH Planning/*.xlsx", required=False)
    if not path:
        log("WARN: IOH POH Planning workbook not found - ioh_poh_records will be empty")
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    # locate the header/year-block row (contains 'NEXT DUE in Months') to find schedule column starts,
    # and the 'Remark' column (its position varies with how many FY schedule blocks are printed).
    year_headers = {}
    remark_col = None
    for r in range(1, 6):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and re.match(r'^\d{4}-\d{2,4}$', v.strip()):
                year_headers[c] = v.strip()
            if isinstance(v, str) and v.strip().lower() == 'remark':
                remark_col = c
    out = []
    # NOTE (fixed 01-Oct-2026): data genuinely starts at row 4 in the current workbook
    # layout (row 1 = main header, row 2 = sub-header, row 3 = a literal column-number
    # footer row "1,2,3..."; row 4 = the first machine, CSM-939). The loop used to start
    # at row 5, silently skipping that first machine from every IOH/POH-sourced figure
    # (overdue table, Fleet Health schedule, etc.) - caught while rebuilding Fleet Health
    # against this same workbook as the explicit primary source.
    for r in range(4, ws.max_row + 1):
        sno = ws.cell(row=r, column=1).value
        machine = ws.cell(row=r, column=2).value
        if not machine:
            continue
        div = ws.cell(row=r, column=3).value
        commission = ws.cell(row=r, column=4).value
        last_ioh_eh = ws.cell(row=r, column=5).value
        last_ioh_date = ws.cell(row=r, column=6).value
        last_poh_eh = ws.cell(row=r, column=7).value
        last_poh_date = ws.cell(row=r, column=8).value
        eh_now = ws.cell(row=r, column=9).value
        eh_run_month = ws.cell(row=r, column=15).value
        next_ioh = ws.cell(row=r, column=16).value
        next_poh = ws.cell(row=r, column=17).value

        def fmt_d(v):
            return v.strftime('%d.%m.%Y') if isinstance(v, (datetime.date, datetime.datetime)) else v

        schedule = {}
        for c, yr in year_headers.items():
            ioh_v = ws.cell(row=r, column=c).value
            poh_v = ws.cell(row=r, column=c + 1).value
            schedule[yr] = [ioh_v, poh_v]
        remark = None
        if remark_col:
            rv = ws.cell(row=r, column=remark_col).value
            if isinstance(rv, str) and rv.strip():
                remark = rv.strip()

        def fmt_commission(v):
            if isinstance(v, (datetime.date, datetime.datetime)):
                return v.strftime('%d-%m-%Y')
            return v

        # Data-quality fix (01-Oct-2026): one row (PCTM-9062) has a stray date value
        # ('09.01.2026', from a merged/shifted cell upstream in the source workbook)
        # in the Division column instead of MYS/SBC/UBL. Division gets backfilled from
        # the Cumulative Progress home-division map in run() once that's available -
        # here it's left as-is if it isn't a recognised division code, rather than
        # silently treated as a real division.
        div_clean = div if div in ('MYS', 'SBC', 'UBL') else None

        out.append(dict(
            SNo=sno, Machine=str(machine).strip(), Division=div_clean, DivisionRaw=div,
            Commission=fmt_commission(commission),
            LastIOH_EH=last_ioh_eh, LastIOH_Date=fmt_d(last_ioh_date),
            LastPOH_EH=last_poh_eh, LastPOH_Date=fmt_d(last_poh_date),
            EH_Now=eh_now, EH_Run_Month=eh_run_month,
            NextDueIOH_Months=next_ioh, NextDuePOH_Months=next_poh,
            Schedule=schedule, Remark=remark,
        ))
    log("IOH/POH Planning records:", len(out), "source=", os.path.basename(path))
    return out


def extract_ioh_poh_duration_norms(base):
    """Second sheet of the same IOH/POH Details 27-30 workbook ('IOH POH
    Duration') - per-machine-type IRTMM service intervals (IOH/POH cadence in
    days, and the Engine-Hour thresholds that trigger each). Shown in Fleet
    Health as reference context alongside the actual schedule - purely
    informational, not used in any computation."""
    path = one_file(f"{base}/IOH POH Planning/*.xlsx", required=False)
    if not path:
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet_name = None
    for name in wb.sheetnames:
        if 'duration' in name.lower():
            sheet_name = name
            break
    if not sheet_name:
        return []
    ws = wb[sheet_name]
    out = []
    for r in range(4, ws.max_row + 1):
        mtype = ws.cell(row=r, column=1).value
        if not mtype:
            continue
        out.append(dict(
            type=str(mtype).strip(),
            iohDays=ws.cell(row=r, column=2).value,
            firstPohDays=ws.cell(row=r, column=3).value,
            subsequentPohDays=ws.cell(row=r, column=4).value,
            iohHrs=ws.cell(row=r, column=5).value,
            pohHrs=ws.cell(row=r, column=6).value,
        ))
    log("IOH/POH Duration norms:", len(out))
    return out


_IOHPOH_MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'april': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'sept': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def parse_iohpoh_month_sort(s):
    """Best-effort chronological sort key for the letter's free-text month
    column ('Aug.26', 'Jun-26', 'Aug- 26', 'Aug-sept.26', 'Jul.26-Aug.26') -
    takes the FIRST month/year mentioned. Unparseable text sorts last rather
    than erroring, consistent with this pipeline's general approach to messy
    narrative fields."""
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
    """Free-text remark on a current-year IOH/POH schedule row -> a clean
    status bucket. Mirrors the classify_exc_category / classify_shortfall /
    classify_infra_responsibility pattern used elsewhere in this pipeline:
    priority-ordered regex rules with a safe fallback (the raw text) rather
    than silently dropping a genuinely new phrasing.

    'Overdue by N months[. ...]' is the single most common remark in the
    05-Oct-2026 "Refined letter to Divisions.docx" (13 of 38 main-schedule
    rows, 2 of 10 CPOH/RYP rows) and previously fell all the way through to
    the raw-text fallback, since nothing here recognized the word "overdue"
    -- every differently-worded "Overdue by N months" row became its own
    unique status string, which silently broke any code that counts/filters
    by status (confirmed 08-Oct-2026: the Fleet Health Sr.DEN release banner
    was showing only 3 pending-release machines fleet-wide when it should
    have shown 23). Checked explicitly for this phrase before the 'due for'
    check below, since 'overdue' and 'due for' are different urgency levels
    and must not collapse into the same bucket."""
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
    """Pulls the numeric month count out of an 'Overdue by N months...'
    remark (used to sort/display IOH/POH release-priority machines by how
    overdue they are); returns None for any other remark shape."""
    m = re.search(r'overdue\s*by\s*(\d+(?:\.\d+)?)\s*month', str(remark or ''), re.I)
    return float(m.group(1)) if m else None


# =====================================================================
# STEP 5b: IOH/POH — current-year authoritative schedule, from the Dy.CE/TM
# office LETTER (not the Exception Sheet, not Cumulative Progress Status).
#
# Added 01-Oct-2026 per the user's explicit instruction: "In that Excel file,
# there are IOH POH details for 27 to 30, and also one letter is there for
# this current year, IOH POH 2026 to 2027... for the current year, you
# follow this letter... for the further coming years and for all the
# machines, you follow strictly the IOH POH details of 27 to 30... It should
# be the primary source for it, not any other sheet."
#
# The "IOH POH Planning" folder holds a Word letter named like
# "IOH POH <FY>.docx" (currently "IOH POH 2026-2027.docx", No.SWR/TM.506/IOH
# & POH, dated 01.10.2026) whose Table 2 is the actual division-by-division,
# machine-by-machine IOH/POH schedule for the current FY - Division, Machine,
# Schedule (IOH or POH), Tentative Month, Duration (days), location
# (ZBD/YPR or STMRD/UBL), and a free-text status remark ("IOH Completed",
# "Yet to release from division", "Under IOH attention", etc). Table 1 is a
# secondary, narrower "POH release priority" list tied to an even earlier
# letter - kept separately rather than merged into the main schedule, since
# it is a different (smaller, overlapping-but-not-identical) machine set.
#
# Never hardcode the filename to this year's: the folder will hold a new
# "IOH POH <next-FY>.docx" letter every year, so the current-FY letter is
# picked by matching the FY computed from as_of against the filename/subject
# line, falling back to the newest-by-mtime docx in the folder.
# =====================================================================
def extract_ioh_poh_letter(base, as_of):
    folder = f"{base}/IOH POH Planning"
    candidates = [p for p in glob.glob(f"{folder}/*.docx") if not os.path.basename(p).startswith('~$')]
    if not candidates:
        log("WARN: no .docx files found in IOH POH Planning - currentYear schedule will be empty")
        return None
    fy_year = as_of.year if as_of.month >= 4 else as_of.year - 1
    fy_next_2 = str(fy_year + 1)[-2:]
    fy_next_4 = str(fy_year + 1)
    # Match both the shorthand ("2026-27") and full-year ("2026-2027") spellings
    # a filename might use for the FY - \D{0,2} alone only covers the shorthand,
    # since "2026-2027" has the digits "20" (not just punctuation) between the
    # "2026" and the trailing "27".
    fy_pat = re.compile(rf'{fy_year}\D{{0,2}}(?:{fy_next_4}|{fy_next_2})')
    # Prefer a filename that (a) matches the current FY and (b) looks like the
    # per-machine schedule letter rather than a POH-proposal covering letter -
    # "ioh" and "poh" both appearing in the name is the simplest reliable signal
    # (the 2027-30 covering letter is named "Lr to CPOH RYP_POH planing...").
    fy_matches = [p for p in candidates if fy_pat.search(os.path.basename(p))]
    schedule_like = [p for p in fy_matches if 'ioh' in os.path.basename(p).lower() and 'poh' in os.path.basename(p).lower()]
    # Fallback tier if the FY regex somehow misses (unexpected naming next year):
    # any candidate with both "ioh" and "poh" in its name is still a much
    # stronger signal than raw mtime (which, after staging/copying files, no
    # longer reliably reflects the original file's real modification time).
    ioh_poh_named = [p for p in candidates if 'ioh' in os.path.basename(p).lower() and 'poh' in os.path.basename(p).lower()]
    filename_tier = (schedule_like or fy_matches or ioh_poh_named or candidates)

    try:
        import docx
    except ImportError:
        log("WARN: python-docx not installed - cannot parse IOH/POH letter, currentYear schedule will be empty")
        return None

    # Content-based override (added 06-Oct-2026): a revision of this same letter can
    # arrive under a filename that doesn't contain "ioh"/"poh" at all - e.g. a
    # "Refined letter to Divisions.docx" superseding "IOH POH 2026-2027.docx" a few
    # days later, same letter number (SWR/TM.506/IOH & POH), newer Date:. Filename
    # keyword/FY matching alone would keep picking the older, superseded letter
    # forever in that case. So: open every .docx candidate in the folder, keep only
    # those that actually have the per-machine schedule table signature ('name of
    # machine' + 'tentative month' header - the same signature schedule_table
    # detection below relies on), and among those prefer the one whose own "Date:"
    # paragraph parses latest; only fall back to filename-tier + mtime when no
    # candidate's content can be read/matched (corrupted file, unexpected layout).
    def _peek_letter_date_and_signature(p):
        try:
            dd = docx.Document(p)
        except Exception:
            return None, False
        has_schedule_table = False
        for t in dd.tables:
            if not t.rows:
                continue
            h = ' '.join(c.text.strip() for c in t.rows[0].cells).lower()
            if 'name of machine' in h and 'tentative month' in h:
                has_schedule_table = True
                break
        dt = None
        for para in dd.paragraphs[:4]:
            m = re.search(r'Date:\s*([\d.]+)', para.text)
            if m:
                dt = parse_ddmmyyyy(m.group(1))
                break
        return dt, has_schedule_table

    content_matches = []  # (path, parsed_date_or_None)
    for p in candidates:
        dt, has_sig = _peek_letter_date_and_signature(p)
        if has_sig:
            content_matches.append((p, dt))

    if content_matches:
        dated = [(p, dt) for p, dt in content_matches if dt is not None]
        if dated:
            path = max(dated, key=lambda x: x[1])[0]
        else:
            path = max((p for p, _ in content_matches), key=os.path.getmtime)
    else:
        path = max(filename_tier, key=os.path.getmtime)

    d = docx.Document(path)
    letter_date, fy_label = None, None
    for p in d.paragraphs[:4]:
        m = re.search(r'Date:\s*([\d.]+)', p.text)
        if m:
            letter_date = m.group(1)
    for p in d.paragraphs[:16]:
        m = re.search(r'for the year\s*([\d\-]+)', p.text, re.I)
        if m:
            fy_label = m.group(1)
            break

    schedule_table, priority_table = None, None
    for t in d.tables:
        if not t.rows:
            continue
        headers_joined = ' '.join(c.text.strip() for c in t.rows[0].cells).lower()
        if 'name of machine' in headers_joined and 'tentative month' in headers_joined:
            schedule_table = t
        elif 'planned date of arrival' in headers_joined:
            priority_table = t

    events = []
    if schedule_table:
        cur_div = None
        for row in schedule_table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if not any(cells):
                continue
            # A division sub-header row repeats the division name across every cell
            # ("SBC Division" | "SBC Division" | ... ) rather than holding real data.
            if cells[0] and cells[0] == cells[1] == cells[2]:
                cur_div = re.sub(r'\s*division\s*$', '', cells[0], flags=re.I).strip().upper()
                continue
            row8 = (cells + [''] * 8)[:8]
            sno, machine, div, sched, month, duration, location, remark = row8
            if not machine:
                continue
            div = (div or cur_div or '').strip().upper()
            sched_type = (sched or '').strip().upper()
            sched_type = 'IOH' if sched_type.startswith('IOH') else ('POH' if sched_type.startswith('POH') else (sched_type or None))
            events.append(dict(
                sno=sno, machine=machine, machineKey=machine_key(machine), division=div if div in ('MYS', 'SBC', 'UBL') else None,
                type=sched_type, month=month, monthSort=list(parse_iohpoh_month_sort(month)),
                durationDays=duration, location=location, remark=remark,
                status=classify_iohpoh_status(remark), overdueMonths=extract_overdue_months(remark),
            ))

    priority_events = []
    if priority_table:
        for row in priority_table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            row5 = (cells + [''] * 5)[:5]
            # Column order confirmed 08-Oct-2026 by reading the actual docx table headers:
            # Sl.No. | Machine No. | Division | Planned date of arrival at CPOH/RYP | Remarks.
            # This used to unpack as (sno, machine, planned_date, remark, div) -- one column
            # short of correct -- which silently fed the real "Division" text into planned_date
            # (so parse_ddmmyyyy() always failed and `month` fell back to showing a division
            # code like "UBL" instead of a month), fed the real planned-date string into
            # `remark` (so classify_iohpoh_status() never saw real remark text like "Overdue by
            # N months" and fell straight to its raw-text fallback, making every one of these
            # 10 CPOH/RYP rows invisible to any status-based filter -- including the Fleet
            # Health Sr.DEN release banner), and fed the real remark text into `division`
            # (so these rows' jurisdiction/srden lookup was also garbage). This was the actual
            # root cause of the user-reported near-empty Sr.DEN banner, not just the missing
            # 'overdue' branch in classify_iohpoh_status().
            sno, machine, div, planned_date, remark = row5
            if not machine:
                continue
            # This table ("Planned date of arrival at CPOH/RYP") is the schedule for
            # machines sent to the Central POH workshop - functionally a third
            # IOH/POH location alongside ZBD/YPR and STMRD/UBL, always a POH (CPOH
            # only performs POH, not IOH). Added 05-Oct-2026: these events get
            # month/monthSort/type/location fields matching the main schedule_table
            # events so run() can fold them into the same current-year events list
            # (and therefore into completedCount/byDivision/the Completed drawer)
            # instead of only existing in this separate priority-list view - a real
            # gap the user caught (a POH "Completed" here wasn't reflected anywhere
            # in the Fleet Health completed counts).
            planned_dt = parse_ddmmyyyy(planned_date)
            priority_events.append(dict(
                sno=sno, machine=machine, machineKey=machine_key(machine),
                division=(div or '').strip().upper() or None, plannedDate=planned_date,
                month=(planned_dt.strftime('%b-%y') if planned_dt else planned_date),
                monthSort=([planned_dt.year, planned_dt.month] if planned_dt else [9999, 99]),
                type='POH', location='CPOH/RYP', durationDays=None,
                remark=remark, status=classify_iohpoh_status(remark), overdueMonths=extract_overdue_months(remark),
            ))

    log(f"IOH/POH letter: {len(events)} current-year schedule events, "
        f"{len(priority_events)} priority-release entries, source={os.path.basename(path)}")
    return dict(sourceFile=os.path.basename(path), letterDate=letter_date, fy=(fy_label or f"{fy_year}-{str(fy_year+1)[-2:]}"),
                events=events, priorityEvents=priority_events)


# =====================================================================
# STEP 5c: CPOH/RYP forward POH planning (added 05-Oct-2026, in response to
# the user asking why POH jobs completed at CPOH/RYP weren't visible anywhere,
# and separately for "CPOH planning").
#
# The "IOH POH Planning" folder holds a second, distinct letter alongside the
# current-year schedule letter: "Lr to CPOH RYP_POH planing_2027-30.docx" -
# Dy.CE/TM's own letter TO Dy.CE/CPOH/RYP (the Central POH workshop at
# Yeshwanthpur) proposing which machines should go there for POH across the
# next three financial years. It has five data tables:
#   1. "Details of Machine proposed for POH during 2027-28" - the firm,
#      detailed proposal: last POH date, engine-hours since last POH,
#      cumulative progress, governing reason, per-engine make/model (a machine
#      can have 2 engines, each on its own row sharing the same Sl.No - these
#      are grouped back into one record per machine, each with an `engines`
#      list), and the probable month the machine will be spared for POH.
#   2 & 3. Tentative lists for 2028-29 and 2029-30 - lighter detail (machine,
#      type, last-POH date/status only), since these years are still
#      provisional.
#   4 & 5. Tamping-unit demand for 2027-28 and 2028-29 - which tamping units
#      specific machines will need and when, separate from the POH schedule
#      itself but travels in the same letter.
#
# Table 1 has TWO header rows in the docx (a merged "Engine details" row above
# the real column names); tables 4-5 likewise have two header rows (a merged
# "Required Tamping unit" row above "CPOH make"/"Plasser make"). Tables 2-3
# have one header row each. Rather than hardcode row-skip counts (fragile if
# the letter's header layout shifts next year), every row is validated by its
# own leading Sl.No cell being a plain number - a row that fails that check
# (a stray header repeat, a blank spacer row) is silently skipped rather than
# included as garbage data.
# =====================================================================
def extract_cpoh_ryp_planning(base):
    folder = f"{base}/IOH POH Planning"
    candidates = [p for p in glob.glob(f"{folder}/*.docx") if not os.path.basename(p).startswith('~$')]
    cpoh_candidates = [p for p in candidates if 'cpoh' in os.path.basename(p).lower()]
    if not cpoh_candidates:
        log("WARN: no CPOH/RYP planning letter (filename containing 'CPOH') found in IOH POH Planning - cpohPlanning will be empty")
        return None
    path = max(cpoh_candidates, key=os.path.getmtime)

    try:
        import docx
    except ImportError:
        log("WARN: python-docx not installed - cannot parse CPOH/RYP planning letter")
        return None

    d = docx.Document(path)
    letter_date = None
    for p in d.paragraphs[:4]:
        m = re.search(r'Date:\s*([\d.]+)', p.text)
        if m:
            letter_date = m.group(1)

    def is_snum(cell):
        return bool(re.match(r'^\d+\.?$', (cell or '').strip()))

    def find_table(required_header_substrings):
        for t in d.tables:
            if not t.rows:
                continue
            h = ' '.join(c.text.strip() for c in t.rows[0].cells).lower()
            if all(k in h for k in required_header_substrings):
                return t
        return None

    proposed_table = find_table(['governing reason'])
    tentative_tables = [t for t in d.tables if t.rows and 'type of m/c' in ' '.join(c.text.strip() for c in t.rows[0].cells).lower()]
    tentative_2829_table = tentative_tables[0] if len(tentative_tables) >= 1 else None
    tentative_2930_table = tentative_tables[1] if len(tentative_tables) >= 2 else None
    tamping_tables = [t for t in d.tables if t.rows and 'tamping unit' in ' '.join(c.text.strip() for c in t.rows[0].cells).lower()]

    # Table 1: group consecutive rows sharing the same Sl.No into one record
    # per machine (handles machines with 2 engines, e.g. DTE 9006).
    proposed_raw = []
    if proposed_table:
        for row in proposed_table.rows:
            cells = [c.text.strip() for c in row.cells]
            if not is_snum(cells[0] if cells else ''):
                continue
            row9 = (cells + [''] * 9)[:9]
            proposed_raw.append(row9)
    proposed = []
    by_sno = {}
    for sno, machine, last_poh, eh_since, cum_prog, reason, make_model, eh2, sparing in proposed_raw:
        if not machine:
            continue
        rec = by_sno.get(sno)
        if rec is None:
            rec = dict(sno=sno, machine=machine, machineKey=machine_key(machine),
                       lastPohDate=last_poh, ehSinceLastPoh=eh_since, cumProgress=cum_prog,
                       governingReason=reason, probableSparingMonth=sparing, engines=[])
            by_sno[sno] = rec
            proposed.append(rec)
        rec['engines'].append(dict(makeModel=make_model, ehSinceLastPoh=eh2))

    def simple_rows(t, n):
        out = []
        if not t:
            return out
        for row in t.rows:
            cells = [c.text.strip() for c in row.cells]
            if not is_snum(cells[0] if cells else ''):
                continue
            out.append((cells + [''] * n)[:n])
        return out

    tentative_2829 = []
    for sno, mcno, typ, last_poh in simple_rows(tentative_2829_table, 4):
        machine = f"{typ} {mcno}".strip()
        if not mcno:
            continue
        tentative_2829.append(dict(sno=sno, machine=machine, machineKey=machine_key(machine), type=typ, lastPohDate=last_poh))

    tentative_2930 = []
    for sno, mcno, typ, last_poh in simple_rows(tentative_2930_table, 4):
        machine = f"{typ} {mcno}".strip()
        if not mcno:
            continue
        tentative_2930.append(dict(sno=sno, machine=machine, machineKey=machine_key(machine), type=typ, lastPohDetails=last_poh))

    # Tamping-unit demand tables aren't keyed by Sl.No - validate instead on a
    # non-empty, non-placeholder ("-") machine number.
    tamping_fy_labels = ['2027-28', '2028-29']
    tamping = []
    for i, t in enumerate(tamping_tables):
        fy_label = tamping_fy_labels[i] if i < len(tamping_fy_labels) else None
        for row in t.rows:
            cells = [c.text.strip() for c in row.cells]
            row5 = (cells + [''] * 5)[:5]
            typ, mcno, month, req1, req2 = row5
            if not mcno or mcno == '-' or typ.lower() == 'type of tamping unit':
                continue
            tamping.append(dict(forFy=fy_label, tampingType=typ, machine=mcno,
                                 month=month, requiredMake=(req2 or req1 or None)))

    log(f"CPOH/RYP planning letter: {len(proposed)} proposed 2027-28, {len(tentative_2829)} tentative 2028-29, "
        f"{len(tentative_2930)} tentative 2029-30, {len(tamping)} tamping-unit demand rows, source={os.path.basename(path)}")
    return dict(sourceFile=os.path.basename(path), letterDate=letter_date,
                proposed2728=proposed, tentative2829=tentative_2829, tentative2930=tentative_2930,
                tampingDemand=tamping)


# =====================================================================
# STEP 6: Double-shift roster (Monthly Machine And Double Shift)
# =====================================================================
def to_hours(v):
    if v is None or v in ('', '-'):
        return None
    if isinstance(v, (datetime.time, datetime.datetime)):
        return v.hour + v.minute / 60 + v.second / 3600
    if isinstance(v, (int, float)):
        h = int(v)
        frac = round((v - h) * 100)
        return h + frac / 60
    if isinstance(v, str):
        m = re.match(r'^(\d{1,2})[.:](\d{1,2})$', v.strip())
        if m:
            return int(m.group(1)) + int(m.group(2)) / 60
        try:
            return float(v)
        except Exception:
            return None
    return None


def display_time(v):
    if v is None or v == '':
        return None
    if v == '-':
        return '-'
    if isinstance(v, (datetime.time, datetime.datetime)):
        return f"{v.hour:02d}:{v.minute:02d}"
    if isinstance(v, (int, float)):
        h = int(v)
        frac = round((v - h) * 100)
        return f"{h:02d}:{frac:02d}"
    return str(v)


def extract_double_shift(base):
    # The user split the old combined "Monthly Machine And Double Shift" folder into two
    # separate folders on 07-Oct-2026: "Monthly Machine" (unrelated monthly progress data,
    # not used by this function) and "Double shift" (double-shift roster data's new home --
    # user confirmed 08-Oct-2026: "I have made a separate folder in the machine monitoring
    # folder named double shift... From now onwards, I will update double shift in that
    # separate folder... that will be the source sheet for double shift details"). The first
    # file placed there, "Double shift machine sheet 07.10.26.xlsx", is a single consolidated
    # workbook covering many weeks of dated blocks (not just one month), same sheet name
    # ('MODIFIED SHEET') and column layout as before -- no extraction-logic changes needed,
    # only the source folder. Still fall back to the old combined folder so an older/archived
    # snapshot of the folder tree continues to work.
    path = one_file(f"{base}/Double shift/*.xlsx", required=False)
    if not path:
        path = one_file(f"{base}/Monthly Machine And Double Shift/*.xlsx", required=False)
    if not path:
        log("WARN: no Double shift workbook found (checked 'Double shift/' and the legacy "
            "'Monthly Machine And Double Shift/') - doubleShift will be empty")
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    # The railway staff have renamed this sheet before (e.g. 'Double shift' -> 'MODIFIED SHEET')
    # without changing its internal layout, so match by known names first, then fall back to
    # scanning every sheet for the distinctive "DOUBLE SHIFT DATED" banner text.
    sheet_name = None
    for candidate in ('Double shift', 'MODIFIED SHEET', 'Double Shift'):
        if candidate in wb.sheetnames:
            sheet_name = candidate
            break
    if sheet_name is None:
        for name in wb.sheetnames:
            probe = wb[name]
            for r in range(1, min(probe.max_row, 6) + 1):
                v = probe.cell(row=r, column=1).value
                if isinstance(v, str) and 'DOUBLE SHIFT DATED' in v.upper():
                    sheet_name = name
                    break
            if sheet_name:
                break
    if sheet_name is None:
        log("WARN: no double-shift sheet found (tried known names + content scan) in", path, "- sheets present:", wb.sheetnames)
        return []
    if sheet_name != 'Double shift':
        log(f"NOTE: double-shift data found in sheet '{sheet_name}' (not the usual 'Double shift' name) in {path}")
    ws = wb[sheet_name]
    date_re = re.compile(r'DOUBLE SHIFT DATED\s+([\d/.]+)', re.I)

    blocks, cur_date, r, maxr = [], None, 1, ws.max_row
    while r <= maxr:
        a = ws.cell(row=r, column=1).value
        if isinstance(a, str) and 'DOUBLE SHIFT' in a.upper():
            m = date_re.search(a)
            cur_date = m.group(1) if m else a
            r += 2
            continue
        sl = ws.cell(row=r, column=1).value
        div = ws.cell(row=r, column=2).value
        srden = ws.cell(row=r, column=3).value
        section = ws.cell(row=r, column=4).value
        mc = ws.cell(row=r, column=5).value
        shift = ws.cell(row=r, column=6).value
        started = ws.cell(row=r, column=7).value
        ready = ws.cell(row=r, column=8).value
        arrived = ws.cell(row=r, column=9).value
        block_t = ws.cell(row=r, column=10).value
        progress = ws.cell(row=r, column=11).value
        staff = ws.cell(row=r, column=12).value
        if sl is None and div is None and srden is None and section is None and mc is None and shift is None:
            r += 1
            continue
        if isinstance(shift, str) and shift.strip().upper() in ('DAY', 'NIGHT'):
            blocks.append(dict(
                date=cur_date, sl=sl, div=div, srden=srden, section=section, machine=mc,
                shift=shift.strip().upper(),
                started_raw=display_time(started), ready_raw=display_time(ready), arrived_raw=display_time(arrived),
                started_hr=to_hours(started), ready_hr=to_hours(ready), arrived_hr=to_hours(arrived),
                blockTimings=block_t, progress=(str(progress) if progress is not None else None),
                staff=(str(staff) if staff is not None else None),
            ))
        r += 1

    filled, last = [], {}
    for b in blocks:
        for k in ('div', 'srden', 'section', 'machine'):
            if b[k] is None and k in last:
                b[k] = last[k]
            elif b[k] is not None:
                last[k] = b[k]
        filled.append(b)

    def duration_hours(started_hr, arrived_hr):
        if started_hr is None or arrived_hr is None:
            return None
        d = arrived_hr - started_hr
        if d < 0:
            d += 24
        return round(d, 2)

    for b in filled:
        b['durationHours'] = duration_hours(b['started_hr'], b['arrived_hr'])
        b['over12h'] = bool(b['durationHours'] is not None and b['durationHours'] > 12)
    log("Double-shift records:", len(filled), "| >12h:", sum(1 for b in filled if b['over12h']))
    return filled


# =====================================================================
# STEP 7: Prior-year cumulative (for YoY) — best effort, static prior FY
# =====================================================================
def extract_prior_year_by_type(base):
    path = one_file(f"{base}/Cumulative Progress/FY 2025-2026/*.xlsx", required=False)
    if not path:
        log("WARN: FY2025-26 workbook not found - YoY comparison will be skipped")
        return None, 0.0, None, None, None
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[-1]]
    prior_by_type = defaultdict(float)
    prior_by_machine = {}  # machine_key(name) -> sum, for the per-machine YoY breakdown
    prior_total = 0.0
    # Secondary-unit (so far: BCM's Turnouts/T/O) continuation rows, same
    # layout quirk as the current-year sheet (extract_cum_progress): no
    # Sl.No/Machine/DIV of their own, just Unit onward, directly below their
    # machine's primary row - accumulated separately so a machine's Km and
    # T/O prior-year progress are never blended into one meaningless sum.
    prior_by_type_secondary = defaultdict(float)
    prior_by_machine_secondary = {}
    last_name = None
    for r in range(4, ws.max_row + 1):
        sl = ws.cell(row=r, column=1).value
        name = ws.cell(row=r, column=2).value
        unit_cell = ws.cell(row=r, column=4).value
        if (not isinstance(sl, (int, float)) or not name):
            div_cell = ws.cell(row=r, column=3).value
            if sl is None and name is None and div_cell is None and unit_cell and last_name:
                vals = [ws.cell(row=r, column=c).value for c in range(6, 12)]
                s = sum(v for v in vals if isinstance(v, (int, float)))
                unit = _norm_unit(unit_cell)
                t = f"{disp_type(mtype_of(last_name))} ({unit})"
                prior_by_type_secondary[t] += s
                prior_by_machine_secondary[f"{machine_key(last_name)}|{unit}"] = \
                    prior_by_machine_secondary.get(f"{machine_key(last_name)}|{unit}", 0) + s
            continue
        vals = [ws.cell(row=r, column=c).value for c in range(6, 12)]
        s = sum(v for v in vals if isinstance(v, (int, float)))
        t = disp_type(mtype_of(name))
        prior_by_type[t] += s
        prior_total += s
        prior_by_machine[machine_key(name)] = prior_by_machine.get(machine_key(name), 0) + s
        last_name = name
    return prior_by_type, prior_total, prior_by_machine, prior_by_type_secondary, prior_by_machine_secondary


# =====================================================================
# STEP 8: 53-Point Track Machine Action Plan (added 29-Sep-2026)
#
# A separate workbook the user maintains and re-shares roughly every 15 days: 53 fixed
# action points (grouped into 8 categories - SAFETY, RELIABILITY, QUALITY, AVAILABILITY,
# EXPENDITURE, STAFF WELFARE, SKILL DEVELOPMENT, MISCELANEOUS), each with a narrative
# free-text "Progress up to <date>" column that gets a NEW column appended to the right
# every review cycle rather than the old one being overwritten - so the sheet accumulates
# the full history over time. Never hardcode which columns exist: auto-discover every
# column whose row-1 header starts with "Progress up to" so this keeps working as future
# 15-day updates add more columns, exactly like the rest of this pipeline auto-discovers
# monthly/dated files instead of hardcoding filenames.
# =====================================================================
def extract_action_plan_53(base):
    path = one_file(f"{base}/53 Point TM Action Plan/*.xlsx", required=False)
    if not path:
        log("WARN: 53 Point TM Action Plan workbook not found - actionPlan53 will be empty")
        return None
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    header1 = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    header2 = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]

    sno_col, cat_col, action_col = 1, 2, 3

    # "Progress up to <date>" columns - auto-discovered, not hardcoded.
    progress_cols = []  # (col_index_1based, date_label_raw, parsed_date_or_None)
    for c, v in enumerate(header1, start=1):
        if isinstance(v, str) and v.strip().lower().startswith('progress up to'):
            date_txt = v.strip()[len('progress up to'):].strip()
            progress_cols.append((c, date_txt, parse_ddmmyyyy(date_txt)))
    first_progress_col = progress_cols[0][0] if progress_cols else (ws.max_column + 1)

    # Supervisor columns sit between the Action Point column and the first progress column
    # (row-2 sub-header gives each one's label, e.g. "SBC\n  Division" -> "SBC Division").
    sup_cols = []  # (col_index, label)
    for c in range(4, first_progress_col):
        v2 = header2[c - 1] if c - 1 < len(header2) else None
        label = re.sub(r'\s+', ' ', str(v2)).strip() if v2 else None
        if label:
            sup_cols.append((c, label))

    # Display/history order should be chronological, not raw column order - the source
    # workbook's own column order has at least one out-of-sequence pair (31.05 before
    # 30.05), so sort by parsed date where possible; anything unparseable keeps its
    # original left-to-right position, appended after the dated ones.
    dated = sorted((pc for pc in progress_cols if pc[2] is not None), key=lambda pc: pc[2])
    undated = [pc for pc in progress_cols if pc[2] is None]
    ordered_progress_cols = dated + undated

    points = []
    cat_counts = Counter()
    for r in range(3, ws.max_row + 1):
        sno = ws.cell(row=r, column=sno_col).value
        if not isinstance(sno, (int, float)):
            continue
        category = ws.cell(row=r, column=cat_col).value
        category = re.sub(r'\s+', ' ', str(category)).strip() if category else 'Uncategorised'
        action_text = ws.cell(row=r, column=action_col).value
        action_text = re.sub(r'\s+', ' ', str(action_text)).strip() if action_text else ''

        supervisors = {}
        for c, label in sup_cols:
            v = ws.cell(row=r, column=c).value
            if v not in (None, ''):
                supervisors[label] = re.sub(r'\s+', ' ', str(v)).strip()

        # A cell can hold whitespace-only content (e.g. a lone '\n') that isn't None/''
        # but carries no real update - treat that as blank too, both for history and for
        # the updatedThisCycle/presence signals below, so a point isn't wrongly credited
        # with "updated" when nothing substantive was actually entered.
        history = []
        presence = []  # aligned to ordered_progress_cols, chronological - True if that checkpoint had real text
        for c, date_txt, dt in ordered_progress_cols:
            v = ws.cell(row=r, column=c).value
            txt = str(v).strip() if v not in (None, '') else ''
            presence.append(bool(txt))
            if txt:
                history.append(dict(date=date_txt, text=txt))
        latest = history[-1] if history else None
        updated_this_cycle = presence[-1] if presence else False

        points.append(dict(
            sno=int(sno), category=category, actionPoint=action_text, supervisors=supervisors,
            history=history, latestDate=(latest['date'] if latest else None),
            latestText=(latest['text'] if latest else None),
            updatedThisCycle=updated_this_cycle, presence=presence,
        ))
        cat_counts[category] += 1

    review_date_labels = [d for _, d, _ in ordered_progress_cols]
    latest_review_label = review_date_labels[-1] if review_date_labels else None

    insights = _build_action_plan_insights(points, review_date_labels)

    log(f"53-Point Action Plan: {len(points)} points, {len(ordered_progress_cols)} review "
        f"checkpoints, latest={latest_review_label}, source={os.path.basename(path)}")

    return dict(
        sourceFile=os.path.basename(path),
        reviewDates=review_date_labels,
        latestReviewDate=latest_review_label,
        totalPoints=len(points),
        categories=[dict(name=k, count=v) for k, v in cat_counts.items()],
        points=points,
        insights=insights,
    )


# ---------------------------------------------------------------------
# 53-Point Action Plan: computed insights (added 29-Sep-2026 per the user:
# "make some useful insights for the officers. Rather than just copy pasting
# the material... brainstorm it and make it better"). Everything here is a
# factual read of the actual data - text-similarity, presence patterns,
# regex-extracted machine IDs and dates, and supervisor-assignment counts -
# never an invented status, sentiment, or quality judgement the narrative
# text itself doesn't support.
# ---------------------------------------------------------------------
_AP_MACHINE_RE = re.compile(
    r'\b(UNIMAT|DUOMATIC|WMDU|BCM|UTV|DGS|MPT|UNI|DUO|CSM|BRM|PBR|SQRS|FRM|SRGM|RGM|RBMV|PCTM|DTE|MDU|WST)'
    r'[\s\-]{0,2}(\d{2,6})\b', re.IGNORECASE)
_AP_MACHINE_ALIAS = {'UNIMAT': 'UNI', 'DUOMATIC': 'DUO', 'WMDU': 'MDU'}
_AP_DATE_RE = re.compile(r'\b(\d{1,2})[./](\d{1,2})[./](\d{2,4})\b')
_AP_NIL_RE = re.compile(r'\b(nil|no\s+progress|no\s+update|no\s+any|not\s+received|not\s+yet\s+started|awaited)\b', re.I)


def _ap_parse_short_date(day, mon, yr):
    try:
        day, mon = int(day), int(mon)
        yr = int(yr)
        if yr < 100:
            yr += 2000
        return datetime.date(yr, mon, day)
    except Exception:
        return None


def _build_action_plan_insights(points, review_date_labels):
    if not points:
        return None
    latest_label = review_date_labels[-1] if review_date_labels else None
    prev_label = review_date_labels[-2] if len(review_date_labels) > 1 else None

    # ---- momentum: updated-with-new-content vs updated-but-near-repeat vs not updated ----
    substantive, repeat, not_updated = [], [], []
    for pt in points:
        h = pt['history']
        if not pt['updatedThisCycle']:
            not_updated.append(pt)
            continue
        if len(h) >= 2:
            ratio = difflib.SequenceMatcher(None, h[-1]['text'], h[-2]['text']).ratio()
            if ratio > 0.80:
                repeat.append(pt)
                continue
        substantive.append(pt)

    # ---- stalled: no entry in the last 2 checkpoints in a row (needs 2+ checkpoints to say so) ----
    stalled = []
    for pt in points:
        pres = pt.get('presence') or []
        if len(pres) >= 2 and not pres[-1] and not pres[-2]:
            stalled.append(pt)

    # ---- explicitly flagged "nil / no progress" in a short latest entry ----
    flagged_nil = []
    for pt in points:
        t = pt.get('latestText') or ''
        if t and len(t) < 200 and _AP_NIL_RE.search(t):
            flagged_nil.append(pt)

    # ---- cross-cutting machines: same machine ID named in >=2 different action points' latest text ----
    # Filtered to 2-4 occurrences spanning >=2 distinct categories: a machine cited in, say,
    # 8-11 points turned out (on inspection) to almost always be a shared roster/example list
    # copy-pasted across several similarly-worded points (e.g. a mock-drill machine list quoted
    # in five SAFETY points) rather than a genuine recurring issue - not a useful signal. A
    # machine named in a handful of points across genuinely different categories (e.g. flagged
    # in both a RELIABILITY point about its IOH and an AVAILABILITY point about lost days) is
    # the real "this machine is touching multiple workstreams right now" signal worth surfacing.
    machine_points = defaultdict(list)  # normalized key -> [{sno, category, actionPoint, matched}]
    for pt in points:
        t = pt.get('latestText') or ''
        if not t:
            continue
        seen_in_point = set()
        for m in _AP_MACHINE_RE.finditer(t):
            code = m.group(1).upper()
            code = _AP_MACHINE_ALIAS.get(code, code)
            num = str(int(m.group(2))) if m.group(2).isdigit() else m.group(2)
            key = f"{code}-{num}"
            if key in seen_in_point:
                continue
            seen_in_point.add(key)
            machine_points[key].append(dict(sno=pt['sno'], category=pt['category'],
                                             actionPoint=pt['actionPoint'], matched=m.group(0)))
    cross_cutting = []
    for k, v in machine_points.items():
        cats = sorted(set(x['category'] for x in v))
        if 2 <= len(v) <= 4 and len(cats) >= 2:
            cross_cutting.append(dict(machine=k, points=v, categories=cats))
    cross_cutting.sort(key=lambda x: (-len(x['categories']), -len(x['points'])))

    # ---- upcoming dates mentioned in the latest cycle's text (next ~45 days from the latest review) ----
    as_of_dt = _ap_parse_short_date(*review_date_labels[-1].split('.')) if review_date_labels else None
    upcoming = []
    if as_of_dt:
        window_end = as_of_dt + datetime.timedelta(days=45)
        for pt in points:
            t = pt.get('latestText') or ''
            if not t:
                continue
            for m in _AP_DATE_RE.finditer(t):
                d = _ap_parse_short_date(*m.groups())
                if d and as_of_dt <= d <= window_end:
                    start = max(0, m.start() - 70)
                    snippet = re.sub(r'\s+', ' ', t[start:m.start()]).strip()
                    upcoming.append(dict(sno=pt['sno'], category=pt['category'],
                                          actionPoint=pt['actionPoint'], date=d.strftime('%d-%b-%Y'),
                                          dateSort=d.strftime('%Y-%m-%d'), snippet=snippet[-90:]))
        upcoming.sort(key=lambda x: x['dateSort'])
        # de-dup near-identical snippets within the same point (e.g. the same date quoted twice)
        seen = set()
        deduped = []
        for u in upcoming:
            k = (u['sno'], u['dateSort'])
            if k in seen:
                continue
            seen.add(k)
            deduped.append(u)
        upcoming = deduped[:15]

    # ---- category momentum: total narrative volume this cycle vs the cycle before ----
    cat_this = defaultdict(int)
    cat_prev = defaultdict(int)
    for pt in points:
        h = pt['history']
        if h and h[-1]['date'] == latest_label:
            cat_this[pt['category']] += len(h[-1]['text'])
        if prev_label:
            prev_entry = next((e for e in h if e['date'] == prev_label), None)
            if prev_entry:
                cat_prev[pt['category']] += len(prev_entry['text'])
    cat_names = sorted(set(list(cat_this.keys()) + list(cat_prev.keys())))
    category_momentum = []
    for c in cat_names:
        cur, prev = cat_this.get(c, 0), cat_prev.get(c, 0)
        pct = round((cur - prev) / prev * 100, 0) if prev else None
        category_momentum.append(dict(category=c, thisCycleChars=cur, prevCycleChars=prev, pctChange=pct))

    # ---- supervisor load: how many distinct action points each nominated person carries ----
    sup_points = defaultdict(set)
    for pt in points:
        for name in (pt.get('supervisors') or {}).values():
            if name:
                sup_points[name].add(pt['sno'])
    supervisor_load = sorted(
        [dict(name=k, pointCount=len(v)) for k, v in sup_points.items()],
        key=lambda x: -x['pointCount'])[:8]

    return dict(
        momentum=dict(substantiveUpdates=len(substantive), repeatUpdates=len(repeat),
                       notUpdated=len(not_updated)),
        stalledPoints=[dict(sno=p['sno'], category=p['category'], actionPoint=p['actionPoint'],
                             latestDate=p['latestDate']) for p in stalled],
        repeatPoints=[dict(sno=p['sno'], category=p['category'], actionPoint=p['actionPoint'])
                      for p in repeat],
        flaggedNilPoints=[dict(sno=p['sno'], category=p['category'], actionPoint=p['actionPoint'],
                                text=p['latestText']) for p in flagged_nil],
        crossCuttingMachines=cross_cutting[:10],
        upcomingDates=upcoming,
        categoryMomentum=category_momentum,
        supervisorLoad=supervisor_load,
    )


# =====================================================================
# STEP 9: SWR Sidings and Rest Houses (Infrastructure Planning) — added
# 01-Oct-2026 per the user's request to add a separate "Infrastructure
# Planning" header/tab covering TM-siding and rest-house availability
# across the three divisions, division-wise / Sr.DEN-wise / CN-unit-wise,
# pending-in-construction and pending-overall, and the FY target picture.
#
# Source: a separate workbook in "SWR Sidings and Rest House" with sheets
# 'Master Summary' (division/overall totals + FY target narrative blocks +
# a footnote about rest houses without sidings), 'Construction ' (trailing
# space in the real sheet name — pending-in-construction list) and
# 'CE TM Sr DEN detail' (exhaustive, station-level breakdown across all
# three divisions with Sr.DEN/section/siding+RH status/responsibility
# remark). Like every other sheet this pipeline reads, nothing here is
# hardcoded beyond the sheet/column layout itself — rows, divisions and
# stations are all discovered from the sheet content. This extraction was
# validated by diffing its output field-for-field against a known-good,
# Playwright-tested payload produced earlier in this project for the same
# workbook (asOf 01.10.2026) until every section matched exactly.
# =====================================================================
INFRA_RESP_PATTERNS = [
    (re.compile(r'not\s*feasible', re.I), 'Not Feasible'),
    (re.compile(r'gati\s*shakti', re.I), 'Gati Shakti'),
    (re.compile(r'\bcon\s*org\b|\bcn\s*org\b|\bby\s*cn\b|\bcn\s*work\b|\bby\s*con\b|\bdone\s*by\s*con\b', re.I), 'CN (Construction Org)'),
    (re.compile(r'propos(?:ed|e)\s*by\s*div|div\w*\s*in\s*irpsm|yet\s*to\s*be\s*propos', re.I), 'Division (IRPSM)'),
]


def classify_infra_responsibility(remark):
    """Free-text 'who is responsible / what's the status' remark on a siding/RH
    row -> a clean bucket label, or None if the remark is blank. Mirrors the
    classify_exc_category / classify_shortfall pattern used elsewhere: an
    ordered list of regexes checked in priority order, with a safe fallback
    rather than silently dropping genuinely new wording."""
    s = str(remark or '').strip()
    if not s:
        return None
    for pat, label in INFRA_RESP_PATTERNS:
        if pat.search(s):
            return label
    return 'Other / Unspecified'


INFRA_SRDEN_WORD_CODE = {
    'N': 'N', 'NORTH': 'N', 'S': 'S', 'SOUTH': 'S', 'E': 'E', 'EAST': 'E',
    'W': 'W', 'WEST': 'W', 'C': 'C', 'CEN': 'C', 'CENTRAL': 'C', 'CENTRE': 'C',
    'HQ': 'HQ', 'HQRS': 'HQ', 'HEADQUARTERS': 'HQ',
}


def canon_srden_word(div, raw):
    """Like canon_srden() but for the Sidings/Rest House sheet, whose Sr.DEN
    column uses full-word labels ('Sr.DEN/North') rather than the single-letter
    codes the Daily Progress sheet uses. Returns 'DIV Sr.DEN/X', or — for the
    rare jurisdiction label that doesn't follow that pattern (e.g. a
    'New Line' construction stretch) — 'DIV <raw text>' so it is still
    bucketed rather than silently dropped."""
    if div not in ('MYS', 'SBC', 'UBL') or not raw:
        return None
    m = re.search(r'DEN[./]?\s*/?\s*([A-Za-z]+)', str(raw), re.I)
    if m:
        code = INFRA_SRDEN_WORD_CODE.get(m.group(1).strip().upper())
        if code:
            return f"{div} Sr.DEN/{code}"
    return f"{div} {str(raw).strip()}"


def parse_station_block(text):
    """'(49)\nUBL, NVU, ...' -> ['UBL','NVU',...] (leading count discarded,
    trailing full stops stripped off the last entry). Also tolerates a bare
    comma-separated list with no leading count."""
    if not text:
        return []
    s = str(text).strip()
    m = re.match(r'^\(?\s*\d+\s*\)?', s)
    if m:
        s = s[m.end():].strip()
    s = s.lstrip(':').strip()
    return [x.strip().rstrip('.') for x in re.split(r'[,\n]+', s) if x.strip()]


def normalize_infra_status(s):
    """Collapses the handful of free-text case/typo variants actually present
    in the source sheet ('Not feasible' / 'Not Feasible', 'Available and
    sanctioned' / 'Available and Sanctioned', '????') onto one canonical
    label per status, while leaving genuinely different statuses untouched."""
    if s is None:
        return None
    t = re.sub(r'\s+', ' ', str(s)).strip()
    if not t or t in ('?', '??', '???', '????'):
        return 'Unknown'
    low = t.lower()
    if low == 'not feasible':
        return 'Not Feasible'
    if low == 'available and sanctioned':
        return 'Available and Sanctioned'
    if low in ('na', 'n/a', 'nil', '-'):
        return 'N/A'
    return t


def parse_infra_target_text(s):
    """Parses the FY target narrative block on the Master Summary sheet, e.g.
    'Target for 2026-27 is 39 Nos.\nUBL: 13\nSBC: 13\nMYS: 13\nTotal no of
    sidings completed till 31.07.2026 is 6 Nos.(SBGA, REPI, URUK, KYND, KJG,
    SBC)' into a structured dict, via regex extraction over the free-text
    narrative (consistent with how this pipeline reads every other narrative
    block, e.g. the shortfall-category reconstruction)."""
    s = str(s or '')
    out = dict(fy=None, total=None, byDivision={}, completedAsOf=None, completedCount=None, completedList=[])
    m = re.search(r'Target for (\d{4}-\d{2,4})', s)
    if m:
        out['fy'] = m.group(1)
    m = re.search(r'Target for [\d\-]+ is (\d+) Nos', s)
    if m:
        out['total'] = int(m.group(1))
    for dcode, cnt in re.findall(r'(UBL|SBC|MYS)\s*:\s*(\d+)', s):
        out['byDivision'][dcode] = int(cnt)
    m = re.search(r'completed till ([\d.]+) is (\d+) Nos\.?\s*\(([^)]*)\)', s)
    if m:
        out['completedAsOf'] = m.group(1)
        out['completedCount'] = int(m.group(2))
        out['completedList'] = [x.strip() for x in m.group(3).split(',') if x.strip()]
    return out


def extract_sidings_rest_houses(base):
    path = (one_file(f"{base}/SWR Sidings and Rest House*/*.xlsx", required=False)
            or one_file(f"{base}/*Sidings*Rest House*/*.xlsx", required=False)
            or one_file(f"{base}/*Sidings*Rest*.xlsx", required=False))
    if not path:
        log("WARN: SWR Sidings and Rest House workbook not found - infra will be omitted")
        return None
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        DIV_NAME = {'Hubballi': 'UBL', 'Bengaluru': 'SBC', 'Mysuru': 'MYS'}

        # ---- Master Summary: division rows + SWR totals + FY target blocks + footnote ----
        ws = wb['Master Summary']

        def div_block(row_idx, swr_row_idx, target_row_idx):
            by_division = []
            for r, (label, code) in zip(range(row_idx, row_idx + 3),
                                         [('Hubballi', 'UBL'), ('Bengaluru', 'SBC'), ('Mysuru', 'MYS')]):
                avail_txt = ws.cell(r, 7).value
                sanc_txt = ws.cell(r, 8).value
                prop_txt = ws.cell(r, 9).value
                remarks = ws.cell(r, 10).value
                by_division.append(dict(
                    division=code, label=label, stations=ws.cell(r, 2).value,
                    required=ws.cell(r, 3).value, available=ws.cell(r, 4).value,
                    sanctioned=ws.cell(r, 5).value, proposed=ws.cell(r, 6).value,
                    availableStations=parse_station_block(avail_txt),
                    sanctionedStations=parse_station_block(sanc_txt),
                    proposedStations=parse_station_block(prop_txt),
                    remarks=(str(remarks).strip() if remarks else None),
                ))
            swr_remarks = ws.cell(swr_row_idx, 10).value
            swr = dict(
                stations=ws.cell(swr_row_idx, 2).value, required=ws.cell(swr_row_idx, 3).value,
                available=ws.cell(swr_row_idx, 4).value, sanctioned=ws.cell(swr_row_idx, 5).value,
                proposed=ws.cell(swr_row_idx, 6).value,
                remarks=(str(swr_remarks).strip() if swr_remarks else None),
            )
            target_txt = ws.cell(target_row_idx, 1).value
            target = parse_infra_target_text(target_txt) if target_txt else None
            return dict(byDivision=by_division, swr=swr, target=target)

        summary = dict(sidings=div_block(4, 7, 8), restHouses=div_block(14, 17, 18))

        note = None
        for r in range(1, ws.max_row + 1):
            a1 = ws.cell(r, 1).value
            if isinstance(a1, str) and a1.strip().lower().startswith('note'):
                note_val = ws.cell(r, 2).value
                note = str(note_val) if note_val is not None else None
                break

        as_of_overall = None
        m = re.search(r'As on ([\d.]+)', str(ws.cell(1, 1).value or ''))
        if m:
            as_of_overall = m.group(1)

        # ---- Construction sheet (trailing space in the real sheet name) ----
        construction = []
        cons_sheet_name = 'Construction ' if 'Construction ' in wb.sheetnames else (
            'Construction' if 'Construction' in wb.sheetnames else None)
        if cons_sheet_name:
            ws2 = wb[cons_sheet_name]
            for r in range(4, ws2.max_row + 1):
                div = ws2.cell(r, 2).value
                if not div:
                    continue
                construction.append(dict(
                    division=div, section=ws2.cell(r, 3).value, sectionLength=ws2.cell(r, 4).value,
                    stations=ws2.cell(r, 5).value, stationRequired=ws2.cell(r, 6).value,
                    sidingAvailable=ws2.cell(r, 7).value, sidingSanctioned=ws2.cell(r, 8).value,
                    sidingProposed=ws2.cell(r, 9).value, rhAvailable=ws2.cell(r, 10).value,
                    rhSanctioned=ws2.cell(r, 11).value, rhProposed=ws2.cell(r, 12).value,
                ))

        # ---- CE TM Sr DEN detail: station-level rows, all 3 divisions ----
        station_rows = []
        if 'CE TM Sr DEN detail' in wb.sheetnames:
            ws3 = wb['CE TM Sr DEN detail']
            cm = ColMerges(ws3)

            def mval(row, col):
                v = ws3.cell(row, col).value
                if v is not None:
                    return v
                vm = cm.vertical_merge_for(row, col)
                if vm:
                    return ws3.cell(vm[0], col).value
                return None

            cur_div, cur_div_asof = None, None
            for r in range(1, ws3.max_row + 1):
                a1 = ws3.cell(r, 1).value
                if isinstance(a1, str) and a1.startswith('Status of TM Sidings'):
                    m = re.search(r'in (\w+) Division \(As on ([\d.]+)\)', a1)
                    if m:
                        cur_div = DIV_NAME.get(m.group(1), m.group(1))
                        cur_div_asof = m.group(2)
                    continue
                slno = ws3.cell(r, 1).value
                if not isinstance(slno, (int, float)):
                    continue
                srden_raw = mval(r, 2)
                section = mval(r, 3)
                station = ws3.cell(r, 4).value
                km = ws3.cell(r, 5).value
                isd = ws3.cell(r, 6).value
                siding_status = normalize_infra_status(ws3.cell(r, 7).value)
                rh_status = normalize_infra_status(ws3.cell(r, 8).value)
                remark = ws3.cell(r, 9).value
                if not station:
                    continue
                srden_key = canon_srden_word(cur_div, srden_raw)
                station_rows.append(dict(
                    division=cur_div, asOf=cur_div_asof, srden=srden_raw, srdenKey=srden_key,
                    section=section, station=station, km=km, interSidingDistance=isd,
                    sidingStatus=siding_status, rhStatus=rh_status, remark=remark,
                    responsibility=classify_infra_responsibility(remark),
                ))

        # ---- Sr.DEN jurisdiction rollup (array, keyed by srdenKey) ----
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

        # ---- Responsibility rollup: overall + by division ----
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

        log(f"Sidings/Rest House: {len(station_rows)} station rows, {len(construction)} construction rows, "
            f"{len(srden_table)} Sr.DEN jurisdictions, source={os.path.basename(path)}")

        return dict(
            asOf=as_of_overall, summary=summary, note=note, construction=construction,
            stationRows=station_rows, srdenTable=srden_table, responsibility=responsibility,
            sourceFile=os.path.basename(path),
        )
    except Exception as e:
        log(f"WARN: Sidings/Rest House extraction failed ({e}) - omitting 'infra' from payload")
        return None


# =====================================================================
# STEP 10: Tenders & Contracts (added 05-Oct-2026) — a brand-new main
# dashboard section covering AMC/Works-contract agreement monitoring,
# OEM/AMC contact persons, the new-AMC/works-proposal approval pipeline,
# tenders under finalization, and the LOA-issued register.
#
# Source: a new device subfolder "Contract Details/Tenders and
# Contracts.xlsx" with five sheets - Agreement Details, Contact Person,
# Status of AMC Works Proposal, Tender Under Finalization, LOA issued.
#
# Nothing here is hardcoded beyond the sheet layout itself (same principle
# as the rest of this pipeline) except for a small, explicit brand-keyword
# registry (BRAND_KEYWORDS below) used to cross-link an Agreement to its
# OEM contact card and, for an expiring agreement, to any renewal proposal
# already in the pipeline. Only ~15 OEM/brand relationships exist in the
# source Contact Person sheet, so a curated keyword list is more reliable
# here than trying to fuzzy-infer it - the same reasoning already used for
# TAMPING_INFO elsewhere in this file. A brand with no safe keyword (e.g. a
# generic "AMC for RMV machine Engines" contact entry with no distinguishing
# company name in the text) is left with an empty keyword list rather than
# guessed - it will simply show as "not yet linked" instead of a wrong link.
# =====================================================================
BRAND_KEYWORDS = [
    dict(id='plasser', label='M/s. Plasser (India) Pvt. Ltd.', keywords=['plasser']),
    dict(id='phooltas', label='M/s. Phooltas Transrail Ltd.', keywords=['phooltas']),
    dict(id='harbour', label='M/s. Harbour Sales Pvt. Ltd.', keywords=['harbour sales', 'harbour']),
    dict(id='omega', label='M/s. Omega Rail & Spares Services Pvt. Ltd. (ORSSPL)', keywords=['omega rail', 'orsspl']),
    dict(id='vandhana', label='Vandhana International', keywords=['vandhana']),
    dict(id='simplex', label='M/s. Simplex Engineering & Foundry Works Pvt. Ltd.', keywords=['simplex']),
    dict(id='san', label='M/s. SAN Engineering and Locomotive Company Ltd.', keywords=['san engineering', 'san make']),
    dict(id='gemac', label='GEMAC (Nagory)', keywords=['gemac', 'nagory']),
    dict(id='gmmco', label='M/s. GMMCO Ltd.', keywords=['gmmco']),
    dict(id='southcalcutta', label='M/s. South Calcutta Diesels Pvt. Ltd.', keywords=['south calcutta diesel', 'south calcutta', 'deutz']),
    dict(id='cummins', label='M/s. Cummins India Ltd.', keywords=['cummins']),
    dict(id='rmv_engine', label='AMC for RMV machine Engines', keywords=[]),  # no safe brand keyword in source text
    dict(id='greaves', label='AMC for Greaves make Engines', keywords=['greaves']),
    dict(id='kirloskar', label='M/s. Kirloskar (Surya Power Services)', keywords=['kirloskar']),
    dict(id='uniwave', label='M/s. Uniwave Engineers LLP (Volvo engines)', keywords=['uniwave', 'volvo']),
]


def _brand_hits(text):
    t = str(text or '').lower()
    return [b['id'] for b in BRAND_KEYWORDS if any(kw in t for kw in b['keywords'])]


def _num(v):
    """Parses Indian currency-formatted strings ('Rs. 5,32,09,612/-',
    '1,90,58,825', '16,32,56,716.58') as well as plain numbers. Strips the
    'Rs.' prefix and trailing '/-' explicitly before dropping thousands
    separators, rather than a blind digit/dot strip - a blind strip would
    keep the dot inside 'Rs.' itself and corrupt the value (e.g. turn
    'Rs. 5,32,09,612/-' into 0.53209612 instead of 53209612)."""
    if v in (None, '', '-'):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v)
    s = re.sub(r'(?i)rs\.?', '', s)
    s = s.replace('₹', '')  # ₹ symbol (seen in the Indent Position workbook)
    s = re.sub(r'/-\s*$', '', s)
    s = re.sub(r'[,\s]', '', s)
    s = s.strip('-').strip()
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None


def _clean_text(v):
    if v is None:
        return None
    s = re.sub(r'\s+', ' ', str(v)).strip()
    return s or None


_DATE_IN_TEXT_RE = re.compile(r'\b(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})\b')


def _dates_in_text(text):
    out = []
    for d, m, y in _DATE_IN_TEXT_RE.findall(str(text or '')):
        dt = parse_ddmmyyyy(f"{d}.{m}.{y}")
        if dt:
            out.append(dt)
    return out


# Ordered most-advanced-stage-first, same pattern as classify_exc_category /
# classify_iohpoh_status elsewhere in this file: a remark narrating several
# events ("approval obtained... concurrence obtained...") must classify by
# its LATEST stage, not whichever phrase happens to match first in the text.
PROPOSAL_STAGE_ORDER = [
    ('tender_prep', 'Tender Document Under Preparation', re.compile(r'tender (?:document )?under preparation', re.I)),
    ('sanction', 'Sanction Obtained', re.compile(r'sanction obtained', re.I)),
    ('estimate_vetted', 'Detailed Estimate Vetted', re.compile(r'estimate vetted', re.I)),
    ('concurrence_obtained', 'Concurrence Obtained', re.compile(r'concurrence obtained', re.I)),
    ('concurrence_sent', 'Sent for Concurrence', re.compile(r'(?:sent|send) for concurr?ence|concurernce', re.I)),
    ('admin_approval_obtained', 'Administrative Approval Obtained', re.compile(r'admin\w* a\w*proval obtained|approval obtained', re.I)),
    ('admin_approval_sent', 'Sent for Administrative Approval', re.compile(r'sent (?:to hq )?for administrative approval', re.I)),
    ('under_preparation', 'Proposal Under Preparation', re.compile(r'under preparation|to be initiated', re.I)),
]


def classify_proposal_stage(remark):
    """Furthest stage reached in the AMC/Works proposal approval pipeline,
    read directly off the Remarks narrative."""
    s = str(remark or '')
    for key, label, pat in PROPOSAL_STAGE_ORDER:
        if pat.search(s):
            return key, label
    return 'pending', (_clean_text(remark) or 'No update recorded')


TENDER_STAGE_ORDER = [
    ('with_accepting_authority', 'With Accepting Authority', 5, re.compile(r'with accepting authority', re.I)),
    ('tcp_preparation', 'Legal Opinion Received — TCP Under Preparation', 4,
     re.compile(r'legal opinion received|tcp is under preparation|tcp under preparation', re.I)),
    ('legal_opinion_pending', 'Legal Opinion Pending', 3,
     re.compile(r'(?:sent|send) for legal opinion|legal opinion sent', re.I)),
    ('documents_scrutiny', 'Documents Under Scrutiny', 2, re.compile(r'under scrutiny', re.I)),
    ('tender_floated', 'Tender Floated — Awaiting Opening', 1, re.compile(r'tender floated', re.I)),
]


def classify_tender_stage(remark, current_holder):
    """Current-holder-aware tender-finalization stage, same priority-ordered-
    regex pattern used throughout this file. Falls back to the Current Holder
    field (Tender section / IREPS / GeM) when the remark text itself doesn't
    name a recognisable stage, and to the raw remark as a last resort."""
    s = str(remark or '')
    for key, label, order, pat in TENDER_STAGE_ORDER:
        if pat.search(s):
            return key, label, order
    holder = str(current_holder or '').strip()
    hl = holder.lower()
    if hl in ('ireps', 'gem'):
        return f'with_{hl}', f'With {holder} — Awaiting Opening', 1
    if 'tender section' in hl:
        return 'with_tender_section', 'In Process at Tender Section', 1
    return 'unclassified', (_clean_text(remark) or 'Status not recorded'), 0


LOA_STATUS_ORDER = [
    ('finalized', 'Agreement Finalized', re.compile(r'agreement (?:is )?finali[sz]ed', re.I)),
    ('sent_to_contractor', 'Agreement Sent to Contractor', re.compile(r'agreement sent to contractor', re.I)),
    ('bg_verified', 'BG Verified', re.compile(r'bg verification done', re.I)),
    ('pg_submitted', 'PG Submitted — Agreement in Process', re.compile(r'pg (?:has been |was )?submitted', re.I)),
    ('pg_awaited', 'Awaiting PG from Firm', re.compile(r'pg to be submitted', re.I)),
]


def classify_loa_status(remark):
    s = str(remark or '')
    for key, label, pat in LOA_STATUS_ORDER:
        if pat.search(s):
            return key, label
    return 'in_process', (_clean_text(remark) or 'Agreement in process')


def extract_tenders_contracts(base, as_of):
    path = one_file(f"{base}/Contract Details/*.xlsx", required=False)
    if not path:
        log("WARN: Contract Details workbook not found - tendersContracts will be omitted")
        return None
    wb = openpyxl.load_workbook(path, data_only=True)

    # ---- Agreement Details ----
    agreements = []
    if 'Agreement Details' in wb.sheetnames:
        ws = wb['Agreement Details']
        for r in range(2, ws.max_row + 1):
            sno = ws.cell(row=r, column=1).value
            if not isinstance(sno, (int, float)):
                continue
            firm = _clean_text(ws.cell(row=r, column=6).value)
            if not firm:
                continue
            ctype = _clean_text(ws.cell(row=r, column=2).value)
            loa_no = _clean_text(ws.cell(row=r, column=3).value)
            allocation = _clean_text(ws.cell(row=r, column=4).value)
            work = _clean_text(ws.cell(row=r, column=5).value)
            loa_date = ws.cell(row=r, column=7).value
            amount = _num(ws.cell(row=r, column=8).value)
            revised = _num(ws.cell(row=r, column=9).value)
            orig_curr = _clean_text(ws.cell(row=r, column=10).value)
            ext_curr = _clean_text(ws.cell(row=r, column=11).value)
            latest_bill = _clean_text(ws.cell(row=r, column=12).value)
            cum_bill = _num(ws.cell(row=r, column=13).value)
            balance = _num(ws.cell(row=r, column=14).value)
            registered = _num(ws.cell(row=r, column=15).value)
            remarks_tmo = _clean_text(ws.cell(row=r, column=16).value)
            remarks_sse = _clean_text(ws.cell(row=r, column=17).value)

            eff_date = parse_ddmmyyyy(ext_curr) or parse_ddmmyyyy(orig_curr)
            days_to_expiry = (eff_date.date() - as_of).days if eff_date else None
            if days_to_expiry is None:
                status = 'Unknown'
            elif days_to_expiry < 0:
                status = 'Expired'
            elif days_to_expiry <= 90:
                status = 'Expiring Soon'
            else:
                status = 'Valid'

            eff_value = revised if revised is not None else amount
            brand_ids = _brand_hits(f"{firm} {work}")

            agreements.append(dict(
                sno=int(sno), contractType=ctype, loaNoAndDate=loa_no, allocation=allocation,
                nameOfWork=work, firmName=firm,
                dateOfIssueOfLoa=(loa_date.strftime('%d-%b-%Y') if isinstance(loa_date, (datetime.date, datetime.datetime)) else _clean_text(loa_date)),
                agreementAmount=amount, revisedAgreementValue=revised, effectiveValue=eff_value,
                originalCurrencyDate=orig_curr, extendedCurrencyDate=ext_curr,
                effectiveCurrencyDate=(eff_date.strftime('%d-%b-%Y') if eff_date else None),
                daysToExpiry=days_to_expiry, expiryStatus=status,
                latestPassedBill=latest_bill, cumulativeBill=cum_bill, balance=balance,
                registeredAmount=registered, remarksTmoAccounts=remarks_tmo, remarksSseProposals=remarks_sse,
                brandIds=brand_ids, renewalProposals=[],
            ))
    log(f"Tenders & Contracts: {len(agreements)} agreement records")

    # ---- Contact Person (grouped by an OEM-group heading row) ----
    contacts = []
    if 'Contact Person' in wb.sheetnames:
        ws = wb['Contact Person']
        cur_group = None
        for r in range(1, ws.max_row + 1):
            a = ws.cell(row=r, column=1).value
            b = ws.cell(row=r, column=2).value
            if isinstance(a, str) and b is None:
                txt = _clean_text(a)
                if txt and not txt.lower().startswith('sl.no'):
                    cur_group = txt
                continue
            if not isinstance(a, (int, float)):
                continue
            amc_name = _clean_text(b)
            if not amc_name:
                continue
            persons = []
            for col in (3, 4, 5):
                raw = ws.cell(row=r, column=col).value
                if raw in (None, '', '-'):
                    continue
                txt = _clean_text(raw)
                phone = None
                m = re.search(r'([\d][\d\s\-]{7,})\s*$', txt)
                namepart = txt
                if m:
                    phone = re.sub(r'\s+', '', m.group(1))
                    namepart = txt[:m.start()]
                role = None
                rm = re.search(r'\(([^)]+)\)', namepart)
                if rm:
                    role = _clean_text(rm.group(1))
                    namepart = namepart[:rm.start()] + namepart[rm.end():]
                namepart = _clean_text(namepart)
                if namepart:
                    namepart = re.sub(r'[\s:\-]+$', '', namepart).strip()
                persons.append(dict(name=namepart, role=role, phone=phone, raw=txt))
            brand_ids = _brand_hits(amc_name)
            contacts.append(dict(sno=int(a), group=cur_group, amcName=amc_name, persons=persons, brandIds=brand_ids))
    log(f"Tenders & Contracts: {len(contacts)} contact-person records")

    # ---- Status of AMC Works Proposal (two stacked sub-tables: Works Contract, AMC) ----
    proposals = []
    status_sheet = next((n for n in wb.sheetnames if n.strip() == 'Status of AMC Works Proposal'), None)
    if status_sheet:
        ws = wb[status_sheet]
        cur_section = None
        r = 1
        while r <= ws.max_row:
            a = ws.cell(row=r, column=1).value
            b = ws.cell(row=r, column=2).value
            if isinstance(a, str) and b is None and a.strip() in ('Works Contract', 'AMC'):
                cur_section = a.strip()
                r += 1
                continue
            if isinstance(a, str) and a.strip().lower().startswith('s.no'):
                r += 1
                continue
            if isinstance(a, (int, float)) and cur_section:
                desc = _clean_text(ws.cell(row=r, column=2).value)
                items = ws.cell(row=r, column=3).value
                value_raw = ws.cell(row=r, column=4).value
                value_num = _num(value_raw)
                remark = _clean_text(ws.cell(row=r, column=5).value)
                currency = _clean_text(ws.cell(row=r, column=6).value)
                stage_key, stage_label = classify_proposal_stage(remark)
                dates = _dates_in_text(remark)
                proposals.append(dict(
                    section=cur_section, sno=int(a), description=desc, items=items,
                    value=(value_num if value_num is not None else _clean_text(value_raw)), valueNum=value_num,
                    remark=remark, currencyOfCurrent=currency,
                    stageKey=stage_key, stageLabel=stage_label,
                    lastDatedMention=(max(dates).strftime('%d-%b-%Y') if dates else None),
                    brandIds=_brand_hits(desc),
                ))
            r += 1
    log(f"Tenders & Contracts: {len(proposals)} AMC/Works proposal records "
        f"({sum(1 for p in proposals if p['section']=='Works Contract')} Works, "
        f"{sum(1 for p in proposals if p['section']=='AMC')} AMC)")

    # ---- Tender Under Finalization ----
    tenders = []
    if 'Tender Under Finalization' in wb.sheetnames:
        ws = wb['Tender Under Finalization']
        for r in range(4, ws.max_row + 1):
            sno = ws.cell(row=r, column=1).value
            if not isinstance(sno, (int, float)):
                continue
            works_amc = _clean_text(ws.cell(row=r, column=2).value)
            ttype = _clean_text(ws.cell(row=r, column=3).value)
            desc = _clean_text(ws.cell(row=r, column=4).value)
            items = _clean_text(ws.cell(row=r, column=5).value)
            value = _clean_text(ws.cell(row=r, column=6).value)
            date_floated = ws.cell(row=r, column=7).value
            date_opening = ws.cell(row=r, column=8).value
            tc_exec = _clean_text(ws.cell(row=r, column=9).value)
            tc_fin = _clean_text(ws.cell(row=r, column=10).value)
            tc_sis = _clean_text(ws.cell(row=r, column=11).value)
            taa = _clean_text(ws.cell(row=r, column=12).value)
            holder = _clean_text(ws.cell(row=r, column=13).value)
            pending_since = _clean_text(ws.cell(row=r, column=14).value)
            pending_days = ws.cell(row=r, column=15).value
            remark = _clean_text(ws.cell(row=r, column=16).value)

            stage_key, stage_label, stage_order = classify_tender_stage(remark, holder)
            dates = _dates_in_text(remark)
            last_dt = max(dates) if dates else None
            days_since = (as_of - last_dt.date()).days if last_dt else None

            tenders.append(dict(
                sno=int(sno), worksOrAmc=works_amc, tenderType=ttype, description=desc, items=items,
                value=value,
                dateFloated=(date_floated.strftime('%d-%b-%Y') if isinstance(date_floated, (datetime.date, datetime.datetime)) else _clean_text(date_floated)),
                dateOpening=(date_opening.strftime('%d-%b-%Y') if isinstance(date_opening, (datetime.date, datetime.datetime)) else _clean_text(date_opening)),
                tcExecutive=tc_exec, tcFinance=tc_fin, tcSister=tc_sis, taa=taa, currentHolder=holder,
                pendingSinceRaw=pending_since, pendingDaysRaw=(pending_days if isinstance(pending_days, (int, float)) else None),
                remark=remark, stageKey=stage_key, stageLabel=stage_label, stageOrder=stage_order,
                lastDatedMention=(last_dt.strftime('%d-%b-%Y') if last_dt else None), daysSinceLastUpdate=days_since,
                brandIds=_brand_hits(f"{works_amc} {desc} {remark}"),
            ))
    log(f"Tenders & Contracts: {len(tenders)} tender-under-finalization records")

    # ---- LOA issued (stacked by "FY <yyyy-yy>" section header rows) ----
    loas = []
    if 'LOA issued' in wb.sheetnames:
        ws = wb['LOA issued']
        cur_fy = None
        r = 1
        while r <= ws.max_row:
            a = ws.cell(row=r, column=1).value
            if isinstance(a, str) and re.match(r'^FY\s*\d{4}-\d{2,4}$', a.strip()):
                cur_fy = a.strip()
                r += 1
                continue
            if isinstance(a, str) and a.strip().lower().startswith('s no'):
                r += 1
                continue
            if isinstance(a, (int, float)) and cur_fy:
                works_amc = _clean_text(ws.cell(row=r, column=2).value)
                ttype = _clean_text(ws.cell(row=r, column=3).value)
                desc = _clean_text(ws.cell(row=r, column=4).value)
                items = _clean_text(ws.cell(row=r, column=5).value)
                value = _clean_text(ws.cell(row=r, column=6).value)
                tc_exec = _clean_text(ws.cell(row=r, column=7).value)
                tc_fin = _clean_text(ws.cell(row=r, column=8).value)
                tc_sis = _clean_text(ws.cell(row=r, column=9).value)
                taa = _clean_text(ws.cell(row=r, column=10).value)
                loa_on = ws.cell(row=r, column=11).value
                remark = _clean_text(ws.cell(row=r, column=12).value)
                status_key, status_label = classify_loa_status(remark)
                value_num = _num(value)
                loas.append(dict(
                    fy=cur_fy, sno=int(a), worksOrAmc=works_amc, tenderType=ttype, description=desc,
                    items=items, value=value, valueNum=value_num,
                    tcExecutive=tc_exec, tcFinance=tc_fin, tcSister=tc_sis, taa=taa,
                    loaIssuedOn=(loa_on.strftime('%d-%b-%Y') if isinstance(loa_on, (datetime.date, datetime.datetime)) else _clean_text(loa_on)),
                    remark=remark, statusKey=status_key, statusLabel=status_label,
                    brandIds=_brand_hits(f"{works_amc} {desc}"),
                ))
            r += 1
    log(f"Tenders & Contracts: {len(loas)} LOA-issued records across "
        f"{len(set(l['fy'] for l in loas))} FY sections")

    # ---- Cross-sheet synthesis: renewal signal for expiring/expired agreements ----
    # For every agreement whose brand is recognised and whose expiry status is
    # Expiring Soon/Expired, check whether that same brand appears in any
    # current AMC/Works proposal - i.e. is a renewal already in the pipeline,
    # and at what stage. This is the one cross-sheet insight the raw sheets
    # don't give directly: "is someone already acting on this expiring contract."
    proposals_by_brand = defaultdict(list)
    for p in proposals:
        for bid in p['brandIds']:
            proposals_by_brand[bid].append(p)
    for ag in agreements:
        if ag['expiryStatus'] in ('Expiring Soon', 'Expired'):
            seen = set()
            for bid in ag['brandIds']:
                for p in proposals_by_brand.get(bid, []):
                    pk = (p['section'], p['sno'])
                    if pk not in seen:
                        seen.add(pk)
                        ag['renewalProposals'].append(p)

    kpi = dict(
        totalAgreements=len(agreements),
        activeAgreements=sum(1 for a in agreements if a['expiryStatus'] in ('Valid', 'Expiring Soon')),
        totalCommittedValue=round(sum(a['effectiveValue'] or 0 for a in agreements), 2),
        expiringSoonCount=sum(1 for a in agreements if a['expiryStatus'] == 'Expiring Soon'),
        expiredCount=sum(1 for a in agreements if a['expiryStatus'] == 'Expired'),
        proposalsInPipeline=len(proposals),
        tendersUnderFinalization=len(tenders),
    )
    if loas:
        current_fy_label = loas[0]['fy']
        kpi['loaIssuedCurrentFy'] = sum(1 for l in loas if l['fy'] == current_fy_label)
        kpi['loaIssuedCurrentFyValue'] = round(sum(l['valueNum'] or 0 for l in loas if l['fy'] == current_fy_label), 2)
        kpi['currentFyLabel'] = current_fy_label

    # ---- "Needs Attention" synthesized alerts (most severe first) ----
    attention = []
    for ag in sorted(agreements, key=lambda a: a['daysToExpiry'] if a['daysToExpiry'] is not None else 99999):
        if ag['expiryStatus'] in ('Expiring Soon', 'Expired') and not ag['renewalProposals']:
            when = (f"expired {abs(ag['daysToExpiry'])} days ago" if ag['expiryStatus'] == 'Expired'
                    else f"expires in {ag['daysToExpiry']} days")
            attention.append(dict(
                severity=('high' if ag['expiryStatus'] == 'Expired' else 'medium'),
                text=f"{ag['firmName']} — {when}, no renewal proposal found in the pipeline.",
                refType='agreement', refSno=ag['sno'],
            ))
    for t in tenders:
        if (t['daysSinceLastUpdate'] or 0) >= 30 and t['stageOrder'] < 5:
            attention.append(dict(
                severity='medium',
                text=f"Tender S.No {t['sno']} ({(t['description'] or '')[:60]}) — no dated update in remarks "
                     f"for {t['daysSinceLastUpdate']} days, currently '{t['stageLabel']}'.",
                refType='tender', refSno=t['sno'],
            ))
    current_fy_label = loas[0]['fy'] if loas else None
    for l in loas:
        if l['fy'] == current_fy_label and l['statusKey'] != 'finalized':
            attention.append(dict(
                severity='low',
                text=f"LOA S.No {l['sno']} ({l['worksOrAmc']}) issued {l['loaIssuedOn']} — agreement not yet "
                     f"finalized ({l['statusLabel']}).",
                refType='loa', refSno=l['sno'],
            ))
    sev_rank = dict(high=0, medium=1, low=2)
    attention.sort(key=lambda x: sev_rank.get(x['severity'], 9))

    log(f"Tenders & Contracts: {kpi['expiringSoonCount']} expiring soon, {kpi['expiredCount']} expired, "
        f"{len(attention)} attention flags raised")

    return dict(
        sourceFile=os.path.basename(path), asOf=as_of.strftime('%d-%b-%Y'),
        kpi=kpi, agreements=agreements, contacts=contacts, proposals=proposals,
        tenders=tenders, loas=loas, attention=attention[:20],
        brandRegistry=BRAND_KEYWORDS,
    )


# =====================================================================
# EXTRACTION: Indent Position (procurement pipeline tracker)
# =====================================================================
# Stage order below follows the Summary sheet's own column order (confirmed
# against its merged-header layout), which also matches the order the user
# asked for ("Tenders to be published ... PO issued ... Material received
# ... Under concurrence etc."). "Safety Items" is NOT a pipeline stage - it's
# a cross-cutting list of safety-critical items that also appear in one of
# the 9 stage sheets below (its own 'Current Status' column names which),
# so it's extracted and surfaced separately rather than folded into the
# funnel.
INDENT_STAGE_KEYS = [
    'tenderToBePublished', 'tenderUnderFinalisation', 'poIssuedPendingReceipt',
    'materialReceived', 'underVetting', 'underConcurrence', 'indentSigning',
    'balanceIndentsToBeProposed', 'indentUnderProcess',
]
INDENT_STAGE_LABELS = {
    'tenderToBePublished': 'Tender to be Published',
    'tenderUnderFinalisation': 'Tender Under Finalisation',
    'poIssuedPendingReceipt': 'PO Issued Pending Receipt',
    'materialReceived': 'Material Received',
    'underVetting': 'Under Vetting',
    'underConcurrence': 'Under Concurrence',
    'indentSigning': 'Indent Signing',
    'balanceIndentsToBeProposed': 'Balance Indents to be Proposed',
    'indentUnderProcess': 'Indent Proposal Under Process',
}
# Sheet names exactly as they appear in the source workbook - openpyxl sheet
# lookup is exact-match, and 'Indent under process ' has a trailing space.
INDENT_STAGE_SHEETS = {
    'tenderToBePublished': 'Tender to be published',
    'tenderUnderFinalisation': 'Tender Under Finalisation',
    'poIssuedPendingReceipt': 'Po Issued Pending Reciept',
    'materialReceived': 'Material Recieved',
    'underVetting': 'Under Vetting',
    'underConcurrence': 'Under Concurrence',
    'indentSigning': 'Indent Signing',
    'balanceIndentsToBeProposed': 'Balance Indents to be proposed',
    'indentUnderProcess': 'Indent under process ',
}
# Per-sheet column layout: (jsonKey, 1-indexed column, kind). Read by FIXED
# column index via ws.cell(row, column) - confirmed cell-by-cell against the
# live workbook (including merged-header ranges) rather than inferred from a
# non-blank-only scan. Several rows here have a genuinely blank middle
# column (e.g. no 'PO issue date' yet, or 'PAC signing Authority' not
# applicable) - a scan that skips None values would silently shift every
# later field in that row onto the wrong key, exactly the class of bug this
# project has been bitten by before (by_mtype_cat, drillMachineNames).
INDENT_STAGE_COLUMNS = {
    'tenderToBePublished': [
        ('sno', 1, 'int'), ('description', 2, 'text'), ('itemCategory', 3, 'text'),
        ('noOfItems', 4, 'num'), ('qtyRequired', 5, 'raw'), ('indentValue', 6, 'money'),
        ('indentNo', 7, 'text'), ('sentToPurchaseUnit', 8, 'date'),
        ('durationInCurrentStatus', 9, 'num'), ('purchaseUnit', 10, 'text'), ('remarks', 11, 'text'),
    ],
    'tenderUnderFinalisation': [
        ('sno', 1, 'int'), ('description', 2, 'text'), ('itemCategory', 3, 'text'),
        ('noOfItems', 4, 'num'), ('qtyRequired', 5, 'raw'), ('indentValue', 6, 'money'),
        ('indentNo', 7, 'text'), ('indentDate', 8, 'date'), ('likelySupplier', 9, 'text'),
        ('tenderNo', 10, 'raw'), ('tenderPublishedDate', 11, 'date'), ('tenderOpeningDate', 12, 'date'),
        ('daysPendingSinceOpening', 13, 'num'),
    ],
    'poIssuedPendingReceipt': [
        ('sno', 1, 'int'), ('description', 2, 'text'), ('demandRaisedDate', 3, 'date'),
        ('indentDate', 4, 'text'), ('noOfItems', 5, 'num'), ('itemCategory', 6, 'text'),
        ('qtyRequired', 7, 'raw'), ('purchaseUnit', 8, 'text'), ('poNumber', 9, 'raw'),
        ('poDate', 10, 'date'), ('firmNameLocation', 11, 'text'), ('firmContact', 12, 'text'),
        ('dueDeliveryDate', 13, 'date'),
    ],
    'materialReceived': [
        ('sno', 1, 'int'), ('description', 2, 'text'), ('noOfItems', 3, 'num'),
        ('poNumber', 4, 'raw'), ('poDate', 5, 'date'), ('qualityReview', 6, 'text'),
        ('latestRemarks', 7, 'text'),
    ],
    'underVetting': [
        ('sno', 1, 'int'), ('description', 2, 'text'), ('noOfItems', 3, 'num'),
        ('itemCategory', 4, 'text'), ('purchaseUnit', 5, 'text'), ('indentNo', 6, 'text'),
        ('indentValue', 7, 'money'), ('indentDate', 8, 'date'), ('sentOn', 9, 'date'),
        ('daysPending', 10, 'num'), ('tdc', 11, 'text'),
    ],
    'underConcurrence': [
        ('sno', 1, 'int'), ('description', 2, 'text'), ('noOfItems', 3, 'num'),
        ('itemCategory', 4, 'text'), ('purchaseUnit', 5, 'text'), ('indentNo', 6, 'text'),
        ('indentDate', 7, 'date'), ('indentValue', 8, 'money'), ('sentOn', 9, 'date'),
        ('daysPending', 10, 'num'), ('tdc', 11, 'text'), ('remarks', 12, 'text'),
    ],
    'indentSigning': [
        ('sno', 1, 'int'), ('description', 2, 'text'), ('noOfItems', 3, 'num'),
        ('itemCategory', 4, 'text'), ('indentValue', 5, 'money'), ('signingAuthority', 6, 'text'),
        ('indentNo', 7, 'text'), ('indentDate', 8, 'date'), ('currentHolder', 9, 'text'),
        ('sentToCurrentHolderOn', 10, 'date'), ('daysPending', 11, 'num'),
        ('expectedSigningDate', 12, 'raw'), ('remarks', 13, 'text'),
    ],
    'balanceIndentsToBeProposed': [
        ('sno', 1, 'int'), ('description', 2, 'text'), ('noOfItems', 3, 'num'),
        ('itemCategory', 4, 'text'), ('dateOfDemandFromField', 5, 'date'),
        ('daysPending', 6, 'num'), ('targetDateForProposal', 7, 'date'), ('remarks', 8, 'text'),
    ],
    'indentUnderProcess': [
        ('sno', 1, 'int'), ('description', 2, 'text'), ('itemCategory', 3, 'text'),
        ('items', 4, 'raw'), ('indentNoIfPlaced', 5, 'raw'), ('indentValue', 6, 'money'),
        ('approvingAuthority', 7, 'text'), ('vettingRequired', 8, 'text'),
        ('concurrenceRequired', 9, 'text'), ('signingAuthority', 10, 'text'),
        ('pacSigningAuthority', 11, 'text'), ('eofficeRequired', 12, 'text'),
        ('currentHolderEoffice', 13, 'text'),
    ],
}
SAFETY_ITEMS_COLUMNS = [
    ('sno', 1, 'int'), ('description', 2, 'text'), ('noOfItems', 3, 'num'),
    ('itemCategory', 4, 'text'), ('qtyRequired', 5, 'raw'), ('indentValue', 6, 'money'),
    ('indentNo', 7, 'text'), ('indentDate', 8, 'date'), ('poIssueDateOrStatus', 9, 'raw'),
    ('firmNameLocation', 10, 'text'), ('firmContact', 11, 'text'), ('currentStatus', 12, 'text'),
    ('currentHolder', 13, 'text'),
]


def _ip_cell(v, kind):
    """Formats one raw openpyxl cell value for the Indent Position payload.
    'money' returns (displayText, numericValue); 'date' returns
    (displayText, isoDateOrNone) so the dashboard can both show the sheet's
    own formatting and do real date math (sort/overdue checks) client-side."""
    if kind == 'money':
        if isinstance(v, (datetime.date, datetime.datetime)):
            return v.strftime('%d-%b-%Y'), None
        return _clean_text(v), _num(v)
    if kind == 'date':
        if isinstance(v, (datetime.date, datetime.datetime)):
            return v.strftime('%d-%b-%Y'), v.strftime('%Y-%m-%d')
        return _clean_text(v), None
    if isinstance(v, (datetime.date, datetime.datetime)):
        v = v.strftime('%d-%b-%Y')  # a stray date in a non-date column (seen in Material Recieved)
    if kind == 'int':
        return int(v) if isinstance(v, (int, float)) else None
    if kind == 'num':
        return v if isinstance(v, (int, float)) else (_clean_text(v) if v is not None else None)
    if kind == 'raw':
        return v if isinstance(v, (int, float)) else _clean_text(v)
    return _clean_text(v)  # 'text'


def _ip_rows(ws, start_row, col_map):
    """Reads data rows for one Indent Position stage sheet starting at
    start_row, stopping at the sheet's own TOTAL row (every stage sheet
    except Summary/Safety Items ends with one) or after a run of blank rows
    for sheets that have none. Handles the 'Indent under process ' sheet's
    stray blank row between its header and first data row the same way as
    any other blank - it's just skipped, not treated as end-of-data."""
    rows, total_raw = [], None
    r, consec_blank = start_row, 0
    while r <= ws.max_row:
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        # The TOTAL row's label sits in column 1 on most stage sheets, but on
        # 'Tender to be published' it's 'Total Indents: N' in column 2 with
        # column 1 blank - check both before treating column 1 as blank.
        if (isinstance(a, str) and 'total' in a.lower()) or (isinstance(b, str) and 'total' in b.lower()):
            total_raw = [ws.cell(row=r, column=c).value for c in range(1, len(col_map) + 2)]
            break
        if a is None:
            consec_blank += 1
            if rows and consec_blank > 20:
                break
            r += 1
            continue
        consec_blank = 0
        if not isinstance(a, (int, float)):
            r += 1
            continue
        row = {}
        for key, col, kind in col_map:
            v = ws.cell(row=r, column=col).value
            if kind in ('money', 'date'):
                row[key], row[key + ('Num' if kind == 'money' else 'Iso')] = _ip_cell(v, kind)
            else:
                row[key] = _ip_cell(v, kind)
        rows.append(row)
        r += 1
    return rows, total_raw


def _ip_total(total_raw):
    if not total_raw:
        return dict(indentCount=None, itemCount=None)
    vals = [v for v in total_raw if v is not None]
    indent_count = None
    for v in vals:
        if isinstance(v, str):
            m = re.search(r'Indents?:\s*(\d+)', v, re.I)
            if m:
                indent_count = int(m.group(1))
    item_count = None
    for v in reversed(vals):
        if isinstance(v, (int, float)):
            item_count = int(v)
            break
    return dict(indentCount=indent_count, itemCount=item_count)


def _ip_summary_table(ws, data_rows):
    out = []
    for r in data_rows:
        store = ws.cell(row=r, column=1).value
        if not store:
            continue
        total = ws.cell(row=r, column=2).value
        stages = {}
        for i, key in enumerate(INDENT_STAGE_KEYS):
            v = ws.cell(row=r, column=3 + i).value
            stages[key] = v if isinstance(v, (int, float)) else 0
        out.append(dict(store=_clean_text(store),
                         total=(total if isinstance(total, (int, float)) else (_num(total) or 0)),
                         stages=stages))
    return out


def extract_indent_position(base, as_of):
    # Folder name carries the as-on date and will change every time the user
    # updates the workbook (they said they'll "keep on updating" it), so
    # match by prefix rather than the exact current folder name.
    path = (one_file(f"{base}/INDENT POSITION*/*.xlsx", required=False)
            or one_file(f"{base}/*Indent Position*/*.xlsx", required=False)
            or one_file(f"{base}/*INDENT*POSITION*/*.xlsx", required=False))
    if not path:
        log("WARN: Indent Position workbook not found - indentPosition will be omitted")
        return None
    wb = openpyxl.load_workbook(path, data_only=True)

    source_as_of = None
    if 'Summary' in wb.sheetnames:
        title = wb['Summary'].cell(row=1, column=1).value
        m = re.search(r'as\s+on\s+([\d.\-/]+)', str(title or ''), re.I)
        if m:
            source_as_of = m.group(1)

    # ---- Summary: two stacked by-division tables (by indent count / by item count) ----
    summary = dict(stageKeys=INDENT_STAGE_KEYS, stageLabels=INDENT_STAGE_LABELS, byIndent=[], byItem=[])
    if 'Summary' in wb.sheetnames:
        ws = wb['Summary']
        summary['byIndent'] = _ip_summary_table(ws, range(4, 7))
        summary['byItem'] = _ip_summary_table(ws, range(13, 16))
    log(f"Indent Position: Summary - {len(summary['byIndent'])} by-indent division rows, "
        f"{len(summary['byItem'])} by-item division rows")

    # ---- Safety Items (cross-cutting; each also appears in its own stage sheet below) ----
    safety_items = []
    if 'Safety Items' in wb.sheetnames:
        safety_items, _ = _ip_rows(wb['Safety Items'], 3, SAFETY_ITEMS_COLUMNS)
    log(f"Indent Position: {len(safety_items)} safety-item records")

    # ---- The 9 pipeline-stage sheets ----
    stages = {}
    for key in INDENT_STAGE_KEYS:
        sheet_name = INDENT_STAGE_SHEETS[key]
        label = INDENT_STAGE_LABELS[key]
        if sheet_name not in wb.sheetnames:
            log(f"WARN: Indent Position sheet '{sheet_name}' not found - '{key}' will be empty")
            stages[key] = dict(label=label, rows=[], indentCount=0, itemCount=0)
            continue
        ws = wb[sheet_name]
        rows, total_raw = _ip_rows(ws, 3, INDENT_STAGE_COLUMNS[key])
        totals = _ip_total(total_raw)
        indent_count = totals['indentCount'] if totals['indentCount'] is not None else len(rows)
        item_count = totals['itemCount']
        if item_count is None:
            item_count = sum((r['noOfItems'] if isinstance(r.get('noOfItems'), (int, float)) else 1) for r in rows)
        if totals['indentCount'] is not None and totals['indentCount'] != len(rows):
            log(f"WARN: Indent Position '{sheet_name}' TOTAL row says {totals['indentCount']} indents "
                f"but {len(rows)} data rows were extracted - check for a layout change before trusting this stage.")
        stages[key] = dict(label=label, rows=rows, indentCount=indent_count, itemCount=item_count)
    for key in INDENT_STAGE_KEYS:
        log(f"Indent Position: '{stages[key]['label']}' - {stages[key]['indentCount']} indents, "
            f"{stages[key]['itemCount']} items")

    # ---- Aging / "Needs Attention" synthesis across the stage sheets' own
    # duration fields and due/target dates - the one cross-sheet view the raw
    # sheets don't give directly: what's stuck, and for how long. ----
    attention = []

    def _overdue_days(iso):
        if not iso:
            return None
        try:
            d = datetime.date.fromisoformat(iso)
        except Exception:
            return None
        return (as_of - d).days

    for key, thr_med, thr_high in (('underVetting', 15, 30), ('underConcurrence', 15, 30), ('indentSigning', 20, 40)):
        for r in stages[key]['rows']:
            days = r.get('daysPending')
            if isinstance(days, (int, float)) and days >= thr_med:
                attention.append(dict(
                    severity=('high' if days >= thr_high else 'medium'), stageKey=key, stageLabel=stages[key]['label'],
                    sno=r.get('sno'), days=int(days),
                    text=f"{(r.get('description') or 'Indent')[:90]} — pending {int(days)} days at {stages[key]['label']}.",
                ))
    for r in stages['tenderUnderFinalisation']['rows']:
        days = r.get('daysPendingSinceOpening')
        if isinstance(days, (int, float)) and days >= 45:
            attention.append(dict(
                severity=('high' if days >= 90 else 'medium'), stageKey='tenderUnderFinalisation',
                stageLabel=stages['tenderUnderFinalisation']['label'], sno=r.get('sno'), days=int(days),
                text=f"{(r.get('description') or 'Tender')[:90]} — {int(days)} days since tender opening, not yet finalized.",
            ))
    for r in stages['tenderToBePublished']['rows']:
        days = r.get('durationInCurrentStatus')
        if isinstance(days, (int, float)) and days >= 180:
            attention.append(dict(
                severity=('high' if days >= 270 else 'medium'), stageKey='tenderToBePublished',
                stageLabel=stages['tenderToBePublished']['label'], sno=r.get('sno'), days=int(days),
                text=f"{(r.get('description') or 'Indent')[:90]} — {int(days)} days since sent to purchase unit, tender not yet published.",
            ))
    for r in stages['poIssuedPendingReceipt']['rows']:
        od = _overdue_days(r.get('dueDeliveryDateIso'))
        if od is not None and od > 0:
            attention.append(dict(
                severity=('high' if od >= 60 else 'medium'), stageKey='poIssuedPendingReceipt',
                stageLabel=stages['poIssuedPendingReceipt']['label'], sno=r.get('sno'), days=od,
                text=f"{(r.get('description') or 'PO')[:90]} — delivery due date passed {od} days ago, material not yet received.",
            ))
    for r in stages['balanceIndentsToBeProposed']['rows']:
        od = _overdue_days(r.get('targetDateForProposalIso'))
        if od is not None and od > 0:
            attention.append(dict(
                severity=('high' if od >= 30 else 'medium'), stageKey='balanceIndentsToBeProposed',
                stageLabel=stages['balanceIndentsToBeProposed']['label'], sno=r.get('sno'), days=od,
                text=f"{(r.get('description') or 'Demand')[:90]} — target date for proposal passed {od} days ago, indent not yet raised.",
            ))
    sev_rank = dict(high=0, medium=1, low=2)
    attention.sort(key=lambda x: (sev_rank.get(x['severity'], 9), -x.get('days', 0)))
    log(f"Indent Position: {len(attention)} aging-attention flags raised")

    return dict(
        sourceFile=os.path.basename(path), asOf=(source_as_of or as_of.strftime('%d.%m.%Y')),
        summary=summary, safetyItems=safety_items, stages=stages,
        stageOrder=INDENT_STAGE_KEYS, attention=attention[:40],
    )


# =====================================================================
# AGGREGATION: base payload skeleton (meta / divRollup / shortfall counts / exceptions / ageBuckets)
# =====================================================================
def build_base_payload(cum_records, shortfall_records, exception_data, age_profile, as_of):
    fy_year = as_of.year if as_of.month >= 4 else as_of.year - 1
    meta = dict(
        asOf=as_of.strftime('%d-%b-%Y'),
        fy=f"FY {fy_year}-{str(fy_year + 1)[-2:]}",
        periodCovered=f"{MONTH_ABBR[0]}–{MONTH_ABBR[-1]} {as_of.year}",
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
        # NOTE: the fleet-level division rollup groups by home/original division
        # (unlike progress.machines' "div" field below, which prioritises where a
        # machine is *currently* working) - confirmed against the known-good payload.
        return r['HomeDivision']

    div_rollup = rollup(cum_records, div_key)
    mtype_rollup = rollup(cum_records, lambda r: r['MachineType'])

    # A machine type whose records carry a `Secondary` continuation line (so
    # far: BCM, tracked in both Km and Turnouts/T/O - see extract_cum_progress)
    # gets its own pseudo-type rollup row, e.g. "BCM (T/O)", appended to
    # mtypeRollup alongside the real "BCM" (Km) row. Every chart/picker that
    # already reads mtypeRollup unit-aware (Cumulative Progress and
    # Year-on-Year type pickers/charts, the Overall % Progress "By Machine
    # Type" breakdown) therefore surfaces it automatically; the one place that
    # must NOT show it (the Reason-for-Shortfall type picker, which is keyed
    # off the shortfall log's own machine types and has no T/O concept) is
    # excluded client-side via the isSecondaryMetric flag set below.
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

    # A row with no Machine identified isn't a real per-machine shortfall entry
    # (confirmed against the known-good payload: excluding these 21 orphan rows,
    # mostly leftover header/blank artifacts from the June sheet's second sub-table,
    # is exactly what the original aggregation did).
    cat_counts = Counter(r['Category'] for r in shortfall_records if r.get('Category') and r.get('Machine'))
    by_month_cat = defaultdict(Counter)
    by_div_cat = defaultdict(Counter)
    by_mtype_cat = defaultdict(Counter)
    machine_counts = Counter()
    def norm_div_for_shortfall(v):
        # Daily Progress "Division" free-text also carries deputed-out machines'
        # away-railway names in many raw spellings (Konkan Rly/KRCL, North Western
        # Rly/NWR, Central Rly/CR) plus stray legend/footnote junk - normalize to the
        # same small bucket set the known-good payload uses.
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
            # disp_type() merges PCTM/PCT into UNI, same as progress.machines/mtypeRollup -
            # keeping this consistent so the Reason-for-Shortfall "by machine type" chart's
            # type keys line up with the Progress-tab-style type picker (added 08-Oct-2026).
            mt = disp_type(mtype_of(r['Machine']))
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
    latest_cat_counts = Counter(classify_exc_category(r['Category'])[1] for r in latest_rows if r.get('Category'))

    # Human-readable snapshot-date labels for the dashboard's captions (fixed 28-Sep-2026 -
    # the template used to hard-code "16-Sep-2026" everywhere; the user flagged that these
    # never reflected the latest uploaded data - these are now computed fresh every run so
    # the dashboard's own captions always match the Exception Sheet actually processed).
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


def run(base, outdir):
    os.makedirs(outdir, exist_ok=True)
    as_of = datetime.date.today()

    home_div, bare_type_div, norm = extract_home_div(base)
    cum_records = extract_cum_progress(base, home_div, bare_type_div, norm)
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

    # Data-quality fix for ioh_poh_records' Division field (see extract_ioh_poh()'s
    # comment): backfill from the Cumulative Progress home-division map (home_div,
    # keyed by norm()) wherever the source sheet didn't give a clean MYS/SBC/UBL value.
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
        nm = norm_machine(r['Machine'])
        if not nm:
            continue
        nm = canon_machine_alias(nm)
        dt = parse_date(r['Date'])
        sortkey = (MONTH_ORDER.get(r['Month'], 0), dt or datetime.datetime.min)
        if nm not in machine_latest or sortkey > machine_latest[nm][0]:
            machine_latest[nm] = (sortkey, r['SrDENKey'])

    # 07-Oct-2026 (per user noticing the Sr.DEN roster summed to 88 against an 87-machine
    # fleet): some machines got logged under 2+ differently-punctuated spellings (e.g.
    # shortfall log has both "DTE 56749" and "DTE-56749" - confirmed the same physical
    # machine, both attributed to MYS Sr.DEN/N). Each spelling became its own key above,
    # so the same machine was counted twice in that jurisdiction's roster. Collapse spellings
    # that share a machine_key() (handles space/hyphen and leading-zero differences) ONLY
    # when they all agree on the same Sr.DEN - a collision spanning different jurisdictions
    # is NOT auto-merged, since picking one would be a guess about which zone actually has
    # the machine; it's left as separate entries and flagged below for manual review instead.
    # (The T-28 case that originally motivated this safety net - "T-28(908 A&B)" /
    # "T-28(908A &B)" / "T-28 (403&404)" all sharing machine_key() "T28" - turned out to be
    # 2 different real machines, confirmed by the user 07-Oct-2026, not 1; the 2 spelling
    # variants of the SAME one are now merged earlier via MACHINE_NAME_ALIASES/
    # canon_machine_alias() above, before this grouping even runs, so what reaches here is
    # correctly just 2 distinct entries, which WILL still group under "T28" and WILL still
    # report as a cross-jurisdiction collision below - that's expected and correct now,
    # not a bug: they really are 2 different machines, each staying in its own jurisdiction.)
    _key_groups = defaultdict(list)
    for nm, (sortkey, k) in machine_latest.items():
        _key_groups[machine_key(nm)].append((nm, sortkey, k))
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
        """Group machine_to_srden by keyfn(); keep only groups whose members all point to
        the SAME Sr.DEN (so a collision across jurisdictions is dropped rather than
        resolved by silent dict-overwrite last-wins, which is what plain alnum() used to
        do here)."""
        groups = defaultdict(set)
        for nm, k in machine_to_srden.items():
            groups[keyfn(nm)].add(k)
        return {key: next(iter(ks)) for key, ks in groups.items() if len(ks) == 1}

    machine_to_srden_alnum = _unambiguous_fallback(alnum)
    # Bridges the Progress-sheet's own spelling quirks against the shortfall log's: a
    # parenthetical location suffix the shortfall log doesn't carry ("UTV-001 (SAN)" vs
    # "UTV 001"), a leading zero ("RBMV-006" vs "RBMV 06"), or (via machine_key_aliased,
    # confirmed by the user 07-Oct-2026) a different type-prefix abbreviation for the same
    # machine type ("FRM-1889" Progress-sheet vs "SBCM 1889" shortfall-log).
    machine_to_srden_key = _unambiguous_fallback(machine_key_aliased)

    def lookup_srden(machine_name):
        nm = canon_machine_alias(norm_machine(machine_name))
        if nm in machine_to_srden:
            return machine_to_srden[nm]
        a = machine_to_srden_alnum.get(alnum(machine_name))
        if a:
            return a
        return machine_to_srden_key.get(machine_key_aliased(machine_name))

    latest_month = max((r['Month'] for r in shortfall_records if r.get('SrDENKey')),
                        key=lambda m: MONTH_ORDER.get(m, 0), default=None)
    srden_month_cat = defaultdict(Counter)
    # Each machine's CURRENT jurisdiction roster (fixed 07-Oct-2026 per the user noticing
    # the dropdown's "(N machines)" counts summed to 227 against an 87-machine fleet).
    # Root cause: this used to add a machine to srden_machines[k] for EVERY shortfall-log
    # row attributed to k across the whole Apr-Sep period, so a machine that worked under
    # 2-5 different Sr.DEN jurisdictions at different points in the period (relocations,
    # or just day-to-day attribution drift in the source sheet) got counted in every one of
    # them - 71 of the fleet's 88 logged machines were double/triple/quadruple-counted this
    # way. Fixed to use machine_to_srden (already computed above: each machine's single,
    # most-recently-recorded Sr.DEN) so every machine belongs to exactly one jurisdiction's
    # roster, matching the "machines working in this jurisdiction" framing in the message
    # text and the "Machines in this jurisdiction" drawer - sum of machineCount across all
    # jurisdictions now equals the fleet's logged-machine count, not a multiple of it.
    srden_machines = defaultdict(set)
    for nm, k in machine_to_srden.items():
        srden_machines[k].add(nm)
    # machine-specific shortfall breakdown per Sr.DEN, latest month (added 28-Sep-2026 -
    # user's most emphasized, twice-repeated requirement: "I don't want 32 entries, 30
    # entries. I want the exact message: this machine number is facing this issue in your
    # section.") - excludes the productive-work category, which isn't a shortfall).
    srden_month_machine_cat = defaultdict(lambda: defaultdict(Counter))
    for r in shortfall_records:
        k = r.get('SrDENKey')
        if not k:
            continue
        if r['Month'] == latest_month:
            srden_month_cat[k][r['Category']] += 1
            if r['Category'] != 'Productive Work - Full Progress' and r.get('Machine'):
                srden_month_machine_cat[k][norm_machine(r['Machine'])][r['Category']] += 1

    srden_exceptions = defaultdict(list)
    if latest_ex_day:
        for row in exception_data.get(latest_ex_day, []):
            nm = norm_machine(row['Machine'])
            k = machine_to_srden.get(nm)
            srden_exceptions[k or f"UNMAPPED|{row.get('Division')}"].append(row)

    # ---- progress_fixed: tamping classification + status correction + IOH/POH dates ----
    progress_fixed = []
    for r in cum_records:
        r = dict(r)
        mt = r.get('MachineType')
        info = TAMPING_INFO.get(mt, ('Non-Tamping', 'Unclassified'))
        r['MachineTypeDisplay'] = disp_type(mt)
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

    # keyed by machine_key (not alnum) - see machine_key() docstring: exception-sheet and
    # cumulative-progress machine spellings can differ by leading zeros / location suffixes.
    ioh_poh_live = {}
    if latest_ex_day:
        for row in exception_data.get(latest_ex_day, []):
            if row.get('Category') == 'M/c under IOH/POH':
                ioh_poh_live[machine_key(row['Machine'])] = row
    for r in progress_fixed:
        if r.get('Status') in ('IOH', 'POH'):
            hit = ioh_poh_live.get(machine_key(r['Machine']))
            if hit:
                r['IOHPOH_StartDate'] = hit.get('UnderSince')
                r['IOHPOH_ExpectedCompletion'] = hit.get('Tentative')
                r['IOHPOH_DaysUnder'] = hit.get('DaysUnder')
                r['IOHPOH_Location'] = hit.get('Location')

    # ---- Status correction pass 2: stale IOH/POH flags (fixed 28-Sep-2026) ----
    # Same pattern as the Deputed-NWR correction above: if a machine's Cumulative Progress
    # Status still reads IOH/POH, but it is absent from today's Exception Sheet AND the
    # IOH/POH Planning sheet's own Last IOH/POH Date shows that overhaul completed RECENTLY
    # (within STALE_IOHPOH_WINDOW_DAYS of as-of), the machine has actually returned to
    # active service and the Cumulative Progress Status column just hasn't caught up yet -
    # correct it here so "Machines Working", Fleet Size and the IOH/POH module all agree.
    # IMPORTANT: the recency window matters - a Last IOH/POH Date from years ago is NOT
    # evidence of a just-finished stint (it's simply the most recent overhaul on record;
    # the machine could well be mid-way through a brand-new one). Only a completion date
    # close to today plausibly explains why today's Exception Sheet has dropped it.
    # User-reported example: UTV-58 (POH) and UTV-003 (IOH) both completed on 6/7-Sep-2026.
    STALE_IOHPOH_WINDOW_DAYS = 60
    _ioh_poh_by_machine_early = {machine_key(r['Machine']): r for r in ioh_poh_records}
    for r in progress_fixed:
        if r.get('Status') in ('IOH', 'POH') and machine_key(r['Machine']) not in ioh_poh_live:
            planning_rec = _ioh_poh_by_machine_early.get(machine_key(r['Machine']))
            if not planning_rec:
                continue
            key = 'LastPOH_Date' if r['Status'] == 'POH' else 'LastIOH_Date'
            completion_dt = parse_ddmmyyyy(planning_rec.get(key))
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

    # ---- IOH/POH reconciled: latest Exception Sheet is authoritative (fixed 28-Sep-2026) ----
    ioh_poh_by_machine = {machine_key(r['Machine']): r for r in ioh_poh_records}
    cp_iohpoh = {machine_key(r['Machine']): r for r in cum_records if r['Status'] in ('IOH', 'POH')}
    ex_iohpoh = {machine_key(row['Machine']): row for row in exception_data.get(latest_ex_day, [])
                 if row.get('Category') == 'M/c under IOH/POH'} if latest_ex_day else {}

    def _last_completion(planning_rec, status):
        if not planning_rec:
            return None
        key = 'LastPOH_Date' if status == 'POH' else 'LastIOH_Date'
        return parse_ddmmyyyy(planning_rec.get(key))

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

    # ---- Sr.DEN messages (longest-pending-first, matches build_final_additions.py) ----
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
            t, label = classify_exc_category(row['Category'])
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
        # Sr.DEN message scope (fixed 06-Oct-2026 per user: "These messages shall be based
        # only on open exceptions") - machine_shortfalls above is EVERY machine with a
        # shortfall-log entry anywhere in the month, whether or not it's still an issue
        # today; that made the message list machines with no current exception at all. The
        # message text (and the Shortfall Flag text below) must only name machines that are
        # in TODAY's open-exception list for this Sr.DEN (exc_rows / exc_by_type_rows) -
        # machine_shortfalls itself is left unfiltered since it also feeds the fleet-wide
        # "Key Insights" and "Way Forward" sections, which are month-level, not today-only.
        open_exc_machines = {norm_machine(r['Machine']) for r in exc_rows}
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

        sf_lines = [f"*Shortfall Flag — {k} — {latest_month} 2026*", ""]
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

    # ---- Reason-for-Shortfall tab: Sr.DEN-wise aggregates, full period + last 7 days
    # (added 08-Oct-2026 per the user: "this shall also be sr den wise. Only shortfall
    # reasons related to their machines should come here over last 7 days. Also in the
    # current FY.") Full-period (Apr-Sep, i.e. "current FY") per-Sr.DEN category/machine
    # counts feed the "current FY" part of the new banner; a rolling 7-day window feeds
    # the "last 7 days" part. The window is anchored to the shortfall log's OWN latest
    # dated entry, not meta.asOf - the separate, more-recent Exception Sheet's as-of date
    # - because Daily Progress (this feature's only source) runs behind Exception Sheet
    # data, and "last 7 days" has to mean the log's own last 7 logged days to be meaningful.
    dated_sf = [(r, parse_ddmmyyyy(r.get('Date'))) for r in shortfall_records if r.get('SrDENKey')]
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
            srden_machine_cat_full[k][norm_machine(r['Machine'])][r['Category']] += 1

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
                    srden_machine_cat_last7[k][norm_machine(r['Machine'])][r['Category']] += 1

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
            srden=machine_to_srden.get(norm_machine(r['Machine'])),
            secondary=_secondary_dict(r.get('Secondary')),
        )
        for abbr in MONTH_ABBR:
            nm[abbr.lower()] = r.get(abbr)
        new_machines.append(nm)
    p['progress']['machines'] = new_machines
    p['progress']['tampingRollup'] = dict(Counter(m['tamping'] for m in new_machines))

    for row in p['exceptions']['latest']:
        bucket, clean_label = classify_exc_category(row['Category'])
        row['Type'] = bucket
        row['CategoryLabel'] = clean_label
        row['SrDEN'] = machine_to_srden.get(norm_machine(row['Machine']))
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
        overdue.append(dict(machine=r['Machine'], division=r['Division'], type=disp_type(typecode), tamping=tamping,
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
    age_by_machine = {norm_machine(a['Machine']): a for a in age_profile}
    for m in p['progress']['machines']:
        a = age_by_machine.get(norm_machine(m['machine']))
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
        cat_mtype_machine[cat][disp_type(mtype_of(mach))][mach] += 1
    p['shortfall']['catMachine'] = {cat: sorted(ctr.items(), key=lambda x: -x[1])[:12] for cat, ctr in cat_machine.items()}
    p['shortfall']['catDivMachine'] = {
        cat: {div: sorted(ctr.items(), key=lambda x: -x[1])[:10] for div, ctr in divs.items()}
        for cat, divs in cat_div_machine.items()
    }
    # By-category x by-machine-type drilldown (added 08-Oct-2026 for the Reason-for-
    # Shortfall tab's redesigned "By Machine Type" chart - mirrors catDivMachine but keyed
    # by machine type instead of division).
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
    # This v3 rebuild replaces mtypeRollup wholesale, which silently dropped the
    # BCM (T/O) pseudo-type row built earlier (around `sec_rollup`) -- re-derive
    # it here from the final p['progress']['machines'] list (which does carry
    # `secondary` after the v2 rebuild) so it survives into the payload.
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
    # 07-Oct-2026: was its own plain-alnum index (m2s_idx), which (a) didn't bridge a
    # parenthetical suffix the shortfall log doesn't carry ("UTV-001 (SAN)" vs "UTV 001")
    # or a leading zero ("RBMV-006" vs "RBMV 06"), and (b) used silent last-wins on any
    # alnum collision. Reuses lookup_srden() (defined above) instead, which tries the
    # same alnum step first but then also tries the ambiguity-aware machine_key()
    # fallback, and never guesses across a jurisdiction-spanning collision.
    for m in p['progress']['machines']:
        if not m.get('srden'):
            hit = lookup_srden(m['machine'])
            if hit:
                m['srden'] = hit
    machine_idx = {alnum(m['machine']): m for m in p['progress']['machines']}

    for r in ds_records:
        r['srdenKey'] = canon_ds_srden(r.get('div'), r.get('srden'))
        hit = machine_idx.get(alnum(r.get('machine') or ''))
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
        # Secondary-unit (BCM Km/T/O split) current-year sums, keyed the same
        # way as prior_by_type_secondary/prior_by_machine_secondary above, so
        # a machine's Turnout progress gets its own YoY row instead of being
        # dropped or blended into its Km figure.
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
                cur_by_machine_secondary[f"{machine_key(m['machine'])}|{sec['unit']}"] = ss
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
                pv = prior_by_machine.get(machine_key(m['machine']))
                pv = round(pv, 2) if pv is not None else 0.0
                cv = round(cv, 2)
                delta = round(cv - pv, 2)
                pct = round((cv - pv) / pv * 100, 1) if pv else None
                yoy_by_machine.append(dict(machine=m['machine'], type=m['type'], div=m['div'],
                                            priorYear=pv, currentYear=cv, delta=delta, pctChange=pct))
                sec = m.get('secondary')
                if sec:
                    key = f"{machine_key(m['machine'])}|{sec['unit']}"
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

    # ---- IOH/POH days-elapsed AND days-designated (fixed 28-Sep-2026) ----
    for m in p['progress']['machines']:
        if m.get('iohPohStart'):
            dt = parse_ddmmyyyy(m['iohPohStart'])
            if dt:
                m['iohPohDaysElapsed'] = (datetime.datetime.combine(as_of, datetime.time()) - dt).days
                exp = parse_ddmmyyyy(m.get('iohPohExpected'))
                if exp:
                    m['iohPohTotalSpan'] = (exp - dt).days
                    m['iohPohDays'] = m['iohPohTotalSpan']  # designated days = target date - start date (TMM/IRTMM)

    # =====================================================================
    # Fleet Health — IOH/POH Schedule (REBUILT 01-Oct-2026). This is now the
    # sole, primary source for the Fleet Health tab: current-year events come
    # from the Dy.CE/TM office letter (extract_ioh_poh_letter, above);
    # FY2027-28 through FY2029-30 events come strictly from the "IOH POH
    # Details 27-30.xlsx" workbook's own per-machine forecast columns
    # (ioh_poh_records[*]['Schedule']). Neither the Exception Sheet nor
    # Cumulative Progress Status feeds any part of this structure. Sr.DEN is
    # attached on a best-effort basis from the same machine_to_srden map used
    # elsewhere (a machine currently under IOH/POH, or scheduled years out,
    # often has no live Daily Progress row to derive a Sr.DEN from - shown as
    # "Not currently mapped" rather than guessed).
    # =====================================================================
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

    # Fold the CPOH/RYP priority-release events into the same current_events list
    # that drives completedCount/byDivision/bySrden/statusCounts and the Fleet
    # Health "Completed" drawer (fixed 05-Oct-2026). Before this, a POH completed
    # at CPOH/RYP (e.g. MPT 2011, DUO 8128) was only visible in the separate "POH
    # Release Priority List" card - invisible to every KPI/total/grouped-by-
    # division view above it, which the user caught by noticing a known CPOH/RYP
    # completion missing from the Completed-by-division breakdown. Checked for
    # machine-overlap with the main depot schedule before merging: one machine
    # (FRM 1889) appears in both, but as two genuinely distinct events (an IOH at
    # ZBD/YPR and a separate POH-at-CPOH/RYP release) - not a duplicate, so a
    # straight append is correct, not a dedupe-by-machine merge.
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
        typecode = disp_type(mt.group(0).upper()) if mt else 'OTHER'
        for fy in FUTURE_FY_LABELS:
            ioh_v, poh_v = (sched.get(fy) or [None, None])
            for kind, val in (('IOH', ioh_v), ('POH', poh_v)):
                if val in (None, '', '-'):
                    continue
                fy_a, fy_b = fy.split('-')
                future_events.append(dict(
                    machine=r['Machine'], machineKey=machine_key(r['Machine']),
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
            source=os.path.basename(one_file(f"{base}/IOH POH Planning/*.xlsx", required=False) or '') or None,
            fyLabels=FUTURE_FY_LABELS,
            events=future_events,
            byDivision=future_div_summary,
            totalCount=len(future_events),
        ),
        norms=ioh_poh_norms,
    )

    # CPOH/RYP forward POH planning (2027-30) — kept as its own structure rather
    # than folded into futureYears, since it's a richer, CPOH/RYP-specific
    # proposal (per-engine detail, governing reason, tamping-unit demand) from a
    # different, more authoritative source than the generic per-division
    # forecast in the "IOH POH Details 27-30" workbook. division/srden attached
    # on the same best-effort basis as the rest of Fleet Health.
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
    if len(sys.argv) != 3:
        print("Usage: python3 run_pipeline.py <MachineMonitoringFolder> <OutputDir>", file=sys.stderr)
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
