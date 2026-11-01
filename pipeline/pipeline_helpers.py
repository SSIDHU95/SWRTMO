"""Pure helper functions and classification tables, lifted verbatim from
claude/pipeline/run_pipeline.py (the local-file pipeline). These have no
file I/O of their own - they just normalize strings, classify free text, or
compute derived values - so the Supabase-backed source adapters and the
aggregation engine both import them unchanged rather than re-implementing
or duplicating this logic. Keeping this module a straight copy of the
originals (not a rewrite) is deliberate: run_pipeline.py is still the
source of truth for the local-file path, and this project has already paid
once for every edge case baked into these functions (BCM's secondary unit,
the FRM/SBCM prefix aliasing, T-28's two physical units, the 19-category
shortfall taxonomy, brand/stage classification for Tenders & Contracts,
etc.) - re-deriving them independently here would just reopen bugs that
were already found and fixed once.

If run_pipeline.py's own copy of any of these functions changes, update
this file to match by hand (there is no shared import between the two
pipelines, since one runs in Cowork against local files and the other runs
in GitHub Actions against Supabase with no access to this project's other
files) - a mismatch here would produce a merge that quietly disagrees with
what the Cowork-side Artifact shows for the same date.
"""
import re
import datetime


MONTH_ABBR = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep']
MONTH_NUM = {'Apr': 4, 'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8, 'Sep': 9}


# ---------------------------------------------------------------------------
# Shortfall-category fallback classifier (used only when a source's own
# tagged category is missing - see run_pipeline.py's long comment above its
# copy of this function for the full rationale).
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


def norm_machine(m):
    if not m:
        return None
    s = re.sub(r'\s+', ' ', str(m)).strip().upper()
    return s.replace(' - ', '-').replace(' -', '-').replace('- ', '-')


def alnum(m):
    return re.sub(r'[^A-Z0-9]', '', str(m).upper())


def machine_key(m):
    s = re.sub(r'\(.*?\)', '', str(m).upper())
    s = re.sub(r'[^A-Z0-9]', '', s)
    mtp = re.match(r'([A-Z]+)(\d+)', s)
    if mtp:
        return mtp.group(1) + str(int(mtp.group(2)))
    return s


MACHINE_PREFIX_ALIASES = {'SBCM': 'FRM', 'BRM': 'PBR', 'PCTM': 'PCT'}


def machine_key_aliased(m):
    s = machine_key(m)
    mtp = re.match(r'([A-Z]+)(\d+)$', s)
    if mtp:
        return MACHINE_PREFIX_ALIASES.get(mtp.group(1), mtp.group(1)) + mtp.group(2)
    return s


MACHINE_NAME_ALIASES = {
    'T-28(908A &B)': 'T-28 (908 A&B)',
    'T-28(908 A&B)': 'T-28 (908 A&B)',
    'T-28 (908A&B)': 'T-28 (908 A&B)',
    'T-28 (03&04)': 'T-28 (403&404)',
}


def canon_machine_alias(nm):
    return MACHINE_NAME_ALIASES.get(nm, nm)


def mtype_of(name):
    m = re.match(r'[A-Za-z]+', str(name).strip())
    return m.group(0).upper() if m else 'OTHER'


UNI_MERGE = {'PCTM': 'UNI', 'PCT': 'UNI'}


def disp_type(t):
    return UNI_MERGE.get(t, t)


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
    'RMV': ('Non-Tamping', 'RMV - maintained by Track Machine Dept but operated by divisions; unique to SWR; no Railway Board target (user-confirmed 21-Oct-2026).'),
}

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
    s = str(raw or '').strip()
    for pat, (bucket, label) in EXC_TYPE_RULES:
        if pat.search(s):
            return bucket, label
    return 'other', (s or 'Uncategorised')


ZONE_LETTER = {'CENTRAL': 'C', 'EAST': 'E', 'WEST': 'W', 'NORTH': 'N', 'SOUTH': 'S', 'HQ': 'HQ'}


def canon_srden(div, raw):
    if div not in ('MYS', 'SBC', 'UBL') or not raw:
        return None
    m = re.search(r'DEN[./]?\s*/?\s*([A-Za-z]+)', str(raw), re.I)
    if not m:
        return None
    code = m.group(1).upper()
    if code not in ('N', 'S', 'E', 'W', 'C', 'HQ'):
        return None
    return f"{div} Sr.DEN/{code}"


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


# ---------------------------------------------------------------------------
# NOT part of the run_pipeline.py copy above - this one is new to the
# Supabase-backed pipeline, needed because of a real difference between the
# two data paths. run_pipeline.py's local-file extractors call
# Utilities.formatDate-equivalent logic (openpyxl) and hand back plain
# 'dd.mm.yyyy' strings for every date cell. The live .gs syncs are not
# always so disciplined: several of them (e.g. exception_sync.gs's
# parseTab(), which reads col.underSince/col.tentative as RAW cell values
# rather than running them through a cellVal()-style formatter first) let
# a Date-typed cell fall straight into JSON.stringify(), which serializes
# it as an ISO-8601 UTC timestamp ("2026-09-01T00:00:00.000Z") instead of
# the 'dd.mm.yyyy' string the aggregation/UI layer expects everywhere else.
# A plain string cell (already 'dd.mm.yyyy', 'dd-mm-yyyy', or free text)
# must still pass through unchanged. This normalizes either shape to the
# one 'dd.mm.yyyy' format run_pipeline.py has always produced, so the
# merged payload looks the same regardless of which path a given day's
# data took.
# ---------------------------------------------------------------------------
_ISO_DATE_RE = re.compile(r'^(\d{4})-(\d{2})-(\d{2})T\d{2}:\d{2}:\d{2}')


def coerce_date_ddmmyyyy(v):
    """Normalizes a live-sync date value (ISO timestamp, 'dd.mm.yyyy',
    'dd-mm-yyyy', 'dd/mm/yyyy', or already-clean text) to 'dd.mm.yyyy'.
    Returns None for empty input, and returns the original trimmed text
    unchanged if it matches none of the recognised date shapes (so a
    genuinely free-text remark in a date column surfaces as-is rather than
    silently disappearing)."""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    m = _ISO_DATE_RE.match(s)
    if m:
        year, mon, day = m.group(1), m.group(2), m.group(3)
        return f"{day}.{mon}.{year}"
    m2 = re.match(r'^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$', s)
    if m2:
        day, mon, year = m2.group(1).zfill(2), m2.group(2).zfill(2), m2.group(3)
        return f"{day}.{mon}.{year}"
    return s


UNIT_CANON = {'KM': 'Km', 'T/O': 'T/O', 'R/S': 'R/S', 'UNIT': 'Unit'}


def norm_unit(v):
    if v is None:
        return v
    s = str(v).strip()
    return UNIT_CANON.get(s.upper(), s)


INFRA_RESP_PATTERNS = [
    (re.compile(r'not\s*feasible', re.I), 'Not Feasible'),
    (re.compile(r'gati\s*shakti', re.I), 'Gati Shakti'),
    (re.compile(r'\bcon\s*org\b|\bcn\s*org\b|\bby\s*cn\b|\bcn\s*work\b|\bby\s*con\b|\bdone\s*by\s*con\b', re.I), 'CN (Construction Org)'),
    (re.compile(r'propos(?:ed|e)\s*by\s*div|div\w*\s*in\s*irpsm|yet\s*to\s*be\s*propos', re.I), 'Division (IRPSM)'),
]


def classify_infra_responsibility(remark):
    s = str(remark or '').strip()
    if not s:
        return None
    for pat, label in INFRA_RESP_PATTERNS:
        if pat.search(s):
            return label
    return 'Other / Unspecified'


def normalize_infra_status(s):
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


def parse_station_block(text):
    if not text:
        return []
    s = str(text).strip()
    m = re.match(r'^\(?\s*\d+\s*\)?', s)
    if m:
        s = s[m.end():].strip()
    s = s.lstrip(':').strip()
    return [x.strip().rstrip('.') for x in re.split(r'[,\n]+', s) if x.strip()]


INFRA_SRDEN_WORD_CODE = {
    'N': 'N', 'NORTH': 'N', 'S': 'S', 'SOUTH': 'S', 'E': 'E', 'EAST': 'E',
    'W': 'W', 'WEST': 'W', 'C': 'C', 'CEN': 'C', 'CENTRAL': 'C', 'CENTRE': 'C',
    'HQ': 'HQ', 'HQRS': 'HQ', 'HEADQUARTERS': 'HQ',
}


def canon_srden_word(div, raw):
    """Like canon_srden() but for the Sidings/Rest House sheet, whose Sr.DEN
    column uses full-word labels ('Sr.DEN/North') rather than the single-letter
    codes the Daily Progress sheet uses. Returns 'DIV Sr.DEN/X', or - for the
    rare jurisdiction label that doesn't follow that pattern (e.g. a
    'New Line' construction stretch) - 'DIV <raw text>' so it is still
    bucketed rather than silently dropped."""
    if div not in ('MYS', 'SBC', 'UBL') or not raw:
        return None
    m = re.search(r'DEN[./]?\s*/?\s*([A-Za-z]+)', str(raw), re.I)
    if m:
        code = INFRA_SRDEN_WORD_CODE.get(m.group(1).strip().upper())
        if code:
            return f"{div} Sr.DEN/{code}"
    return f"{div} {str(raw).strip()}"


def parse_infra_target_text(s):
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
    dict(id='rmv_engine', label='AMC for RMV machine Engines', keywords=[]),
    dict(id='greaves', label='AMC for Greaves make Engines', keywords=['greaves']),
    dict(id='kirloskar', label='M/s. Kirloskar (Surya Power Services)', keywords=['kirloskar']),
    dict(id='uniwave', label='M/s. Uniwave Engineers LLP (Volvo engines)', keywords=['uniwave', 'volvo']),
]


def brand_hits(text):
    t = str(text or '').lower()
    return [b['id'] for b in BRAND_KEYWORDS if any(kw in t for kw in b['keywords'])]


def num(v):
    if v in (None, '', '-'):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v)
    s = re.sub(r'(?i)rs\.?', '', s)
    s = s.replace('₹', '')
    s = re.sub(r'/-\s*$', '', s)
    s = re.sub(r'[,\s]', '', s)
    s = s.strip('-').strip()
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None


def clean_text(v):
    if v is None:
        return None
    s = re.sub(r'\s+', ' ', str(v)).strip()
    return s or None


_DATE_IN_TEXT_RE = re.compile(r'\b(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})\b')


def dates_in_text(text):
    out = []
    for d, m, y in _DATE_IN_TEXT_RE.findall(str(text or '')):
        dt = parse_ddmmyyyy(f"{d}.{m}.{y}")
        if dt:
            out.append(dt)
    return out


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
    s = str(remark or '')
    for key, label, pat in PROPOSAL_STAGE_ORDER:
        if pat.search(s):
            return key, label
    return 'pending', (clean_text(remark) or 'No update recorded')


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
    return 'unclassified', (clean_text(remark) or 'Status not recorded'), 0


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
    return 'in_process', (clean_text(remark) or 'Agreement in process')
