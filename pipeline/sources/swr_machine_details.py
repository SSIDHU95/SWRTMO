"""Adapter for the `swr_machine_details` table.

Live sync: swr_machine_details_sync.gs. One row, pushed daily, shaped:
    {
      "ageProfile": [
        {"slNo": 1, "machineType": "BCM", "machine": "BCM-01",
         "commissioningDate": "12.03.2010", "age": "16 years"},
        ...
      ],
      "originalAllocation": {
        "mys": ["BCM-01", ...], "sbc": [...], "ubl": [...]
      }
    }
Covers input-data-specification.md items 1 ("SWR Machine Details --
Original Allocation") and 2 ("...Commissioning / Age"), which live as two
tabs of this one Google Sheet and are pushed together in one payload.

NOTE (confirmed live, 25-Nov-2026): the shape above is the documented/
intended one, but swr_machine_details_sync.gs is currently pushing
'commissioningDate' and 'age' SWAPPED -- every real row looks like
{"commissioningDate": "-", "age": "Fri Apr 01 2005 13:30:00 GMT+0530
(India Standard Time)", ...}, i.e. the real commissioning date (as a JS
Date.toString() dump) in 'age', and a literal '-' placeholder in
'commissioningDate'. extract_age_profile() below is written defensively
against this -- it tries to parse an actual date out of EITHER field --
so it keeps working whether or not/whenever that .gs bug gets fixed on
the Sheets side.

This replaces TWO local-file functions in run_pipeline.py:
  - extract_home_div(base)   -> (home_div, bare_type_div, norm)
  - extract_age_profile(base) -> [{MachineType, Machine, CommissioningDate,
                                    AgeYears}, ...]
Both read the same live row here, so fetch_payload() is called once and
the two extract_* functions below each take that same payload dict,
rather than each doing their own REST call.
"""
import datetime
import re

import pipeline_helpers as ph
import supabase_client

TABLE = "swr_machine_details"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


# ---------------------------------------------------------------------------
# Fixed 25-Nov-2026 per Shruti ("I am able to see only the first 0 to 10
# years, not further" -- Fleet Age Distribution showing ALL 99 machines in
# the "0-10 yrs" bucket and nothing in any other one). Root cause confirmed
# via a direct live Supabase pull of this table's real payload: swr_machine_
# details_sync.gs currently pushes the 'age' and 'commissioningDate' fields
# SWAPPED -- every row's real commissioning date (as a JS Date.toString()
# dump, e.g. "Fri Apr 01 2005 13:30:00 GMT+0530 (India Standard Time)")
# lands in 'age', while 'commissioningDate' is left as a literal '-' for
# every row. The old extract_age_profile() below took 'age' as a
# pre-computed "N years" string and pulled the first run of digits out of
# it with a bare regex; since nearly every commissioning date in the source
# sheet is month/year-only, Sheets/Apps Script normalizes it to the 1st of
# that month, so the regex always found "01" first and reported age=1 for
# virtually every machine -- exactly the observed symptom.
#
# Fixed defensively rather than assuming the Apps Script swap gets
# corrected upstream (same approach already used for the Original
# Allocation/Cumulative Progress BRM/PBR aliasing above): try to parse
# EITHER field as an actual date (commissioningDate first -- the
# documented/intended field -- then age, which is where today's swapped
# value actually lives) and compute AgeYears ourselves from whichever one
# parses, using today's date rather than trusting a possibly-stale
# precomputed number. Only if neither field parses as a date at all do we
# fall back to the OLD "number embedded in text" reading of 'age' (e.g. a
# sheet that really does write "16 years" by hand one day), so a future,
# differently-shaped sync still degrades sensibly instead of going straight
# to None.
#
# Date shapes handled (all seen live, across this table and the newly
# shared IOH/POH Planning sheet that uses the same "Date of Commissioning"
# column): a JS Date.toString() dump; 'Mon-YY'/'Mon.YY' (e.g. "Aug-93",
# "Dec.15"); 'dd.mm.yy'/'dd.mm.yyyy' (and '-'/'/' separators); and an ISO
# timestamp. Two-digit years use the same 69-99->19xx / 00-68->20xx pivot
# as the rest of this project.
# ---------------------------------------------------------------------------
_MONTH_ABBR = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}
_JS_DATE_TOSTRING_RE = re.compile(
    r'^\w{3}\s+([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})\s+\d{1,2}:\d{2}:\d{2}'
)
_ISO_DATE_RE = re.compile(r'^(\d{4})-(\d{1,2})-(\d{1,2})')
_MON_YY_RE = re.compile(r'^([A-Za-z]{3})[.\-/](\d{2})$')
_DMY_RE = re.compile(r'^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$')


def _pivot_year(yy):
    return 1900 + yy if yy >= 69 else 2000 + yy


def _parse_commission_date(v):
    """Best-effort parse of a 'date of commissioning'-ish value into a
    datetime.date, or None if nothing recognised is found (blank, '-', or
    genuinely free text like a hand-typed "16 years")."""
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = str(v).strip()
    if not s or s == '-':
        return None

    m = _JS_DATE_TOSTRING_RE.match(s)
    if m:
        mon = _MONTH_ABBR.get(m.group(1).lower())
        if mon:
            try:
                return datetime.date(int(m.group(3)), mon, int(m.group(2)))
            except ValueError:
                return None
        return None

    m = _ISO_DATE_RE.match(s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None

    m = _MON_YY_RE.match(s)
    if m:
        mon = _MONTH_ABBR.get(m.group(1).lower())
        if mon:
            return datetime.date(_pivot_year(int(m.group(2))), mon, 1)
        return None

    m = _DMY_RE.match(s)
    if m:
        day, mon, year_txt = int(m.group(1)), int(m.group(2)), m.group(3)
        year = int(year_txt) if len(year_txt) == 4 else _pivot_year(int(year_txt))
        try:
            return datetime.date(year, mon, day)
        except ValueError:
            return None

    return None


def _age_years_from_date(d, today=None):
    """Whole years between a commissioning date and today (or None for a
    bad/future date)."""
    today = today or datetime.date.today()
    if d is None or d > today:
        return None
    return today.year - d.year - ((today.month, today.day) < (d.month, d.day))


# ---------------------------------------------------------------------------
# norm_alloc_name() - copied verbatim from inside run_pipeline.py's
# extract_home_div(). Kept here (no longer used for the home_div/
# bare_type_div keys themselves - see 18-Nov-2026 fix below - but still
# exported/available in case a future caller wants a plain alnum-only
# normalization of an Original Allocation cell without the type-alias
# step _home_div_key() below adds).
# ---------------------------------------------------------------------------
def norm_alloc_name(name):
    if name is None:
        return None
    s = str(name).upper()
    s = re.sub(r'\(.*?\)', '', s)
    return re.sub(r'[^A-Z0-9]', '', s)


# ---------------------------------------------------------------------------
# Fixed 18-Nov-2026 per Shruti's corrected SWR_Machine_Details_1.xlsx: the
# Original Allocation list (this table's own "originalAllocation" sheet)
# always names BRM machines "BRM" (confirmed directly in her workbook's
# Div Wise Machine Original Alloc sheet: "BRM-33", "BRM-56669", "BRM-27",
# "BRM 56680", "BRM-04" - never "PBR"), but the Cumulative Progress sheet
# sometimes names those same physical machines "PBR" instead (same
# FRM/SBCM, PBR/BRM aliasing Shruti described, and the same
# MACHINE_PREFIX_ALIASES pipeline_helpers.machine_key_aliased() already
# uses for the Sr.DEN and Unit Cost joins elsewhere in this pipeline).
#
# extract_cum_progress() (cumulative_progress.py) looks a machine's home
# division up in `home_div` by its own raw name - plain norm_alloc_name()
# used to key that dict, so "BRM-04" (Original Allocation) and "PBR 04"
# (Cumulative Progress) normalized to two DIFFERENT keys ("BRM04" vs
# "PBR04") and never matched. For a machine currently away at CPOH/RYP
# for IOH/POH (CurrentDIV is "IOH"/"POH", not MYS/SBC/UBL), that lookup
# miss is exactly what made its division show as "UNKNOWN" on the live
# site - confirmed against Shruti's own screenshot of "PBR" bucketed
# under "UNKNOWN" division.
#
# The fix: key home_div/bare_type_div with the SAME alias-aware function
# the rest of this pipeline already trusts for this exact problem
# (machine_key_aliased(norm_machine(...))), instead of the plain
# alnum-only norm_alloc_name(). "BRM-04" and "PBR 04" both alias to the
# same "PBR4" key this way (machine_key_aliased always normalizes a
# BRM/SBCM/PCTM-family raw type to its alias form), so the lookup now
# succeeds regardless of which naming convention either sheet happens to
# use for a given machine.
# ---------------------------------------------------------------------------
def _home_div_key(name):
    return ph.machine_key_aliased(ph.norm_machine(name))


def extract_home_div(payload):
    """-> (home_div, bare_type_div, norm). Mirrors run_pipeline.py's
    extract_home_div(base) exactly, except the source is already three
    plain machine-name lists (one per division) instead of a raw sheet
    grid - the live sync's parseOriginalAllocation() already dropped the
    "-" placeholders and the TOTAL/GRAND TOTAL summary rows on the Apps
    Script side, so no further filtering is needed here.

    Fixed 18-Nov-2026: keys are now built with _home_div_key() (alias-
    aware) instead of plain norm_alloc_name() - see that function's own
    comment above for why. `norm` (the 3rd return value) is also now
    _home_div_key, since every caller re-uses it to build the matching
    lookup key for a machine from a DIFFERENT sheet (extract_cum_progress's
    `home_div.get(norm(machine))`) - it must stay the exact same function
    the keys above were built with."""
    alloc = (payload or {}).get("originalAllocation") or {}
    home_div = {}
    bare_type_div = {}
    for div, key in (("MYS", "mys"), ("SBC", "sbc"), ("UBL", "ubl")):
        for val in alloc.get(key) or []:
            val = str(val).strip()
            if not val or val == "-":
                continue
            home_div[_home_div_key(val)] = div
            if re.match(r'^[A-Za-z]+$', val):
                # Fixed 18-Nov-2026: disp_type() too, so a bare-type
                # Original Allocation row using a raw alias prefix (e.g.
                # a lone "PBR" row, were one ever entered) still matches
                # the canonical type code bare_type_div.get() is looked
                # up with in extract_cum_progress().
                bare_type_div[ph.disp_type(val.upper())] = div
    return home_div, bare_type_div, _home_div_key


def extract_age_profile(payload, today=None):
    """-> [{MachineType, MachineTypeFullName, Machine, CommissioningDate,
    AgeYears}, ...]. Mirrors run_pipeline.py's extract_age_profile(base).
    The live sync's parseAgeProfile() already skipped the machine-type
    banner rows (e.g. "BCM  (7 machines)") on the Apps Script side, so
    every row here is already a real machine.

    Fixed 18-Nov-2026 per Shruti ("the machine details are utilized
    elsewhere also, in age profile and all -- you connect it there
    also"): MachineType is now disp_type()'d (so a sheet that still
    types a machine's raw alias prefix - FRM/PBR/SQRS/T/WST - reports its
    real canonical type here too, consistent with every other tab), and
    MachineTypeFullName is attached from the same pipeline_helpers.
    MTYPE_FULL_NAME table the rest of the dashboard now uses.

    Fixed 25-Nov-2026 (see the module-level comment above _parse_commission_
    date for the full root cause -- a swr_machine_details_sync.gs field-swap
    bug was putting every machine's real commissioning date into 'age'
    instead of 'commissioningDate'): AgeYears is no longer read verbatim off
    'age' at all. We now try to parse an actual date out of
    'commissioningDate' first (the documented, intended field), then 'age'
    (where the live bug's real date currently lands), and compute AgeYears
    ourselves from whichever one parses -- so this keeps working correctly
    regardless of whether/when that Apps Script bug gets fixed on the
    Sheets side. `today` is only for tests that need a fixed reference
    date; production calls always get a live payload and leave it as the
    real today."""
    rows = (payload or {}).get("ageProfile") or []
    as_of = today or datetime.date.today()
    out = []
    for r in rows:
        machine = r.get("machine")
        if not machine:
            continue

        commissioning_raw = r.get("commissioningDate")
        age_raw = r.get("age")

        commission_date = _parse_commission_date(commissioning_raw)
        if commission_date is None:
            commission_date = _parse_commission_date(age_raw)

        if commission_date is not None:
            age = _age_years_from_date(commission_date, as_of)
            disp_commission_date = commission_date.strftime('%d.%m.%Y')
        else:
            # Neither field parsed as a date -- last-resort fallback to the
            # original behaviour (a hand-typed "16 years"-style age text).
            agem = re.search(r'(\d+)', str(age_raw)) if age_raw else None
            age = int(agem.group(1)) if agem else None
            disp_commission_date = (
                str(commissioning_raw).strip()
                if commissioning_raw not in (None, '-', '') else None
            )

        raw_type = r.get("machineType")
        disp_mtype = ph.disp_type(str(raw_type).strip().upper()) if raw_type else None
        out.append(dict(
            MachineType=disp_mtype,
            MachineTypeFullName=ph.MTYPE_FULL_NAME.get(disp_mtype) if disp_mtype else None,
            Machine=str(machine).strip(),
            CommissioningDate=disp_commission_date,
            AgeYears=age,
        ))
    return out
