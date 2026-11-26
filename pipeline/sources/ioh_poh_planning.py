"""Adapter for the `ioh_poh_planning` table.

Live sync: ioh_poh_planning_sync.gs. One row, pushed daily, covering the
IOH/POH Planning workbook's two sheets run_pipeline.py's extract_ioh_poh
(base) and extract_ioh_poh_duration_norms(base) read, PLUS a third,
"Progress 2026-2027", that neither of those reads today (not in
input-data-specification.md - synced anyway per the user, same as several
other "for later use" sources this round, left untouched by this adapter).
Payload shape:
    {
      "forecast": {"sheetName": "IOH POH Calculation", "headers": [...],
          "rows": [{"S. No.": 1, "Name of Machines": "CSM-939", "Divn.":
              "SBC", "Date of Commissioning": "Aug-93", "Last IOH Engine
              Hrs.": 41746.0, "Last IOH Date": "31.12.2025", "Last POH
              Engine Hrs.": ..., "Last POH Date": ..., "Engine Hrs. as on
              <dd.mm.yyyy>": ..., "Engine Hr Run/month": ..., "NEXT DUE in
              Months IOH": ..., "NEXT DUE in Months POH": ..., "<FY> IOH":
              ..., "<FY> POH": ..., "Remark": ...}, ...]},
      "durationNorms": {"headers": [...], "rows": [{"Machine Type": "DTE",
          "Duration (in Days) IOH": 30, "Duration (in Days) 1st POH": 75,
          "Duration (in Days) Subsequent POHs": 90, "Duration (in Days)
          IOH Hrs": 2000, "Duration (in Days) POH Hrs": "8000 , 14000"}]},
      "currentYearProgress": {"poh": {...}|None, "ioh": {...}|None}  # not
          read by run_pipeline.py's two functions - kept raw, unused here
    }

Confirmed against the real uploaded workbook (the same file run_pipeline.
py's local-file path reads): the live sheet's forecast-table COLUMN LAYOUT
has drifted from what run_pipeline.py's extract_ioh_poh(base) assumes (two
extra "Previous year engine Hr" columns and an extra "NEXT DUE in Months
2000/Avg Run" column have been inserted since that function's hardcoded
column numbers were written - confirmed: local column 15 would now land
on "NEXT DUE in Months IOH", not "Engine Hr Run/month"). This adapter
reads every field by its own header TEXT instead (same defensive approach
already used for tenders_contracts), which also means it survives a
further column insertion without a code change. The year-schedule columns
("<FY> IOH"/"<FY> POH") are discovered the same way - any header matching
_YEAR_COL_RE, not a fixed year list - mirroring run_pipeline.py's own
year_headers scan (which looks for a bare 'YYYY-YYYY' string rather than a
hardcoded year list) but working off header text instead of a column scan.
"""
import datetime
import re

import pipeline_helpers as ph
import supabase_client

TABLE = "ioh_poh_planning"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def _g(row, *keys):
    for k in keys:
        v = row.get(k)
        if v not in (None, ''):
            return v
    return None


def _g_prefix(row, prefix):
    """First key in row starting with `prefix` (case-insensitive) - used for
    the forecast sheet's 'Engine Hrs. as on <dd.mm.yyyy>' column, whose
    exact header text changes with whatever date the source sheet's own
    author last typed into it."""
    pl = prefix.lower()
    for k, v in row.items():
        if k and k.lower().startswith(pl) and v not in (None, ''):
            return v
    return None


def _fmt_date(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime('%d.%m.%Y')
    dt = ph.parse_ddmmyyyy(v)
    return dt.strftime('%d.%m.%Y') if dt else ph.clean_text(v)


def _fmt_commission(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime('%d-%m-%Y')
    return v


_YEAR_COL_RE = re.compile(r'^(\d{4}-\d{2,4})\s+(IOH|POH)$', re.I)


# ---------------------------------------------------------------------------
# Added 26-Nov-2026 per Shruti's detailed Fleet Health rebuild request: the
# "Division-wise FY 2026-27 Schedule" / KPI tiles ("Planned This Year",
# "Overdue for Release", "Completed", "Ongoing / Under Attention",
# "Scheduled / Upcoming") used to come from extract_ioh_poh_letter(), which
# has NEVER had a live source (a Word .docx office letter, never synced --
# see that function's own comment) and so was always empty. She pointed out
# that this table's own live payload already carries a third sheet,
# "Progress 2026-2027" (forecast/durationNorms are the other two), synced
# into payload.currentYearProgress.poh/.ioh -- confirmed live via Supabase,
# just never read by this adapter before now ("kept raw, unused here").
# That sheet is exactly the real, up-to-date "who's been released, who's
# still pending" picture for the current FY, so it replaces the never-
# connected office letter as the source for all of the above.
#
# Read defensively by header TEXT with BOTH the shape currently live in
# Supabase (SL.NO/MACHINE/Belongs to DIV/Attend at/Actually planned/FROM/TO/
# ...) and the shape of Shruti's own freshly-cleaned-up sheet (S. No./
# Machine No./Parent Division/Planned At/Month of Planning/Date of Arrival
# of machine to depot/Actual Date of Completion of IOH/POH/...), same
# defensive approach as extract_ioh_poh() above -- whichever shape
# ioh_poh_planning_sync.gs is actually emitting on a given day, this keeps
# working without a code change, and survives the sync catching up to her
# sheet edits on its own schedule.
#
# Classification (Shruti's own rules, 26-Nov-2026):
#   - 'Completed': a real completion date is on record (not a "TDC: ..."
#     estimate -- that's still pending, just with a tentative target).
#   - 'Ongoing / Under Attention': the machine has arrived/started (a real
#     arrival date) but has no real completion date yet.
#   - no arrival date at all -> the machine hasn't been released from its
#     division yet. Split by the machine's own planned month vs as_of:
#       - 'Overdue for Release': planned month is before the current month
#       - 'Upcoming': planned month is this month or later
#   This must be dynamic (recomputed from `as_of`, not hardcoded to any one
#   month) so a machine automatically moves from Upcoming to Overdue for
#   Release the month after its planned month passes, with no code change
#   needed run to run.
# ---------------------------------------------------------------------------
_MONTH_ABBR_CYP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def _parse_any_date(v):
    """dd.mm.yyyy / dd.mm.yy / dd-mm-yyyy / an already-native date -> a
    datetime.date, or None. Unlike pipeline_helpers.parse_ddmmyyyy(), this
    also accepts a 2-digit year ('28.05.26') -- both shapes appear in this
    sheet's own FROM/TO-equivalent columns. Returns None (rather than a
    bogus date) for a "TDC: ..." tentative-estimate string -- that is
    deliberately NOT a real date for classification purposes; see
    _extract_tentative() below for pulling the estimate back out for
    display."""
    if v is None:
        return None
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.date() if isinstance(v, datetime.datetime) else v
    s = str(v).strip()
    if not s or s == '-' or 'tdc' in s.lower():
        return None
    m = re.match(r'^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$', s)
    if not m:
        return None
    day, mon, year_txt = int(m.group(1)), int(m.group(2)), m.group(3)
    year = int(year_txt) if len(year_txt) == 4 else (1900 + int(year_txt) if int(year_txt) >= 69 else 2000 + int(year_txt))
    try:
        return datetime.date(year, mon, day)
    except ValueError:
        return None


def _parse_month_year(v):
    """A 'month of planning' cell -> (year, month), or None. Handles a
    clean 'dd.mm.yyyy'/'dd.mm.yy' string (the live sync's own format --
    Apps Script normalizes the day to the 1st), a 'Mon-YY'/'MON-YY' text
    cell (seen on a couple of rows even in the live payload, e.g.
    'SEPT-26'), or an already-native date. The day component, whatever it
    is, is NEVER trusted as a real planning date -- only year/month matter
    here, matching how swr_machine_details.py already treats these
    month/year-only sheet dates."""
    if v is None:
        return None
    if isinstance(v, (datetime.date, datetime.datetime)):
        return (v.year, v.month)
    s = str(v).strip()
    if not s or s == '-':
        return None
    m = re.match(r'^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$', s)
    if m:
        mon, year_txt = int(m.group(2)), m.group(3)
        year = int(year_txt) if len(year_txt) == 4 else (1900 + int(year_txt) if int(year_txt) >= 69 else 2000 + int(year_txt))
        return (year, mon)
    m = re.match(r'^([A-Za-z]{3,4})[.\-/](\d{2})$', s)
    if m:
        mon = _MONTH_ABBR_CYP.get(m.group(1)[:3].lower())
        if mon:
            yy = int(m.group(2))
            return (1900 + yy if yy >= 69 else 2000 + yy, mon)
    return None


_TDC_RE = re.compile(r'tdc\s*[:\-]?\s*([\d./\-]{4,})', re.I)


def _extract_tentative(*texts):
    """First 'TDC: dd.mm.yy'-style tentative-completion estimate found
    across any of the given strings (checked in order), formatted as
    dd.mm.yyyy, or None. Pulled out purely for display (e.g. "expected
    around ...") -- never used for status classification, which only ever
    trusts a REAL completion date."""
    for t in texts:
        if not t:
            continue
        m = _TDC_RE.search(str(t))
        if m:
            d = _parse_any_date(m.group(1).strip())
            if d:
                return d.strftime('%d.%m.%Y')
    return None


def _classify_cyp_status(arrival_raw, completion_raw, month_raw, as_of):
    if _parse_any_date(completion_raw) is not None:
        return 'Completed'
    if _parse_any_date(arrival_raw) is not None:
        return 'Ongoing / Under Attention'
    my = _parse_month_year(month_raw)
    if my is None:
        return None  # genuinely nothing to classify from -- stay honest, not guessed
    return 'Overdue for Release' if my < (as_of.year, as_of.month) else 'Upcoming'


def extract_current_year_progress(payload, as_of=None):
    """-> [{Machine, Kind, Division, DivisionRaw, PlannedAt, MonthOfPlanning,
    DateOfArrival, DateOfCompletion, Tentative, TotalDays, StipulatedDays,
    Term, Remarks, Status}, ...] from payload.currentYearProgress.poh/.ioh
    -- the live sync's own day-by-day "who's been released, who's pending"
    picture for the current FY (see the module comment above this
    function for the full background and the Completed/Ongoing/Overdue/
    Upcoming rules). `as_of` defaults to today; tests pass a fixed date so
    the Overdue-vs-Upcoming cutoff is deterministic."""
    as_of = as_of or datetime.date.today()
    cyp = (payload or {}).get('currentYearProgress') or {}
    out = []
    for kind, key in (('POH', 'poh'), ('IOH', 'ioh')):
        rows = ((cyp.get(key) or {}) or {}).get('rows') or []
        for r in rows:
            machine = _g(r, 'MACHINE', 'Machine No.', 'Machine')
            if not machine:
                continue
            div_raw = _g(r, 'Belongs to DIV', 'Parent Division')
            div_clean = div_raw if div_raw in ('MYS', 'SBC', 'UBL') else None
            planned_at = _g(r, 'Attend at', 'Planned At')
            month_raw = _g(r, 'Actually planned', 'Month of Planning')
            arrival_raw = _g(r, 'FROM', 'Date of Arrival of machine to depot')
            completion_raw = _g(r, 'TO', 'Actual Date of Completion of IOH/POH')
            remarks_raw = _g(r, 'Remarks')
            status = _classify_cyp_status(arrival_raw, completion_raw, month_raw, as_of)
            my = _parse_month_year(month_raw)
            completion_dt = _parse_any_date(completion_raw)
            arrival_dt = _parse_any_date(arrival_raw)
            out.append(dict(
                Machine=str(machine).strip(), Kind=kind,
                Division=div_clean, DivisionRaw=ph.clean_text(div_raw),
                PlannedAt=ph.clean_text(planned_at),
                MonthOfPlanning=('%04d-%02d' % my) if my else None,
                MonthOfPlanningLabel=(datetime.date(my[0], my[1], 1).strftime('%b-%y') if my else None),
                DateOfArrival=arrival_dt.strftime('%d.%m.%Y') if arrival_dt else None,
                DateOfCompletion=completion_dt.strftime('%d.%m.%Y') if completion_dt else None,
                Tentative=_extract_tentative(completion_raw, remarks_raw),
                TotalDays=_g(r, 'Total days taken for completion'),
                StipulatedDays=_g(r, 'Stipulated days as per RDSO'),
                Term=_g(r, 'POH Term', 'IOH Term', 'Term'),
                Remarks=ph.clean_text(remarks_raw),
                Status=status,
            ))
    return out


def extract_ioh_poh(payload):
    """-> [record, ...], same shape run_pipeline.py's extract_ioh_poh(base)
    returns (SNo, Machine, Division, DivisionRaw, Commission, LastIOH_EH,
    LastIOH_Date, LastPOH_EH, LastPOH_Date, EH_Now, EH_Run_Month,
    NextDueIOH_Months, NextDuePOH_Months, Schedule, Remark)."""
    rows = ((payload or {}).get('forecast') or {}).get('rows') or []
    year_cols = set()
    for r in rows:
        for k in r.keys():
            if k and _YEAR_COL_RE.match(k):
                year_cols.add(k)

    out = []
    for r in rows:
        sno = _g(r, 'S. No.', 'S.No.', 'S. No', 'SNo')
        machine = _g(r, 'Name of Machines', 'Name of Machine')
        if not machine:
            continue
        div = _g(r, 'Divn.', 'Division')
        div_clean = div if div in ('MYS', 'SBC', 'UBL') else None
        commission = _g(r, 'Date of Commissioning')
        remark_raw = _g(r, 'Remark')
        remark = ph.clean_text(remark_raw) if isinstance(remark_raw, str) else remark_raw

        schedule = {}
        for col in year_cols:
            m = _YEAR_COL_RE.match(col)
            yr, kind = m.group(1), m.group(2).upper()
            schedule.setdefault(yr, [None, None])
            schedule[yr][0 if kind == 'IOH' else 1] = r.get(col)

        out.append(dict(
            SNo=sno, Machine=str(machine).strip(), Division=div_clean, DivisionRaw=div,
            Commission=_fmt_commission(commission),
            LastIOH_EH=_g(r, 'Last IOH Engine Hrs.'), LastIOH_Date=_fmt_date(_g(r, 'Last IOH Date')),
            LastPOH_EH=_g(r, 'Last POH Engine Hrs.'), LastPOH_Date=_fmt_date(_g(r, 'Last POH Date')),
            EH_Now=_g_prefix(r, 'engine hrs. as on'), EH_Run_Month=_g(r, 'Engine Hr Run/month'),
            NextDueIOH_Months=_g(r, 'NEXT DUE in Months IOH'), NextDuePOH_Months=_g(r, 'NEXT DUE in Months POH'),
            Schedule=schedule, Remark=remark,
        ))
    return out


def extract_ioh_poh_duration_norms(payload):
    """-> [record, ...], same shape run_pipeline.py's
    extract_ioh_poh_duration_norms(base) returns (type, iohDays,
    firstPohDays, subsequentPohDays, iohHrs, pohHrs)."""
    rows = ((payload or {}).get('durationNorms') or {}).get('rows') or []
    out = []
    for r in rows:
        mtype = _g(r, 'Machine Type')
        if not mtype:
            continue
        out.append(dict(
            type=str(mtype).strip(),
            iohDays=_g(r, 'Duration (in Days) IOH'),
            firstPohDays=_g(r, 'Duration (in Days) 1st POH'),
            subsequentPohDays=_g(r, 'Duration (in Days) Subsequent POHs'),
            iohHrs=_g(r, 'Duration (in Days) IOH Hrs'),
            pohHrs=_g(r, 'Duration (in Days) POH Hrs'),
        ))
    return out
