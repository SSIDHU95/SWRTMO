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
