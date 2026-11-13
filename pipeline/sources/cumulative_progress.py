"""Adapter for the `cumulative_progress` table.

Live sync: cumulative_progress_sync.gs. One row, pushed daily, covering
BOTH financial-year tabs in one payload:
    {
      "years": {
        "2025-2026": {"title": "...", "notes": [...], "rows": [
            {"Sl. No": 1, "Machine": "SQRS 37&38", "DIV": "UBL",
             "Unit": "Km", "Target": 50.04, "Target / Month": 4.17,
             "Actual prog.": 35.63, "% of Prog.": 0.712,
             "Apr": 3.89, "May": 1.3, ..., "Mar": 3.65,
             "RowType": "machine"},
            {"Unit": "T/O", "Target": 21.48, ..., "RowType": "secondary",
             "SecondaryOf": "BCM-351"},
            {"Sl. No": "Cumulative", "RowType": "total", ...},
            ...
        ]},
        "2026-2027": {"title": "...", "rows": [... same shape, PLUS an
            "Annual Target" key on machine/secondary rows ...]}
      }
    }
Confirmed against the real live payload (not just the .gs source) that
both tabs carry exactly these keys - 'Target' is the ProportionateTarget
(target-to-date) column, distinct from 'Target / Month' and (current-year
only) 'Annual Target'; '% of Prog.' can come through as the literal
string "#DIV/0!" on a secondary row with a zero ProportionateTarget.

This replaces TWO local-file functions in run_pipeline.py:
  - extract_cum_progress(base, home_div, bare_type_div, norm) -> [record, ...]
  - extract_prior_year_by_type(base) -> (prior_by_type, prior_total,
        prior_by_machine, prior_by_type_secondary, prior_by_machine_secondary)
Both read the same live row here (fetch_payload() is called once), so
each extract_* function below takes that same payload dict rather than
doing its own REST call.

ONE DELIBERATE CARRYOVER LIMITATION (not introduced by this adapter -
already flagged by cumulative_progress_sync.gs's own header comment):
run_pipeline.py's MONTH_ABBR is hardcoded to ['Apr'..'Sep'], so both
functions below only sum/copy the April-September columns even though the
live payload already carries the full Apr-Mar year. This keeps this
adapter's output byte-for-byte comparable to the Cowork-side Artifact
(which still runs the unmodified local-file pipeline) for this rebuild;
extending past September is a follow-up, not part of this port.
"""
import re

import pipeline_helpers as ph
import supabase_client

TABLE = "cumulative_progress"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def _num_or_none(v):
    """Same contract as run_pipeline.py's module-level _num_or_none(): a
    real number passes through, anything else (blank, a hand-typed
    annotation like 'New M/c', or the literal '#DIV/0!' Excel error
    string that can appear in a secondary row's '% of Prog.' cell when
    its ProportionateTarget is 0) becomes None."""
    return v if isinstance(v, (int, float)) else None


def _current_and_prior_fy(years):
    """years is {"2025-2026": {...}, "2026-2027": {...}}. The current FY
    is always the lexicographically-largest key ("2026-2027" > "2025-2026"
    as plain strings, since both share the same "YYYY-YYYY" shape), the
    prior FY the smallest - equivalent to run_pipeline.py's hardcoded
    'FY 2026-2027' vs 'FY 2025-2026' folder paths, but derived from
    whatever two years the live sync actually has rather than a frozen
    pair of literals."""
    keys = sorted(years.keys())
    if not keys:
        return None, None
    return keys[-1], (keys[0] if len(keys) > 1 else None)


def _row_fields(r):
    out = {
        'Unit': ph.norm_unit(r.get('Unit')),
        'AnnualTarget': _num_or_none(r.get('Annual Target')),
        'TargetPerMonth': _num_or_none(r.get('Target / Month')),
        'ProportionateTarget': _num_or_none(r.get('Target')),
        'ActualProg': _num_or_none(r.get('Actual prog.')),
        'PctProg': _num_or_none(r.get('% of Prog.')),
    }
    for abbr in ph.MONTH_ABBR:
        out[abbr] = _num_or_none(r.get(abbr))
    return out


def extract_cum_progress(payload, home_div, bare_type_div, norm):
    """-> [record, ...], one per real machine, each carrying the same
    keys run_pipeline.py's extract_cum_progress(base, ...) produces
    (SlNo, Machine, MachineType, CurrentDIV, HomeDivision, Status, Unit,
    AnnualTarget, TargetPerMonth, ProportionateTarget, ActualProg,
    PctProg, Apr..Sep, and Secondary when a BCM-style continuation row
    exists for that machine). Row classification (machine/secondary/
    total, and which machine a secondary row belongs to) was already done
    server-side by cumulative_progress_sync.gs's RowType/SecondaryOf
    tagging, so this is simpler than the local-file version (which has to
    infer that from raw cell positions) - just split by RowType."""
    years = (payload or {}).get("years") or {}
    current_fy, _prior_fy = _current_and_prior_fy(years)
    if not current_fy:
        return []
    rows = years[current_fy].get("rows") or []

    secondary_by_machine = {
        r.get("SecondaryOf"): _row_fields(r)
        for r in rows
        if r.get("RowType") == "secondary" and r.get("SecondaryOf")
    }

    records = []
    for r in rows:
        if r.get("RowType") != "machine":
            continue
        machine = r.get("Machine")
        if not machine:
            continue
        machine = str(machine).strip()
        div_cur = r.get("DIV")
        fields = _row_fields(r)
        mtype_match = re.match(r'^([A-Za-z]+)', machine)
        mtype = mtype_match.group(1).upper() if mtype_match else 'UNKNOWN'
        hd = home_div.get(norm(machine)) or bare_type_div.get(mtype)
        if div_cur in ('MYS', 'SBC', 'UBL'):
            status = 'Active'
        elif div_cur in ('IOH', 'POH'):
            status = div_cur
        else:
            status = f'Deputed-{div_cur}' if div_cur else 'UNKNOWN'
        rec = {
            'SlNo': r.get('Sl. No'), 'Machine': machine, 'MachineType': mtype,
            'CurrentDIV': div_cur,
            'HomeDivision': hd if hd else (div_cur if div_cur in ('MYS', 'SBC', 'UBL') else 'UNKNOWN'),
            'Status': status,
        }
        rec.update(fields)
        sec = secondary_by_machine.get(machine)
        if sec:
            rec['Secondary'] = sec
        records.append(rec)
    return records


def extract_prior_year_by_type(payload):
    """-> (prior_by_type, prior_total, prior_by_machine,
    prior_by_type_secondary, prior_by_machine_secondary), matching
    run_pipeline.py's extract_prior_year_by_type(base) exactly - summing
    only the Apr-Sep columns per machine (the YoY comparison is always
    "same period, year over year"), grouped by *display* type
    (disp_type(mtype_of(name))), not the raw regex prefix used for
    MachineType in extract_cum_progress above - this is run_pipeline.py's
    own existing (and slightly inconsistent) behaviour, reproduced as-is
    rather than "fixed" here, since fixing it would make this adapter's
    output disagree with the Cowork-side Artifact's YoY numbers for the
    same date. Returns (None, 0.0, None, None, None) if there is no prior
    FY tab yet, matching the local-file version's own "not found" return
    shape so aggregate.py can check `prior_by_type is not None` exactly
    like run_pipeline.py's run() already does."""
    years = (payload or {}).get("years") or {}
    _current_fy, prior_fy = _current_and_prior_fy(years)
    if not prior_fy:
        return None, 0.0, None, None, None

    rows = years[prior_fy].get("rows") or []

    prior_by_type = {}
    prior_by_machine = {}
    prior_total = 0.0
    prior_by_type_secondary = {}
    prior_by_machine_secondary = {}

    def month_sum(r):
        return sum(v for v in (r.get(abbr) for abbr in ph.MONTH_ABBR) if isinstance(v, (int, float)))

    for r in rows:
        if r.get("RowType") == "machine":
            name = r.get("Machine")
            if not name:
                continue
            name = str(name).strip()
            s = month_sum(r)
            t = ph.disp_type(ph.mtype_of(name))
            prior_by_type[t] = prior_by_type.get(t, 0) + s
            prior_total += s
            key = ph.machine_key(name)
            prior_by_machine[key] = prior_by_machine.get(key, 0) + s
        elif r.get("RowType") == "secondary":
            sec_of = r.get("SecondaryOf")
            if not sec_of:
                continue
            s = month_sum(r)
            unit = ph.norm_unit(r.get("Unit"))
            t = f"{ph.disp_type(ph.mtype_of(sec_of))} ({unit})"
            prior_by_type_secondary[t] = prior_by_type_secondary.get(t, 0) + s
            key = f"{ph.machine_key(sec_of)}|{unit}"
            prior_by_machine_secondary[key] = prior_by_machine_secondary.get(key, 0) + s

    return prior_by_type, prior_total, prior_by_machine, prior_by_type_secondary, prior_by_machine_secondary
