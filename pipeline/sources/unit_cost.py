"""Adapter for the `unit_cost` table.

Live sync: unit_cost_sync.gs, covering TWO separate Google Sheets (one per
financial year) in one push. Payload shape:
    {
      "years": {
        "2024-2025": {"machines": [{slNo, machine, codalLife, initialCost,
                                     rateOfInterest, crf, ageLabel,
                                     capitalRecoveryCost, staffExpenditure,
                                     overheadsExpenditure, fixedCostPerDay,
                                     consumables, spares, iohPohProportional,
                                     daysWorked, variableCostPerDay,
                                     annualOutput, variableCostPerUnit,
                                     outputPerDay, variableCostTotalPerDay,
                                     dailyCost, unitCost}, ...]},
        "2025-2026": {"machines": [...]}
      }
    }
The .gs script's own header comment is explicit that it is raw extraction
only - "matching each machine against the fleet list ... stays in the
Python pipeline, not duplicated here" - so this adapter, like
extract_unit_cost_year()/extract_unit_cost() in run_pipeline.py, only
reshapes the raw per-machine cost columns and computes machineKey; the
actual cross-reference against progress.machines (to attach unit, display
type, division, status) happens later in aggregate.py, same as it does in
run_pipeline.py's run().
"""
import pipeline_helpers as ph
import supabase_client

TABLE = "unit_cost"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def _num(v):
    return v if isinstance(v, (int, float)) else None


def extract_unit_cost_year(year_block):
    """One FY's {"machines": [...]} block -> the list of per-machine dicts
    run_pipeline.py's extract_unit_cost_year(path, fy_label) returns.
    Row filtering (skipping "Sub total/Average <TYPE>" rows, the
    "1,2,3...,22=21/19" column-label row, and footnotes) already happened
    on the Apps Script side in parseUnitCostSheet() - every row here is
    already a real machine - so this only needs to reshape field names
    and compute machineKey, exactly as the local-file version does."""
    out = []
    for r in (year_block or {}).get("machines") or []:
        machine = r.get("machine")
        if not machine or not str(machine).strip():
            continue
        machine_name = str(machine).strip()
        out.append(dict(
            machine=machine_name,
            machineKey=ph.machine_key_aliased(ph.norm_machine(machine_name)),
            codalLife=_num(r.get("codalLife")),
            initialCost=_num(r.get("initialCost")),
            ageLabel=(str(r.get("ageLabel")).strip() if r.get("ageLabel") is not None else None),
            capitalRecoveryCost=_num(r.get("capitalRecoveryCost")),
            staffExpenditure=_num(r.get("staffExpenditure")),
            overheadsExpenditure=_num(r.get("overheadsExpenditure")),
            fixedCostPerDay=_num(r.get("fixedCostPerDay")),
            consumables=_num(r.get("consumables")),
            spares=_num(r.get("spares")),
            iohPohProportional=_num(r.get("iohPohProportional")),
            daysWorked=_num(r.get("daysWorked")),
            variableCostPerDay=_num(r.get("variableCostPerDay")),
            annualOutput=_num(r.get("annualOutput")),
            variableCostPerUnit=_num(r.get("variableCostPerUnit")),
            outputPerDay=_num(r.get("outputPerDay")),
            dailyCost=_num(r.get("dailyCost")),
            unitCost=_num(r.get("unitCost")),
        ))
    return out


def extract_unit_cost(payload):
    """-> {years: {fyLabel: {sourceFile, machines}}, latestFy}. Mirrors
    run_pipeline.py's extract_unit_cost(base), minus the filesystem glob
    (every FY this sync knows about is already a key in payload["years"]).
    `sourceFile` has no literal equivalent on the live-sync path - it is
    kept as a key (rather than dropped) only because the dashboard's
    Unit Cost tab type may read it; it is filled with a label naming the
    live sync rather than a filename."""
    years_in = (payload or {}).get("years") or {}
    years_out = {}
    for fy_label, block in years_in.items():
        recs = extract_unit_cost_year(block)
        if recs:
            years_out[fy_label] = dict(sourceFile=f"live sync ({fy_label})", machines=recs)
    if not years_out:
        return {}
    latest_fy = sorted(years_out.keys())[-1]
    return dict(years=years_out, latestFy=latest_fy)
