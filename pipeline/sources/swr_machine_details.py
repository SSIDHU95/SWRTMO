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

This replaces TWO local-file functions in run_pipeline.py:
  - extract_home_div(base)   -> (home_div, bare_type_div, norm)
  - extract_age_profile(base) -> [{MachineType, Machine, CommissioningDate,
                                    AgeYears}, ...]
Both read the same live row here, so fetch_payload() is called once and
the two extract_* functions below each take that same payload dict,
rather than each doing their own REST call.
"""
import re

import pipeline_helpers as ph
import supabase_client

TABLE = "swr_machine_details"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


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


def extract_age_profile(payload):
    """-> [{MachineType, MachineTypeFullName, Machine, CommissioningDate,
    AgeYears}, ...]. Mirrors run_pipeline.py's extract_age_profile(base).
    The live sync's parseAgeProfile() already skipped the machine-type
    banner rows (e.g. "BCM  (7 machines)") on the Apps Script side, so
    every row here is already a real machine - only AgeYears needs the
    same free-text -> int extraction run_pipeline.py does ("16 years" ->
    16).

    Fixed 18-Nov-2026 per Shruti ("the machine details are utilized
    elsewhere also, in age profile and all -- you connect it there
    also"): MachineType is now disp_type()'d (so a sheet that still
    types a machine's raw alias prefix - FRM/PBR/SQRS/T/WST - reports its
    real canonical type here too, consistent with every other tab), and
    MachineTypeFullName is attached from the same pipeline_helpers.
    MTYPE_FULL_NAME table the rest of the dashboard now uses. Neither
    field is read by aggregate.py's own age-bucket rollup today (that
    only needs AgeYears/Machine/CommissioningDate - see its own fix for
    the actual age/commissioning-date JOIN, which is where the real
    FRM/SBCM and PBR/BRM mismatch was silently dropping machines), but
    both are carried through so a future consumer never has to re-derive
    them from a raw, possibly-aliased type string."""
    rows = (payload or {}).get("ageProfile") or []
    out = []
    for r in rows:
        machine = r.get("machine")
        if not machine:
            continue
        age_txt = r.get("age")
        agem = re.search(r'(\d+)', str(age_txt)) if age_txt else None
        age = int(agem.group(1)) if agem else None
        raw_type = r.get("machineType")
        disp_mtype = ph.disp_type(str(raw_type).strip().upper()) if raw_type else None
        out.append(dict(
            MachineType=disp_mtype,
            MachineTypeFullName=ph.MTYPE_FULL_NAME.get(disp_mtype) if disp_mtype else None,
            Machine=str(machine).strip(),
            CommissioningDate=r.get("commissioningDate"),
            AgeYears=age,
        ))
    return out
