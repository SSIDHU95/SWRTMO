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

import supabase_client

TABLE = "swr_machine_details"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


# ---------------------------------------------------------------------------
# norm() - copied verbatim from inside run_pipeline.py's extract_home_div().
# It is local to that function there (not one of the module-level helpers
# lifted into pipeline_helpers.py), so it is reproduced here instead of
# imported - this is the ONE place in the whole adapter package it is
# needed, since it is specific to matching Original-Allocation machine
# names, not a general-purpose machine-name normalizer (norm_machine() in
# pipeline_helpers.py is the general one, used everywhere else).
# ---------------------------------------------------------------------------
def norm_alloc_name(name):
    if name is None:
        return None
    s = str(name).upper()
    s = re.sub(r'\(.*?\)', '', s)
    return re.sub(r'[^A-Z0-9]', '', s)


def extract_home_div(payload):
    """-> (home_div, bare_type_div, norm). Mirrors run_pipeline.py's
    extract_home_div(base) exactly, except the source is already three
    plain machine-name lists (one per division) instead of a raw sheet
    grid - the live sync's parseOriginalAllocation() already dropped the
    "-" placeholders and the TOTAL/GRAND TOTAL summary rows on the Apps
    Script side, so no further filtering is needed here."""
    alloc = (payload or {}).get("originalAllocation") or {}
    home_div = {}
    bare_type_div = {}
    for div, key in (("MYS", "mys"), ("SBC", "sbc"), ("UBL", "ubl")):
        for val in alloc.get(key) or []:
            val = str(val).strip()
            if not val or val == "-":
                continue
            home_div[norm_alloc_name(val)] = div
            if re.match(r'^[A-Za-z]+$', val):
                bare_type_div[val.upper()] = div
    return home_div, bare_type_div, norm_alloc_name


def extract_age_profile(payload):
    """-> [{MachineType, Machine, CommissioningDate, AgeYears}, ...].
    Mirrors run_pipeline.py's extract_age_profile(base). The live sync's
    parseAgeProfile() already skipped the machine-type banner rows (e.g.
    "BCM  (7 machines)") on the Apps Script side, so every row here is
    already a real machine - only AgeYears needs the same free-text ->
    int extraction run_pipeline.py does ("16 years" -> 16)."""
    rows = (payload or {}).get("ageProfile") or []
    out = []
    for r in rows:
        machine = r.get("machine")
        if not machine:
            continue
        age_txt = r.get("age")
        agem = re.search(r'(\d+)', str(age_txt)) if age_txt else None
        age = int(agem.group(1)) if agem else None
        out.append(dict(
            MachineType=r.get("machineType"),
            Machine=str(machine).strip(),
            CommissioningDate=r.get("commissioningDate"),
            AgeYears=age,
        ))
    return out
