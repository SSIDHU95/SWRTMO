"""
Tests for the daily_progress_srden live sync -- Sr.DEN (jurisdiction) ONLY
from the Daily Progress Google Sheet (Shruti's explicitly chosen narrower
scope, 14-Nov-2026; the full shortfall-category log from this same sheet
remains Task #53, separate and still open).

Two layers:

  1. TestForwardFillParsing - a pure-Python mirror of
     daily_progress_srden_sync.gs's parseTab() (the Apps Script itself is
     the source of truth that actually runs in Shruti's Google account;
     this mirror exists only so the forward-fill/Konkan-exclusion/header-
     skip logic can be exercised here, under pytest, without a JS
     runtime), run against tests/fixtures/daily_progress_srden_raw_rows.json
     -- the REAL raw rows of the 12.11.2026 tab of the uploaded
     Daily_Progress.xlsx (dumped once via openpyxl, not hand-typed), so
     every real-world quirk (inconsistent "OL Sr.DEN.W." vs "OL Sr. DEN W"
     spacing/punctuation, multi-row continuation blocks, the Konkan
     Railway section's different title format with no Sr.DEN banner that
     fits the SWR model) is exercised against real data, not an idealized
     fake.

  2. TestExtractSrden - the Python adapter's extract_srden(), fed a
     payload built from that same real-data parse, proving the
     machine_key_aliased lookup dict + canon_srden() canonicalization
     come out correct for real machine names/Sr.DEN text.

Run from the swr_pipeline directory:
    python3 -m unittest tests.test_daily_progress_srden -v
"""
import json
import os
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pipeline_helpers as ph
from sources import daily_progress_srden

FIXTURE_PATH = os.path.join(os.path.dirname(__file__), "fixtures", "daily_progress_srden_raw_rows.json")


# =====================================================================
# Pure-Python mirror of daily_progress_srden_sync.gs's parseTab(). Kept
# byte-for-byte equivalent in logic to the real .gs file -- see that
# file's own comments for the reasoning behind each step (the Konkan
# title regex NOT requiring a trailing " DIV" suffix, the header_passed
# numeric-Sl.No gate, forward-fill only overwriting on a non-blank
# banner cell).
# =====================================================================
_DIV_TITLE_RE = re.compile(r'PERFORMANCE OF TRACK MACHINES IN\s+(.+)', re.I)


def parse_tab_mirror(rows_vals):
    current_div = None
    skip_section = False
    current_machine = None
    current_srden = None
    header_passed = False
    out_rows = []
    seen = set()

    for row in rows_vals:
        colA_raw = row[0] if len(row) > 0 else None
        colA = str(colA_raw).strip() if colA_raw is not None else ""
        colB = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        colC = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""

        title_text = colA or colB or colC
        mTitle = _DIV_TITLE_RE.search(title_text)
        if mTitle:
            raw_div = mTitle.group(1).strip()
            skip_section = bool(re.search(r'konkan', raw_div, re.I))
            div_name = re.sub(r'\s+DIV\s*$', '', raw_div, flags=re.I).strip()
            current_div = None if skip_section else div_name.upper()
            current_machine = None
            current_srden = None
            header_passed = False
            continue

        if colB.lower() == 'note' or colA.lower() == 'note':
            break

        if skip_section or not current_div:
            continue

        if not header_passed:
            if isinstance(colA_raw, (int, float)):
                header_passed = True
            else:
                continue

        if colC:
            current_machine = colC
        if not current_machine:
            continue
        if colB:
            current_srden = colB

        key = current_machine.upper()
        if key not in seen:
            seen.add(key)
            out_rows.append(dict(machine=current_machine, srDen=current_srden, division=current_div))

    return out_rows


class TestForwardFillParsing(unittest.TestCase):
    """Exercises parse_tab_mirror() against the REAL raw rows of the
    12.11.2026 tab (4 division sections: UBL, SBC, MYS, Konkan Railway --
    confirmed by direct openpyxl inspection before this suite was
    written)."""

    @classmethod
    def setUpClass(cls):
        if not os.path.exists(FIXTURE_PATH):
            raise unittest.SkipTest(f"missing fixture {FIXTURE_PATH}")
        with open(FIXTURE_PATH) as f:
            fixture = json.load(f)
        cls.rows_vals = fixture["rows"]
        cls.parsed = parse_tab_mirror(cls.rows_vals)
        cls.by_machine = {r["machine"]: r for r in cls.parsed}

    def test_no_header_rows_leak_through_as_fake_machines(self):
        """The literal header-row text ("Machine", "M/c worked in OL/CNS")
        must never appear as a machine name -- this was a real bug caught
        and fixed while building this sync (the header block under each
        division title has no numeric Sl.No, so it must be gated out, not
        just skipped by non-blank-Machine-cell alone)."""
        self.assertNotIn("Machine", self.by_machine)
        self.assertNotIn("M/c worked in OL/CNS", self.by_machine)

    def test_konkan_railway_excluded_entirely(self):
        """BCM 384, the Konkan Railway section's only machine, must not
        appear at all -- it is a different railway's machine, confirmed by
        direct inspection (its own jurisdiction banner reads "Sr. DEN RN",
        which doesn't fit the SWR Sr.DEN West/East/Central/North/South/HQ
        model at all). This also exercises the title-regex fix: the
        Konkan title has no trailing " DIV" the way UBL/SBC/MYS's titles
        do, so a naive regex requiring that suffix would wrongly keep
        Konkan folded into the preceding MYS section instead of excluding
        it (the bug this suite's own authoring caught and fixed)."""
        self.assertNotIn("BCM 384", self.by_machine)
        for r in self.parsed:
            self.assertNotEqual(r["division"], "KONKAN RAILWAY")

    def test_three_real_swr_divisions_present(self):
        divisions = set(r["division"] for r in self.parsed)
        self.assertEqual(divisions, {"UBL", "SBC", "MYS"})

    def test_forward_fill_carries_across_blank_rows_within_a_block(self):
        """MDU 848, MDU 849, MDU 850, MDU 57220, MDU 57222 all sit between
        two banner-bearing rows (BCM 712's "OL Sr.DEN. E." and the next
        restated "OL Sr.DEN. E." banner further down) with a BLANK Sr.DEN
        cell of their own in the real sheet -- confirmed by direct
        inspection. Every one of them must inherit "OL Sr.DEN. E." via
        forward-fill, not come back None."""
        for machine in ("BCM 712", "MDU 848", "MDU 849", "MDU 850", "MDU 57220", "MDU 57222"):
            self.assertIn(machine, self.by_machine, f"{machine} missing from parse")
            self.assertEqual(self.by_machine[machine]["srDen"], "OL Sr.DEN. E.",
                              f"{machine} did not inherit the forward-filled Sr.DEN banner")

    def test_forward_fill_resets_at_next_banner(self):
        """RGM 96 -07 carries a NEW banner ("OL Sr.DEN.C.") immediately
        after the "OL Sr.DEN. E." run above -- the forward-fill must
        switch to the new value, not keep carrying the old one forward."""
        self.assertEqual(self.by_machine["RGM 96 -07"]["srDen"], "OL Sr.DEN.C.")
        self.assertEqual(self.by_machine["DUO 57434"]["srDen"], "OL Sr.DEN.C.")

    def test_forward_fill_resets_at_next_division_block_not_carried_over(self):
        """UNI 8277 (the last real banner-holder of the UBL block, "OL Sr.
        DEN. HQ.") must NOT leak its jurisdiction into SBC's first real
        machine, SQRS 7&8 -- a brand-new division section must start with
        no inherited banner from the previous one."""
        self.assertEqual(self.by_machine["UNI 8277"]["srDen"], "OL Sr.DEN. HQ.")
        self.assertIn("SQRS 7&8", self.by_machine)
        self.assertNotEqual(self.by_machine["SQRS 7&8"]["srDen"], "OL Sr.DEN. HQ.")

    def test_forward_fill_carries_to_end_of_block(self):
        """BCM 400, BCM 351 and DUO 8090 (the last three machines of the
        UBL block, all blank-banner continuation rows) must all inherit
        UNI 8277's "OL Sr.DEN. HQ." banner all the way to the end of the
        block -- not just the row or two immediately after it."""
        for machine in ("BCM 400", "BCM 351", "DUO 8090"):
            self.assertEqual(self.by_machine[machine]["srDen"], "OL Sr.DEN. HQ.",
                              f"{machine} should inherit HQ through end of UBL block")

    def test_every_machine_has_a_division(self):
        for r in self.parsed:
            self.assertIn(r["division"], ("UBL", "SBC", "MYS"))

    def test_realistic_machine_count(self):
        # Confirmed directly against the real file: 81 real SWR machines
        # on the 12.11.2026 tab (40 UBL + 20 SBC + 21 MYS), Konkan excluded.
        self.assertEqual(len(self.parsed), 81)


class TestExtractSrden(unittest.TestCase):
    """The Python adapter's extract_srden(), fed a payload shaped exactly
    like what daily_progress_srden_sync.gs would push (built from the same
    real-data parse above, not hand-typed)."""

    @classmethod
    def setUpClass(cls):
        if not os.path.exists(FIXTURE_PATH):
            raise unittest.SkipTest(f"missing fixture {FIXTURE_PATH}")
        with open(FIXTURE_PATH) as f:
            fixture = json.load(f)
        parsed = parse_tab_mirror(fixture["rows"])
        cls.payload = dict(asOfDate=fixture["tabName"], machines=parsed)
        cls.result = daily_progress_srden.extract_srden(cls.payload)

    def test_as_of_date_passed_through(self):
        self.assertEqual(self.result["asOfDate"], "12.11.2026")

    def test_canonical_keys_shaped_correctly(self):
        """Every value in the lookup dict must be the "<DIV> Sr.DEN/<code>"
        canonical shape the rest of the pipeline (progress.machines[].srden,
        machinesForSrden(), the dashboard's own Sr.DEN jump buttons) already
        expects -- not the sheet's raw, inconsistently-punctuated banner
        text verbatim."""
        srden_map = self.result["srden"]
        self.assertTrue(srden_map, "expected at least one real machine to canonicalize")
        for key, val in srden_map.items():
            self.assertRegex(val, r'^(UBL|SBC|MYS) Sr\.DEN/(N|S|E|W|C|HQ)$',
                              f"{key} -> {val} is not a canonical Sr.DEN key")

    def test_known_machine_canonicalizes_correctly(self):
        """BCM 712 (UBL, raw banner "OL Sr.DEN. E.") must canonicalize to
        'UBL Sr.DEN/E', keyed by machine_key_aliased(norm_machine(...)) the
        same way every other adapter in this pipeline keys machines."""
        key = ph.machine_key_aliased(ph.norm_machine("BCM 712"))
        self.assertEqual(self.result["srden"].get(key), "UBL Sr.DEN/E")

    def test_hq_jurisdiction_canonicalizes_correctly(self):
        key = ph.machine_key_aliased(ph.norm_machine("UNI 8277"))
        self.assertEqual(self.result["srden"].get(key), "UBL Sr.DEN/HQ")

    def test_konkan_machine_never_appears(self):
        key = ph.machine_key_aliased(ph.norm_machine("BCM 384"))
        self.assertNotIn(key, self.result["srden"])

    def test_missing_payload_degrades_to_empty(self):
        result = daily_progress_srden.extract_srden(None)
        self.assertEqual(result["srden"], {})
        self.assertIsNone(result["asOfDate"])

    def test_machine_with_no_srden_that_day_is_simply_omitted(self):
        """A machine with no Sr.DEN banner at all (None) must not produce
        a dict entry (canon_srden() returns None for a falsy raw value) --
        confirmed separately against other real date-tabs (07/06/05-Nov-
        2026) where the very first machine of a block has no banner
        stated; here it's exercised directly via a synthetic gap, since
        the 12.11.2026 fixture tab happens to have none."""
        payload = dict(asOfDate="x", machines=[dict(machine="TEST 999", srDen=None, division="UBL")])
        result = daily_progress_srden.extract_srden(payload)
        self.assertEqual(result["srden"], {})


if __name__ == "__main__":
    unittest.main()
