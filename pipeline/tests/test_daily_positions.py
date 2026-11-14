"""Targeted regression test for sources/daily_positions.py's 14-Nov-2026
key fix.

Pure-function test, synthetic data only (no network, no fixtures
required) -- same style as test_network_map.py.

Background: extract_positions() used to key its output dict on
pipeline_helpers.machine_key_aliased(norm_machine(name)) -- the
TYPE+NUMBER key the rest of this pipeline uses to join a machine against
progress/exception data. That key strips parenthetical suffixes, which
is fine for a normal machine but silently collapsed two genuinely
distinct real machines in the live Daily Progress sheet ("T-28
(403&404)" and "T-28(908 A&B)", both -> "T28") into a single dict entry
-- one of the two would overwrite the other, and the Network Map's own
machine count would be one short every single run, with no error or
warning anywhere. Neither of extract_positions()'s two real callers
(aggregate.py's rawSources.dailyPositions, network_map.py's live map
build) actually needs the type-keyed join this was originally built
for -- both just iterate `.values()` -- so the fix keys on the plain
normalized machine name instead, which cannot collapse two differently
-labelled machines together.

Run from the swr_pipeline directory:
    python3 -m unittest tests.test_daily_positions -v
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sources import daily_positions


class TestExtractPositionsKeying(unittest.TestCase):

    def test_two_distinct_machines_sharing_a_type_number_are_both_kept(self):
        # The exact real-world pair that was colliding before the fix --
        # both would machine_key_aliased() to "T28", so one used to
        # silently overwrite the other.
        payload = {
            "asOfDate": "13.11.2026",
            "machines": [
                {"machine": "T-28 (403&404)", "div": "UBL", "stabledAt": "AAA"},
                {"machine": "T-28(908 A&B)", "div": "SBC", "stabledAt": "BBB"},
            ],
        }
        result = daily_positions.extract_positions(payload)
        self.assertEqual(len(result['positions']), 2)
        machines = {p['machine'] for p in result['positions'].values()}
        self.assertEqual(machines, {"T-28 (403&404)", "T-28(908 A&B)"})

    def test_realistic_87_machine_payload_keeps_every_distinct_name(self):
        # Mirrors the real shape confirmed live on 13-Nov-2026: 87 real
        # machine rows (after the header-leakage row is excluded),
        # including the one colliding T-28 pair -- must now yield 87
        # distinct entries, not 86.
        names = [f"BCM {i}" for i in range(50)] + [f"MPT {i}" for i in range(35)]
        names += ["T-28 (403&404)", "T-28(908 A&B)"]
        payload = {"machines": [{"machine": n, "div": "UBL", "stabledAt": "X"} for n in names]}
        result = daily_positions.extract_positions(payload)
        self.assertEqual(len(result['positions']), 87)

    def test_header_leakage_row_still_passes_through_unharmed(self):
        # extract_positions() itself has never filtered the "Machine"
        # header-leakage row (that filtering happens one layer up, in
        # network_map.py's _is_header_leakage()) -- confirm the key
        # change didn't accidentally start/stop that here.
        payload = {"machines": [{"machine": "Machine", "div": None, "stabledAt": "STABLED AT"}]}
        result = daily_positions.extract_positions(payload)
        self.assertEqual(len(result['positions']), 1)

    def test_detail_fields_are_passed_through_when_present(self):
        payload = {"machines": [{
            "machine": "DUO 8154", "div": "UBL", "stabledAt": "TNGL",
            "blockSection": "TNGL-DAJ Dn M/L", "majorSection": "HPT-BAY",
            "durationMinutes": 105.0, "actualTotal": 0.7,
            "remarks": "Night: 700 mtrs tamped.",
        }]}
        result = daily_positions.extract_positions(payload)
        entry = list(result['positions'].values())[0]
        self.assertEqual(entry['blockSection'], "TNGL-DAJ Dn M/L")
        self.assertEqual(entry['durationMinutes'], 105.0)
        self.assertEqual(entry['actualTotal'], 0.7)
        self.assertEqual(entry['remarks'], "Night: 700 mtrs tamped.")

    def test_missing_detail_fields_are_simply_absent_not_fabricated(self):
        # A machine from a section where daily_progress_sync.gs could not
        # locate the Stabled At column has no detail fields pushed at
        # all -- extract_positions() must not invent nulls for them.
        payload = {"machines": [{"machine": "BCM 400", "div": "MYS", "stabledAt": "XYZ"}]}
        result = daily_positions.extract_positions(payload)
        entry = list(result['positions'].values())[0]
        for field in daily_positions.DETAIL_FIELDS:
            self.assertNotIn(field, entry)

    def test_whitespace_and_dash_variants_of_the_same_name_still_collapse(self):
        # norm_machine() itself still normalizes whitespace runs and
        # " - "/" -"/"- " dash spacing -- this is deliberate (the same
        # real machine typed slightly differently on different days
        # should still be one entry), only the TYPE+NUMBER collapsing is
        # gone.
        payload = {"machines": [
            {"machine": "BRM - 04", "div": "SBC", "stabledAt": "X"},
            {"machine": "BRM  -  04", "div": "SBC", "stabledAt": "Y"},
        ]}
        result = daily_positions.extract_positions(payload)
        self.assertEqual(len(result['positions']), 1)


if __name__ == '__main__':
    unittest.main()
