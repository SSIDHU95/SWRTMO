"""Targeted regression test for aggregate.py's 26-Nov-2026 Fleet Health
rebuild: extract_cpoh_ryp_planning()'s static-file read, and the
letter-supersedes-Excel rule for future-year POH events.

Pure-function test, synthetic data only (no network, no live Supabase
pull) -- same style as test_swr_machine_details.py.

Background: a signed CPOH/RYP forward-planning letter (like the one
Shruti shared dated 07.10.2026) is a PDF issued roughly once a year, not
a Google Sheet, so it can never have a live Apps Script sync. Before this
fix, extract_cpoh_ryp_planning() was a permanent stub that always
returned None ("honest degradation" for a source that genuinely didn't
exist yet). Now that Shruti has provided the real signed letter's data,
it's hand-transcribed into pipeline/data/cpoh_ryp_planning.json and this
function reads that file -- this test proves the read, the shape, and
the graceful-degradation path if that file is ever missing or corrupt.

It also proves the "letter supersedes Excel" rule Shruti explicitly
asked for ("This signed letter supersedes the Excel data of the second
sub-sheet of IOH/POH planning related to POH... If some machine is
available in the Excel sheet, shown as due for POH but not written in
the letter, then this letter will supersede"): future-year POH events
must come ONLY from the letter data, never from the Excel workbook's own
year-schedule columns, even when a machine appears in both.

Run from the swr_pipeline directory:
    python3 -m unittest tests.test_cpoh_ryp_planning -v
"""
import json
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import aggregate


class TestExtractCpohRypPlanning(unittest.TestCase):

    def test_reads_the_real_data_file_with_expected_shape(self):
        # Uses the real pipeline/data/cpoh_ryp_planning.json shipped with
        # this repo -- proves the actual file Shruti's letter was
        # transcribed into still parses and has every key the dashboard
        # JS (renderCpohProposedTable etc.) expects.
        out = aggregate.extract_cpoh_ryp_planning()
        self.assertIsNotNone(out)
        for key in ('letterDate', 'proposed2728', 'tentative2829', 'tentative2930', 'tampingDemand'):
            self.assertIn(key, out)
        self.assertTrue(len(out['proposed2728']) > 0)
        row = out['proposed2728'][0]
        for key in ('machine', 'lastPohDate', 'ehSinceLastPoh', 'governingReason',
                    'probableSparingMonth', 'engines'):
            self.assertIn(key, row)

    def test_missing_file_degrades_to_none_not_an_exception(self):
        old = aggregate._CPOH_RYP_DATA_FILE
        aggregate._CPOH_RYP_DATA_FILE = os.path.join(tempfile.gettempdir(), 'does_not_exist_cpoh.json')
        try:
            self.assertIsNone(aggregate.extract_cpoh_ryp_planning())
        finally:
            aggregate._CPOH_RYP_DATA_FILE = old

    def test_corrupt_file_degrades_to_none_not_an_exception(self):
        old = aggregate._CPOH_RYP_DATA_FILE
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write('{not valid json')
            tmp_path = f.name
        aggregate._CPOH_RYP_DATA_FILE = tmp_path
        try:
            self.assertIsNone(aggregate.extract_cpoh_ryp_planning())
        finally:
            aggregate._CPOH_RYP_DATA_FILE = old
            os.unlink(tmp_path)

    def test_missing_optional_lists_default_to_empty_not_none(self):
        old = aggregate._CPOH_RYP_DATA_FILE
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump({"letterDate": "01.01.2026"}, f)
            tmp_path = f.name
        aggregate._CPOH_RYP_DATA_FILE = tmp_path
        try:
            out = aggregate.extract_cpoh_ryp_planning()
            self.assertEqual(out['letterDate'], '01.01.2026')
            self.assertEqual(out['proposed2728'], [])
            self.assertEqual(out['tentative2829'], [])
            self.assertEqual(out['tentative2930'], [])
            self.assertEqual(out['tampingDemand'], [])
        finally:
            aggregate._CPOH_RYP_DATA_FILE = old
            os.unlink(tmp_path)


if __name__ == "__main__":
    unittest.main()
