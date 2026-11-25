"""Targeted regression test for sources/exception_reports.py's 25-Nov-2026
DaysUnder fallback.

Pure-function test, synthetic data only (no network, no fixtures
required) -- same style as test_daily_positions.py.

Background: Shruti uploaded the live Exception Sheet after noticing some
machines (e.g. "SQRS 51 &52") show no day-count ("-- d") on the dashboard
even though the sheet does give a date for when the machine went under.
Direct inspection of that sheet confirmed the "No of Days under repair/
IOH/ POH till date" column is filled in BY HAND and is blank on most rows
(24-Nov-2026: 9 of 34) -- extract_exceptions() used to take that column's
value completely verbatim and leave DaysUnder as None whenever it was
blank, even for a row that DOES have an "Under Repair/IOH/POH/ planned
from date" to count from. Fixed by falling back to
(as-of date - UnderSince date), using the sheet's own +1 (inclusive)
counting -- verified against the 9 rows that already had both columns
filled in by hand on that sheet (7 match this exactly; the one mismatch
is a single-day manual entry slip, not a different convention). A row
with no UnderSince at all ("SQRS 37 & 38") still has nothing to compute
from and correctly stays absent rather than guessed.

Run from the swr_pipeline directory:
    python3 -m unittest tests.test_exception_reports -v
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sources import exception_reports


class TestExtractExceptionsDaysUnderFallback(unittest.TestCase):

    def _rows(self, rows, as_of="24.11.2026"):
        payload = {"asOfDate": as_of, "reportTitle": "t", "rows": rows}
        out = exception_reports.extract_exceptions([(payload, "2026-11-24T12:00:00Z")])
        return out["24-Nov"]

    def test_computes_days_under_from_under_since_when_sheet_column_blank(self):
        # The exact real row from Shruti's uploaded sheet: UnderSince
        # "23.11.2026" given, "No of Days" column left blank by whoever
        # filled the sheet in.
        rows = self._rows([
            {"category": "c", "machine": "SQRS 51 &52", "division": "UBL",
             "location": "At SED", "remarks": "Depot work",
             "underSince": "23.11.2026", "tentative": None, "daysUnder": None},
        ])
        self.assertEqual(rows[0]["DaysUnder"], 2)  # 24.11 - 23.11, inclusive (+1)

    def test_leaves_days_under_absent_when_there_is_no_under_since_either(self):
        # The other real row from the same sheet: neither column filled in
        # at all -- nothing to compute from, must stay honestly absent
        # rather than guessed at.
        rows = self._rows([
            {"category": "c", "machine": "SQRS 37 & 38", "division": "UBL",
             "location": "At TGT", "remarks": "AT under prepartion & site not ready",
             "underSince": None, "tentative": None, "daysUnder": None},
        ])
        self.assertIsNone(rows[0]["DaysUnder"])

    def test_sheets_own_days_under_value_is_never_overridden_when_present(self):
        # A row where the sheet DOES give its own day-count must keep that
        # exact value, not get silently replaced by the computed one (even
        # if they happen to differ, e.g. the one hand-entry slip found on
        # the live sheet).
        rows = self._rows([
            {"category": "c", "machine": "MDU 845", "division": "MYS",
             "location": "At JRU", "remarks": "x",
             "underSince": "11.11.2026", "tentative": None, "daysUnder": 14.0},
        ])
        self.assertEqual(rows[0]["DaysUnder"], 14.0)

    def test_computed_fallback_ignores_an_under_since_date_after_the_as_of_date(self):
        # Defensive: a bad/garbled date later than the report's own as-of
        # date must not produce a nonsensical negative day count.
        rows = self._rows([
            {"category": "c", "machine": "XYZ 1", "division": "UBL",
             "location": "At ABC", "remarks": "r",
             "underSince": "30.11.2026", "tentative": None, "daysUnder": None},
        ])
        self.assertIsNone(rows[0]["DaysUnder"])


if __name__ == "__main__":
    unittest.main()
