"""Targeted regression test for sources/swr_machine_details.py's 25-Nov-2026
fix to extract_age_profile().

Pure-function test, synthetic data only (no network) -- same style as
test_exception_reports.py.

Background: Shruti's live Fleet Age Distribution chart showed ALL 99
machines bucketed into "0-10 yrs" and nothing in any other bucket. A
direct live Supabase pull of the `swr_machine_details` table's real
payload confirmed swr_machine_details_sync.gs currently pushes the
'commissioningDate'/'age' fields SWAPPED: the real commissioning date (as
a JS Date.toString() dump) lands in 'age', while 'commissioningDate' is a
literal '-' for every row. The old extract_age_profile() took 'age' as a
precomputed "N years" string and grabbed the first run of digits out of it
with a bare regex -- since nearly every commissioning date is month/
year-only, Sheets/Apps Script normalizes it to the 1st of the month, so
that regex always found "01" and reported age=1 for virtually everyone.

Fixed by trying to parse an actual date out of EITHER field
(commissioningDate first, then age) and computing AgeYears ourselves from
whichever one parses, so this keeps working correctly whether or not, or
whenever, the Apps Script swap itself gets fixed.

Run from the swr_pipeline directory:
    python3 -m unittest tests.test_swr_machine_details -v
"""
import datetime
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sources import swr_machine_details as smd

TODAY = datetime.date(2026, 11, 25)


class TestExtractAgeProfileDateFix(unittest.TestCase):

    def _rows(self, rows, today=TODAY):
        return smd.extract_age_profile({"ageProfile": rows}, today=today)

    def test_live_swapped_fields_js_date_tostring_in_age(self):
        # The exact real row confirmed live on the swr_machine_details
        # table: commissioningDate is a placeholder '-', the real
        # commissioning date (1-Apr-2005) is sitting in 'age' as a raw JS
        # Date.toString() dump.
        out = self._rows([
            {"slNo": 1, "machine": "BCM-351", "machineType": "Non-Tamping",
             "commissioningDate": "-",
             "age": "Fri Apr 01 2005 13:30:00 GMT+0530 (India Standard Time)"},
        ])
        self.assertEqual(out[0]["AgeYears"], 21)
        self.assertEqual(out[0]["CommissioningDate"], "01.04.2005")

    def test_well_formed_row_commissioning_date_dd_mm_yyyy(self):
        # The ORIGINAL documented shape (not currently live, but must keep
        # working once/if the Apps Script swap is fixed upstream).
        out = self._rows([
            {"slNo": 2, "machine": "BCM-02", "machineType": "BCM",
             "commissioningDate": "12.03.2010", "age": "16 years"},
        ])
        self.assertEqual(out[0]["AgeYears"], 16)
        self.assertEqual(out[0]["CommissioningDate"], "12.03.2010")

    def test_mon_yy_commissioning_date(self):
        # Same "Mon-YY" shape confirmed in the newly uploaded IOH/POH
        # Planning sheet's own "Date of Commissioning" column.
        out = self._rows([
            {"slNo": 3, "machine": "DTE-05", "machineType": "DTE",
             "commissioningDate": "Aug-93", "age": None},
        ])
        self.assertEqual(out[0]["AgeYears"], 33)
        self.assertEqual(out[0]["CommissioningDate"], "01.08.1993")

    def test_two_digit_year_dmy_with_pivot(self):
        out = self._rows([
            {"slNo": 4, "machine": "MPT-09", "machineType": "MPT",
             "commissioningDate": "24.12.18", "age": None},
        ])
        self.assertEqual(out[0]["AgeYears"], 7)
        self.assertEqual(out[0]["CommissioningDate"], "24.12.2018")

    def test_falls_back_to_old_regex_reading_when_neither_field_is_a_date(self):
        # Neither column parses as a date at all (both placeholders/free
        # text) -- must still fall back to the original "16 years" -> 16
        # behaviour rather than going straight to None.
        out = self._rows([
            {"slNo": 5, "machine": "BCM-77", "machineType": "BCM",
             "commissioningDate": "-", "age": "16 years"},
        ])
        self.assertEqual(out[0]["AgeYears"], 16)
        self.assertIsNone(out[0]["CommissioningDate"])

    def test_future_or_garbled_date_does_not_produce_a_negative_age(self):
        out = self._rows([
            {"slNo": 6, "machine": "BCM-88", "machineType": "BCM",
             "commissioningDate": "01.01.2099", "age": None},
        ])
        self.assertIsNone(out[0]["AgeYears"])

    def test_machine_type_display_mapping_is_unchanged(self):
        out = self._rows([
            {"slNo": 7, "machine": "BRM-04", "machineType": "PBR",
             "commissioningDate": "01.01.2000", "age": None},
        ])
        self.assertEqual(out[0]["MachineType"], "BRM")

    def test_rows_missing_a_machine_name_are_skipped(self):
        out = self._rows([
            {"slNo": 8, "machine": None, "machineType": "BCM",
             "commissioningDate": "01.01.2000", "age": None},
        ])
        self.assertEqual(out, [])


if __name__ == "__main__":
    unittest.main()
