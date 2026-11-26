"""Targeted regression test for sources/ioh_poh_planning.py's 26-Nov-2026
Fleet Health rebuild: extract_current_year_progress() and its supporting
month/date parsing and status-classification helpers.

Pure-function test, synthetic data only (no network) -- same style as
test_swr_machine_details.py.

Background: Shruti's detailed Fleet Health rebuild spec asked for the
"Planned This Year" tab to classify every machine from the Progress
2026-2027 sheet into exactly four states, and for that classification to
be DYNAMIC -- a machine must move from "Upcoming" into "Overdue for
Release" automatically the month after its planned month passes, with no
code change needed each month:

  - "Completed"             -> has a clearly-written actual completion date
  - "Ongoing / Under Attention" -> not completed, completion date blank,
                                   but has an arrival date
  - "Overdue for Release"   -> no arrival date, and planned for a month
                               before the given `as_of` month
  - "Upcoming"              -> no arrival date, planned for `as_of`'s month
                               or later

This file proves that classification is correct and genuinely driven by
`as_of` (not hardcoded to any one month), proves the live sync's two known
header shapes (old Dy.CE/TM-letter-derived shape vs. Shruti's cleaned-up
"Progress 2026-2027" sheet shape) are both read correctly via the same
code path, and proves the month/year-only parsing used for "Month of
Planning" genuinely ignores a spurious day component however it got there.

Run from the swr_pipeline directory:
    python3 -m unittest tests.test_ioh_poh_current_year_progress -v
"""
import datetime
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sources import ioh_poh_planning as iohp

AS_OF_NOV26 = datetime.date(2026, 11, 26)


def _payload(poh_rows=None, ioh_rows=None):
    return {
        "currentYearProgress": {
            "poh": {"rows": poh_rows or []},
            "ioh": {"rows": ioh_rows or []},
        }
    }


class TestMonthYearParsing(unittest.TestCase):
    """_parse_month_year() must trust only (year, month) -- never the day --
    whatever shape the day component arrives in."""

    def test_clean_live_sync_date_with_normalized_day(self):
        # The live JSON sync's own format: Apps Script normalizes the day
        # to 01 before pushing.
        self.assertEqual(iohp._parse_month_year("01.04.2026"), (2026, 4))

    def test_raw_uploaded_xlsx_spurious_day_component(self):
        # The raw uploaded workbook shows a Sheets/Excel auto-fill quirk:
        # typing "Apr-26" into a date-formatted cell can store the CURRENT
        # EDIT DAY instead of defaulting to the 1st (e.g. day=26). Must
        # still read as April 2026, not be thrown off by the day.
        self.assertEqual(iohp._parse_month_year(datetime.date(2026, 4, 26)), (2026, 4))

    def test_mon_yy_text_cell(self):
        self.assertEqual(iohp._parse_month_year("SEPT-26"), (2026, 9))
        self.assertEqual(iohp._parse_month_year("Feb-27"), (2027, 2))

    def test_blank_or_dash_is_none(self):
        self.assertIsNone(iohp._parse_month_year(None))
        self.assertIsNone(iohp._parse_month_year("-"))
        self.assertIsNone(iohp._parse_month_year(""))


class TestAnyDateParsing(unittest.TestCase):
    """_parse_any_date() backs the arrival/completion columns -- must
    never mistake a 'TDC: ...' tentative estimate for a real date."""

    def test_well_formed_dates(self):
        self.assertEqual(iohp._parse_any_date("05.11.2026"), datetime.date(2026, 11, 5))
        self.assertEqual(iohp._parse_any_date("05.11.26"), datetime.date(2026, 11, 5))

    def test_tdc_estimate_is_not_a_real_date(self):
        self.assertIsNone(iohp._parse_any_date("TDC: 10.12.26"))
        self.assertIsNone(iohp._parse_any_date("tdc-10.12.2026"))

    def test_dash_and_none_are_not_dates(self):
        self.assertIsNone(iohp._parse_any_date("-"))
        self.assertIsNone(iohp._parse_any_date(None))


class TestExtractTentative(unittest.TestCase):

    def test_pulls_tdc_estimate_from_completion_field(self):
        self.assertEqual(iohp._extract_tentative("TDC: 10.12.26", None), "10.12.2026")

    def test_pulls_tdc_estimate_from_remarks_when_completion_has_none(self):
        self.assertEqual(iohp._extract_tentative("-", "ETA per remarks TDC:15.01.2027"), "15.01.2027")

    def test_no_tdc_anywhere_is_none(self):
        self.assertIsNone(iohp._extract_tentative("-", "all good, no delay"))


class TestClassification(unittest.TestCase):
    """The core four-state rule, and its dynamism across as_of dates."""

    def test_completed_wins_even_with_an_arrival_date_present(self):
        status = iohp._classify_cyp_status("01.09.2026", "20.10.2026", "Sep-26", AS_OF_NOV26)
        self.assertEqual(status, 'Completed')

    def test_ongoing_when_arrived_but_not_completed(self):
        status = iohp._classify_cyp_status("05.11.2026", "-", "Nov-26", AS_OF_NOV26)
        self.assertEqual(status, 'Ongoing / Under Attention')

    def test_ongoing_even_with_a_tdc_estimate_in_completion_field(self):
        # A TDC estimate is not a real completion date -- must still read
        # as Ongoing, not Completed.
        status = iohp._classify_cyp_status("05.11.2026", "TDC: 15.12.26", "Nov-26", AS_OF_NOV26)
        self.assertEqual(status, 'Ongoing / Under Attention')

    def test_overdue_when_planned_month_before_as_of_and_no_arrival(self):
        status = iohp._classify_cyp_status("-", "-", "Aug-26", AS_OF_NOV26)
        self.assertEqual(status, 'Overdue for Release')

    def test_upcoming_when_planned_month_is_as_of_month_or_later(self):
        self.assertEqual(iohp._classify_cyp_status("-", "-", "Nov-26", AS_OF_NOV26), 'Upcoming')
        self.assertEqual(iohp._classify_cyp_status("-", "-", "Feb-27", AS_OF_NOV26), 'Upcoming')

    def test_classification_is_dynamic_across_a_month_rollover(self):
        # Same machine, same "Dec-26" planning month: Upcoming as of
        # November, Overdue as of January -- proving the cutoff is
        # computed from as_of every run, never hardcoded.
        self.assertEqual(
            iohp._classify_cyp_status("-", "-", "Dec-26", datetime.date(2026, 11, 26)),
            'Upcoming')
        self.assertEqual(
            iohp._classify_cyp_status("-", "-", "Dec-26", datetime.date(2027, 1, 15)),
            'Overdue for Release')

    def test_none_when_nothing_to_classify_from(self):
        # No arrival, no completion, and an unparseable/blank planning
        # month -- stay honest and report None rather than guessing.
        self.assertIsNone(iohp._classify_cyp_status("-", "-", "-", AS_OF_NOV26))


class TestExtractCurrentYearProgressOldShape(unittest.TestCase):
    """The OLD live-sync header shape (Dy.CE/TM-letter-derived column
    names), confirmed still live via a direct Supabase pull as of
    25-Nov-2026."""

    def test_reads_old_shape_headers_and_classifies(self):
        payload = _payload(
            poh_rows=[{
                "SL.NO": 1, "MACHINE": "CSM-945", "Belongs to DIV": "MYS",
                "Attend at": "MYS", "Actually planned": "Aug-26",
                "FROM": "-", "TO": "-",
                "Stipulated days as per RDSO": 90, "POH Term": "1st POH",
                "Remarks": "pending release",
            }],
            ioh_rows=[{
                "SL.NO": 1, "MACHINE": "DTE-9006", "Belongs to DIV": "SBC",
                "Attend at": "SBC", "Actually planned": "15.10.2026",
                "FROM": "20.10.2026", "TO": "05.11.2026",
                "Total days taken for completion": 16,
                "Stipulated days as per RDSO": 20, "Remarks": "-",
            }],
        )
        out = iohp.extract_current_year_progress(payload, as_of=AS_OF_NOV26)
        self.assertEqual(len(out), 2)
        poh = next(r for r in out if r['Machine'] == 'CSM-945')
        self.assertEqual(poh['Kind'], 'POH')
        self.assertEqual(poh['Division'], 'MYS')
        self.assertEqual(poh['Status'], 'Overdue for Release')
        self.assertEqual(poh['MonthOfPlanning'], '2026-08')
        self.assertEqual(poh['MonthOfPlanningLabel'], 'Aug-26')
        ioh = next(r for r in out if r['Machine'] == 'DTE-9006')
        self.assertEqual(ioh['Kind'], 'IOH')
        self.assertEqual(ioh['Status'], 'Completed')
        self.assertEqual(ioh['DateOfArrival'], '20.10.2026')
        self.assertEqual(ioh['DateOfCompletion'], '05.11.2026')
        self.assertEqual(ioh['TotalDays'], 16)


class TestExtractCurrentYearProgressNewShape(unittest.TestCase):
    """Shruti's cleaned-up "Progress 2026-2027" sheet header shape -- must
    be read via the exact same function, with no schema flag needed."""

    def test_reads_new_shape_headers_and_classifies(self):
        payload = _payload(
            poh_rows=[{
                "S. No.": 1, "Machine No.": "BCM-392", "Parent Division": "UBL",
                "Planned At": "UBL", "Month of Planning": "Dec-26",
                "Date of Arrival of machine to depot": "-",
                "Actual Date of Completion of IOH/POH": "-",
                "Stipulated days as per RDSO": 90, "POH Term": "Subsequent POH",
                "Remarks": "awaiting release from division",
            }],
        )
        out = iohp.extract_current_year_progress(payload, as_of=AS_OF_NOV26)
        self.assertEqual(len(out), 1)
        r = out[0]
        self.assertEqual(r['Machine'], 'BCM-392')
        self.assertEqual(r['Division'], 'UBL')
        self.assertEqual(r['Status'], 'Upcoming')
        self.assertEqual(r['Term'], 'Subsequent POH')

    def test_division_outside_the_known_three_is_kept_raw_not_misclassified(self):
        payload = _payload(poh_rows=[{
            "Machine No.": "MPT-2014", "Parent Division": "Some Other Text",
            "Month of Planning": "Jan-27",
        }])
        out = iohp.extract_current_year_progress(payload, as_of=AS_OF_NOV26)
        self.assertIsNone(out[0]['Division'])
        self.assertEqual(out[0]['DivisionRaw'], 'Some Other Text')

    def test_rows_missing_a_machine_name_are_skipped(self):
        payload = _payload(poh_rows=[{"Parent Division": "MYS", "Month of Planning": "Jan-27"}])
        out = iohp.extract_current_year_progress(payload, as_of=AS_OF_NOV26)
        self.assertEqual(out, [])

    def test_empty_or_missing_currentyearprogress_degrades_to_empty_list(self):
        self.assertEqual(iohp.extract_current_year_progress({}, as_of=AS_OF_NOV26), [])
        self.assertEqual(iohp.extract_current_year_progress(None, as_of=AS_OF_NOV26), [])
        self.assertEqual(iohp.extract_current_year_progress({"currentYearProgress": {}}, as_of=AS_OF_NOV26), [])


if __name__ == "__main__":
    unittest.main()
