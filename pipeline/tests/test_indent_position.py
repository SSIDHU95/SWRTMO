"""
Tests for indent_position.extract_summary's totals, added 27-Nov-2026 after
Shruti noticed the Indent Position tab's "Total Indents" tile showing 0.
The live sync's Summary rows carry only the 9 stage columns (not the
sheet's "Total Indents of SWR" / "Total items" column), so the old parser
-- which only accepted a header literally named "Total" -- left total at 0.
On the sheet that column is exactly the sum of the 9 stages (checked
against her INDENT_POSITION.xlsx), so a missing total now falls back to
that sum; an explicit total column is still used when it does arrive.
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sources import indent_position as ip


def _block(title, rows):
    return {"title": title, "rows": rows}


class TestSummaryTotals(unittest.TestCase):
    def test_missing_total_falls_back_to_sum_of_stages(self):
        payload = {"summary": [_block("Indent position as on 25.11.2026", [
            {"store": "UBL", "Tender to Be Published": 22, "Tender Under Finalisation": 11,
             "PO Issued Pending Receipt": 16, "Material Received": 19, "Under Vetting": 4,
             "Under Concurrence": 2, "Indent Signing": 2, "Balance Indents to be proposed": 34,
             "Indent  proposal under process": 4},
            {"store": "SBC", "PO Issued Pending Receipt": 1, "Material Received": 18},
        ])]}
        s = ip.extract_summary(payload)
        by_store = {d["store"]: d for d in s["byIndent"]}
        self.assertEqual(by_store["UBL"]["total"], 114)
        self.assertEqual(by_store["SBC"]["total"], 19)

    def test_explicit_total_column_is_used_when_present(self):
        payload = {"summary": [_block("Items position as on 25.11.2026", [
            {"store": "Total", "Total items": 676, "Items to be Published": 103},
        ])]}
        s = ip.extract_summary(payload)
        self.assertEqual(s["byItem"][0]["total"], 676)

    def test_sheet_as_on_date_is_read_from_indent_block_title(self):
        payload = {"summary": [
            _block("Indent position as on 16.11.2026", [{"store": "UBL", "Material Received": 1}]),
            _block("Items position as on 16.11.2026", [{"store": "UBL", "Material Received": 1}]),
        ]}
        self.assertEqual(ip.extract_summary(payload)["sheetAsOn"], "16.11.2026")

    def test_no_title_date_gives_none(self):
        payload = {"summary": [_block("Indent position", [{"store": "UBL", "Material Received": 1}])]}
        self.assertIsNone(ip.extract_summary(payload)["sheetAsOn"])


if __name__ == "__main__":
    unittest.main()


class TestDaysInStatus(unittest.TestCase):
    """27-Nov-2026: one "days in current status" figure per indent. The
    sheet's own day count wins (decoded when the sheet formats it as a
    date); the status-start date is used only when that count is blank,
    unreadable or 0. Changed after Shruti found S.No 17 (Tender to be
    Published) showing 240 days computed from 'Sent to Purchase Unit'
    01.04.2026 while her sheet said 15 -- the indent had restarted that
    status after an LOA was withdrawn, so the old date was misleading."""
    AS_OF = __import__('datetime').date(2026, 11, 27)

    def _stage(self, key, row):
        return {key: {'rows': [row]}}

    def test_computed_from_status_date(self):
        st = self._stage('indentSigning', {'sno': 1, 'sentToCurrentHolderOn': '25.11.2026',
                                          'sentToCurrentHolderOnIso': '2026-11-25', 'daysPending': 0})
        ip.add_days_in_status(st, self.AS_OF)
        self.assertEqual(st['indentSigning']['rows'][0]['daysInStatus'], 2)

    def test_date_formatted_day_count_is_decoded(self):
        st = self._stage('poIssuedPendingReceipt', {'sno': 1, 'poDate': None, 'poDateIso': None, 'daysSincePo': '17.07.1900'})
        ip.add_days_in_status(st, self.AS_OF)
        self.assertEqual(st['poIssuedPendingReceipt']['rows'][0]['daysInStatus'], 199)

    def test_two_dates_in_one_cell_uses_the_first(self):
        st = self._stage('poIssuedPendingReceipt', {'sno': 11, 'poDate': '08/10/26 23/10/26', 'poDateIso': None})
        ip.add_days_in_status(st, self.AS_OF)
        self.assertEqual(st['poIssuedPendingReceipt']['rows'][0]['daysInStatus'], 50)

    def test_tender_not_yet_opened_gets_a_note_not_zero(self):
        st = self._stage('tenderUnderFinalisation', {'sno': 4, 'tenderOpeningDate': '30.11.2026',
                                                    'tenderOpeningDateIso': '2026-11-30', 'daysPendingSinceOpening': 0})
        ip.add_days_in_status(st, self.AS_OF)
        r = st['tenderUnderFinalisation']['rows'][0]
        self.assertIsNone(r['daysInStatus'])
        self.assertEqual(r['daysNote'], 'opens 30.11.2026')

    def test_future_typo_date_falls_back_to_sheet_value(self):
        st = self._stage('underVetting', {'sno': 1, 'sentOn': '17.11.2027', 'sentOnIso': '2027-11-17', 'daysPending': 9})
        ip.add_days_in_status(st, self.AS_OF)
        self.assertEqual(st['underVetting']['rows'][0]['daysInStatus'], 9)


    def test_sheet_count_wins_over_older_status_date(self):
        st = self._stage('tenderToBePublished', {'sno': 17, 'sentToPurchaseUnit': '01.04.2026',
                                                'sentToPurchaseUnitIso': '2026-04-01',
                                                'durationInCurrentStatus': 15})
        ip.add_days_in_status(st, self.AS_OF)
        r = st['tenderToBePublished']['rows'][0]
        self.assertEqual(r['daysInStatus'], 15)
        self.assertIsNone(r['statusSince'])  # the date doesn't explain 15 days

    def test_matching_sheet_count_keeps_its_status_date(self):
        st = self._stage('tenderToBePublished', {'sno': 1, 'sentToPurchaseUnit': '06.02.2026',
                                                'sentToPurchaseUnitIso': '2026-02-06',
                                                'durationInCurrentStatus': 294})
        ip.add_days_in_status(st, self.AS_OF)
        r = st['tenderToBePublished']['rows'][0]
        self.assertEqual((r['daysInStatus'], r['statusSince']), (294, '06.02.2026'))

    def test_opened_tender_with_zero_uses_opening_date(self):
        st = self._stage('tenderUnderFinalisation', {'sno': 2, 'tenderOpeningDate': '12.10.2026',
                                                    'tenderOpeningDateIso': '2026-10-12', 'daysPendingSinceOpening': 0})
        ip.add_days_in_status(st, self.AS_OF)
        self.assertEqual(st['tenderUnderFinalisation']['rows'][0]['daysInStatus'], 46)


class TestHeaderAlternates(unittest.TestCase):
    def test_renamed_po_date_header_is_still_read(self):
        rows = ip._build_rows([{'Sl. No.': 1, 'PO/modified PO Date': '12.05.2026'}],
                              ip.INDENT_STAGE_COLUMNS['poIssuedPendingReceipt'])
        self.assertEqual(rows[0]['poDateIso'], '2026-05-12')

    def test_case_and_spacing_differences_are_tolerated(self):
        rows = ip._build_rows([{'Sl. No.': 1, 'REMARKS': 'x'}], ip.INDENT_STAGE_COLUMNS['underVetting'])
        self.assertEqual(rows[0]['remarks'], 'x')
