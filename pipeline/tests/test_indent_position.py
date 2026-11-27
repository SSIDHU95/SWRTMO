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
