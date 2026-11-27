"""Tests for sources/tenders_contracts.py, added 27-Nov-2026 after Shruti
reported the Tenders & Contracts tab showing 0 agreements / Rs 0 committed
value / 0 expiring / 0 expired, with an empty Contract Runway.

Root cause (confirmed against the live table): the Agreement Details sheet's
serial-number column has a spaces-only header, so the live sync drops that
column and every agreement row arrives without an S No. The adapter used to
skip any row without a numeric S No -- i.e. all 29 of them. Synthetic
fixtures only; no network.
"""
import datetime
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sources import tenders_contracts as tc
import pipeline_helpers as ph

AS_OF = datetime.date(2026, 11, 27)


def _ag(firm, ctype='AMC', amount='1000000', revised='-', orig='31.03.2027', ext='-', **extra):
    r = {'Contract Type': ctype, 'Firm Name': firm, 'Name of work': f'Work for {firm}',
         'Date of Issue of LOA': '01.04.2025', 'Agreement Amount': amount,
         'Revised agreement value': revised, 'Original Currency Date': orig,
         'Extended Currency Date': ext}
    r.update(extra)
    return r


class TestAgreementsWithoutSerialNumber(unittest.TestCase):
    def test_rows_without_s_no_are_kept_and_numbered_in_sheet_order(self):
        ags = tc._extract_agreements([_ag('Firm A'), _ag('Firm B'), _ag('Firm C')], AS_OF)
        self.assertEqual([(a['sno'], a['firmName']) for a in ags], [(1, 'Firm A'), (2, 'Firm B'), (3, 'Firm C')])

    def test_blank_rows_are_still_skipped(self):
        rows = [_ag('Firm A'), {'Firm Name': None}, {'Remarks from TMO/Accounts': 'note'}, _ag('Firm B')]
        self.assertEqual([a['sno'] for a in tc._extract_agreements(rows, AS_OF)], [1, 2])

    def test_explicit_s_no_still_used_when_present(self):
        ags = tc._extract_agreements([_ag('Firm A', **{'S No': 7})], AS_OF)
        self.assertEqual(ags[0]['sno'], 7)

    def test_kpis_no_longer_zero(self):
        payload = {'agreementDetails': {'rows': [
            _ag('Firm A', amount='48485828.27', orig='31.03.2028'),
            _ag('Firm B', amount='86558275', revised='100655891', orig='31.03.2025'),           # expired
            _ag('Firm C', ctype='Works Contract', amount='23065765', orig='19.12.2026'),        # expiring soon
            _ag('Firm D', ctype='Works Contract', amount='9775200', orig='30.09.2026', ext='31-03-2027'),
        ]}}
        k = tc.extract_tenders_contracts(payload, AS_OF)['kpi']
        self.assertEqual(k['totalAgreements'], 4)
        self.assertEqual(k['expiredCount'], 1)
        self.assertEqual(k['expiringSoonCount'], 1)
        self.assertAlmostEqual(k['totalCommittedValue'], 48485828.27 + 100655891 + 23065765 + 9775200, places=2)


class TestProposalValuesAndStages(unittest.TestCase):
    def test_lakh_text_value_is_read(self):
        self.assertEqual(tc._value_num('54.92 Lakhs'), 5492000.0)
        self.assertEqual(tc._value_num('Rs. 3,09,10,205/-'), 30910205.0)
        self.assertEqual(tc._value_num('1.5 Cr'), 15000000.0)
        self.assertIsNone(tc._value_num('-'))

    def test_misspelt_administrative_still_classified(self):
        """AMC proposal S.No 9 says 'sent for admnistrative approval' and
        used to show as 'Step ? of 8'."""
        key, _ = ph.classify_proposal_stage('New AMC Proposal sent for admnistrative approval on 23.11.2026')
        self.assertEqual(key, 'admin_approval_sent')


if __name__ == '__main__':
    unittest.main()
