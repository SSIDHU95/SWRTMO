"""
Integrity test suite for the SWR Track Machine Dashboard pipeline
(Task #55, explicitly requested by the user alongside the Python rebuild
and the "proper structure" requirement).

Three layers, cheapest/most-isolated first:

  1. TestPipelineHelpers      - unit tests for the pure functions in
                                 pipeline_helpers.py (fuzzy matching,
                                 classification, date parsing). These
                                 never touch the network or a fixture
                                 file.
  2. TestSourceAdapters       - each adapter's extract_*() function run
                                 against its cached /tmp/*_latest.json
                                 fixture (where one exists in this
                                 sandbox), asserting STRUCTURAL invariants
                                 (required keys present, lists are lists,
                                 numeric-or-None fields never come back as
                                 some other type, counts are non-negative)
                                 rather than brittle exact-count
                                 assertions that would break the moment
                                 the live data changes shape slightly.
                                 Still network-free: these only exercise
                                 extract_*(), never fetch_payload().
  3. TestAggregateEndToEnd    - aggregate.run() driven end-to-end against
                                 every adapter monkeypatched to its cached
                                 fixture (see tests/_validate_aggregate.py
                                 for the same patching, reused here),
                                 asserting cross-cutting invariants the
                                 final dashboard_payload.json must hold
                                 regardless of which day's data produced
                                 it: fleetSize matches the machine list
                                 length and the status-count sum, every
                                 machine record carries its required
                                 keys, the whole payload round-trips
                                 through JSON, and known-closed/open gaps
                                 degrade exactly as documented (gap #53
                                 empty, no live IOH/POH letter -> None).

Run from the swr_pipeline directory:
    python3 -m unittest tests.test_integrity -v
or, for a quick pass/fail without -v noise:
    python3 -m unittest tests.test_integrity

This suite is network-free by design (no SUPABASE_URL/SUPABASE_ANON_KEY
needed) so it can run in GitHub Actions on every push, independent of
whether that run also does a live dry-run against Supabase.
"""
import datetime
import json
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pipeline_helpers as ph
import aggregate
from sources import (
    swr_machine_details, cumulative_progress, daily_positions, double_shift_roster,
    exception_reports, indent_position, ioh_poh_planning, monthly_machine,
    shortfall_reasons, sidings_rest_house, tenders_contracts, unit_cost,
    action_plan_53, cumulative_failure,
)
# hq_summary and rdps_points are intentionally not imported/stubbed here -
# aggregate.py no longer fetches either (01-Nov-2026, see its own module
# docstring), so there is nothing in it left to monkeypatch.

FIXTURE_DIR = "/tmp"


def _rest_dump_payload(path):
    """Cached /tmp/*_latest.json files are raw REST responses: a list of
    [{"payload": {...}, "updated_at": "..."}]. Returns just the payload
    dict of the first (only) row, or None if the fixture file is absent -
    callers must skip (not fail) when a fixture is missing, since this
    sandbox was never actually synced for every one of the 16 tables."""
    if not os.path.exists(path):
        return None
    with open(path) as f:
        rows = json.load(f)
    if not rows:
        return None
    return rows[0]["payload"]


def skip_if_missing(path):
    def deco(test_fn):
        def wrapper(self):
            payload = _rest_dump_payload(path)
            if payload is None:
                self.skipTest(f"no cached fixture at {path}")
            return test_fn(self, payload)
        wrapper.__name__ = test_fn.__name__
        return wrapper
    return deco


# =====================================================================
# 1. pipeline_helpers.py unit tests
# =====================================================================
class TestPipelineHelpers(unittest.TestCase):

    def test_machine_key_leading_zero_and_parenthetical(self):
        self.assertEqual(ph.machine_key('UTV-058'), ph.machine_key('UTV-58'))
        self.assertEqual(ph.machine_key('UTV-058'), ph.machine_key('UTV 058'))
        self.assertEqual(ph.machine_key('UTV-003'), ph.machine_key('UTV-003 (SAN)'))

    def test_machine_key_aliased_bridges_progress_vs_shortfall_prefixes(self):
        # 07-Oct-2026 user-confirmed cross-sheet prefix aliases.
        self.assertEqual(ph.machine_key_aliased('SBCM-1889'), ph.machine_key_aliased('FRM 1889'))
        self.assertEqual(ph.machine_key_aliased('BRM-12'), ph.machine_key_aliased('PBR 12'))
        self.assertEqual(ph.machine_key_aliased('PCTM-5'), ph.machine_key_aliased('PCT 5'))

    def test_canon_machine_alias_t28_variants(self):
        self.assertEqual(ph.canon_machine_alias('T-28(908A &B)'), 'T-28 (908 A&B)')
        self.assertEqual(ph.canon_machine_alias('T-28(908 A&B)'), 'T-28 (908 A&B)')

    def test_disp_type_merges_pctm_and_pct_into_uni(self):
        self.assertEqual(ph.disp_type('PCTM'), 'UNI')
        self.assertEqual(ph.disp_type('PCT'), 'UNI')
        self.assertEqual(ph.disp_type('BCM'), 'BCM')  # untouched type passes through

    def test_disp_type_merges_18nov2026_corrected_aliases(self):
        # Shruti's SWR_Machine_Details_1.xlsx (18-Nov-2026): every SBCM
        # machine is sometimes written "FRM", every BRM machine is
        # sometimes written "PBR", every TLE machine is always written
        # "SQRS", every PCCM machine is always written "T-28", and WST-01
        # is the 10th machine listed under the DUO section - none of
        # these six raw prefixes are real types of their own.
        self.assertEqual(ph.disp_type('FRM'), 'SBCM')
        self.assertEqual(ph.disp_type('PBR'), 'BRM')
        self.assertEqual(ph.disp_type('SQRS'), 'TLE')
        self.assertEqual(ph.disp_type('T'), 'PCCM')
        self.assertEqual(ph.disp_type('WST'), 'DUO')
        # A machine recorded either way must land on the exact same
        # canonical type:
        self.assertEqual(ph.disp_type(ph.mtype_of('PBR 04')), ph.disp_type(ph.mtype_of('BRM-04')))
        self.assertEqual(ph.disp_type(ph.mtype_of('FRM 1889')), ph.disp_type(ph.mtype_of('SBCM 1889')))

    def test_tamping_info_and_full_name_cover_every_display_type(self):
        # Every code disp_type() can ever RETURN (i.e. every value in
        # UNI_MERGE, plus every type disp_type() leaves untouched) must
        # have both a TAMPING_INFO entry and an MTYPE_FULL_NAME entry -
        # otherwise a machine of that type silently falls back to
        # ('Non-Tamping', 'Unclassified') and no full name, which is
        # exactly the 18-Nov-2026 bug this table exists to prevent.
        display_codes = set(ph.UNI_MERGE.values()) | {
            'BCM', 'CSM', 'DGS', 'DTE', 'MDU', 'MPT', 'RBMV', 'RGM', 'UTV', 'RMV',
        }
        for code in sorted(display_codes):
            self.assertIn(code, ph.TAMPING_INFO, f"{code} missing from TAMPING_INFO")
            self.assertIn(code, ph.MTYPE_FULL_NAME, f"{code} missing from MTYPE_FULL_NAME")
        # None of the raw alias prefixes should be keys any more - keeping
        # both the raw and canonical key around is exactly what let this
        # bug happen in the first place (see pipeline_helpers.py's
        # TAMPING_INFO comment).
        for raw in ('FRM', 'PBR', 'SQRS', 'T', 'WST', 'PCT'):
            self.assertNotIn(raw, ph.TAMPING_INFO, f"{raw} should not be its own TAMPING_INFO key any more")

    def test_elapsed_fy_months_dynamic_window(self):
        # 18-Nov-2026 fix: the YoY comparison used to hardcode
        # ['apr'..'sep'] regardless of how far the live Cumulative
        # Progress sheet's own data actually extended. _elapsed_fy_months
        # is the shared helper that now derives "Apr through however far
        # the sheet reaches" for BOTH sides of that comparison, mirroring
        # build_base_payload()'s own cum_period_label thresholds exactly.
        months, phrase = aggregate._elapsed_fy_months(datetime.date(2026, 11, 13), 2026)
        self.assertEqual(months, ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov'])
        self.assertEqual(phrase, 'Apr – 13 Nov')

        # The 28-Sep-or-later special case collapses to the short "Apr-Sep"
        # phrase, same threshold build_base_payload() uses for periodCovered.
        months, phrase = aggregate._elapsed_fy_months(datetime.date(2026, 9, 29), 2026)
        self.assertEqual(months, ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep'])
        self.assertEqual(phrase, 'Apr–Sep')

        # No derivable as-of date (e.g. no ProportionateTarget/TargetPerMonth
        # samples) falls back to the full 12-month FY, never silently
        # reverting to the old 6-month Apr-Sep cap.
        months, phrase = aggregate._elapsed_fy_months(None, 2026)
        self.assertEqual(months, ph.MONTH_ABBR)
        self.assertEqual(phrase, 'Apr–Mar')

        # A cum_as_of outside fy_year's own Apr-Mar span is treated as
        # undated input, same defensive fallback.
        months, _ = aggregate._elapsed_fy_months(datetime.date(2099, 1, 1), 2026)
        self.assertEqual(months, ph.MONTH_ABBR)

    def test_mtype_of_extracts_leading_letters(self):
        self.assertEqual(ph.mtype_of('BCM-351'), 'BCM')
        self.assertEqual(ph.mtype_of('T-28 (403&404)'), 'T')

    def test_classify_shortfall_known_categories(self):
        self.assertEqual(ph.classify_shortfall('Block not permitted by controller due to heavy traffic'),
                          'Block Not Permitted by Traffic')
        self.assertEqual(ph.classify_shortfall('Machine under repair, breakdown'),
                          'Breakdown / Under Repair')
        self.assertEqual(ph.classify_shortfall('tamped 200m, ballast cleaning done'),
                          'Productive Work - Full Progress')
        self.assertIsNone(ph.classify_shortfall(''))
        self.assertIsNone(ph.classify_shortfall(None))

    def test_classify_exc_category_fuzzy_match_and_fallback(self):
        bucket, label = ph.classify_exc_category('GSU/CN/K- RIDE')
        self.assertEqual(bucket, 'block')
        self.assertIn('GSU', label)
        bucket, label = ph.classify_exc_category('Some brand new unrecognised text')
        self.assertEqual(bucket, 'other')
        self.assertEqual(label, 'Some brand new unrecognised text')

    def test_canon_srden_valid_and_invalid(self):
        # canon_srden() is for the Daily Progress sheet's single-LETTER
        # Sr.DEN codes ("Sr.DEN/N"); the full-WORD form ("Sr.DEN/North",
        # used by the Sidings/Rest House sheet) is canon_srden_word()'s
        # job, covered by test_canon_srden_word_full_word_label above.
        self.assertEqual(ph.canon_srden('MYS', 'Sr.DEN/N'), 'MYS Sr.DEN/N')
        self.assertIsNone(ph.canon_srden('XYZ', 'Sr.DEN/N'))  # not a real division
        self.assertIsNone(ph.canon_srden('MYS', ''))
        self.assertIsNone(ph.canon_srden('MYS', 'Sr.DEN/North'))  # full word -- not this function's job

    def test_canon_srden_word_full_word_label(self):
        self.assertEqual(ph.canon_srden_word('SBC', 'Sr.DEN/South'), 'SBC Sr.DEN/S')
        # Unrecognised jurisdiction label still buckets rather than dropping.
        self.assertEqual(ph.canon_srden_word('MYS', 'New Line Construction'), 'MYS New Line Construction')

    def test_parse_ddmmyyyy_formats(self):
        for s in ('17.09.2026', '17-09-2026', '17/09/2026'):
            dt = ph.parse_ddmmyyyy(s)
            self.assertIsNotNone(dt, f"failed to parse {s!r}")
            self.assertEqual((dt.day, dt.month, dt.year), (17, 9, 2026))
        self.assertIsNone(ph.parse_ddmmyyyy(''))
        self.assertIsNone(ph.parse_ddmmyyyy(None))

    def test_coerce_date_ddmmyyyy_handles_iso_and_ddmmyyyy(self):
        # Only a full ISO *timestamp* (with a T and time-of-day, exactly
        # what a real Google Sheets Date cell serializes to in Apps
        # Script's JSON push) is recognised as the ISO case; a bare
        # 'YYYY-MM-DD' date string isn't one of the shapes this function
        # promises to recognise, so it passes through unchanged (by
        # design -- see its own docstring: "a genuinely free-text remark
        # ... surfaces as-is rather than silently disappearing").
        self.assertEqual(ph.coerce_date_ddmmyyyy('2026-09-17T00:00:00.000Z'), '17.09.2026')
        self.assertEqual(ph.coerce_date_ddmmyyyy('17.09.2026'), '17.09.2026')
        self.assertEqual(ph.coerce_date_ddmmyyyy('17/9/2026'), '17.09.2026')
        self.assertEqual(ph.coerce_date_ddmmyyyy('2026-09-17'), '2026-09-17')  # passthrough, not ISO-timestamp shaped
        self.assertIsNone(ph.coerce_date_ddmmyyyy(None))
        self.assertIsNone(ph.coerce_date_ddmmyyyy(''))

    def test_norm_unit_canonicalizes_case(self):
        self.assertEqual(ph.norm_unit('km'), 'Km')
        self.assertEqual(ph.norm_unit('KM'), 'Km')
        self.assertEqual(ph.norm_unit('t/o'), 'T/O')
        self.assertIsNone(ph.norm_unit(None))

    def test_classify_proposal_stage_latest_stage_wins(self):
        key, label = ph.classify_proposal_stage('Sent for concurrence. Concurrence obtained on 12.05.2026.')
        self.assertEqual(key, 'concurrence_obtained')  # furthest stage, not first match

    def test_classify_tender_stage_falls_back_to_holder(self):
        key, label, order = ph.classify_tender_stage('', 'IREPS')
        self.assertEqual(key, 'with_ireps')
        self.assertGreater(order, 0)

    def test_classify_loa_status_default(self):
        key, label = ph.classify_loa_status('')
        self.assertEqual(key, 'in_process')

    def test_brand_hits_case_insensitive(self):
        hits = ph.brand_hits('Supplied by M/s. PLASSER (India) Pvt. Ltd. as per AMC')
        self.assertIn('plasser', hits)

    def test_num_parses_indian_currency_strings(self):
        self.assertEqual(ph.num('Rs. 5,32,09,612/-'), 53209612.0)
        self.assertEqual(ph.num('1,90,58,825'), 19058825.0)
        self.assertIsNone(ph.num('-'))
        self.assertIsNone(ph.num(None))

    def test_classify_infra_responsibility_buckets(self):
        self.assertEqual(ph.classify_infra_responsibility('Not feasible due to space constraint'), 'Not Feasible')
        self.assertEqual(ph.classify_infra_responsibility('Done by CN org'), 'CN (Construction Org)')
        self.assertIsNone(ph.classify_infra_responsibility(''))

    def test_normalize_infra_status_collapses_case_variants(self):
        self.assertEqual(ph.normalize_infra_status('Not feasible'), 'Not Feasible')
        self.assertEqual(ph.normalize_infra_status('NOT FEASIBLE'.title()), 'Not Feasible')
        self.assertEqual(ph.normalize_infra_status('????'), 'Unknown')

    def test_parse_station_block_strips_leading_count(self):
        self.assertEqual(ph.parse_station_block('(3)\nUBL, NVU, SBC'), ['UBL', 'NVU', 'SBC'])
        self.assertEqual(ph.parse_station_block(''), [])


# =====================================================================
# 2. Source adapter extraction structural tests (fixture-driven)
# =====================================================================
class TestSourceAdapters(unittest.TestCase):

    @skip_if_missing(f"{FIXTURE_DIR}/swr_machine_details_latest.json")
    def test_swr_machine_details(self, payload):
        home_div, bare_type_div, norm = swr_machine_details.extract_home_div(payload)
        self.assertIsInstance(home_div, dict)
        self.assertTrue(all(v in ('MYS', 'SBC', 'UBL') for v in home_div.values()))
        age_profile = swr_machine_details.extract_age_profile(payload)
        self.assertIsInstance(age_profile, list)
        for rec in age_profile:
            self.assertIn('Machine', rec)
            self.assertIn('AgeYears', rec)
            self.assertTrue(rec['AgeYears'] is None or isinstance(rec['AgeYears'], int))

    def test_extract_home_div_resolves_brm_pbr_alias(self):
        # Regression test for the 18-Nov-2026 "PBR shown as UNKNOWN
        # division" bug (Shruti's screenshot): Original Allocation always
        # names these machines "BRM" (confirmed in her own Div Wise
        # Machine Original Alloc sheet), but Cumulative Progress
        # sometimes names the SAME physical machine "PBR" instead. Before
        # the fix, home_div was keyed by a plain alnum normalization, so
        # "BRM-04" and "PBR 04" produced two different keys and never
        # matched - a machine away at IOH/POH (CurrentDIV not
        # MYS/SBC/UBL) then fell all the way through to HomeDivision =
        # 'UNKNOWN'. No live fixture needed: this is a synthetic payload
        # built the same shape swr_machine_details_sync.gs produces.
        payload = {"originalAllocation": {
            "mys": ["BCM-351"],
            "sbc": ["BRM-04", "BRM-56669"],
            "ubl": ["FRM-1889"],
        }}
        home_div, bare_type_div, norm = swr_machine_details.extract_home_div(payload)
        records = cumulative_progress.extract_cum_progress(
            {"years": {"2026-2027": {"rows": [
                {"Machine": "PBR 04", "DIV": "POH", "RowType": "machine"},
                {"Machine": "FRM 1889", "DIV": "MYS", "RowType": "machine"},
            ]}}},
            home_div, bare_type_div, norm)
        by_machine = {r['Machine']: r for r in records}
        self.assertEqual(by_machine['PBR 04']['HomeDivision'], 'SBC',
                          "PBR 04 should resolve to SBC via its BRM-04 Original Allocation entry, not UNKNOWN")
        self.assertNotEqual(by_machine['PBR 04']['HomeDivision'], 'UNKNOWN')
        # FRM 1889 is listed in Original Allocation under "ubl" (as
        # "FRM-1889"), so its HomeDivision is UBL regardless of its
        # CurrentDIV ("MYS", e.g. on deputation) - home_div always wins
        # over CurrentDIV when a match is found (see extract_cum_progress).
        self.assertEqual(by_machine['FRM 1889']['HomeDivision'], 'UBL')

    def test_extract_prior_year_by_type_months_window(self):
        # 18-Nov-2026 fix: extract_prior_year_by_type() used to always sum
        # the FULL prior FY regardless of how much of the current FY had
        # elapsed, which overstated the prior-year side of the YoY "same
        # period" comparison once the current-year side stopped being
        # hardcoded to Apr-Sep. No live fixture needed: a synthetic
        # payload with the real "years" shape, full 12 months filled in
        # for both FY tabs (as a completed prior FY would have).
        row = {"Machine": "BCM-351", "RowType": "machine",
               "Apr": 1.0, "May": 1.0, "Jun": 1.0, "Jul": 1.0, "Aug": 1.0, "Sep": 1.0,
               "Oct": 1.0, "Nov": 1.0, "Dec": 1.0, "Jan": 1.0, "Feb": 1.0, "Mar": 1.0}
        payload = {"years": {"2025-2026": {"rows": [row]}, "2026-2027": {"rows": [row]}}}

        # Default (no months passed): full 12-month total, for backward
        # compatibility with callers that want a plain full-year sum.
        prior_by_type, prior_total, _, _, _ = cumulative_progress.extract_prior_year_by_type(payload)
        self.assertEqual(prior_total, 12.0)

        # An explicit window sums only those columns - this is what
        # aggregate.py's run() now passes (the same dynamically-derived
        # "Apr through however far the sheet currently reaches" window
        # used on the current-year side), so both sides of the YoY
        # comparison are always summing the identical span.
        prior_by_type, prior_total, _, _, _ = cumulative_progress.extract_prior_year_by_type(
            payload, months=['Apr', 'May', 'Jun'])
        self.assertEqual(prior_total, 3.0)
        self.assertEqual(prior_by_type['BCM'], 3.0)

    def test_extract_monthly_machine_by_machine(self):
        # 18-Nov-2026 (new feature per Shruti): synthetic payload in the
        # real live shape (months as a DICT, confirmed against the live
        # sync - this adapter used to claim/assume a list). Covers: (a)
        # the BRM/PBR cross-sheet join-key aliasing, (b) a blank-machine
        # "Total Progress" row being skipped, (c) a numeric value and a
        # remark-string value both surviving, (d) '-'/empty being treated
        # as "no entry that day", and (e) the "NOV 26" (space, not
        # hyphen) month-tab naming slip on her own sheet still parsing.
        payload = {"months": {
            "OCT-26": {"categories": [{"category": "BRM", "title": "BRM...", "periods": [
                {"rows": [
                    {"machine": "BRM 04", "daily": {"1": 0.26, "2": "Depot work", "3": "-", "4": None}},
                    {"machine": "Total Progress of all machines", "daily": {"1": 0.26}},
                ]},
            ]}]},
            "NOV 26": {"categories": [{"category": "BRM", "title": "BRM...", "periods": [
                {"rows": [{"machine": "BRM 04", "daily": {"1": 9.6}}]},
            ]}]},
        }}
        by_machine = monthly_machine.extract_monthly_machine_by_machine(payload)
        self.assertIn('PBR4', by_machine)  # BRM 04 -> aliased join key, same as Cumulative Progress's "PBR 04"
        self.assertNotIn('BRM4', by_machine)
        log = by_machine['PBR4']
        self.assertEqual(log, [
            {"date": "2026-10-01", "value": 0.26},
            {"date": "2026-10-02", "value": "Depot work"},
            {"date": "2026-11-01", "value": 9.6},
        ])
        # The blank-named "Total Progress" summary row must never show up
        # as its own machine.
        self.assertTrue(all('total' not in k.lower() for k in by_machine.keys()))

    @skip_if_missing(f"{FIXTURE_DIR}/cumulative_progress_latest.json")
    def test_cumulative_progress(self, payload):
        home_div, bare_type_div, norm = ({}, {}, lambda x: x)
        records = cumulative_progress.extract_cum_progress(payload, home_div, bare_type_div, norm)
        self.assertIsInstance(records, list)
        self.assertGreater(len(records), 0)
        required = {'SlNo', 'Machine', 'MachineType', 'CurrentDIV', 'HomeDivision', 'Status', 'Unit'}
        valid_status = ('Active', 'IOH', 'POH', 'UNKNOWN')
        for rec in records:
            self.assertTrue(required.issubset(rec.keys()))
            status = rec['Status']
            self.assertTrue(
                status in valid_status or status.startswith('Deputed-'),
                f"unexpected Status {status!r} for machine {rec.get('Machine')!r}")
        prior = cumulative_progress.extract_prior_year_by_type(payload)
        self.assertEqual(len(prior), 5)

    @skip_if_missing(f"{FIXTURE_DIR}/exception_reports_latest.json")
    def test_exception_reports_structural(self, _unused_single_payload):
        with open(f"{FIXTURE_DIR}/exception_reports_latest.json") as f:
            rows = json.load(f)
        all_rows = [(r["payload"], r.get("updated_at")) for r in rows]
        ex = exception_reports.extract_exceptions(all_rows)
        self.assertIsInstance(ex, dict)
        for day_key, day_rows in ex.items():
            self.assertIsInstance(day_rows, list)
            for r in day_rows:
                self.assertIn('Category', r)
                self.assertIn('Machine', r)
                self.assertTrue(r['DaysUnder'] is None or isinstance(r['DaysUnder'], (int, float)))

    @skip_if_missing(f"{FIXTURE_DIR}/double_shift_roster_latest.json")
    def test_double_shift_roster(self, payload):
        records = double_shift_roster.extract_double_shift(payload)
        self.assertIsInstance(records, list)
        for r in records:
            self.assertIn(r['shift'], ('DAY', 'NIGHT'))
            if r.get('durationHours') is not None:
                self.assertGreaterEqual(r['durationHours'], 0)
            self.assertEqual(r['over12h'], bool(r.get('durationHours') is not None and r['durationHours'] > 12))

    @skip_if_missing(f"{FIXTURE_DIR}/indent_position_latest.json")
    def test_indent_position(self, payload):
        as_of = datetime.date(2026, 10, 31)
        result = indent_position.extract_indent_position(payload, as_of)
        for key in ('sourceFile', 'asOf', 'summary', 'safetyItems', 'stages', 'stageOrder', 'attention'):
            self.assertIn(key, result)
        self.assertEqual(set(result['stages'].keys()), set(result['stageOrder']))
        severities = [a['severity'] for a in result['attention']]
        rank = dict(high=0, medium=1, low=2)
        self.assertEqual(severities, sorted(severities, key=lambda s: rank.get(s, 9)),
                          "attention list must be severity-sorted, most severe first")

    @skip_if_missing(f"{FIXTURE_DIR}/sidings_rest_house_latest.json")
    def test_sidings_rest_house(self, payload):
        result = sidings_rest_house.extract_sidings_rest_houses(payload)
        self.assertIsNotNone(result)
        for key in ('summary', 'stationRows', 'srdenTable', 'responsibility'):
            self.assertIn(key, result)
        for row in result['stationRows']:
            self.assertIn(row['division'], ('MYS', 'SBC', 'UBL', None))

    @skip_if_missing(f"{FIXTURE_DIR}/tenders_contracts_latest.json")
    def test_tenders_contracts(self, payload):
        as_of = datetime.date(2026, 10, 31)
        result = tenders_contracts.extract_tenders_contracts(payload, as_of)
        self.assertIsNotNone(result)
        for agreement in result['agreements']:
            self.assertIn(agreement['expiryStatus'], ('Valid', 'Expiring Soon', 'Expired', 'Unknown'))
        sev_rank = dict(high=0, medium=1, low=2)
        sevs = [a['severity'] for a in result['attention']]
        self.assertEqual(sevs, sorted(sevs, key=lambda s: sev_rank.get(s, 9)))

    @skip_if_missing(f"{FIXTURE_DIR}/unit_cost_latest.json")
    def test_unit_cost(self, payload):
        result = unit_cost.extract_unit_cost(payload)
        self.assertIn('years', result)
        self.assertIn('latestFy', result)
        self.assertEqual(result['latestFy'], sorted(result['years'].keys())[-1])
        for fy, block in result['years'].items():
            for m in block['machines']:
                self.assertIn('machine', m)
                self.assertIn('machineKey', m)

    @skip_if_missing(f"{FIXTURE_DIR}/action_plan_53_latest.json")
    def test_action_plan_53(self, payload):
        result = action_plan_53.extract_action_plan_53(payload)
        self.assertIsNotNone(result)
        self.assertEqual(result['totalPoints'], len(result['points']))
        for pt in result['points']:
            self.assertIn('sno', pt)
            self.assertIn('category', pt)
            self.assertIsInstance(pt['history'], list)

    @skip_if_missing(f"{FIXTURE_DIR}/daily_positions_latest.json")
    def test_daily_positions(self, payload):
        result = daily_positions.extract_positions(payload)
        self.assertIn('positions', result)
        self.assertIsInstance(result['positions'], dict)


# =====================================================================
# 3. aggregate.run() end-to-end integrity
# =====================================================================
class TestAggregateEndToEnd(unittest.TestCase):
    """Mirrors tests/_validate_aggregate.py's fixture patching, but asserts
    hard invariants instead of just printing a summary."""

    REQUIRED_FIXTURES = [
        f"{FIXTURE_DIR}/swr_machine_details_latest.json",
        f"{FIXTURE_DIR}/cumulative_progress_latest.json",
        f"{FIXTURE_DIR}/exception_reports_latest.json",
    ]

    @classmethod
    def setUpClass(cls):
        # This end-to-end class needs the cached fixture files this
        # sandbox captured earlier in the project. A CI runner (GitHub
        # Actions) has no such files on disk and no live Supabase
        # credentials either -- skip the whole class there rather than
        # erroring, so the pure pipeline_helpers/adapter-structural tests
        # above still run as a CI smoke check (see .github/workflows'
        # comment on this).
        missing = [p for p in cls.REQUIRED_FIXTURES if not os.path.exists(p)]
        if missing:
            raise unittest.SkipTest(f"missing required fixture(s) for end-to-end test: {missing}")
        cls._patch_with_fixtures()
        cls.outdir = "/tmp/aggregate_integrity_out"
        cls.as_of = datetime.date(2026, 10, 31)
        cls.out_path = aggregate.run(cls.outdir, as_of=cls.as_of)
        with open(cls.out_path) as f:
            cls.payload = json.load(f)

    @staticmethod
    def _patch_with_fixtures():
        def rest_dump(path):
            def _fetch():
                with open(path) as f:
                    rows = json.load(f)
                return rows[0]["payload"], rows[0].get("updated_at", "fixture")
            return _fetch

        swr_machine_details.fetch_payload = rest_dump(f"{FIXTURE_DIR}/swr_machine_details_latest.json")
        cumulative_progress.fetch_payload = rest_dump(f"{FIXTURE_DIR}/cumulative_progress_latest.json")
        daily_positions.fetch_payload = rest_dump(f"{FIXTURE_DIR}/daily_positions_latest.json")
        double_shift_roster.fetch_payload = rest_dump(f"{FIXTURE_DIR}/double_shift_roster_latest.json")
        indent_position.fetch_payload = rest_dump(f"{FIXTURE_DIR}/indent_position_latest.json")
        sidings_rest_house.fetch_payload = rest_dump(f"{FIXTURE_DIR}/sidings_rest_house_latest.json")
        tenders_contracts.fetch_payload = rest_dump(f"{FIXTURE_DIR}/tenders_contracts_latest.json")
        unit_cost.fetch_payload = rest_dump(f"{FIXTURE_DIR}/unit_cost_latest.json")
        action_plan_53.fetch_payload = rest_dump(f"{FIXTURE_DIR}/action_plan_53_latest.json")

        sim_path = "/tmp/ioh_poh_planning_simulated.json"
        if os.path.exists(sim_path):
            with open(sim_path) as f:
                sim_payload = json.load(f)

            def _fetch_ioh_poh():
                return sim_payload, "simulated-fixture"
            ioh_poh_planning.fetch_payload = _fetch_ioh_poh
        else:
            ioh_poh_planning.fetch_payload = lambda: ({}, "empty-fixture")

        def _fetch_all_exceptions():
            with open(f"{FIXTURE_DIR}/exception_reports_latest.json") as f:
                rows = json.load(f)
            return [(r["payload"], r.get("updated_at", "fixture")) for r in rows]
        exception_reports.fetch_all = _fetch_all_exceptions

        empty = lambda: ({}, "no-fixture")
        monthly_machine.fetch_payload = empty
        cumulative_failure.fetch_payload = empty
        shortfall_reasons.fetch_payload = empty

    # ---- structural invariants ----

    def test_output_is_json_serializable_and_was_written(self):
        self.assertTrue(os.path.exists(self.out_path))
        self.assertGreater(os.path.getsize(self.out_path), 1000)

    def test_top_level_keys_present(self):
        for key in ('meta', 'progress', 'shortfall', 'exceptions', 'fleet'):
            self.assertIn(key, self.payload)

    def test_fleet_size_matches_machine_list_and_status_counts(self):
        meta = self.payload['meta']
        machines = self.payload['progress']['machines']
        self.assertEqual(meta['fleetSize'], len(machines))
        self.assertEqual(sum(meta['statusCounts'].values()), meta['fleetSize'])

    def test_rmv_wst_synthetic_fleet_included(self):
        machines = {m['machine'] for m in self.payload['progress']['machines']}
        self.assertIn('RMV-0332', machines)
        self.assertIn('WST 01', machines)

    def test_every_machine_record_has_required_keys(self):
        required = {'type', 'machine', 'div', 'status', 'unit', 'annualTarget',
                    'targetToDate', 'actualToDate', 'pct'}
        for m in self.payload['progress']['machines']:
            self.assertTrue(required.issubset(m.keys()), f"{m.get('machine')} missing keys")

    def test_div_rollup_covers_all_three_divisions(self):
        divs = {r['div'] for r in self.payload['progress']['divRollup']}
        self.assertEqual(divs, {'MYS', 'SBC', 'UBL'})

    def test_gap_53_degrades_to_empty_shortfall(self):
        """Task #53's confirmed-still-open gap: no live source for the
        Daily Progress shortfall-category log. Every shortfall-dependent
        section must degrade to empty/zero, not crash or silently fill
        with wrong data."""
        self.assertEqual(self.payload['shortfall']['totalEntries'], 0)
        self.assertEqual(self.payload['shortfall']['catCounts'], {})
        self.assertEqual(self.payload['shortfall']['srdenBundle'], {})

    def test_no_live_ioh_poh_letter_degrades_current_year_to_empty(self):
        cy = self.payload['fleet']['iohPohSchedule']['currentYear']
        self.assertEqual(cy['totalCount'], 0)
        self.assertEqual(cy['events'], [])
        self.assertIsNone(self.payload['fleet']['cpohPlanning'])

    def test_ioh_poh_future_years_still_populated_from_live_source(self):
        """futureYears comes from ioh_poh_planning, which IS live-synced -
        must NOT degrade just because the currentYear letter has no live
        source."""
        fy = self.payload['fleet']['iohPohSchedule']['futureYears']
        self.assertGreater(fy['totalCount'], 0)

    def test_attention_lists_are_severity_sorted(self):
        rank = dict(high=0, medium=1, low=2)
        for section_path in (
            (self.payload.get('indentPosition') or {}).get('attention', []),
            (self.payload.get('tendersContracts') or {}).get('attention', []),
        ):
            sevs = [a['severity'] for a in section_path]
            self.assertEqual(sevs, sorted(sevs, key=lambda s: rank.get(s, 9)))

    def test_raw_sources_attached_for_for_later_use_tables(self):
        # cumulative_failure/monthly_machine/shortfall_reasons were
        # stubbed to {} in this test's fixtures -- extract_* on an empty
        # payload must still succeed and simply produce an empty shape
        # (never raise), so these keys should all still be present.
        # hq_summary/rdps_points are deliberately excluded from
        # rawSources entirely (01-Nov-2026, see aggregate.py's module
        # docstring) and must NOT reappear here.
        raw = self.payload.get('rawSources') or {}
        for key in ('cumulativeFailure', 'dailyPositions', 'monthlyMachine', 'shortfallReasonsRaw'):
            self.assertIn(key, raw, f"rawSources missing '{key}' - an empty-payload extract_* must not raise")

    def test_every_machine_has_daily_log_key_even_with_no_monthly_machine_data(self):
        # 18-Nov-2026: m['dailyLog'] (feeding the Network Map popup's "last
        # team's progress" and the machine detail drawer's "last 15 days")
        # must degrade to [] per machine, never a missing key or a crash,
        # when monthly_machine has no live data this run (stubbed to {}
        # above) - same "never raise on an empty for-later-use source"
        # discipline as test_raw_sources_attached_for_for_later_use_tables.
        for m in self.payload['progress']['machines']:
            self.assertIn('dailyLog', m)
            self.assertEqual(m['dailyLog'], [])
        for key in ('hqSummary', 'rdpsPoints'):
            self.assertNotIn(key, raw, f"rawSources should no longer include '{key}' (intentionally dropped)")

    def test_unit_cost_year_rollups_consistent(self):
        uc = self.payload.get('unitCost')
        if not uc:
            self.skipTest("unitCost not present in this fixture run")
        for fy, block in uc['years'].items():
            self.assertEqual(block['kpi']['machineCount'], len(block['machines']))

    def test_mtype_rollup_counts_sum_to_fleet_size(self):
        total = sum(r['count'] for r in self.payload['progress']['mtypeRollup'] if not r.get('isSecondaryMetric'))
        self.assertEqual(total, self.payload['meta']['fleetSize'])

    def test_double_shift_duration_flag_consistent(self):
        ds = self.payload['fleet']['doubleShift']
        for r in ds['over12hRecords']:
            self.assertTrue(r['over12h'])

    def test_network_map_degrades_gracefully_without_rdps_credentials(self):
        """Added 13-Nov-2026 alongside sources/network_map.py. This
        sandbox's TestAggregateEndToEnd run has no SUPABASE_URL/ANON_KEY
        and no rdps_points fixture to monkeypatch with, so
        extract_network_map()'s rdps_points.fetch_rows() call must fail
        -- and the run must still complete cleanly with 'networkMap'
        simply OMITTED (not present as an empty/broken shape, and
        certainly not a crash), exactly like every other optional live
        source tested above. The website's own baked-literal NMDATA
        fallback is what covers this case in production."""
        self.assertNotIn('networkMap', self.payload)


if __name__ == '__main__':
    unittest.main(verbosity=2)
