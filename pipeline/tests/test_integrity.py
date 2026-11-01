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
    exception_reports, hq_summary, indent_position, ioh_poh_planning, monthly_machine,
    rdps_points, shortfall_reasons, sidings_rest_house, tenders_contracts, unit_cost,
    action_plan_53, cumulative_failure,
)

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
        hq_summary.fetch_payload = empty
        monthly_machine.fetch_payload = empty
        cumulative_failure.fetch_payload = empty
        shortfall_reasons.fetch_payload = empty
        rdps_points.fetch_rows = lambda **kw: []

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
        # cumulative_failure/hq_summary/monthly_machine/shortfall_reasons
        # were stubbed to {} in this test's fixtures -- extract_* on an
        # empty payload must still succeed and simply produce an empty
        # shape (never raise), so these keys should all still be present.
        raw = self.payload.get('rawSources') or {}
        for key in ('cumulativeFailure', 'dailyPositions', 'hqSummary', 'monthlyMachine', 'shortfallReasonsRaw'):
            self.assertIn(key, raw, f"rawSources missing '{key}' - an empty-payload extract_* must not raise")

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
            self.assertGreater(r['durationHours'], 12)


if __name__ == '__main__':
    unittest.main(verbosity=2)
