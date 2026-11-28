"""Targeted unit tests for sources/infra_map.py -- the Infra tab's
Sidings & Rest House map + Sr.DEN breakdown, added 26-Nov-2026.

Pure-function tests, fed entirely synthetic sidings_rest_house-shaped and
rdps_points-shaped data (no network, no live fixtures required) --
mirrors this project's other same-day-feature test style (see
test_network_map.py). Covers:

  - decoding 'SrDEN Detailed's live key-naming scheme (TM/RGM/TM (2)/
    RGM (2)/TM (3)/RGM (3)/Available/Sanctioned/Proposed),
  - recovering the one stray data row a division's block-parser
    mis-files under 'total' (empty Label = not a real total row),
  - merging a division's 3 Sr.DEN-breakdown sub-blocks (TM siding/RGM
    siding/RH) back into one row per Sl No,
  - attaching each section's live per-station remark (from SrDEN
    Detailed) and the static supplement's strikeout flag,
  - station geo resolution: exact code, km-suffix-stripped ("SBC (355)"),
    no-space variant ("E UBL"/"EUBL"), and the APPROX_STATIONS fallback
    only applying when RDPS genuinely has no point for that code,
  - map vs. unmapped split (available/sanctioned -> map; proposed-only
    and geo-missing -> left off the map, still counted),
  - KPI figures come from the static supplement (corrected), not the
    live (confirmed-buggy) masterSummary, when a supplement is given,
  - division-level sanctioned-execution stage rollup,
  - sanctioned-depot geo resolution,
  - honest degradation: payload=None -> None; supplement=None -> still
    builds from live data alone.

Run from the swr_pipeline directory:
    python3 -m unittest tests.test_infra_map -v
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "sources"))

from sources import infra_map


def synthetic_rdps_by_div():
    return {
        'UBL': [
            dict(route='UBL-LWR', km=490, distanceM=0, lat=15.40, lon=75.40,
                 stationCode='DWR', stationName='Dharwad', reviewNotes=None),
            dict(route='UBL-LWR', km=503, distanceM=0, lat=15.45, lon=75.45,
                 stationCode='MGD', stationName='Mugad', reviewNotes=None),
            # no RDPS point at all for 'SBC' (code, not division) -- should
            # fall back to network_map.APPROX_STATIONS if present there,
            # else stay genuinely unmapped.
        ],
        'SBC': [
            dict(route='SBC-JTJ', km=355, distanceM=0, lat=12.90, lon=77.60,
                 stationCode='SBC', stationName='Bengaluru City', reviewNotes=None),
        ],
    }


def synthetic_payload():
    """Shaped exactly as sidings_rest_house.fetch_payload() returns it,
    trimmed to one division's worth of 'SrDEN Detailed' and
    'UBL Sr.DEN' data, enough to exercise every code path."""
    return dict(
        masterSummary=dict(asOf='24.11.2026', siding=dict(swr=dict(stations=343)), restHouse=dict(swr=dict())),
        sanctionedSidingStatus=[
            dict(division='UBL', stage='Completed'),
            dict(division='UBL', stage='Track Work Pending'),
        ],
        supplementary={
            'SrDEN Detailed': [
                dict(
                    title='Status of TM Sidings & Rest House in Hubballi Division (As on 24.11.2026)',
                    rows=[
                        dict(**{'S. No': 1, 'Sr.DEN': 'Sr.DEN/HQ', 'Station': 'DWR',
                                'Section (Km)': 'UBL (469) - LWR (525)', 'KM': 490,
                                'TM': 'Available', 'Available': 'Available'}),
                        dict(**{'S. No': 2, 'Sr.DEN': 'Sr.DEN/HQ', 'Station': 'MGD', 'KM': 503,
                                'TM (2)': 'Sanctioned', 'Sanctioned': 'Sanctioned',
                                'Remark': 'MGD - converted to BT siding.'}),
                        # Proposed-only station with no RDPS point at all --
                        # should count toward totals but never appear on the map.
                        dict(**{'S. No': 3, 'Sr.DEN': 'Sr.DEN/HQ', 'Station': 'GNJ', 'KM': 573,
                                'TM (3)': 'Proposed', 'Proposed': 'Proposed'}),
                        # A station whose code carries a stray km-suffix.
                        dict(**{'S. No': 4, 'Sr.DEN': 'Sr.DEN/East', 'Station': 'SBC (355)', 'KM': 355,
                                'RGM': 'Available', 'Available': 'Available'}),
                    ],
                    # The quirk: this division's sheet has no real trailing
                    # Total row, so the block-parser mis-filed one more
                    # genuine data row here -- Label is empty, never a real
                    # division name, which is how this is told apart from an
                    # actual total row.
                    total=dict(Label='', Station='E UBL', KM=5, Remark='RH- Not Feasible',
                               **{'TM (2)': 'Sanctioned', 'Sanctioned': 'Sanctioned'}),
                ),
            ],
            'UBL Sr.DEN': [
                dict(title='Status of TM siding in Hubballi Div',
                     rows=[dict(**{'Sl No': 1, 'Sr.DEN': 'Sr.DEN/HQ', 'Major Section': 'UBL-LWR',
                                   'Section Length': 60, 'No. of stations': 10, 'Siding Required': 2,
                                   'Available': 1, 'Available (2)': 'DWR', 'Sanctioned': 1,
                                   'Sanctioned (2)': 'MGD', 'Proposed': 0, 'Proposed (2)': '-'})]),
                dict(title='Status of RGM siding in Hubballi Div',
                     rows=[dict(**{'Sl No': 1, 'Sr.DEN': 'Sr.DEN/HQ', 'Major Section': 'UBL-LWR',
                                   'Section Length': 60, 'No. of stations': 10, 'Siding Required': 0,
                                   'Available': 0, 'Available (2)': '-', 'Sanctioned': 0,
                                   'Sanctioned (2)': '-', 'Proposed': 0, 'Proposed (2)': '-'})]),
                dict(title='Status of RH in Hubballi Div',
                     rows=[dict(**{'Sl No': 1, 'Sr.DEN': 'Sr.DEN/HQ', 'Major Section': 'UBL-LWR',
                                   'Section Length': 60, 'No. of stations': 10, 'RH Required': 1,
                                   'Available': 1, 'Available (2)': 'DWR', 'Sanctioned': 0,
                                   'Sanctioned (2)': '-', 'Proposed': 0, 'Proposed (2)': '-'})]),
            ],
        },
    )


def synthetic_supplement():
    return dict(
        masterSummaryCorrected=dict(
            asOf='24.11.2026',
            siding=dict(swr=dict(stations=343, required=180, available=98, sanctioned=54, proposed=28)),
            restHouse=dict(swr=dict(stations=343, required=188, available=63, sanctioned=65, proposed=60)),
        ),
        rbTargets=dict(stablingSiding=dict(required=106, available=52, target=15),
                       satDepot=dict(required=3, available=1, target=2)),
        strikeouts={'UBL': ['MGD']},
        rhSanctionedStatus=[
            dict(division='UBL', workStatus='Slab has been casted.'),
            dict(division='UBL', workStatus='Foundation work completed.'),
        ],
        sanctionedDepots=[
            dict(slNo=1, description='New Satellite TM Siding Depot at Hubballi', station='DWR',
                 division='UBL', srdenKey='UBL Sr.DEN/HQ', status='Sanctioned.'),
        ],
    )


class TestFeatureStatus(unittest.TestCase):
    """_feature_status() priority within ONE feature (never blended
    across features -- that's covered by TestBuildInfraMap instead)."""

    def test_available_wins(self):
        self.assertEqual(infra_map._feature_status(True, True, True), 'available')

    def test_sanctioned_when_not_available(self):
        self.assertEqual(infra_map._feature_status(False, True, True), 'sanctioned')

    def test_proposed_when_only_proposed(self):
        self.assertEqual(infra_map._feature_status(False, False, True), 'proposed')

    def test_none_when_nothing_set(self):
        self.assertIsNone(infra_map._feature_status(False, False, False))


class TestStationCodeAlias(unittest.TestCase):
    """Added 27-Nov-2026 per Shruti: 'E UBL is BMAP' -- a manual code
    alias so this station resolves to a real RDPS position instead of
    staying permanently unmapped."""

    def test_e_ubl_resolves_via_bmap_alias(self):
        rdps = synthetic_rdps_by_div()
        rdps['UBL'].append(dict(route='UBL-LWR', km=5, distanceM=0, lat=15.33, lon=75.12,
                                 stationCode='BMAP', stationName='Bellary Main Area Point', reviewNotes=None))
        stations_geo = infra_map._build_stations_geo(rdps)
        geo = infra_map._resolve_geo('E UBL', stations_geo)
        self.assertIsNotNone(geo)
        self.assertEqual(geo['lat'], 15.33)

    def test_without_bmap_point_still_honestly_unmapped(self):
        # No BMAP point in the plain fixture -- confirms the alias doesn't
        # invent a position, it just gives the real one another name to
        # match against when one exists.
        stations_geo = infra_map._build_stations_geo(synthetic_rdps_by_div())
        self.assertIsNone(infra_map._resolve_geo('E UBL', stations_geo))


class TestSrdenDetailedDecoding(unittest.TestCase):
    """The live generic block-parser's key-naming scheme for this sheet."""

    def setUp(self):
        self.stations_geo = infra_map._build_stations_geo(synthetic_rdps_by_div())
        self.rows = infra_map._parse_srden_detailed(synthetic_payload(), self.stations_geo)
        self.by_station = {r['station']: r for r in self.rows}

    def test_tm_available_decoded(self):
        self.assertTrue(self.by_station['DWR']['sidingAvailTM'])
        self.assertFalse(self.by_station['DWR']['sidingSancTM'])

    def test_tm_sanctioned_decoded_from_suffix_2(self):
        self.assertTrue(self.by_station['MGD']['sidingSancTM'])
        self.assertFalse(self.by_station['MGD']['sidingAvailTM'])
        self.assertEqual(self.by_station['MGD']['remark'], 'MGD - converted to BT siding.')

    def test_tm_proposed_decoded_from_suffix_3(self):
        self.assertTrue(self.by_station['GNJ']['sidingPropTM'])

    def test_rgm_available_decoded(self):
        # station code carried a stray km-suffix in the raw 'Station' field
        self.assertIn('SBC', self.by_station)
        self.assertTrue(self.by_station['SBC']['sidingAvailRGM'])

    def test_bare_available_is_rest_house(self):
        self.assertTrue(self.by_station['DWR']['rhAvail'])

    def test_stray_total_row_recovered_as_real_station(self):
        """E UBL was mis-filed under this division's 'total' key (empty
        Label, unlike a real 'X Total' row) -- it must still show up as
        a genuine station with its own remark, not be silently dropped."""
        self.assertIn('E UBL', self.by_station)
        self.assertEqual(self.by_station['E UBL']['remark'], 'RH- Not Feasible')

    def test_proposed_only_station_has_no_geo_by_design(self):
        # GNJ was never given an RDPS point in the fixture -- confirms
        # the parser doesn't invent a coordinate.
        self.assertIsNone(self.by_station['GNJ']['lat'])


class TestStationGeoResolution(unittest.TestCase):
    def setUp(self):
        self.stations_geo = infra_map._build_stations_geo(synthetic_rdps_by_div())

    def test_exact_code_match(self):
        self.assertEqual(infra_map._resolve_geo('DWR', self.stations_geo)['lat'], 15.40)

    def test_km_suffix_stripped(self):
        geo = infra_map._resolve_geo('SBC (355)', self.stations_geo)
        self.assertIsNotNone(geo)
        self.assertEqual(geo['lat'], 12.90)

    def test_no_space_variant(self):
        # 'E UBL' has no RDPS point and isn't in APPROX_STATIONS either in
        # this synthetic fixture -- confirms an honest miss, not a guess.
        self.assertIsNone(infra_map._resolve_geo('E UBL', self.stations_geo))

    def test_unresolvable_code_is_none_not_guessed(self):
        self.assertIsNone(infra_map._resolve_geo('ZZZZZ', self.stations_geo))


class TestSrdenBreakdownMerge(unittest.TestCase):
    def test_three_subblocks_merge_by_sl_no(self):
        payload = synthetic_payload()
        stations_geo = infra_map._build_stations_geo(synthetic_rdps_by_div())
        station_rows = infra_map._parse_srden_detailed(payload, stations_geo)
        remark_by_station = {r['station']: r['remark'] for r in station_rows if r.get('remark')}
        strikeouts = {'UBL': {'MGD'}}
        breakdown = infra_map._parse_srden_breakdown(payload, remark_by_station, strikeouts)
        self.assertEqual(len(breakdown['UBL']), 1)
        section = breakdown['UBL'][0]
        self.assertEqual(section['majorSection'], 'UBL-LWR')
        self.assertIsNotNone(section['tm'])
        self.assertIsNotNone(section['rgm'])
        self.assertIsNotNone(section['rh'])
        self.assertEqual(section['tm']['availableStations'], ['DWR'])
        self.assertEqual(section['tm']['sanctionedStations'], ['MGD'])

    def test_remark_attached_from_srden_detailed(self):
        payload = synthetic_payload()
        stations_geo = infra_map._build_stations_geo(synthetic_rdps_by_div())
        station_rows = infra_map._parse_srden_detailed(payload, stations_geo)
        remark_by_station = {r['station']: r['remark'] for r in station_rows if r.get('remark')}
        breakdown = infra_map._parse_srden_breakdown(payload, remark_by_station, {'UBL': set()})
        # MGD's remark came from the live per-station SrDEN Detailed row,
        # not from any Sr.DEN-sheet column (which doesn't carry one live).
        self.assertIn('MGD', breakdown['UBL'][0]['tm']['remark'])

    def test_strikeout_flag_attached(self):
        payload = synthetic_payload()
        stations_geo = infra_map._build_stations_geo(synthetic_rdps_by_div())
        station_rows = infra_map._parse_srden_detailed(payload, stations_geo)
        remark_by_station = {r['station']: r['remark'] for r in station_rows if r.get('remark')}
        breakdown = infra_map._parse_srden_breakdown(payload, remark_by_station, {'UBL': {'MGD'}})
        self.assertIn('MGD', breakdown['UBL'][0]['tm']['struckStations'])


class TestBuildInfraMap(unittest.TestCase):
    def setUp(self):
        self.result = infra_map.build_infra_map(
            synthetic_payload(), synthetic_rdps_by_div(), synthetic_supplement())

    def test_none_payload_degrades_to_none(self):
        self.assertIsNone(infra_map.build_infra_map(None, synthetic_rdps_by_div(), synthetic_supplement()))

    def test_missing_supplement_still_builds_from_live_data(self):
        result = infra_map.build_infra_map(synthetic_payload(), synthetic_rdps_by_div(), None)
        self.assertIsNotNone(result)
        # falls back to the live (possibly wrong) masterSummary rather than crashing
        self.assertEqual(result['kpi']['stationsCovered'], 343)

    def test_kpi_uses_corrected_supplement_not_live_master_summary(self):
        # live masterSummary.restHouse.swr is an empty dict in the fixture
        # (standing in for the confirmed-buggy live figure); the
        # supplement's corrected value must be what's actually used.
        self.assertEqual(self.result['kpi']['restHouse']['available'], 63)

    def test_rb_targets_surfaced(self):
        self.assertEqual(self.result['kpi']['sidingTargetRB']['target'], 15)
        self.assertEqual(self.result['kpi']['satDepotRB']['target'], 2)

    def test_available_and_sanctioned_stations_are_mapped(self):
        codes = {s['station'] for s in self.result['mapStations']}
        self.assertIn('DWR', codes)
        self.assertIn('MGD', codes)

    def test_proposed_only_station_now_counted_not_dropped(self):
        # Fixed 27-Nov-2026 per Shruti ("the proposed ones you have shown
        # as zero, which is incorrect"): a Proposed-only entry (no
        # Available/Sanctioned flag at all) used to be silently excluded
        # entirely -- not even counted as unmapped. It's a real pipeline
        # stage now, so GNJ (siding Proposed-only, no RDPS point in this
        # fixture) must show up in unmappedStations, not vanish.
        unmapped_codes = {s['station'] for s in self.result['unmappedStations']}
        self.assertIn('GNJ', unmapped_codes)
        gnj = next(s for s in self.result['unmappedStations'] if s['station'] == 'GNJ')
        self.assertEqual(gnj['sidingStatus'], 'proposed')
        self.assertEqual(gnj['rhStatus'], 'proposed')

    def test_available_station_with_no_geo_is_unmapped_not_dropped(self):
        unmapped_codes = {s['station'] for s in self.result['unmappedStations']}
        self.assertIn('E UBL', unmapped_codes)

    def test_division_rollup_stage_counts(self):
        rollup = self.result['divisionRollup']['UBL']
        self.assertEqual(rollup['sidingExecution']['total'], 2)
        self.assertEqual(rollup['sidingExecution']['byStage']['Completed'], 1)
        self.assertEqual(rollup['rhExecution']['total'], 2)

    def test_depot_geo_resolved(self):
        depot = self.result['depots'][0]
        self.assertEqual(depot['station'], 'DWR')
        self.assertEqual(depot['lat'], 15.40)

    def test_siding_and_rh_status_computed_independently(self):
        # Added 27-Nov-2026 per Shruti: "sanctioned doesn't mean it is
        # available... we need not merge both of these things." Each
        # station's siding status and rest-house status must each come
        # from only that feature's own flags.
        by_station = {s['station']: s for s in self.result['mapStations']}
        self.assertEqual(by_station['DWR']['sidingStatus'], 'available')
        self.assertEqual(by_station['DWR']['rhStatus'], 'available')
        self.assertTrue(by_station['DWR']['hasSiding'])
        self.assertTrue(by_station['DWR']['hasRH'])
        self.assertEqual(by_station['MGD']['sidingStatus'], 'sanctioned')
        self.assertEqual(by_station['MGD']['rhStatus'], 'sanctioned')
        self.assertEqual(by_station['SBC']['sidingStatus'], 'available')
        self.assertEqual(by_station['SBC']['rhStatus'], 'available')

    def test_tracks_by_div_reuses_network_map_builder(self):
        # Added 26-Nov-2026 per Shruti's feature-parity request: the Infra
        # map must show the same route polylines as the Overview tab's
        # Network Map, built from the exact same live rdps_by_div input
        # this module already receives -- no second network round trip.
        from sources import network_map
        expected = network_map.build_tracks_from_rdps(synthetic_rdps_by_div())
        self.assertEqual(self.result['tracksByDiv'], expected)
        self.assertIn('UBL', self.result['tracksByDiv'])
        self.assertIn('SBC', self.result['tracksByDiv'])

    def test_div_color_surfaced_for_frontend_reuse(self):
        from sources import network_map
        self.assertEqual(self.result['divColor'], network_map.DIV_COLOR)
        self.assertEqual(self.result['divColor']['UBL'], '#0284C7')


if __name__ == '__main__':
    unittest.main()


class TestSheetTextInStationCells(unittest.TestCase):
    """28-Nov-2026: 8 Mysuru Sr.DEN/West 'Station' cells in SrDEN Detailed
    were overwritten with sentences of pasted text, which then appeared on
    the site as fake stations in "Not mapped yet". They must be left out
    of the stations and reported as rows to correct instead."""
    def _payload(self):
        return {'supplementary': {'SrDEN Detailed': [{
            'title': 'Status of TM Sidings & Rest House in Mysuru Division (As on 26.11.2026)',
            'rows': [
                {'Sr.DEN': 'Sr.DEN/West', 'Section (Km)': 'HLN (87)-MYS (0)', 'Station': 'CMNR', 'Proposed': 'Proposed'},
                {'Sr.DEN': 'Sr.DEN/West', 'Station': 'It is looking good now, but there is something I want to explain to you:', 'TM (2)': 'Sanctioned'},
                {'Sr.DEN': 'Sr.DEN/West', 'Station': 'KRYP  KKMI', 'TM (3)': 'Proposed'},
            ]}]}}

    def test_sentence_is_not_a_station(self):
        codes = [r['station'] for r in infra_map._parse_srden_detailed(self._payload(), {})]
        self.assertEqual(codes, ['CMNR', 'KRYP  KKMI'])

    def test_sentence_is_reported(self):
        issues = infra_map.sheet_issues(self._payload())
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0]['division'], 'MYS')
        self.assertTrue(issues[0]['text'].startswith('It is looking good'))


class TestStationRowsFromSrdenDetailed(unittest.TestCase):
    """28-Nov-2026: the old 'CE TM Sr DEN detail' tab syncs no rows any
    more, which left Stations Covered = 0, Pending - CN Units = 0 and an
    empty Sr.DEN table. Station rows now come from 'SrDEN Detailed'."""
    def test_fallback_builds_station_rows_and_rollups(self):
        from sources import sidings_rest_house as sr
        payload = {'masterSummary': {}, 'ceTmSrDenDetail': [], 'supplementary': {'SrDEN Detailed': [{
            'title': 'Status of TM Sidings & Rest House in Hubballi Division (As on 26.11.2026)',
            'rows': [
                {'Sr.DEN': 'Sr.DEN/East', 'Section (Km)': 'A-B', 'Station': 'KJDI', 'TM (2)': 'Sanctioned', 'Proposed': 'Proposed', 'Remark': 'To be done by CN Org.'},
                {'Sr.DEN': 'Sr.DEN/East', 'Station': 'GDG', 'TM': 'TM', 'RGM': 'RGM', 'Available': 'Available'},
                {'Sr.DEN': 'Sr.DEN/East', 'Station': 'XYZ'},   # no entry at all -> not counted
            ]}]}}
        out = sr.extract_sidings_rest_houses(payload)
        rows = {r['station']: r for r in out['stationRows']}
        self.assertEqual(set(rows), {'KJDI', 'GDG'})
        self.assertEqual((rows['KJDI']['sidingStatus'], rows['KJDI']['rhStatus']), ('Sanctioned', 'Proposed'))
        self.assertEqual(rows['GDG']['sidingType'], 'TM + RGM')
        self.assertEqual(rows['KJDI']['asOf'], '26.11.2026')
        self.assertEqual(out['responsibility']['overall'], {'CN (Construction Org)': 1})
        self.assertEqual(out['srdenTable'][0]['stations'], 2)
