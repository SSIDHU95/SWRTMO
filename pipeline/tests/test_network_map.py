"""Targeted unit tests for sources/network_map.py -- the live SWR
Network Map payload builder added 13-Nov-2026.

Pure-function tests, fed entirely synthetic rdps_points-shaped and
daily_positions-shaped data (no network, no fixtures required) --
mirrors how this project proved its other same-day fixes (see
test_integrity.py's TestPipelineHelpers). Covers:

  - station grouping/dedup from rdps_points rows (one entry per
    stationCode, first occurrence wins on a repeat),
  - track-polyline grouping by (division, route), sorted by km,
  - the hand-curated DEPOTS/APPROX_STATIONS supplement merging in with
    the right confidence tiers and not clobbering real RDPS stations,
  - machine-to-station matching (exact code, then each '/'-split part),
  - end-to-end build_network_map(): placed machines land under the
    right station with only the live-shape fields, unmatched machines
    land in `unmatched` (never silently dropped), and source-sheet
    header-row leakage ("Machine" / "STABLED AT") is filtered out.

Run from the swr_pipeline directory:
    python3 -m unittest tests.test_network_map -v
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "sources"))

from sources import network_map
import rdps_points


# ---------------------------------------------------------------------
# Synthetic fixtures
# ---------------------------------------------------------------------
def synthetic_rdps_rows():
    """Raw rows, as rdps_points.fetch_rows() would return them (raw
    snake_case keys) -- two divisions, two routes in UBL (so track
    grouping has something to group), one repeated stationCode (to
    exercise the dedup-first-occurrence rule)."""
    return [
        # UBL, route A, 3 points in descending km order (tests sorting).
        # Kept close together (~150m apart) -- not realistic chainage
        # spacing (real survey points are ~10-100m apart; see
        # GAP_THRESHOLD_KM's comment in network_map.py), but the point
        # here is just to stay well UNDER that 10km gap-split threshold
        # so this fixture exercises sorting, not splitting. AAA's
        # lat/lon must stay exactly 15.10/75.10 -- TestStationBuilding's
        # test_first_occurrence_wins_on_a_repeated_code asserts it.
        dict(division='UBL', route='UBL-MRJ DN', km=10, distance_m=0,
             lat=15.10, lon=75.10, station_code='AAA', station_name='Alpha', review_notes=None),
        dict(division='UBL', route='UBL-MRJ DN', km=0, distance_m=0,
             lat=15.099, lon=75.099, station_code='BBB', station_name='Beta', review_notes=None),
        dict(division='UBL', route='UBL-MRJ DN', km=20, distance_m=0,
             lat=15.101, lon=75.101, station_code=None, station_name=None, review_notes=None),
        # UBL, route B
        dict(division='UBL', route='UBL-LWR', km=5, distance_m=0,
             lat=15.50, lon=75.50, station_code='CCC', station_name='Gamma', review_notes=None),
        # a repeated stationCode on a different route (second occurrence
        # carries a deliberately different, "wrong" coordinate so the
        # test can confirm it's ignored)
        dict(division='UBL', route='UBL-LWR', km=6, distance_m=0,
             lat=99.0, lon=99.0, station_code='AAA', station_name='Alpha-again', review_notes=None),
        # MYS division, one route/one station
        dict(division='MYS', route='MYS-ASK', km=1, distance_m=0,
             lat=12.30, lon=76.60, station_code='DDD', station_name='Delta', review_notes=None),
    ]


def synthetic_positions_payload():
    """As daily_positions.fetch_payload()'s payload element (the dict
    BEFORE extract_positions() reshapes it) -- includes a header-leakage
    row, a slash-coded stabling, an exact match, and a genuinely
    unmatched machine."""
    return dict(asOfDate='13.11.2026', machineCount=6, machines=[
        dict(machine='Machine', div='', stabledAt='STABLED AT'),       # header leakage -> must be dropped
        dict(machine='BCM 001', div='UBL', stabledAt='AAA',             # exact match, WITH detail fields
             blockSection='AAA-BBB', majorSection='UBL-MRJ', durationMinutes=105.0,
             actualTotal=0.7, remarks='Night: 700 mtrs tamped.'),
        dict(machine='DUO 002', div='UBL', stabledAt='BBB/CCC'),        # slash-coded, BBB should win (first part)
        dict(machine='MPT 003', div='MYS', stabledAt='DDD'),            # exact match, different division
        dict(machine='UTV 004', div='SBC', stabledAt='ZZZ-UNKNOWN'),    # no station anywhere -> unmatched
        dict(machine='RMV 005', div='SBC', stabledAt=None),             # no code at all -> unmatched, reason no_code
    ])


def synthetic_srden_payload():
    """As daily_progress_srden.fetch_payload()'s payload element (the
    dict BEFORE extract_srden() reshapes it) -- see that module's own
    docstring for the exact shape. Covers BCM 001 and MPT 003 (both
    placed machines in synthetic_positions_payload above) and
    deliberately leaves DUO 002 out, to exercise the lookup-miss path
    (no fabricated 'srDen')."""
    return dict(asOfDate='13.11.2026', machines=[
        dict(machine='BCM 001', division='UBL', srDen='OL Sr.DEN.W.'),
        dict(machine='MPT 003', division='MYS', srDen='OL Sr.DEN.N.'),
    ])


class TestStationBuilding(unittest.TestCase):

    def setUp(self):
        self.rdps_by_div = rdps_points.extract_points(synthetic_rdps_rows())

    def test_one_entry_per_distinct_station_code(self):
        stations = network_map.build_stations_from_rdps(self.rdps_by_div)
        self.assertEqual(set(stations.keys()), {'AAA', 'BBB', 'CCC', 'DDD'})

    def test_rows_without_a_station_code_are_skipped(self):
        stations = network_map.build_stations_from_rdps(self.rdps_by_div)
        self.assertEqual(len(stations), 4)  # the 5th UBL-MRJ DN point (no code) contributes nothing

    def test_first_occurrence_wins_on_a_repeated_code(self):
        stations = network_map.build_stations_from_rdps(self.rdps_by_div)
        # AAA appears twice (lat 15.10/75.10 then the "wrong" 99.0/99.0) --
        # the first (real) coordinate must win, not the later one.
        self.assertAlmostEqual(stations['AAA']['lat'], 15.10)
        self.assertAlmostEqual(stations['AAA']['lon'], 75.10)

    def test_every_station_tagged_surveyed_and_rdps_station(self):
        stations = network_map.build_stations_from_rdps(self.rdps_by_div)
        for s in stations.values():
            self.assertEqual(s['confidence'], 'surveyed')
            self.assertEqual(s['kind'], 'rdps_station')

    def test_division_carried_through_per_station(self):
        stations = network_map.build_stations_from_rdps(self.rdps_by_div)
        self.assertEqual(stations['AAA']['div'], 'UBL')
        self.assertEqual(stations['DDD']['div'], 'MYS')


class TestTrackBuilding(unittest.TestCase):

    def setUp(self):
        self.rdps_by_div = rdps_points.extract_points(synthetic_rdps_rows())

    def test_grouped_by_division_then_route(self):
        tracks = network_map.build_tracks_from_rdps(self.rdps_by_div)
        self.assertEqual(set(tracks.keys()), {'UBL', 'MYS'})
        ubl_sheets = {s['sheet'] for s in tracks['UBL']}
        self.assertEqual(ubl_sheets, {'UBL-MRJ DN', 'UBL-LWR'})

    def test_points_sorted_by_km_ascending(self):
        tracks = network_map.build_tracks_from_rdps(self.rdps_by_div)
        dn_sheet = next(s for s in tracks['UBL'] if s['sheet'] == 'UBL-MRJ DN')
        # source rows were km 10, 0, 20 in that order -- must come out 0,10,20
        self.assertEqual(dn_sheet['pts'], [[15.099, 75.099], [15.10, 75.10], [15.101, 75.101]])

    def test_rows_missing_lat_lon_are_excluded_from_tracks(self):
        rdps_by_div = {'UBL': [dict(route='R1', km=1, distanceM=0, lat=None, lon=None,
                                     stationCode=None, stationName=None, reviewNotes=None)]}
        tracks = network_map.build_tracks_from_rdps(rdps_by_div)
        self.assertEqual(tracks['UBL'], [])


class TestTrackDecimation(unittest.TestCase):
    """Fixed 14-Nov-2026: a route with far more points than
    MAX_POINTS_PER_ROUTE must be thinned down, not embedded whole --
    see sources/network_map.py's 14-Nov-2026 docstring entry for why
    (this is what was ballooning `networkMap` to ~14.5 MB and breaking
    the live Supabase push with a statement timeout, once
    rdps_points.fetch_rows() started returning the FULL ~430,370-row
    survey set instead of a silently-truncated 1000)."""

    def _route_with_n_points(self, n):
        # Tight, realistic spacing (~11m lat-step) so this fixture tests
        # decimation alone, with no gap large enough to trigger the
        # 18-Nov-2026 gap-splitting fix (GAP_THRESHOLD_KM=10) -- that
        # behavior has its own TestTrackGapSplitting below.
        return {'UBL': [dict(route='R1', km=i, distanceM=0,
                              lat=15.0 + i * 0.0001, lon=75.0 + i * 0.0002,
                              stationCode=None, stationName=None, reviewNotes=None)
                         for i in range(n)]}

    def test_route_under_the_cap_is_left_untouched(self):
        rdps_by_div = self._route_with_n_points(50)
        tracks = network_map.build_tracks_from_rdps(rdps_by_div)
        self.assertEqual(len(tracks['UBL'][0]['pts']), 50)

    def test_route_over_the_cap_is_thinned_to_the_cap(self):
        rdps_by_div = self._route_with_n_points(50000)
        tracks = network_map.build_tracks_from_rdps(rdps_by_div)
        pts = tracks['UBL'][0]['pts']
        self.assertEqual(len(pts), network_map.MAX_POINTS_PER_ROUTE)

    def test_thinned_route_keeps_its_first_and_last_point(self):
        rdps_by_div = self._route_with_n_points(50000)
        tracks = network_map.build_tracks_from_rdps(rdps_by_div)
        pts = tracks['UBL'][0]['pts']
        self.assertEqual(pts[0], [15.0, 75.0])
        self.assertAlmostEqual(pts[-1][0], 15.0 + 49999 * 0.0001)
        self.assertAlmostEqual(pts[-1][1], 75.0 + 49999 * 0.0002)

    def test_decimate_helper_never_exceeds_max_points(self):
        rows = [{'v': i} for i in range(12345)]
        thinned = network_map._decimate(rows, max_points=37)
        self.assertEqual(len(thinned), 37)
        self.assertEqual(thinned[0]['v'], 0)
        self.assertEqual(thinned[-1]['v'], 12344)

    def test_decimate_helper_is_a_no_op_under_the_cap(self):
        rows = [{'v': i} for i in range(5)]
        thinned = network_map._decimate(rows, max_points=400)
        self.assertEqual(thinned, rows)


class TestTrackGapSplitting(unittest.TestCase):
    """Fixed 18-Nov-2026, reported directly by Shruti off the live map:
    a genuine gap in a route's surveyed chainage (a stretch that was
    simply never surveyed/synced between two km values that ARE
    recorded) used to be drawn as one long, dead-straight line jumping
    clear across unrelated geography -- e.g. her screenshots of a line
    from the Chikjajur/Davangere area all the way to Chikmagalur/
    Tumakuru (MYS "DRU-CMGR"), and another near Jolarpet/Malur (SBC
    "SBC-JTJ DN"). Confirmed directly against the live rdps_points
    table: normal point-to-point spacing is a few metres (median
    ~10m), with a long tail of legitimate sparse-but-real stretches
    climbing gradually up to 16.4km -- then a dead zone with NOT ONE
    gap anywhere between 16.4km and 51.2km, then exactly 3 genuine
    breaks far past it (51.2km, 91.8km, 152.2km). GAP_THRESHOLD_KM=20
    sits in the middle of that dead zone. See network_map.py's
    GAP_THRESHOLD_KM comment for the full writeup."""

    def _pt(self, route, km, lat, lon):
        return dict(route=route, km=km, distanceM=0, lat=lat, lon=lon,
                     stationCode=None, stationName=None, reviewNotes=None)

    def test_haversine_matches_known_one_degree_latitude_distance(self):
        # 1 degree of latitude is ~111.2km anywhere on Earth.
        d = network_map._haversine_km(0.0, 0.0, 1.0, 0.0)
        self.assertAlmostEqual(d, 111.2, delta=0.5)

    def test_route_with_no_large_gap_stays_a_single_segment(self):
        # Every consecutive pair here is ~110m apart -- well under the
        # 20km threshold.
        rdps_by_div = {'UBL': [self._pt('R1', i, 15.0 + i * 0.001, 75.0) for i in range(5)]}
        tracks = network_map.build_tracks_from_rdps(rdps_by_div)
        self.assertEqual(len(tracks['UBL']), 1)
        self.assertEqual(len(tracks['UBL'][0]['pts']), 5)

    def test_gap_just_under_threshold_is_not_split(self):
        # ~16.4km gap -- right at the top of the legitimate sparse-tail
        # the live data showed (UBL "GDG-HG UP" and others have real,
        # honest gaps this big), which this fix must leave alone. Must
        # NOT split.
        rdps_by_div = {'UBL': [
            self._pt('R1', 0, 15.00, 75.0),
            self._pt('R1', 1, 15.147, 75.0),   # ~16.3km from the point above
        ]}
        tracks = network_map.build_tracks_from_rdps(rdps_by_div)
        self.assertEqual(len(tracks['UBL']), 1)
        self.assertEqual(len(tracks['UBL'][0]['pts']), 2)

    def test_one_large_gap_splits_route_into_two_segments(self):
        # Two tight 2-point clusters ~55km apart (mirrors the real
        # DRU-CMGR/SBC-JTJ DN breaks, just a smaller gap for a fast
        # test) -- must come back as 2 separate {"sheet": "R1", ...}
        # entries, never one line connecting across the 55km jump.
        rdps_by_div = {'MYS': [
            self._pt('R1', 0, 15.000, 75.0),
            self._pt('R1', 1, 15.001, 75.0),   # ~110m from the point above
            self._pt('R1', 2, 15.500, 75.0),   # ~55.5km from the point above -- THE gap
            self._pt('R1', 3, 15.501, 75.0),   # ~110m from the point above
        ]}
        tracks = network_map.build_tracks_from_rdps(rdps_by_div)
        segments = tracks['MYS']
        self.assertEqual(len(segments), 2)
        self.assertTrue(all(s['sheet'] == 'R1' for s in segments))
        self.assertEqual(segments[0]['pts'], [[15.000, 75.0], [15.001, 75.0]])
        self.assertEqual(segments[1]['pts'], [[15.500, 75.0], [15.501, 75.0]])

    def test_two_large_gaps_split_route_into_three_segments(self):
        rdps_by_div = {'MYS': [
            self._pt('R1', 0, 15.0, 75.0),
            self._pt('R1', 1, 15.5, 75.0),   # gap 1 (~55.5km)
            self._pt('R1', 2, 16.0, 75.0),   # gap 2 (~55.5km)
        ]}
        tracks = network_map.build_tracks_from_rdps(rdps_by_div)
        self.assertEqual(len(tracks['MYS']), 3)
        self.assertEqual([s['pts'] for s in tracks['MYS']],
                          [[[15.0, 75.0]], [[15.5, 75.0]], [[16.0, 75.0]]])

    def test_each_segment_is_decimated_independently(self):
        # A small segment (5 points) and a huge one (50000 points),
        # separated by one big gap -- the small segment must come back
        # untouched while only the huge one gets capped, proving
        # decimation runs AFTER splitting, per segment, not on the
        # whole unsplit route (which would have let the huge segment's
        # own decimate() stride pick an anchor point from the small
        # segment across the gap).
        small = [self._pt('R1', i, 15.0 + i * 0.0001, 75.0) for i in range(5)]
        huge = [self._pt('R1', 100 + i, 20.0 + i * 0.0001, 80.0) for i in range(50000)]
        rdps_by_div = {'UBL': small + huge}
        tracks = network_map.build_tracks_from_rdps(rdps_by_div)
        self.assertEqual(len(tracks['UBL']), 2)
        self.assertEqual(len(tracks['UBL'][0]['pts']), 5)
        self.assertEqual(len(tracks['UBL'][1]['pts']), network_map.MAX_POINTS_PER_ROUTE)


class TestSupplementMerge(unittest.TestCase):

    def test_depots_added_with_correct_confidence_tiers(self):
        merged = network_map.merge_supplement_stations({})
        self.assertEqual(merged['CPOH/RYP']['confidence'], 'approximate')
        self.assertEqual(merged['STMRD/UBL']['confidence'], 'confirmed')
        self.assertEqual(merged['ZBD/YPR']['confidence'], 'confirmed')
        for code in ('CPOH/RYP', 'STMRD/UBL', 'ZBD/YPR'):
            self.assertEqual(merged[code]['kind'], 'depot')
            self.assertEqual(merged[code]['div'], 'SHARED')

    def test_approx_stations_added_with_their_own_confidence(self):
        merged = network_map.merge_supplement_stations({})
        self.assertEqual(merged['HPT']['confidence'], 'confirmed')
        self.assertEqual(merged['NAN']['confidence'], 'approximate')
        self.assertIsNone(merged['NAN']['div'])  # filled in per-machine, not here
        self.assertEqual(merged['NAN']['kind'], 'web_geocoded')

    def test_real_rdps_stations_are_preserved_alongside_the_supplement(self):
        live = {'AAA': dict(lat=1, lon=2, code='AAA', div='UBL', confidence='surveyed',
                             kind='rdps_station', depotLabel=None, source='x')}
        merged = network_map.merge_supplement_stations(live)
        self.assertIn('AAA', merged)
        self.assertIn('HPT', merged)
        self.assertIn('STMRD/UBL', merged)
        self.assertEqual(len(merged), 1 + len(network_map.DEPOTS) + len(network_map.APPROX_STATIONS))


class TestMachineStationMatching(unittest.TestCase):

    def setUp(self):
        self.stations = {'AAA': {'confidence': 'surveyed'}, 'BBB': {'confidence': 'surveyed'}}

    def test_exact_match(self):
        code, info = network_map.match_machine_to_station('AAA', self.stations)
        self.assertEqual(code, 'AAA')

    def test_slash_split_first_part_matches(self):
        code, info = network_map.match_machine_to_station('AAA/XXX', self.stations)
        self.assertEqual(code, 'AAA')

    def test_slash_split_second_part_matches(self):
        code, info = network_map.match_machine_to_station('XXX/BBB', self.stations)
        self.assertEqual(code, 'BBB')

    def test_no_match_returns_none_none(self):
        code, info = network_map.match_machine_to_station('ZZZ', self.stations)
        self.assertIsNone(code)
        self.assertIsNone(info)

    def test_empty_or_none_code_returns_none_none(self):
        self.assertEqual(network_map.match_machine_to_station(None, self.stations), (None, None))
        self.assertEqual(network_map.match_machine_to_station('', self.stations), (None, None))


class TestBuildNetworkMapEndToEnd(unittest.TestCase):

    def setUp(self):
        self.result = network_map.build_network_map(
            synthetic_rdps_rows(), synthetic_positions_payload())

    def test_header_leakage_row_is_dropped_entirely(self):
        all_machines = []
        for s in self.result['stations'].values():
            all_machines.extend(m['machine'] for m in s['machines'])
        all_machines.extend(m['machine'] for m in self.result['unmatched'])
        self.assertNotIn('Machine', all_machines)

    def test_exact_match_places_machine_at_right_station(self):
        self.assertIn('AAA', self.result['stations'])
        placed = [m['machine'] for m in self.result['stations']['AAA']['machines']]
        self.assertIn('BCM 001', placed)

    def test_slash_coded_stabling_places_machine_at_first_matching_part(self):
        self.assertIn('BBB', self.result['stations'])
        placed = [m['machine'] for m in self.result['stations']['BBB']['machines']]
        self.assertIn('DUO 002', placed)
        # and NOT under CCC, since BBB (the first slash part) matched first
        ccc_placed = [m['machine'] for m in self.result['stations'].get('CCC', {}).get('machines', [])]
        self.assertNotIn('DUO 002', ccc_placed)

    def test_unmatched_machines_are_kept_not_dropped(self):
        unmatched_names = {m['machine'] for m in self.result['unmatched']}
        self.assertIn('UTV 004', unmatched_names)
        self.assertIn('RMV 005', unmatched_names)

    def test_unmatched_reason_distinguishes_no_code_from_station_not_found(self):
        by_name = {m['machine']: m for m in self.result['unmatched']}
        self.assertEqual(by_name['UTV 004']['reason'], 'station_not_found')
        self.assertEqual(by_name['RMV 005']['reason'], 'no_code')

    def test_placed_machine_entries_carry_only_the_honest_live_fields(self):
        # BCM 001 (station AAA) is the one synthetic machine given detail
        # fields -- fixed 14-Nov-2026: these are now expected to come
        # through (see daily_positions.DETAIL_FIELDS), on top of the
        # original base fields.
        m = self.result['stations']['AAA']['machines'][0]
        self.assertEqual(set(m.keys()), {'machine', 'type', 'typeName', 'cat', 'div', 'stabledAt',
                                          'station', 'posConfidence', 'source',
                                          'blockSection', 'majorSection', 'durationMinutes',
                                          'actualTotal', 'remarks'})
        self.assertEqual(m['blockSection'], 'AAA-BBB')
        self.assertEqual(m['actualTotal'], 0.7)
        # explicitly must still NOT carry any of the OLD baked snapshot's
        # fields that remain genuinely unavailable -- this is the
        # still-standing part of the honest-degradation contract (see
        # this module's docstring's "Still genuinely absent" paragraph).
        for stale_field in ('corridorTiming', 'kmFrom', 'kmTo', 'stipulatedKm', 'actualKm',
                             'pctAchieved', 'dutyStart', 'lossReasons', 'workDone', 'srDen',
                             'iohPohType', 'status'):
            self.assertNotIn(stale_field, m)

    def test_placed_machine_without_detail_fields_does_not_fabricate_them(self):
        # DUO 002 (station BBB/CCC match) was given no detail fields in
        # the synthetic fixture -- must come through with only the base
        # fields, not nulls for every DETAIL_FIELDS entry.
        m = self.result['stations']['BBB']['machines'][0]
        self.assertEqual(m['machine'], 'DUO 002')
        import sources.daily_positions as daily_positions
        for field in daily_positions.DETAIL_FIELDS:
            self.assertNotIn(field, m)

    def test_top_level_shape_is_self_contained_for_the_dynamic_index_html_path(self):
        for key in ('mtype', 'catMeta', 'divColor', 'divName', 'typeGroup', 'kpi'):
            self.assertIn(key, self.result)
        self.assertEqual(self.result['divColor'], network_map.DIV_COLOR)
        self.assertIn('BCM', self.result['mtype'])
        self.assertEqual(self.result['mtype']['BCM']['cat'], 'nontamping')

    def test_kpi_rollup_matches_the_actual_synthetic_data(self):
        kpi = self.result['kpi']
        self.assertEqual(kpi['total'], 5)  # 6 synthetic rows minus the 1 header-leakage row
        self.assertEqual(kpi['placed'], 3)
        self.assertEqual(kpi['unplaced'], 2)
        self.assertEqual(kpi['asOf'], '13.11.2026')

    def test_machine_type_derived_from_leading_uppercase_letters(self):
        m = self.result['stations']['AAA']['machines'][0]
        self.assertEqual(m['machine'], 'BCM 001')
        self.assertEqual(m['type'], 'BCM')

    def test_as_of_date_passed_through(self):
        self.assertEqual(self.result['asOfDate'], '13.11.2026')

    def test_depot_and_approx_stations_available_for_matching_even_with_no_machines_there(self):
        # none of the synthetic machines stable at a depot/approx code,
        # but they must still exist in the geometry the map would draw,
        # just with an empty machines list if actually rendered as a
        # station -- here we only assert they're resolvable at all.
        code, info = network_map.match_machine_to_station('ZBD/YPR', {})  # not in live stations
        self.assertIsNone(code)  # confirms match_machine_to_station needs the merged dict
        merged = network_map.merge_supplement_stations(
            network_map.build_stations_from_rdps(rdps_points.extract_points(synthetic_rdps_rows())))
        code, info = network_map.match_machine_to_station('ZBD/YPR', merged)
        self.assertEqual(code, 'ZBD/YPR')
        self.assertEqual(info['kind'], 'depot')

    def test_tracks_by_div_present_and_shaped_correctly(self):
        self.assertIn('UBL', self.result['tracksByDiv'])
        self.assertIn('MYS', self.result['tracksByDiv'])
        for sheets in self.result['tracksByDiv'].values():
            for sheet in sheets:
                self.assertIn('sheet', sheet)
                self.assertIn('pts', sheet)
                for pt in sheet['pts']:
                    self.assertEqual(len(pt), 2)


class TestSrDenJoin(unittest.TestCase):
    """Fixed 18-Nov-2026: see build_network_map's docstring entry on
    this date -- this join is what feeds index.html's SRDEN_CODES check
    and its "Jump closer by Sr.DEN jurisdiction" box, which used to
    always read empty because this function never looked at the
    daily_progress_srden table at all."""

    def setUp(self):
        self.result = network_map.build_network_map(
            synthetic_rdps_rows(), synthetic_positions_payload(), synthetic_srden_payload())

    def test_matched_machine_gets_the_bare_jurisdiction_letter(self):
        # BCM 001 -> UBL, canon_srden gives "UBL Sr.DEN/W" -- only the
        # trailing "W" is what SRDEN_CODES (on the JS side) checks for.
        m = self.result['stations']['AAA']['machines'][0]
        self.assertEqual(m['machine'], 'BCM 001')
        self.assertEqual(m['srDen'], 'W')

    def test_a_different_machine_and_division_gets_its_own_code(self):
        m = self.result['stations']['DDD']['machines'][0]
        self.assertEqual(m['machine'], 'MPT 003')
        self.assertEqual(m['srDen'], 'N')

    def test_lookup_miss_leaves_srden_off_rather_than_fabricating_it(self):
        # DUO 002 is deliberately absent from synthetic_srden_payload().
        m = self.result['stations']['BBB']['machines'][0]
        self.assertEqual(m['machine'], 'DUO 002')
        self.assertNotIn('srDen', m)

    def test_no_srden_payload_at_all_leaves_every_machine_without_it(self):
        # The default (no 3rd argument) must behave exactly as it did
        # before this fix -- confirmed separately by
        # TestBuildNetworkMapEndToEnd's own
        # test_placed_machine_entries_carry_only_the_honest_live_fields,
        # which asserts 'srDen' is absent; this test instead checks the
        # explicit srden_payload=None case gives the identical result.
        result_default = network_map.build_network_map(synthetic_rdps_rows(), synthetic_positions_payload())
        result_explicit_none = network_map.build_network_map(
            synthetic_rdps_rows(), synthetic_positions_payload(), None)
        for res in (result_default, result_explicit_none):
            m = res['stations']['AAA']['machines'][0]
            self.assertNotIn('srDen', m)


if __name__ == '__main__':
    unittest.main()
