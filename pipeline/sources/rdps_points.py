"""Adapter for the `rdps_points` table.

Live sync: rdps_sync.gs - three standalone copies (one per division:
MYS/SBC/UBL), run BY HAND whenever a fresh survey-correction pass is
ready, not on a daily trigger. Static track-survey data: one real row per
KM/station point (not the {id, payload jsonb, updated_at} shape every
other sync in this project uses), columns: division, route, km,
distance_m, lat, lon, station_code, station_name, review_notes. Feeds the
SWR Network Map tab's station database - not read by run_pipeline.py's
run() at all (the Network Map is a separate concept doc/prototype -
claude/network-map-tab-concept.md - not yet wired into the main
aggregation payload).

Uses supabase_client.fetch_rows() rather than fetch_latest_payload() since
this table has no `payload` column to unwrap.
"""
import supabase_client

TABLE = "rdps_points"


def fetch_rows(division=None):
    """-> list of raw rdps_points rows. Pass division='MYS'/'SBC'/'UBL' to
    fetch just one division's points (cheaper - this table can hold tens
    of thousands of rows across all three); omit for everything.

    Fixed 14-Nov-2026: this table holds ~430,000 rows total (UBL 210,502
    / SBC 173,018 / MYS 46,850) - far past supabase_client.fetch_rows()'s
    old default `limit=20000`, which silently truncated every call here
    to a small fraction of one division. `limit` is now raised well past
    the real total (the fetch still stops as soon as a page comes back
    short, via the pagination fix in supabase_client.fetch_rows() itself,
    so this is just a safety ceiling, not the actual request size).

    Fixed 26-Nov-2026: passes a stable `order` through to
    supabase_client.fetch_rows() now - see that function's own 26-Nov-2026
    docstring entry for why an unordered paginated fetch of this many rows
    could silently skip/duplicate rows (confirmed: it was)."""
    params = f"&division=eq.{division}" if division else ""
    return supabase_client.fetch_rows(
        TABLE,
        select="division,route,km,distance_m,lat,lon,station_code,station_name,review_notes",
        params=params,
        limit=500000,
        order="division,route,km",
    )


def extract_points(rows):
    """-> {division: [{"route","km","distanceM","lat","lon","stationCode",
    "stationName","reviewNotes"}, ...], ...}, grouped by division and
    reshaped to camelCase keys (matching this pipeline's own output-field
    naming convention elsewhere) rather than the raw snake_case column
    names Supabase returns."""
    out = {}
    for r in rows or []:
        div = r.get('division')
        if not div:
            continue
        out.setdefault(div, []).append(dict(
            route=r.get('route'), km=r.get('km'), distanceM=r.get('distance_m'),
            lat=r.get('lat'), lon=r.get('lon'), stationCode=r.get('station_code'),
            stationName=r.get('station_name'), reviewNotes=r.get('review_notes'),
        ))
    return out
