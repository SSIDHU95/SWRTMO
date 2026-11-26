"""Thin read-only Supabase REST client for the live-sync tables.

Every one of the 16 Apps Script syncs POSTs one new row per run into its own
table, shaped {id, payload (jsonb), updated_at}. We only ever need the
LATEST row per table (no history needed for the merge), so this is the one
function every source adapter calls.

Auth: reads use the anon key (every sync table grants anon SELECT per the
project's standard grants convention - see live-sheet-cron-notes.md), so no
service_role key is needed or read by this module. The service_role key is
only needed for the final WRITE to dashboard_payload, which lives in
push_to_supabase.py, kept deliberately separate so a bug in a read path can
never accidentally touch write credentials.

Env vars (same two names used by every .gs script in this project):
  SUPABASE_URL
  SUPABASE_ANON_KEY
"""
import os
import sys
import urllib.request
import urllib.error
import json


class SupabaseFetchError(RuntimeError):
    pass


def _base_url():
    url = os.environ.get("SUPABASE_URL")
    if not url:
        raise SupabaseFetchError("SUPABASE_URL is not set")
    return url.rstrip("/")


def _anon_key():
    key = os.environ.get("SUPABASE_ANON_KEY")
    if not key:
        raise SupabaseFetchError("SUPABASE_ANON_KEY is not set")
    return key


def fetch_latest_payload(table, required=True):
    """Fetches the single most-recently-updated row's `payload` column from
    `table`. Returns None if the table is empty and required=False; raises
    SupabaseFetchError if empty and required=True, or on any HTTP/network
    error - a source adapter should let this propagate rather than catch it
    silently, since a missing source should fail the whole run loudly (same
    philosophy as run_pipeline.py's one_file(required=True) for local files).
    """
    url = (
        f"{_base_url()}/rest/v1/{table}"
        f"?select=payload,updated_at&order=updated_at.desc&limit=1"
    )
    req = urllib.request.Request(
        url,
        headers={
            "apikey": _anon_key(),
            "Authorization": f"Bearer {_anon_key()}",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = resp.read()
    except urllib.error.HTTPError as e:
        raise SupabaseFetchError(
            f"Fetching '{table}' failed: HTTP {e.code} -- {e.read().decode(errors='replace')}"
        ) from e
    except urllib.error.URLError as e:
        raise SupabaseFetchError(f"Fetching '{table}' failed: {e}") from e

    rows = json.loads(body)
    if not rows:
        if required:
            raise SupabaseFetchError(f"Table '{table}' has no rows yet")
        return None
    return rows[0]["payload"], rows[0]["updated_at"]


def fetch_all_payloads(table, limit=2000):
    """Fetches every row's (payload, updated_at) from `table`, oldest first.

    Most of the 16 tables only ever need fetch_latest_payload() above - one
    new row overwrites the previous day's meaning entirely. A few sources
    are different: the live sync pushes one NEW snapshot per day without
    deleting the earlier ones (e.g. exception_sync.gs's one-tab-per-day
    payload, or daily_progress_sync.gs's per-day shortfall log), so the
    table is itself the accumulating history that run_pipeline.py used to
    get by reading one local file per day. Those adapters call this instead
    and fold the rows into whatever day-keyed shape the aggregation step
    expects. Returns [] if the table is empty (never raises on empty - an
    accumulating source legitimately starts empty on day one).
    """
    url = (
        f"{_base_url()}/rest/v1/{table}"
        f"?select=payload,updated_at&order=updated_at.asc&limit={int(limit)}"
    )
    req = urllib.request.Request(
        url,
        headers={
            "apikey": _anon_key(),
            "Authorization": f"Bearer {_anon_key()}",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = resp.read()
    except urllib.error.HTTPError as e:
        raise SupabaseFetchError(
            f"Fetching all rows of '{table}' failed: HTTP {e.code} -- {e.read().decode(errors='replace')}"
        ) from e
    except urllib.error.URLError as e:
        raise SupabaseFetchError(f"Fetching all rows of '{table}' failed: {e}") from e

    rows = json.loads(body)
    return [(row["payload"], row["updated_at"]) for row in rows]


def fetch_rows(table, select="*", params="", limit=20000, page_size=1000, order=None):
    """Fetches raw rows from a table that does NOT follow the standard
    {id, payload jsonb, updated_at} shape - so far, only rdps_points (one
    real row per survey point: division, route, km, distance_m, lat, lon,
    station_code, station_name, review_notes - pushed by rdps_sync.gs,
    which is run by hand per-division rather than on the daily trigger
    every other sync uses). `params` is any extra raw querystring
    (already URL-encoded) to append, e.g. "&division=eq.MYS". Returns []
    if the table is empty - never raises on empty, since this is static
    survey data with no "day one" accumulation expectation the way
    fetch_all_payloads' sources have.

    Fixed 14-Nov-2026: Supabase's PostgREST caps every single response at
    1000 rows (its own server-side max-rows setting) NO MATTER what
    `?limit=` asks for - confirmed directly against the live project (a
    plain `?limit=20000` request and an explicit `Range: 0-1999` request
    both silently came back with exactly 1000 rows, `Content-Range:
    0-999/*`). rdps_points alone holds ~430,000 rows across the three
    divisions, so the un-paginated version of this function was silently
    reading only the first 1000 of them every single run - in practice,
    whichever single route happened to be inserted first - which is why
    the Network Map was only ever placing a small handful of machines
    ("station_not_found" for the rest) no matter how fresh the daily sync
    data was. Now pages through with the PostgREST `Range`/`Range-Unit`
    headers (0-999, 1000-1999, ...) until a page comes back with fewer
    than `page_size` rows (the real end of the table), accumulating every
    page. `limit` is kept as an overall safety cap (stops paging once this
    many rows have been collected, even if more remain) rather than a
    per-request value now.

    Fixed 26-Nov-2026: found while building the new Infra Map feature
    (pipeline/sources/infra_map.py) -- two back-to-back unfiltered fetches
    of this same ~430,000-row table came back with a DIFFERENT division
    split each time (221,326/162,063/47,000 vs. 184,442/195,879/50,068 -
    same total, different per-division counts), and a set of station
    codes confirmed present in the live data (KNP, MGD, GPB, SBH, BJP,
    JRKT, SVNR, GIN, all Hubballi division) were silently MISSING from one
    of the two fetches' results entirely. Root cause: offset/Range-based
    pagination against a query with no `order=` clause has no guarantee
    from Postgres that row order stays stable between separate requests
    (no explicit sort = whatever order the query planner's scan happens
    to produce, which can and does shift) - so paging through ~430 pages
    this way can silently skip or duplicate rows at page boundaries,
    differently on every call. This was already a live, if less visible,
    risk for the Network Map tab's own build_stations_from_rdps() (it
    keeps whichever occurrence of a station code happens to be read
    first, so an unstable fetch order could already have been silently
    varying which single survey reading won for a station from one
    pipeline run to the next). `order`, when given (now always passed by
    rdps_points.fetch_rows() as "division,route,km" - the natural sort
    this data already implicitly assumes elsewhere, e.g. in route/track
    reconstruction), makes every page's boundary deterministic, so the
    full result is identical and complete regardless of how many pages
    it takes or when it's fetched."""
    page_size = min(int(page_size), 1000)  # the server-side cap itself
    order_param = f"&order={order}" if order else ""
    base_url = f"{_base_url()}/rest/v1/{table}?select={select}{params}{order_param}"
    rows = []
    offset = 0
    while True:
        req = urllib.request.Request(
            base_url,
            headers={
                "apikey": _anon_key(),
                "Authorization": f"Bearer {_anon_key()}",
                "Range-Unit": "items",
                "Range": f"{offset}-{offset + page_size - 1}",
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                body = resp.read()
        except urllib.error.HTTPError as e:
            # A 416 here means the table has fewer rows than `offset` -
            # i.e. we already read every row on the previous page - treat
            # it as a clean end-of-table, not an error.
            if e.code == 416:
                break
            raise SupabaseFetchError(
                f"Fetching rows of '{table}' failed: HTTP {e.code} -- {e.read().decode(errors='replace')}"
            ) from e
        except urllib.error.URLError as e:
            raise SupabaseFetchError(f"Fetching rows of '{table}' failed: {e}") from e
        page = json.loads(body)
        rows.extend(page)
        if len(page) < page_size or len(rows) >= limit:
            break
        offset += page_size
    return rows[:limit]


def log(*a):
    print(*a, file=sys.stderr)
