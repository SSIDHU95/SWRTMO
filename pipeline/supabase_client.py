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


def fetch_rows(table, select="*", params="", limit=20000):
    """Fetches raw rows from a table that does NOT follow the standard
    {id, payload jsonb, updated_at} shape - so far, only rdps_points (one
    real row per survey point: division, route, km, distance_m, lat, lon,
    station_code, station_name, review_notes - pushed by rdps_sync.gs,
    which is run by hand per-division rather than on the daily trigger
    every other sync uses). `params` is any extra raw querystring
    (already URL-encoded) to append, e.g. "&division=eq.MYS". Returns []
    if the table is empty - never raises on empty, since this is static
    survey data with no "day one" accumulation expectation the way
    fetch_all_payloads' sources have."""
    url = f"{_base_url()}/rest/v1/{table}?select={select}&limit={int(limit)}{params}"
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
            f"Fetching rows of '{table}' failed: HTTP {e.code} -- {e.read().decode(errors='replace')}"
        ) from e
    except urllib.error.URLError as e:
        raise SupabaseFetchError(f"Fetching rows of '{table}' failed: {e}") from e
    return json.loads(body)


def log(*a):
    print(*a, file=sys.stderr)
