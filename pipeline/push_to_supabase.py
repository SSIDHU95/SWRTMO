"""
Pushes a built dashboard_payload.json to the `dashboard_payload` Supabase
table that the live SAMARTH website (GitHub Pages `index.html`, Supabase-
auth-gated) reads from (`select * order by updated_at desc limit 1`).

This is the ONE place in the whole project that ever touches the
service_role key -- deliberately kept separate from every read-only
adapter in sources/*.py (which only ever use the anon key via
supabase_client.fetch_latest_payload/fetch_all_payloads/fetch_rows), so a
bug in a read path can never accidentally touch write credentials. This
mirrors the hand-rolled `curl` push pattern used throughout this project's
history (see claude/supabase-live-deployment-notes.md) -- same endpoint,
same request shape (`{"payload": <dashboard_payload.json>}`), same table
-- just in Python instead of a one-off shell command, so it can run
unattended inside GitHub Actions (see .github/workflows/daily-refresh.yml,
handed to the user to add to the SWRTMO repo herself).

Why this has to be a separate script, run by GitHub Actions rather than by
this Cowork session directly: Claude's own safety system blocks
persisting a service_role key anywhere for later unattended reuse (hard
rule, confirmed repeatedly across this project's history -- see the
"Daily automation" section of supabase-live-deployment-notes.md). Moving
the secret into GitHub's own Secrets vault and letting Actions read it at
run time sidesteps that entirely: this script never has the key typed
into Claude's own context, Claude never sees or handles it, and nothing
about it is persisted inside this chat/session.

Env vars required:
    SUPABASE_URL
    SUPABASE_SERVICE_ROLE_KEY   (NOT the anon key -- the whole point of
                                 this script is writing, which the anon
                                 key's RLS policy does not permit; see
                                 supabase-live-deployment-notes.md's GRANT
                                 statements)

Usage:
    python3 push_to_supabase.py <path-to-dashboard_payload.json>

---
Added 26-Nov-2026: this table was never trimmed -- every daily run just
added a new row and never removed the old one (see project-status.md
section 4), and a real GitHub Actions failure confirmed the consequence:
the INSERT itself started hitting a Postgres statement timeout (HTTP 500,
code 57014 "canceling statement due to statement timeout") because the
table had grown too large. The website only ever reads the single newest
row (`order=updated_at.desc limit 1`), so nothing is lost by trimming.

Two things changed:
1. On every successful push, this script now deletes every OTHER row in
   the table right after the insert, so the table settles at exactly 1
   row going forward and this can't recur.
2. If the insert itself times out (the state the table was already in),
   this script self-heals once: it finds the current newest existing row,
   deletes every row EXCEPT that one (shrinking the table while still
   leaving a fallback row in place -- if the retry below also fails, the
   website keeps showing that last-good day's data, exactly as it would
   have before this change), then retries the insert exactly once. It
   never retries more than that -- a second failure is raised loudly
   rather than looped on, so a different underlying problem doesn't get
   silently masked.
"""
import json
import os
import sys
import urllib.error
import urllib.request

TABLE = "dashboard_payload"


def log(*a):
    print(*a, file=sys.stderr)


def _headers(service_key, extra=None):
    h = {"apikey": service_key, "Authorization": f"Bearer {service_key}"}
    if extra:
        h.update(extra)
    return h


def _insert(base_url, service_key, table, dashboard_payload, timeout=30):
    body = json.dumps({"payload": dashboard_payload}).encode("utf-8")
    url = f"{base_url.rstrip('/')}/rest/v1/{table}"
    req = urllib.request.Request(
        url,
        data=body,
        method="POST",
        headers=_headers(service_key, {
            "Content-Type": "application/json",
            # Ask PostgREST to hand back the inserted row (id/updated_at)
            # so this script can confirm and log what landed, same as the
            # project's prior hand-run curl pushes always verified with a
            # follow-up GET.
            "Prefer": "return=representation",
        }),
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.status, resp.read()


def _get_newest_id(base_url, service_key, table, timeout=30):
    """The current newest row's id, via the identity PK's own index
    (`order=id.desc`, not `updated_at`, so this stays fast even on a
    table that's already in the bloated state this module exists to fix).
    None if the table is already empty."""
    url = f"{base_url.rstrip('/')}/rest/v1/{table}?select=id&order=id.desc&limit=1"
    req = urllib.request.Request(url, method="GET", headers=_headers(service_key))
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        rows = json.loads(resp.read() or b"[]")
    return rows[0]["id"] if rows else None


def _delete_other_rows(base_url, service_key, table, keep_id, timeout=60):
    """Deletes every row in `table` except `keep_id`. `id` is a positive
    identity column, so `id=neq.<keep_id>` alone safely matches every
    other row with no second SELECT needed."""
    url = f"{base_url.rstrip('/')}/rest/v1/{table}?id=neq.{keep_id}"
    req = urllib.request.Request(
        url, method="DELETE", headers=_headers(service_key, {"Prefer": "return=minimal"})
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.status


def _is_statement_timeout(http_code, body_text):
    if http_code != 500:
        return False
    try:
        return json.loads(body_text).get("code") == "57014"
    except (ValueError, AttributeError):
        return False


def push(payload_path, table=TABLE):
    base_url = os.environ.get("SUPABASE_URL")
    if not base_url:
        raise RuntimeError("SUPABASE_URL is not set")
    service_key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
    if not service_key:
        raise RuntimeError("SUPABASE_SERVICE_ROLE_KEY is not set")

    with open(payload_path) as f:
        dashboard_payload = json.load(f)

    try:
        status, resp_body = _insert(base_url, service_key, table, dashboard_payload)
    except urllib.error.HTTPError as e:
        err_body = e.read().decode(errors="replace")
        if not _is_statement_timeout(e.code, err_body):
            raise RuntimeError(f"Push to '{table}' failed: HTTP {e.code} -- {err_body}") from e
        log(f"INSERT into '{table}' hit a statement timeout (57014) -- the table has "
            f"likely grown too large (never trimmed before this fix -- see "
            f"project-status.md section 4). Trimming old rows (keeping the most "
            f"recent one as a fallback) and retrying the insert once.")
        try:
            keep_id = _get_newest_id(base_url, service_key, table)
            if keep_id is not None:
                _delete_other_rows(base_url, service_key, table, keep_id)
        except (urllib.error.HTTPError, urllib.error.URLError) as de:
            raise RuntimeError(
                f"Push to '{table}' failed (insert timed out, and the recovery "
                f"cleanup also failed: {de}). The live site should still be showing "
                f"its last successful day's data, untouched."
            ) from de
        try:
            status, resp_body = _insert(base_url, service_key, table, dashboard_payload)
        except urllib.error.HTTPError as e2:
            err_body2 = e2.read().decode(errors="replace")
            raise RuntimeError(
                f"Push to '{table}' failed even after trimming old rows: "
                f"HTTP {e2.code} -- {err_body2}. The live site should still be "
                f"showing its last successful day's data, untouched."
            ) from e2
    except urllib.error.URLError as e:
        raise RuntimeError(f"Push to '{table}' failed: {e}") from e

    rows = json.loads(resp_body) if resp_body else []
    row = rows[0] if rows else {}
    new_id = row.get("id")

    # Keep the table at exactly 1 row going forward, now that the new row
    # is safely in -- this is what stops the timeout above from recurring.
    # Best-effort: a failure here doesn't undo the successful push; it
    # just means the table stays a bit larger until the next run trims it.
    if new_id is not None:
        try:
            _delete_other_rows(base_url, service_key, table, new_id)
        except (urllib.error.HTTPError, urllib.error.URLError) as e:
            log(f"Warning: push succeeded, but trimming old '{table}' rows failed: {e}")

    log(f"Pushed to '{table}': HTTP {status}, new row id={new_id}, "
        f"updated_at={row.get('updated_at')}")
    return row


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 push_to_supabase.py <path-to-dashboard_payload.json>", file=sys.stderr)
        sys.exit(1)
    push(sys.argv[1])
