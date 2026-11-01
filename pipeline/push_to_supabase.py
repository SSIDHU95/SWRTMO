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
"""
import json
import os
import sys
import urllib.error
import urllib.request

TABLE = "dashboard_payload"


def log(*a):
    print(*a, file=sys.stderr)


def push(payload_path, table=TABLE):
    base_url = os.environ.get("SUPABASE_URL")
    if not base_url:
        raise RuntimeError("SUPABASE_URL is not set")
    service_key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
    if not service_key:
        raise RuntimeError("SUPABASE_SERVICE_ROLE_KEY is not set")

    with open(payload_path) as f:
        dashboard_payload = json.load(f)

    body = json.dumps({"payload": dashboard_payload}).encode("utf-8")
    url = f"{base_url.rstrip('/')}/rest/v1/{table}"
    req = urllib.request.Request(
        url,
        data=body,
        method="POST",
        headers={
            "apikey": service_key,
            "Authorization": f"Bearer {service_key}",
            "Content-Type": "application/json",
            # Ask PostgREST to hand back the inserted row (id/updated_at)
            # so this script can confirm and log what landed, same as the
            # project's prior hand-run curl pushes always verified with a
            # follow-up GET.
            "Prefer": "return=representation",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            status = resp.status
            resp_body = resp.read()
    except urllib.error.HTTPError as e:
        raise RuntimeError(
            f"Push to '{table}' failed: HTTP {e.code} -- {e.read().decode(errors='replace')}"
        ) from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Push to '{table}' failed: {e}") from e

    rows = json.loads(resp_body) if resp_body else []
    row = rows[0] if rows else {}
    log(f"Pushed to '{table}': HTTP {status}, new row id={row.get('id')}, "
        f"updated_at={row.get('updated_at')}")
    return row


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 push_to_supabase.py <path-to-dashboard_payload.json>", file=sys.stderr)
        sys.exit(1)
    push(sys.argv[1])
