# SAMARTH live aggregation pipeline (Python, Supabase-backed)

This is the full rebuild of the SWR Track Machine dashboard's data
pipeline against all 16 live-synced Supabase tables, replacing the old
local-Excel-file version (`run_pipeline.py`, which still exists in this
repo as a historical reference / fallback and is untouched by this work).

## What's here

```
pipeline/
  aggregate.py            <- the pipeline itself: fetches all 16 live
                              Supabase tables, aggregates them into the
                              dashboard_payload.json shape the website
                              already expects, writes it to a local file.
  push_to_supabase.py      <- pushes that file to the dashboard_payload
                              table (the ONLY script that uses the
                              service_role write key).
  pipeline_helpers.py       <- shared pure helper functions (fuzzy machine-
                              name matching, classification, date parsing)
                              used by both aggregate.py and the adapters.
  supabase_client.py        <- thin read-only REST client (anon key only).
  sources/                  <- one adapter module per live-synced table
                              (16 total). Each exposes fetch_payload()/
                              fetch_rows()/fetch_all() (the one REST call)
                              and extract_*() (pure reshaping, no network)
                              separately, so extract_*() is unit-testable
                              without hitting Supabase at all.
  tests/
    test_integrity.py        <- the integrity test suite (unittest).
    _validate_aggregate.py   <- a manual/offline validation harness used
                              during development (not part of CI).
  README.md                 <- this file.

.github/workflows/
  daily-refresh.yml          <- the GitHub Actions workflow that runs this
                              pipeline daily and pushes the result to
                              Supabase. See its own header comment for the
                              one-time setup (3 repo secrets) a human needs
                              to do once.
```

## Running it locally / manually

```bash
export SUPABASE_URL="https://xxxxx.supabase.co"
export SUPABASE_ANON_KEY="..."          # read-only, for aggregate.py
python3 aggregate.py /tmp/out            # writes /tmp/out/dashboard_payload.json

export SUPABASE_SERVICE_ROLE_KEY="..."  # write access, for push_to_supabase.py
python3 push_to_supabase.py /tmp/out/dashboard_payload.json
```

Pass `--as-of YYYY-MM-DD` to `aggregate.py` to override today's date (used
for FY/period-label computation, expiry/overdue day-counts, etc.) -- useful
for re-running against an earlier day's data for comparison.

## Running the integrity tests

```bash
python3 -m unittest tests.test_integrity -v
```

No network access or credentials needed -- the adapter-structural tests
run against cached fixture files (if present; they skip themselves
cleanly if not), and the pure `pipeline_helpers.py` unit tests always run
regardless. This is also what `.github/workflows/daily-refresh.yml` runs
as a pre-flight smoke check before it spends a live Supabase round trip.

## Known gaps (by design, not bugs)

Two sources genuinely have no live-sync equivalent yet and degrade
gracefully (see `aggregate.py`'s own module docstring for the full
reasoning):

- **Daily Progress shortfall-category log** -- none of the 16 live tables
  cover this (confirmed by checking every candidate table's own data
  shape against it). `shortfall.*`/`srdenBundle`/the Reason-for-Shortfall
  tab all degrade to empty/zero until a live source exists for this.
- **IOH/POH current-year office letter + CPOH/RYP forward-planning
  letter** -- both are Word (.docx) documents with no Supabase table
  behind them. Fleet Health's `currentYear` section and `cpohPlanning`
  degrade to `None`/empty; the FY2027-28..2029-30 **forecast** (from the
  live `ioh_poh_planning` table) is unaffected.

The other 6 live tables with no dashboard consumption yet
(`cumulative_failure`, `daily_positions`, `hq_summary`, `monthly_machine`,
`rdps_points`, `shortfall_reasons`) are still fetched and reshaped by this
pipeline (so all 16 sources really are live) and attached to the output
under `payload.rawSources`, ready for whenever a future dashboard tab
wants to read them.

## Adding this to the SWRTMO repo

This Cowork session has no GitHub write access (confirmed repeatedly
throughout this project -- see `claude/supabase-live-deployment-notes.md`),
so a human needs to add these files to the repo by hand, once:

1. Copy the `pipeline/` folder and `.github/workflows/daily-refresh.yml`
   into the `SWRTMO` repo at the same relative paths.
2. Add the three repository secrets listed in `daily-refresh.yml`'s own
   header comment (`SUPABASE_URL`, `SUPABASE_ANON_KEY`,
   `SUPABASE_SERVICE_ROLE_KEY` -- the service_role one should already
   exist from the earlier manual-push era; reuse it).
3. Done. The workflow runs itself daily from then on (and can also be
   triggered manually from the Actions tab's "Run workflow" button).
