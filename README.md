# SWR Track Machine Watch

A dashboard that tracks South Western Railway's track machines — progress, exceptions, fleet health, indent position, and more.

## Files in this repository

- **index.html** — The complete, ready-to-view dashboard, with the latest data already built in. This is the one file you need if you just want to open or host the dashboard as-is.
- **dashboard_template2.html** — The dashboard's design/layout "template," without any data filled in yet. Used together with `run_pipeline.py` and `merge_payload.py` to rebuild `index.html` whenever the source data changes.
- **run_pipeline.py** — A Python script that reads the Excel/Word files from the Machine Monitoring folder and turns them into a single data file (`dashboard_payload.json`).
- **merge_payload.py** — A Python script that combines `dashboard_template2.html` (the template) with `dashboard_payload.json` (the data) to produce an updated `index.html`.

## To view the dashboard right now

Just open `index.html` in any web browser (double-click it, or drag it into a browser tab).

## To host it as a live website (GitHub Pages)

1. Make sure `index.html` is in the root of this repository.
2. Go to the repository's **Settings** tab → **Pages** (left sidebar).
3. Under "Build and deployment," set Source to **Deploy from a branch**, pick the **main** branch and the **/ (root)** folder, then click **Save**.
4. GitHub will give you a web address (something like `https://your-username.github.io/SWRTMO/`) — that's your live site.

Note: GitHub Pages only works for free on a **public** repository (or on a paid GitHub plan for private repositories). Since this repository is currently private, you'll need to either make it public or upgrade your plan before Pages will work.
