"""Adapter for the `action_plan_53` table.

Live sync: action_plan_53_sync.gs. One row, re-pushed daily (the sheet
itself only changes about once every 15 days - the daily trigger just
harmlessly re-pushes the same data on days nothing has changed), covering
the WHOLE 53-Point TM Action Plan sheet in one payload:
    {
      "sheetName": "...",
      "actionPlan53": {
        "checkpointLabels": ["Progress up to 20.04.2026", ...],   # sheet's
            own left-to-right column order - NOT guaranteed chronological
            (confirmed: '31.05.2026' appears before '30.05.2026' in the
            live sheet, same out-of-order pair run_pipeline.py's own
            docstring already flags)
        "supervisorLabels": ["SBC Division", "MYS Division", "UBL Division",
            "Co-ordinator"],
        "rows": [
          {"slNo": 1, "category": "SAFETY", "actionPoint": "...",
           "supervisors": {"SBC Division": "...", ...},
           "checkpoints": {"Progress up to 20.04.2026": "...", ...}},
          ...
        ]
      }
    }

Replaces run_pipeline.py's extract_action_plan_53(base) (and its helper
_build_action_plan_insights, ported verbatim below as _build_insights) -
the live sync already does the column-discovery/header-parsing work the
local-file version did with openpyxl (auto-finding "Progress up to ..."
columns and the Nominated-Supervisor sub-header columns), so this adapter
only has to re-sort the checkpoint columns chronologically and replicate
the per-point history/momentum/insight computation.
"""
import datetime
import difflib
import re
from collections import Counter, defaultdict

import pipeline_helpers as ph
import supabase_client

TABLE = "action_plan_53"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def _checkpoint_date_text(label):
    """'Progress up to 20.04.2026' -> '20.04.2026', matching run_pipeline.
    py's own date_txt = v.strip()[len('progress up to'):].strip()."""
    return re.sub(r'(?i)^\s*progress up to\s*', '', str(label or '')).strip()


def _ordered_checkpoints(checkpoint_labels):
    """-> [(label, date_txt, parsed_date_or_None), ...], sorted chronologically
    where parseable (the live sheet's own column order can have an
    out-of-sequence pair - see module docstring), unparseable ones kept at
    the end in their original left-to-right order, exactly like
    run_pipeline.py's own dated/undated split."""
    cols = [(label, _checkpoint_date_text(label)) for label in (checkpoint_labels or [])]
    cols = [(label, date_txt, ph.parse_ddmmyyyy(date_txt)) for label, date_txt in cols]
    dated = sorted((c for c in cols if c[2] is not None), key=lambda c: c[2])
    undated = [c for c in cols if c[2] is None]
    return dated + undated


def extract_action_plan_53(payload):
    ap = (payload or {}).get('actionPlan53') or {}
    checkpoint_labels = ap.get('checkpointLabels') or []
    supervisor_labels = ap.get('supervisorLabels') or []
    rows_in = ap.get('rows') or []

    ordered_cols = _ordered_checkpoints(checkpoint_labels)

    points = []
    cat_counts = Counter()
    for r in rows_in:
        sno = r.get('slNo')
        if not isinstance(sno, (int, float)):
            continue
        category = re.sub(r'\s+', ' ', str(r.get('category') or '')).strip() or 'Uncategorised'
        action_text = re.sub(r'\s+', ' ', str(r.get('actionPoint') or '')).strip()

        raw_sup = r.get('supervisors') or {}
        supervisors = {}
        for label in supervisor_labels:
            v = raw_sup.get(label)
            if v not in (None, ''):
                supervisors[label] = re.sub(r'\s+', ' ', str(v)).strip()

        raw_cp = r.get('checkpoints') or {}
        history, presence = [], []
        for label, date_txt, _dt in ordered_cols:
            v = raw_cp.get(label)
            txt = str(v).strip() if v not in (None, '') else ''
            presence.append(bool(txt))
            if txt:
                history.append(dict(date=date_txt, text=txt))
        latest = history[-1] if history else None
        updated_this_cycle = presence[-1] if presence else False

        points.append(dict(
            sno=int(sno), category=category, actionPoint=action_text, supervisors=supervisors,
            history=history, latestDate=(latest['date'] if latest else None),
            latestText=(latest['text'] if latest else None),
            updatedThisCycle=updated_this_cycle, presence=presence,
        ))
        cat_counts[category] += 1

    review_date_labels = [date_txt for _label, date_txt, _dt in ordered_cols]
    latest_review_label = review_date_labels[-1] if review_date_labels else None

    insights = _build_insights(points, review_date_labels)

    sheet_name = (payload or {}).get('sheetName')
    return dict(
        sourceFile=f"live sync (sheet: {sheet_name})" if sheet_name else "live sync (action_plan_53)",
        reviewDates=review_date_labels,
        latestReviewDate=latest_review_label,
        totalPoints=len(points),
        categories=[dict(name=k, count=v) for k, v in cat_counts.items()],
        points=points,
        insights=insights,
    )


# ---------------------------------------------------------------------
# Ported verbatim from run_pipeline.py's _build_action_plan_insights() -
# a pure function of `points`/`review_date_labels`, no file I/O, so it
# carries over unchanged (see pipeline_helpers.py's module docstring for
# why this project keeps such ports byte-for-byte rather than rewriting).
# ---------------------------------------------------------------------
_AP_MACHINE_RE = re.compile(
    r'\b(UNIMAT|DUOMATIC|WMDU|BCM|UTV|DGS|MPT|UNI|DUO|CSM|BRM|PBR|SQRS|FRM|SRGM|RGM|RBMV|PCTM|DTE|MDU|WST)'
    r'[\s\-]{0,2}(\d{2,6})\b', re.IGNORECASE)
_AP_MACHINE_ALIAS = {'UNIMAT': 'UNI', 'DUOMATIC': 'DUO', 'WMDU': 'MDU'}
_AP_DATE_RE = re.compile(r'\b(\d{1,2})[./](\d{1,2})[./](\d{2,4})\b')
_AP_NIL_RE = re.compile(r'\b(nil|no\s+progress|no\s+update|no\s+any|not\s+received|not\s+yet\s+started|awaited)\b', re.I)


def _ap_parse_short_date(day, mon, yr):
    try:
        day, mon = int(day), int(mon)
        yr = int(yr)
        if yr < 100:
            yr += 2000
        return datetime.date(yr, mon, day)
    except Exception:
        return None


def _build_insights(points, review_date_labels):
    if not points:
        return None
    latest_label = review_date_labels[-1] if review_date_labels else None
    prev_label = review_date_labels[-2] if len(review_date_labels) > 1 else None

    substantive, repeat, not_updated = [], [], []
    for pt in points:
        h = pt['history']
        if not pt['updatedThisCycle']:
            not_updated.append(pt)
            continue
        if len(h) >= 2:
            ratio = difflib.SequenceMatcher(None, h[-1]['text'], h[-2]['text']).ratio()
            if ratio > 0.80:
                repeat.append(pt)
                continue
        substantive.append(pt)

    stalled = []
    for pt in points:
        pres = pt.get('presence') or []
        if len(pres) >= 2 and not pres[-1] and not pres[-2]:
            stalled.append(pt)

    flagged_nil = []
    for pt in points:
        t = pt.get('latestText') or ''
        if t and len(t) < 200 and _AP_NIL_RE.search(t):
            flagged_nil.append(pt)

    machine_points = defaultdict(list)
    for pt in points:
        t = pt.get('latestText') or ''
        if not t:
            continue
        seen_in_point = set()
        for m in _AP_MACHINE_RE.finditer(t):
            code = m.group(1).upper()
            code = _AP_MACHINE_ALIAS.get(code, code)
            num = str(int(m.group(2))) if m.group(2).isdigit() else m.group(2)
            key = f"{code}-{num}"
            if key in seen_in_point:
                continue
            seen_in_point.add(key)
            machine_points[key].append(dict(sno=pt['sno'], category=pt['category'],
                                             actionPoint=pt['actionPoint'], matched=m.group(0)))
    cross_cutting = []
    for k, v in machine_points.items():
        cats = sorted(set(x['category'] for x in v))
        if 2 <= len(v) <= 4 and len(cats) >= 2:
            cross_cutting.append(dict(machine=k, points=v, categories=cats))
    cross_cutting.sort(key=lambda x: (-len(x['categories']), -len(x['points'])))

    as_of_dt = _ap_parse_short_date(*review_date_labels[-1].split('.')) if review_date_labels else None
    upcoming = []
    if as_of_dt:
        window_end = as_of_dt + datetime.timedelta(days=45)
        for pt in points:
            t = pt.get('latestText') or ''
            if not t:
                continue
            for m in _AP_DATE_RE.finditer(t):
                d = _ap_parse_short_date(*m.groups())
                if d and as_of_dt <= d <= window_end:
                    start = max(0, m.start() - 70)
                    snippet = re.sub(r'\s+', ' ', t[start:m.start()]).strip()
                    upcoming.append(dict(sno=pt['sno'], category=pt['category'],
                                          actionPoint=pt['actionPoint'], date=d.strftime('%d-%b-%Y'),
                                          dateSort=d.strftime('%Y-%m-%d'), snippet=snippet[-90:]))
        upcoming.sort(key=lambda x: x['dateSort'])
        seen = set()
        deduped = []
        for u in upcoming:
            k = (u['sno'], u['dateSort'])
            if k in seen:
                continue
            seen.add(k)
            deduped.append(u)
        upcoming = deduped[:15]

    cat_this = defaultdict(int)
    cat_prev = defaultdict(int)
    for pt in points:
        h = pt['history']
        if h and h[-1]['date'] == latest_label:
            cat_this[pt['category']] += len(h[-1]['text'])
        if prev_label:
            prev_entry = next((e for e in h if e['date'] == prev_label), None)
            if prev_entry:
                cat_prev[pt['category']] += len(prev_entry['text'])
    cat_names = sorted(set(list(cat_this.keys()) + list(cat_prev.keys())))
    category_momentum = []
    for c in cat_names:
        cur, prev = cat_this.get(c, 0), cat_prev.get(c, 0)
        pct = round((cur - prev) / prev * 100, 0) if prev else None
        category_momentum.append(dict(category=c, thisCycleChars=cur, prevCycleChars=prev, pctChange=pct))

    sup_points = defaultdict(set)
    for pt in points:
        for name in (pt.get('supervisors') or {}).values():
            if name:
                sup_points[name].add(pt['sno'])
    supervisor_load = sorted(
        [dict(name=k, pointCount=len(v)) for k, v in sup_points.items()],
        key=lambda x: -x['pointCount'])[:8]

    return dict(
        momentum=dict(substantiveUpdates=len(substantive), repeatUpdates=len(repeat),
                       notUpdated=len(not_updated)),
        stalledPoints=[dict(sno=p['sno'], category=p['category'], actionPoint=p['actionPoint'],
                             latestDate=p['latestDate']) for p in stalled],
        repeatPoints=[dict(sno=p['sno'], category=p['category'], actionPoint=p['actionPoint'])
                      for p in repeat],
        flaggedNilPoints=[dict(sno=p['sno'], category=p['category'], actionPoint=p['actionPoint'],
                                text=p['latestText']) for p in flagged_nil],
        crossCuttingMachines=cross_cutting[:10],
        upcomingDates=upcoming,
        categoryMomentum=category_momentum,
        supervisorLoad=supervisor_load,
    )
