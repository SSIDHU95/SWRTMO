"""Adapter for the `indent_position` table.

Live sync: indent_sync.gs. One row, pushed daily, shaped:
    {
      "summary": [{"title": "Indent position as on 26.10.2026",
                    "rows": [{"store": "UBL", "Under Vetting": 3, ...}]},
                   {"title": "Items position as on 26.10.2026", "rows": [...]}],
      "safetyItems": {"title": "...", "rows": [{<header-text>: value, ...}]},
      "stages": {
        "Tender to be published": {"title": "...", "totalRowText": [...],
                                    "rows": [{<header-text>: value, ...}]},
        "Under Vetting": {...}, ... (9 stage sheets, exact Google-Sheet names
        as keys - including the typos and trailing space in
        "Indent under process ")
      }
    }
Unlike every other adapter so far, the live sync already builds each row
as a NAME-keyed object (parseGenericSheet()'s `obj[header] = value`),
not a fixed-position array - so this adapter's job is purely a rename:
map each sheet's own header text onto the internal field names
run_pipeline.py's extract_indent_position(base, as_of) used (sno,
description, indentValue, etc.), then run the same money/date formatting
those fields got there (a 'money' field becomes (text, numericValue), a
'date' field becomes (text, isoDateOrNone)).

The header-text mapping below was built two ways: for columns actually
populated in the live row sampled during discovery (31-Oct-2026), the
exact live header string is used; for a column that happened to be
blank in every sampled row (indent_sync.gs's parseGenericSheet() omits a
blank cell's key entirely, so a sparsely-filled column's exact header
text isn't always visible from one snapshot), the display text from
run_pipeline.py's own IP_FIELD_LABELS (dashboard_template2.html) is used
as the best available stand-in. Task #57 (dry-run against live data)
should re-check these once a row that actually populates every column
comes through - a wrong guess here just means that one field stays None,
never a crash or a misattributed value, since every lookup is a plain
dict .get() by name.
"""
import re

import pipeline_helpers as ph
import supabase_client

TABLE = "indent_position"

INDENT_STAGE_KEYS = [
    'tenderToBePublished', 'tenderUnderFinalisation', 'poIssuedPendingReceipt',
    'materialReceived', 'underVetting', 'underConcurrence', 'indentSigning',
    'balanceIndentsToBeProposed', 'indentUnderProcess',
]
INDENT_STAGE_LABELS = {
    'tenderToBePublished': 'Tender to be Published',
    'tenderUnderFinalisation': 'Tender Under Finalisation',
    'poIssuedPendingReceipt': 'PO Issued Pending Receipt',
    'materialReceived': 'Material Received',
    'underVetting': 'Under Vetting',
    'underConcurrence': 'Under Concurrence',
    'indentSigning': 'Indent Signing',
    'balanceIndentsToBeProposed': 'Balance Indents to be Proposed',
    'indentUnderProcess': 'Indent Proposal Under Process',
}
# Exact live Google-Sheet tab names (indent_sync.gs's STAGE_SHEET_NAMES) -
# these are the keys under payload["stages"].
INDENT_STAGE_SHEETS = {
    'tenderToBePublished': 'Tender to be published',
    'tenderUnderFinalisation': 'Tender Under Finalisation',
    'poIssuedPendingReceipt': 'Po Issued Pending Reciept',
    'materialReceived': 'Material Recieved',
    'underVetting': 'Under Vetting',
    'underConcurrence': 'Under Concurrence',
    'indentSigning': 'Indent Signing',
    'balanceIndentsToBeProposed': 'Balance Indents to be proposed',
    'indentUnderProcess': 'Indent under process ',
}

# Per-stage column map: (jsonKey, liveHeaderText, kind). kind matches
# run_pipeline.py's _ip_cell() contract: 'money' -> (text, num), 'date' ->
# (text, iso), 'int'/'num'/'raw'/'text' -> a single value.
INDENT_STAGE_COLUMNS = {
    'tenderToBePublished': [
        ('sno', 'Sl. No.', 'int'), ('description', 'Description of Items', 'text'),
        ('itemCategory', 'Item Category', 'text'),
        ('noOfItems', 'No. of Items for which tender to be published', 'num'),
        ('qtyRequired', 'Quantity Req', 'raw'), ('indentValue', 'Indent Value', 'money'),
        ('indentNo', 'Indent No', 'text'), ('sentToPurchaseUnit', 'Sent to Purchase unit', 'date'),
        ('durationInCurrentStatus', ('Duration in Current Status', 'Duration from Current Status', 'Days since current status'), 'raw'),
        ('purchaseUnit', 'Purchase Unit', 'text'), ('remarks', 'Remarks', 'text'),
        ('machines', ('Machines', 'Machine'), 'text'),
    ],
    'tenderUnderFinalisation': [
        ('sno', 'Sl. No.', 'int'), ('description', 'Description of Items', 'text'),
        ('itemCategory', 'Item Category', 'text'), ('noOfItems', 'No. of Items for which  tender is floated and to be finalized', 'num'),
        ('qtyRequired', 'Quantity Required', 'raw'), ('indentValue', 'Indent Value', 'money'),
        ('indentNo', 'Indent No', 'text'), ('indentDate', 'Indent Date', 'date'),
        ('likelySupplier', 'Likely Supplier', 'text'), ('tenderNo', 'Tender No', 'raw'),
        ('tenderPublishedDate', 'Tender Published Date', 'date'),
        ('tenderOpeningDate', 'Tender Opening Date', 'date'),
        ('daysPendingSinceOpening', 'Days pending since opening', 'raw'),
        ('purchaseUnit', 'Purchase Unit', 'text'), ('remarks', 'Remarks', 'text'),
        ('machines', ('Machines', 'Machine'), 'text'),
    ],
    'poIssuedPendingReceipt': [
        ('sno', 'Sl. No.', 'int'), ('description', 'Description of Items', 'text'),
        ('demandRaisedDate', 'Demand raised date', 'date'), ('indentDate', 'Indent date', 'text'),
        ('noOfItems', 'No. of Items in Indent', 'num'), ('itemCategory', 'Item Category', 'text'),
        ('qtyRequired', 'Quantity Required', 'raw'), ('purchaseUnit', 'Purchase Unit', 'text'),
        ('poNumber', 'PO Number', 'raw'), ('poDate', ('PO/modification Date', 'PO/modified PO Date', 'PO Date'), 'date'),
        ('poValue', ('PO VALUE', 'PO Value'), 'money'),
        ('daysSincePo', ('Days Passed since PO issued', 'Days passed since PO issue'), 'raw'),
        ('deliveryDateQty', ('Deilvery date/quantity', 'Delivery date/quantity'), 'text'),
        ('machines', ('Machines', 'Machine'), 'text'), ('remarks', 'Remarks', 'text'),
        ('firmNameLocation', 'Firm Name with Location & Contact', 'text'),
        ('firmContact', 'Firm Contact Details With Name and No.', 'text'),
        ('dueDeliveryDate', 'Due Delivery date', 'date'),
    ],
    'materialReceived': [
        ('sno', 'Sl. No.', 'int'), ('description', 'Description of Items', 'text'),
        ('noOfItems', 'No. of Items', 'num'), ('poNumber', 'PO Number', 'raw'),
        ('poDate', 'PO Date', 'date'), ('qualityReview', 'Quality Review', 'text'),
        ('latestRemarks', 'Latest Remarks /  Entry in UDM in case of  Bad Quality material', 'text'),
    ],
    'underVetting': [
        ('sno', 'Sl. No.', 'int'), ('description', 'Description of Items', 'text'),
        ('noOfItems', 'No. of Items in Indent', 'num'), ('itemCategory', 'Item Category', 'text'),
        ('purchaseUnit', 'Purchase Unit', 'text'), ('indentNo', 'Indent No', 'text'),
        ('indentValue', 'Indent Value', 'money'), ('indentDate', 'Indent Date', 'date'),
        ('sentOn', 'Sent for Vetting On', 'date'),
        ('daysPending', 'Days for which Pending under vetting', 'raw'),
        ('tdc', 'TDC for completion of vetting', 'text'),
        ('machines', ('Machine name', 'Machine', 'Machines'), 'text'), ('remarks', 'Remarks', 'text'),
    ],
    'underConcurrence': [
        ('sno', 'Sl. No.', 'int'), ('description', 'Description of Items', 'text'),
        ('noOfItems', 'No. of Items in Indent', 'num'), ('itemCategory', 'Item Category', 'text'),
        ('purchaseUnit', 'Purchase Unit', 'text'), ('indentNo', 'Indent No', 'text'),
        ('indentDate', 'Indent Date', 'date'), ('indentValue', 'Indent Value', 'money'),
        ('sentOn', 'Sent for Concurrence On', 'date'),
        ('daysPending', 'Days for which Pending under concurrence', 'raw'),
        ('tdc', 'TDC for completion of concurrence', 'text'), ('remarks', 'Remarks', 'text'),
        ('machines', ('Machine', 'Machines', 'Machine name'), 'text'),
    ],
    'indentSigning': [
        ('sno', 'Sl. No.', 'int'), ('description', 'Description of Items', 'text'),
        ('noOfItems', 'No. of Items in Indent', 'num'), ('itemCategory', 'Item Category', 'text'),
        ('indentValue', 'Indent Value', 'money'), ('signingAuthority', 'Indent Signing Authority', 'text'),
        ('indentNo', 'Indent No', 'text'), ('indentDate', 'Indent Date', 'date'),
        ('currentHolder', 'Current Holder', 'text'),
        ('sentToCurrentHolderOn', 'Sent to Current Holder on', 'date'),
        ('daysPending', 'No of Days Passed with current holder', 'raw'),
        ('expectedSigningDate', 'Expected Signing Date', 'raw'), ('remarks', 'Remarks', 'text'),
        ('machines', ('Machine', 'Machines', 'Machine name'), 'text'),
    ],
    'balanceIndentsToBeProposed': [
        ('sno', 'Sl. No.', 'int'), ('description', 'Description', 'text'),
        ('noOfItems', 'No. of Items', 'num'), ('itemCategory', 'Item Category', 'text'),
        ('dateOfDemandFromField', 'Date of Demand from Field', 'date'),
        ('daysPending', 'Days Pending', 'num'),
        ('targetDateForProposal', 'Target date for Putting up proposal to AXEN/XEN/DyCe', 'date'),
        ('remarks', 'Remarks', 'text'),
    ],
    'indentUnderProcess': [
        ('sno', 'Sl. No.', 'int'), ('description', 'Description', 'text'),
        ('itemCategory', 'Item Category', 'text'), ('items', 'Items', 'raw'),
        ('indentNoIfPlaced', 'Indent No', 'raw'), ('indentValue', 'Indent Val', 'money'),
        ('approvingAuthority', 'Appr Auth', 'text'), ('vettingRequired', 'Vetting Required', 'text'),
        ('concurrenceRequired', 'Concurrence Required', 'text'),
        ('signingAuthority', 'Indent Signing Authority', 'text'),
        ('pacSigningAuthority', 'PAC signing Authority', 'text'),
        ('eofficeRequired', 'E-Office Required', 'text'),
        ('currentHolderEoffice', 'Current Holder in e-office/UDM', 'text'),
    ],
}
SAFETY_ITEMS_COLUMNS = [
    ('sno', 'Sl. No.', 'int'), ('description', 'Description of Items', 'text'),
    ('noOfItems', 'No. of Items in Indent', 'num'), ('itemCategory', 'Item Category', 'text'),
    ('qtyRequired', 'Quantity Required', 'raw'), ('indentValue', 'Indent Value', 'money'),
    ('indentNo', 'Indent No', 'text'), ('indentDate', 'Indent Date', 'date'),
    ('poIssueDateOrStatus', 'PO issue date', 'raw'),
    ('firmNameLocation', 'Firm Name with Location & Contact', 'text'),
    ('firmContact', 'Firm Contact Details With Name and No.', 'text'),
    ('currentStatus', 'Current Status', 'text'), ('currentHolder', 'Current Holder', 'text'),
]

# Keyword stems used to match a Summary-sheet column header (which varies
# in wording between the by-indent-count block, "Tender to Be Published",
# and the by-item-count block, "Items to be Published", and in casing/
# spacing from INDENT_STAGE_LABELS) onto its stage key. Matched against a
# whitespace-collapsed, lowercased header - order matters (more specific
# stems first) since e.g. "finalisation" must be checked before a looser
# "tender" stem would also match it.
_SUMMARY_STAGE_STEMS = [
    ('tenderUnderFinalisation', ('finalis',)),
    ('tenderToBePublished', ('publish',)),
    ('poIssuedPendingReceipt', ('po issued', 'pending receipt')),
    ('materialReceived', ('material', 'received')),
    ('underVetting', ('vetting',)),
    ('underConcurrence', ('concurrence',)),
    ('indentSigning', ('signing',)),
    ('balanceIndentsToBeProposed', ('balance',)),
    ('indentUnderProcess', ('under process',)),
]


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def _clean_text(v):
    if v is None:
        return None
    s = re.sub(r'\s+', ' ', str(v)).strip()
    return s or None


_MONEY_RE = re.compile(r'[\d,]+(?:\.\d+)?')


def _parse_money_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return None
    m = _MONEY_RE.search(str(v))
    if not m:
        return None
    try:
        return float(m.group(0).replace(',', ''))
    except ValueError:
        return None


def _ip_date(v):
    if v is None or v == '':
        return None, None
    text = _clean_text(v)
    dt = ph.parse_ddmmyyyy(text) if text else None
    return text, (dt.date().isoformat() if dt else None)


def _ip_cell(v, kind):
    if kind == 'money':
        return _clean_text(v), _parse_money_num(v)
    if kind == 'date':
        return _ip_date(v)
    if kind == 'int':
        return int(v) if isinstance(v, (int, float)) else None
    if kind == 'num':
        return v if isinstance(v, (int, float)) else (_clean_text(v) if v is not None else None)
    if kind == 'raw':
        return v if isinstance(v, (int, float)) else _clean_text(v)
    return _clean_text(v)  # 'text'


def _norm_header(h):
    return re.sub(r'\s+', ' ', str(h or '')).strip().lower()


def _build_rows(sheet_rows, col_map):
    """sheet_rows: the live payload's list of name-keyed row dicts for one
    sheet. col_map: one of the INDENT_STAGE_COLUMNS[...]/SAFETY_ITEMS_COLUMNS
    lists above. -> [{jsonKey: value, ...}, ...], with a 'money' or 'date'
    column contributing BOTH its plain key and a Num/Iso companion key,
    exactly like run_pipeline.py's _ip_rows()."""
    out = []
    for raw_row in sheet_rows or []:
        row = {}
        # header lookup tolerant of case/spacing and of the alternative
        # wordings listed in the column map (the sheet's headers get edited
        # from time to time -- e.g. "PO/modification Date" became "PO/modified
        # PO Date", which silently blanked that column until 27-Nov-2026)
        norm = {_norm_header(k): v for k, v in (raw_row or {}).items()}
        for key, header, kind in col_map:
            alts = header if isinstance(header, (tuple, list)) else (header,)
            v = None
            for h in alts:
                if h in raw_row:
                    v = raw_row[h]
                elif _norm_header(h) in norm:
                    v = norm[_norm_header(h)]
                else:
                    continue
                if v not in (None, ''):
                    break
            if kind in ('money', 'date'):
                row[key], row[key + ('Num' if kind == 'money' else 'Iso')] = _ip_cell(v, kind)
            else:
                row[key] = _ip_cell(v, kind)
        out.append(row)
    return out


def _total_counts(total_row_text):
    """total_row_text is the live sync's `totalRowText` array (whatever it
    captured from the TOTAL row's first 5 cells) - mirrors run_pipeline.py's
    _ip_total(total_raw): pulls an explicit 'Indents: N'/'Main Indents: N'
    count from any text cell, and the item count from the last numeric
    cell in the row (that position varies stage to stage - some stages put
    the item total in column D, others in column E, which is exactly why
    the local-file version takes "the last number in the row" rather than
    a fixed column)."""
    if not total_row_text:
        return None, None
    vals = [v for v in total_row_text if v not in (None, '')]
    indent_count = None
    for v in vals:
        if isinstance(v, str):
            m = re.search(r'Indent[s1]?:\s*(\d+)', v, re.I)
            if m:
                indent_count = int(m.group(1))
    item_count = None
    for v in reversed(vals):
        if isinstance(v, (int, float)):
            item_count = int(v)
            break
    return indent_count, item_count


def extract_stages(payload):
    """-> {stageKey: {label, rows, indentCount, itemCount}, ...} for all 9
    INDENT_STAGE_KEYS, matching run_pipeline.py's per-stage output exactly
    (indentCount falls back to len(rows), itemCount falls back to summing
    noOfItems across rows, same as the local-file version)."""
    stages_in = (payload or {}).get("stages") or {}
    out = {}
    for key in INDENT_STAGE_KEYS:
        sheet_name = INDENT_STAGE_SHEETS[key]
        label = INDENT_STAGE_LABELS[key]
        block = stages_in.get(sheet_name)
        if not block:
            out[key] = dict(label=label, rows=[], indentCount=0, itemCount=0)
            continue
        rows = _build_rows(block.get("rows") or [], INDENT_STAGE_COLUMNS[key])
        # drop the sheet's own TOTAL line (e.g. 'Main indents: 04'), which the
        # sync sometimes passes through as a data row with no S.No/description
        rows = [r for r in rows if r.get('sno') is not None or r.get('description')]
        indent_count, item_count = _total_counts(block.get("totalRowText"))
        if indent_count is None:
            indent_count = len(rows)
        if item_count is None:
            item_count = sum(
                (r['noOfItems'] if isinstance(r.get('noOfItems'), (int, float)) else 1) for r in rows
            )
        out[key] = dict(label=label, rows=rows, indentCount=indent_count, itemCount=item_count)
    return out


def _match_stage_key(header):
    h = re.sub(r'\s+', ' ', str(header)).strip().lower()
    for key, stems in _SUMMARY_STAGE_STEMS:
        if all(stem in h for stem in stems):
            return key
    return None


def _summary_table(summary_rows):
    out = []
    for row in summary_rows or []:
        store = row.get("store")
        if not store:
            continue
        stages = {key: 0 for key in INDENT_STAGE_KEYS}
        total = 0
        for header, v in row.items():
            if header == "store":
                continue
            if not isinstance(v, (int, float)):
                continue
            hl = re.sub(r'\s+', ' ', str(header)).strip().lower()
            # "Total" / "Total Indents of SWR" / "Total items" -- the Summary
            # sheet's own total column, whatever its exact wording.
            if hl == "total" or hl.startswith("total "):
                total = v
                continue
            key = _match_stage_key(header)
            if key:
                stages[key] = v
        # Fixed 27-Nov-2026: the live sync's Summary rows carry only the 9
        # stage columns, not column B ("Total Indents of SWR" / "Total
        # items"), so the website's "Total Indents" tile read 0. On the
        # sheet itself that column is exactly the sum of the 9 stages
        # (checked against Shruti's INDENT_POSITION.xlsx: UBL 114, SBC 19,
        # TOTAL 133 indents; 655/21/676 items) -- so when no explicit total
        # arrives, use that same sum rather than showing 0.
        if not total:
            total = sum(v for v in stages.values() if isinstance(v, (int, float)))
        out.append(dict(store=_clean_text(store), total=total, stages=stages))
    return out


def _as_on_date(title):
    """'Indent position as on 16.11.2026' -> '16.11.2026' (the sheet's own
    date for its Summary figures), else None."""
    m = re.search(r'as\s+on\s+([0-9]{1,2}[./-][0-9]{1,2}[./-][0-9]{2,4})', str(title or ''), re.I)
    return m.group(1) if m else None


def extract_summary(payload):
    """-> {stageKeys, stageLabels, byIndent, byItem}, matching
    run_pipeline.py's summary block shape. The live sync's two "as on"
    blocks come through in sheet order (by-indent-count table first,
    by-item-count table second, per parseSummarySheet()'s top-to-bottom
    scan) - picked here by matching each block's own title text rather
    than assuming that order holds, so a future re-layout of the Summary
    sheet degrades gracefully instead of silently swapping the two
    tables."""
    blocks = (payload or {}).get("summary") or []
    by_indent, by_item = [], []
    sheet_as_on = None
    for block in blocks:
        raw_title = block.get("title") or ""
        title = str(raw_title).lower()
        rows = _summary_table(block.get("rows"))
        if "item" in title:
            by_item = rows
        else:
            by_indent = rows
            sheet_as_on = _as_on_date(raw_title) or sheet_as_on
    return dict(stageKeys=INDENT_STAGE_KEYS, stageLabels=INDENT_STAGE_LABELS,
                byIndent=by_indent, byItem=by_item, sheetAsOn=sheet_as_on)


def extract_safety_items(payload):
    """-> [row, ...], matching run_pipeline.py's safety_items list."""
    block = (payload or {}).get("safetyItems") or {}
    return _build_rows(block.get("rows") or [], SAFETY_ITEMS_COLUMNS)



# ---------------------------------------------------------------------
# "Days since current status" (added 27-Nov-2026 per Shruti: it's the key
# figure in every stage list, sorted descending). The sheet's own day
# columns are unreliable as delivered by the sync -- some are formatted as
# DATES in the sheet, so "199 days" arrives as "17.07.1900"; some are typed
# 0 or left stale; one stage's header was renamed. So, per stage, the
# figure is worked out the same way the sheet's own TODAY()-date formulas
# do: days from the date the indent entered its current status to the run
# date. Only when that date is missing or unusable (blank, typo'd into the
# future) does it fall back to the sheet's own number (decoded if it came
# through as a 1899/1900 "date").
# ---------------------------------------------------------------------
_STATUS_START_FIELD = {
    'tenderToBePublished': ('sentToPurchaseUnit', 'durationInCurrentStatus'),
    'tenderUnderFinalisation': ('tenderOpeningDate', 'daysPendingSinceOpening'),
    'poIssuedPendingReceipt': ('poDate', 'daysSincePo'),
    'underVetting': ('sentOn', 'daysPending'),
    'underConcurrence': ('sentOn', 'daysPending'),
    'indentSigning': ('sentToCurrentHolderOn', 'daysPending'),
}


def _sheet_days(v):
    """The sheet's own day count -> int, or None. Handles real numbers,
    numeric text, and day counts that arrived formatted as dates near the
    spreadsheet epoch ('17.07.1900' == 199 days, '31.12.1899' == 1 day)."""
    import datetime
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(round(v)) if v >= 0 else None
    t = _clean_text(v)
    if not t:
        return None
    if re.fullmatch(r'\d+(\.\d+)?', t):
        return int(round(float(t)))
    m = re.fullmatch(r'(\d{1,2})[./-](\d{1,2})[./-](1899|1900|1901)', t)
    if m:
        try:
            d = datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            return (d - datetime.date(1899, 12, 30)).days
        except ValueError:
            return None
    return None


def _first_date(text):
    import datetime
    m = re.search(r'(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})', str(text or ''))
    if not m:
        return None
    dd, mm, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yy < 100:
        yy += 2000
    try:
        return datetime.date(yy, mm, dd)
    except ValueError:
        return None


def add_days_in_status(stages, as_of):
    """Adds daysInStatus (int or None), statusSince (the date text used) and
    daysNote (short text when there is no day count, e.g. a tender not yet
    opened) to every row of the stages that have a status-start date."""
    import datetime
    for key, (date_field, sheet_field) in _STATUS_START_FIELD.items():
        for r in (stages.get(key) or {}).get('rows') or []:
            iso = r.get(date_field + 'Iso')
            days, since, note = None, None, None
            d = None
            if iso:
                try:
                    d = datetime.date.fromisoformat(iso)
                except ValueError:
                    d = None
            if d is None:
                # a cell holding two dates (e.g. PO and modified PO:
                # '08/10/26 23/10/26') -- the first is when this status began
                d = _first_date(r.get(date_field))
            if d is not None and key == 'tenderUnderFinalisation' and d > as_of:
                note = f"opens {r.get(date_field)}"
            elif d is not None and d <= as_of:
                days, since = (as_of - d).days, r.get(date_field)
            if days is None and note is None:
                days = _sheet_days(r.get(sheet_field))
                if days is not None and key == 'tenderUnderFinalisation' and days == 0:
                    days = None  # 0 here means "not yet opened"
            r['daysInStatus'] = days
            r['statusSince'] = since
            r['daysNote'] = note
    return stages

def extract_indent_position(payload, as_of):
    """-> {sourceFile, asOf, summary, safetyItems, stages, stageOrder,
    attention}, the same top-level shape run_pipeline.py's
    extract_indent_position(base, as_of) returns. Assembles the three
    pieces above and ports that function's aging/"Needs Attention"
    synthesis verbatim (it's a pure function of the already-extracted
    stage rows, so it carries over unchanged regardless of where those
    rows came from)."""
    summary = extract_summary(payload)
    safety_items = extract_safety_items(payload)
    stages = extract_stages(payload)
    add_days_in_status(stages, as_of)

    import datetime

    attention = []

    def _overdue_days(iso):
        if not iso:
            return None
        try:
            d = datetime.date.fromisoformat(iso)
        except Exception:
            return None
        return (as_of - d).days

    for key, thr_med, thr_high in (('underVetting', 15, 30), ('underConcurrence', 15, 30), ('indentSigning', 20, 40)):
        for r in stages[key]['rows']:
            days = r.get('daysPending')
            if isinstance(days, (int, float)) and days >= thr_med:
                attention.append(dict(
                    severity=('high' if days >= thr_high else 'medium'), stageKey=key, stageLabel=stages[key]['label'],
                    sno=r.get('sno'), days=int(days),
                    text=f"{(r.get('description') or 'Indent')[:90]} — pending {int(days)} days at {stages[key]['label']}.",
                ))
    for r in stages['tenderUnderFinalisation']['rows']:
        days = r.get('daysPendingSinceOpening')
        if isinstance(days, (int, float)) and days >= 45:
            attention.append(dict(
                severity=('high' if days >= 90 else 'medium'), stageKey='tenderUnderFinalisation',
                stageLabel=stages['tenderUnderFinalisation']['label'], sno=r.get('sno'), days=int(days),
                text=f"{(r.get('description') or 'Tender')[:90]} — {int(days)} days since tender opening, not yet finalized.",
            ))
    for r in stages['tenderToBePublished']['rows']:
        days = r.get('durationInCurrentStatus')
        if isinstance(days, (int, float)) and days >= 180:
            attention.append(dict(
                severity=('high' if days >= 270 else 'medium'), stageKey='tenderToBePublished',
                stageLabel=stages['tenderToBePublished']['label'], sno=r.get('sno'), days=int(days),
                text=f"{(r.get('description') or 'Indent')[:90]} — {int(days)} days since sent to purchase unit, tender not yet published.",
            ))
    for r in stages['poIssuedPendingReceipt']['rows']:
        od = _overdue_days(r.get('dueDeliveryDateIso'))
        if od is not None and od > 0:
            attention.append(dict(
                severity=('high' if od >= 60 else 'medium'), stageKey='poIssuedPendingReceipt',
                stageLabel=stages['poIssuedPendingReceipt']['label'], sno=r.get('sno'), days=od,
                text=f"{(r.get('description') or 'PO')[:90]} — delivery due date passed {od} days ago, material not yet received.",
            ))
    for r in stages['balanceIndentsToBeProposed']['rows']:
        od = _overdue_days(r.get('targetDateForProposalIso'))
        if od is not None and od > 0:
            attention.append(dict(
                severity=('high' if od >= 30 else 'medium'), stageKey='balanceIndentsToBeProposed',
                stageLabel=stages['balanceIndentsToBeProposed']['label'], sno=r.get('sno'), days=od,
                text=f"{(r.get('description') or 'Demand')[:90]} — target date for proposal passed {od} days ago, indent not yet raised.",
            ))
    sev_rank = dict(high=0, medium=1, low=2)
    attention.sort(key=lambda x: (sev_rank.get(x['severity'], 9), -x.get('days', 0)))

    return dict(
        sourceFile="live sync (indent_position)", asOf=as_of.strftime('%d.%m.%Y'),
        summary=summary, safetyItems=safety_items, stages=stages,
        stageOrder=INDENT_STAGE_KEYS, attention=attention[:40],
    )
