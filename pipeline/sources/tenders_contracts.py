"""Adapter for the `tenders_contracts` table.

Live sync: tenders_contracts_sync.gs. One row, pushed daily, covering the
5 sheets run_pipeline.py's extract_tenders_contracts(base, as_of) reads
(Agreement Details, Contact Person, Status of AMC Works Proposal, Tender
Under Finalization, LOA issued) PLUS a 6th, "Tender to be floated", that
nothing reads yet (not in input-data-specification.md at all - synced
anyway per the user, same as sidings_rest_house's 7 unused supplementary
tabs). Payload shape, per sheet:
    {
      "agreementDetails":       {"headers": [...], "rows": [{<header text
          as the key>: value, ...}, ...]},
      "contactPerson":          {"headers": [...], "rows": [{..., "group":
          "OEMs/ Authorised dealers of different machines"|"...of engines"}]},
      "amcWorksProposal":       {"headers": [...], "rows": [{..., "group":
          "Works Contract"|"AMC"}]},
      "tenderUnderFinalization":{"headers": [...], "rows": [...]},
      "tenderToBeFloated":      {"headers": [...], "rows": [...]},   # unused
      "loaIssued":              {"headers": [...], "rows": [{..., "FY":
          "FY 2026-27"}]}
    }

Unlike every other adapter in this project, this live sync keys each row
by its OWN header text (confirmed: the two "Status of AMC Works Proposal"
groups actually use slightly different header wording for the same
logical column, e.g. "Currency of current works" vs "Currency of current
AMC", "S.No." vs "S.No", "Items/ Number of machines" vs "Item/ no of
machines") rather than by a fixed column index the way run_pipeline.py's
own openpyxl-based extraction does - so _g() below tries every header
spelling actually seen in the live data for a given logical field, instead
of ws.cell(row, column). Everything downstream (expiry-status/stage/brand
classification, the cross-sheet renewal-proposal linking, the KPI block
and the "Needs Attention" synthesis) is ported verbatim from
run_pipeline.py's extract_tenders_contracts(), since none of that depends
on how the raw sheet was read.
"""
import datetime
import re
from collections import defaultdict

import pipeline_helpers as ph
import supabase_client

TABLE = "tenders_contracts"


def fetch_payload():
    return supabase_client.fetch_latest_payload(TABLE)


def _g(row, *keys):
    """First non-blank value among the given header-text spellings."""
    for k in keys:
        v = row.get(k)
        if v not in (None, ''):
            return v
    return None


def _fmt_date(v):
    """A live-sync date cell is already 'dd.mm.yyyy' text (cellVal()'s own
    Utilities.formatDate call) - reformat to 'dd-Mon-YYYY' to match
    run_pipeline.py's strftime('%d-%b-%Y') display convention. Free text
    that isn't a real date (e.g. '-') passes through via clean_text()."""
    dt = ph.parse_ddmmyyyy(v)
    return dt.strftime('%d-%b-%Y') if dt else ph.clean_text(v)


def _extract_agreements(rows, as_of):
    agreements = []
    for r in rows:
        sno = _g(r, 'S No', 'S.No', 'S.No.', 'Sl.No.', 'Sl No', 'Sl.No')
        firm = ph.clean_text(_g(r, 'Firm Name'))
        if not isinstance(sno, (int, float)) or not firm:
            continue
        ctype = ph.clean_text(_g(r, 'Contract Type'))
        loa_no = ph.clean_text(_g(r, 'LOA/Agreement No and Date'))
        allocation = ph.clean_text(_g(r, 'Allocation'))
        work = ph.clean_text(_g(r, 'Name of work'))
        loa_date_raw = _g(r, 'Date of Issue of LOA')
        amount = ph.num(_g(r, 'Agreement Amount'))
        revised = ph.num(_g(r, 'Revised agreement value'))
        orig_curr = ph.clean_text(_g(r, 'Original Currency Date'))
        ext_curr = ph.clean_text(_g(r, 'Extended Currency Date'))
        latest_bill = ph.clean_text(_g(r, 'Latest Passed Bill (Rs.)'))
        cum_bill = ph.num(_g(r, 'Cumulative bill (Rs.)'))
        balance = ph.num(_g(
            r,
            'Balance (Rs.) ( Agreement Balance-Pending with Acc-Bill under eMb- Bills waiting)',
            'Agreement Balance (Rs.) (Agm Value-Cumulative bill)',
        ))
        remarks_tmo = ph.clean_text(_g(r, 'Remarks from TMO/Accounts'))
        remarks_sse = ph.clean_text(_g(r, 'Remarks from SSE/Proposals'))

        eff_date = ph.parse_ddmmyyyy(ext_curr) or ph.parse_ddmmyyyy(orig_curr)
        days_to_expiry = (eff_date.date() - as_of).days if eff_date else None
        if days_to_expiry is None:
            status = 'Unknown'
        elif days_to_expiry < 0:
            status = 'Expired'
        elif days_to_expiry <= 90:
            status = 'Expiring Soon'
        else:
            status = 'Valid'

        eff_value = revised if revised is not None else amount
        brand_ids = ph.brand_hits(f"{firm} {work}")

        agreements.append(dict(
            sno=int(sno), contractType=ctype, loaNoAndDate=loa_no, allocation=allocation,
            nameOfWork=work, firmName=firm,
            dateOfIssueOfLoa=_fmt_date(loa_date_raw),
            agreementAmount=amount, revisedAgreementValue=revised, effectiveValue=eff_value,
            originalCurrencyDate=orig_curr, extendedCurrencyDate=ext_curr,
            effectiveCurrencyDate=(eff_date.strftime('%d-%b-%Y') if eff_date else None),
            daysToExpiry=days_to_expiry, expiryStatus=status,
            latestPassedBill=latest_bill, cumulativeBill=cum_bill, balance=balance,
            registeredAmount=None, remarksTmoAccounts=remarks_tmo, remarksSseProposals=remarks_sse,
            brandIds=brand_ids, renewalProposals=[],
        ))
    return agreements


_PHONE_TRAILING_RE = re.compile(r'([\d][\d\s\-]{7,})\s*$')
_ROLE_PAREN_RE = re.compile(r'\(([^)]+)\)')


def _parse_contact_cell(txt):
    phone = None
    namepart = txt
    m = _PHONE_TRAILING_RE.search(txt)
    if m:
        phone = re.sub(r'\s+', '', m.group(1))
        namepart = txt[:m.start()]
    role = None
    rm = _ROLE_PAREN_RE.search(namepart)
    if rm:
        role = ph.clean_text(rm.group(1))
        namepart = namepart[:rm.start()] + namepart[rm.end():]
    namepart = ph.clean_text(namepart)
    if namepart:
        namepart = re.sub(r'[\s:\-]+$', '', namepart).strip()
    return dict(name=namepart, role=role, phone=phone, raw=txt)


def _extract_contacts(rows):
    contacts = []
    for r in rows:
        sno = _g(r, 'Sl.No.', 'Sl.No', 'Sl No', 'S.No.', 'S.No', 'S No')
        amc_name = ph.clean_text(_g(r, 'Name of the AMC'))
        if not isinstance(sno, (int, float)) or not amc_name:
            continue
        persons = []
        for key in ('Contact person Name 1', 'Contact person Name 2', 'Contact person Name 3'):
            raw = r.get(key)
            if raw in (None, '', '-'):
                continue
            txt = ph.clean_text(raw)
            if txt:
                persons.append(_parse_contact_cell(txt))
        contacts.append(dict(sno=int(sno), group=r.get('group'), amcName=amc_name,
                              persons=persons, brandIds=ph.brand_hits(amc_name)))
    return contacts


def _extract_proposals(rows):
    proposals = []
    for r in rows:
        sno = _g(r, 'S.No.', 'S.No', 'S No', 'Sl.No.', 'Sl No')
        if not isinstance(sno, (int, float)):
            continue
        section = r.get('group')
        desc = ph.clean_text(_g(r, 'Description of Item'))
        items = _g(r, 'Items/ Number of machines', 'Item/ no of machines')
        value_raw = _g(r, 'Value')
        value_num = ph.num(value_raw)
        remark = ph.clean_text(_g(r, 'Remarks'))
        currency = ph.clean_text(_g(r, 'Currency of current works', 'Currency of current AMC',
                                      'Currency of current'))
        stage_key, stage_label = ph.classify_proposal_stage(remark)
        dates = ph.dates_in_text(remark)
        proposals.append(dict(
            section=section, sno=int(sno), description=desc, items=items,
            value=(value_num if value_num is not None else ph.clean_text(value_raw)), valueNum=value_num,
            remark=remark, currencyOfCurrent=currency,
            stageKey=stage_key, stageLabel=stage_label,
            lastDatedMention=(max(dates).strftime('%d-%b-%Y') if dates else None),
            brandIds=ph.brand_hits(desc),
        ))
    return proposals


def _extract_tenders(rows, as_of):
    tenders = []
    for r in rows:
        sno = _g(r, 'S No', 'S.No', 'S.No.')
        if not isinstance(sno, (int, float)):
            continue
        works_amc = ph.clean_text(_g(r, 'Works / AMC', 'Works/AMC'))
        ttype = ph.clean_text(_g(r, 'Type of Tender'))
        desc = ph.clean_text(_g(r, 'Description of Item'))
        items = ph.clean_text(_g(r, 'Item/ no of machines'))
        value = ph.clean_text(_g(r, 'Value'))
        date_floated = _g(r, 'Date of floating of tender')
        date_opening = _g(r, 'Date of opening')
        tc_exec = ph.clean_text(_g(r, 'TC Members Executive'))
        tc_fin = ph.clean_text(_g(r, 'TC Members Finance'))
        tc_sis = ph.clean_text(_g(r, 'TC Members Sister'))
        taa = ph.clean_text(_g(r, 'TAA'))
        holder = ph.clean_text(_g(r, 'Current Holder'))
        pending_since = ph.clean_text(_g(r, 'Pending with Current holder since'))
        pending_days_raw = _g(r, 'No of days pending with current holder')
        remark = ph.clean_text(_g(r, 'Remarks'))

        stage_key, stage_label, stage_order = ph.classify_tender_stage(remark, holder)
        dates = ph.dates_in_text(remark)
        last_dt = max(dates) if dates else None
        days_since = (as_of - last_dt.date()).days if last_dt else None

        tenders.append(dict(
            sno=int(sno), worksOrAmc=works_amc, tenderType=ttype, description=desc, items=items,
            value=value, dateFloated=_fmt_date(date_floated), dateOpening=_fmt_date(date_opening),
            tcExecutive=tc_exec, tcFinance=tc_fin, tcSister=tc_sis, taa=taa, currentHolder=holder,
            pendingSinceRaw=pending_since,
            pendingDaysRaw=(pending_days_raw if isinstance(pending_days_raw, (int, float)) else None),
            remark=remark, stageKey=stage_key, stageLabel=stage_label, stageOrder=stage_order,
            lastDatedMention=(last_dt.strftime('%d-%b-%Y') if last_dt else None), daysSinceLastUpdate=days_since,
            brandIds=ph.brand_hits(f"{works_amc} {desc} {remark}"),
        ))
    return tenders


def _extract_loas(rows):
    loas = []
    for r in rows:
        sno = _g(r, 'S No', 'S.No', 'S.No.')
        fy = r.get('FY')
        if not isinstance(sno, (int, float)) or not fy:
            continue
        works_amc = ph.clean_text(_g(r, 'Works/AMC', 'Works / AMC'))
        ttype = ph.clean_text(_g(r, 'Type of Tender'))
        desc = ph.clean_text(_g(r, 'Description of Item'))
        items = ph.clean_text(_g(r, 'Item/ no of machines'))
        value = ph.clean_text(_g(r, 'Value'))
        tc_exec = ph.clean_text(_g(r, 'TC Members Executive'))
        tc_fin = ph.clean_text(_g(r, 'TC Members Finance'))
        tc_sis = ph.clean_text(_g(r, 'TC Members Sister'))
        taa = ph.clean_text(_g(r, 'TAA'))
        loa_on = _g(r, 'LOA issued on')
        remark = ph.clean_text(_g(r, 'Remarks'))
        status_key, status_label = ph.classify_loa_status(remark)
        value_num = ph.num(value)
        loas.append(dict(
            fy=fy, sno=int(sno), worksOrAmc=works_amc, tenderType=ttype, description=desc,
            items=items, value=value, valueNum=value_num,
            tcExecutive=tc_exec, tcFinance=tc_fin, tcSister=tc_sis, taa=taa,
            loaIssuedOn=_fmt_date(loa_on),
            remark=remark, statusKey=status_key, statusLabel=status_label,
            brandIds=ph.brand_hits(f"{works_amc} {desc}"),
        ))
    return loas


def extract_tenders_contracts(payload, as_of):
    """-> the same dict shape run_pipeline.py's extract_tenders_contracts
    (base, as_of) returns (sourceFile, asOf, kpi, agreements, contacts,
    proposals, tenders, loas, attention, brandRegistry), or None if the
    live payload is empty (mirrors the local-file version's own
    "workbook not found" -> None degrade path). `as_of` is a datetime.date,
    same contract as the local-file version's own as_of parameter."""
    if not payload:
        return None

    agreements = _extract_agreements(((payload.get('agreementDetails') or {}).get('rows')) or [], as_of)
    contacts = _extract_contacts(((payload.get('contactPerson') or {}).get('rows')) or [])
    proposals = _extract_proposals(((payload.get('amcWorksProposal') or {}).get('rows')) or [])
    tenders = _extract_tenders(((payload.get('tenderUnderFinalization') or {}).get('rows')) or [], as_of)
    loas = _extract_loas(((payload.get('loaIssued') or {}).get('rows')) or [])

    # ---- Cross-sheet synthesis: renewal signal for expiring/expired agreements ----
    proposals_by_brand = defaultdict(list)
    for p in proposals:
        for bid in p['brandIds']:
            proposals_by_brand[bid].append(p)
    for ag in agreements:
        if ag['expiryStatus'] in ('Expiring Soon', 'Expired'):
            seen = set()
            for bid in ag['brandIds']:
                for p in proposals_by_brand.get(bid, []):
                    pk = (p['section'], p['sno'])
                    if pk not in seen:
                        seen.add(pk)
                        ag['renewalProposals'].append(p)

    kpi = dict(
        totalAgreements=len(agreements),
        activeAgreements=sum(1 for a in agreements if a['expiryStatus'] in ('Valid', 'Expiring Soon')),
        totalCommittedValue=round(sum(a['effectiveValue'] or 0 for a in agreements), 2),
        expiringSoonCount=sum(1 for a in agreements if a['expiryStatus'] == 'Expiring Soon'),
        expiredCount=sum(1 for a in agreements if a['expiryStatus'] == 'Expired'),
        proposalsInPipeline=len(proposals),
        tendersUnderFinalization=len(tenders),
    )
    if loas:
        current_fy_label = loas[0]['fy']
        kpi['loaIssuedCurrentFy'] = sum(1 for l in loas if l['fy'] == current_fy_label)
        kpi['loaIssuedCurrentFyValue'] = round(sum(l['valueNum'] or 0 for l in loas if l['fy'] == current_fy_label), 2)
        kpi['currentFyLabel'] = current_fy_label

    attention = []
    for ag in sorted(agreements, key=lambda a: a['daysToExpiry'] if a['daysToExpiry'] is not None else 99999):
        if ag['expiryStatus'] in ('Expiring Soon', 'Expired') and not ag['renewalProposals']:
            when = (f"expired {abs(ag['daysToExpiry'])} days ago" if ag['expiryStatus'] == 'Expired'
                    else f"expires in {ag['daysToExpiry']} days")
            attention.append(dict(
                severity=('high' if ag['expiryStatus'] == 'Expired' else 'medium'),
                text=f"{ag['firmName']} — {when}, no renewal proposal found in the pipeline.",
                refType='agreement', refSno=ag['sno'],
            ))
    for t in tenders:
        if (t['daysSinceLastUpdate'] or 0) >= 30 and t['stageOrder'] < 5:
            attention.append(dict(
                severity='medium',
                text=f"Tender S.No {t['sno']} ({(t['description'] or '')[:60]}) — no dated update in remarks "
                     f"for {t['daysSinceLastUpdate']} days, currently '{t['stageLabel']}'.",
                refType='tender', refSno=t['sno'],
            ))
    current_fy_label = loas[0]['fy'] if loas else None
    for l in loas:
        if l['fy'] == current_fy_label and l['statusKey'] != 'finalized':
            attention.append(dict(
                severity='low',
                text=f"LOA S.No {l['sno']} ({l['worksOrAmc']}) issued {l['loaIssuedOn']} — agreement not yet "
                     f"finalized ({l['statusLabel']}).",
                refType='loa', refSno=l['sno'],
            ))
    sev_rank = dict(high=0, medium=1, low=2)
    attention.sort(key=lambda x: sev_rank.get(x['severity'], 9))

    return dict(
        sourceFile="live sync (tenders_contracts)", asOf=as_of.strftime('%d-%b-%Y'),
        kpi=kpi, agreements=agreements, contacts=contacts, proposals=proposals,
        tenders=tenders, loas=loas, attention=attention[:20],
        brandRegistry=ph.BRAND_KEYWORDS,
    )
