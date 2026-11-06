"""Read-only operational analytics; first issues and reissues stay distinct."""
from collections import Counter
from datetime import date, datetime, timedelta
from statistics import median
import re

from certificate_core import parse_date, status_value


ISSUED = {'generated', 'generated_manual', 'completed'}
TYPE_NAMES = {'standard': 'Recycling', 'destruction': 'Data Destruction',
              'siemens': 'Medical', 'sita': 'International'}


def _cod(value):
    value = str(value or '').strip().upper()
    return str(int(value)) if re.fullmatch(r'\d+', value) else value


def _day(value):
    if isinstance(value, datetime):
        return value.astimezone().date() if value.tzinfo else value.date()
    if isinstance(value, date):
        return value
    try:
        stamp = datetime.fromisoformat(str(value).replace('Z', '+00:00'))
        return stamp.astimezone().date() if stamp.tzinfo else stamp.date()
    except (ValueError, TypeError):
        return parse_date(value)


def build_analytics(rows, history, *, period='month', plant='all', today=None, scan_complete=True, saved_pdfs=()):
    if period not in {'today', 'week', 'month', 'quarter'}:
        raise ValueError('Choose a valid analytics period.')
    if plant not in {'all', 'plant1', 'plant2'}:
        raise ValueError('Choose a valid analytics plant.')
    today = today or date.today()
    start = {'today': today, 'week': today - timedelta(days=6),
             'month': today.replace(day=1), 'quarter': today - timedelta(days=89)}[period]
    span = (today - start).days + 1
    previous_start, previous_end = start - timedelta(days=span), start - timedelta(days=1)
    rows = [r for r in rows if (r.get('plant') in {'plant1', 'plant2'} or r.get('source_kind') == 'manual') and r.get('source_kind') != 'system']
    all_rows = rows
    rows = [r for r in rows if plant == 'all' or r.get('plant') == plant]
    # CODs are global. Corrections retain the same identity, even across paths.
    first_issues, revisions, recent, seen_history, undated_issues = {}, [], [], set(), {}
    completed = [h for h in history if h.get('state') == 'complete' and _cod(h.get('cod'))
                 and _day(h.get('completed_at') or h.get('started_at'))]
    completed.sort(key=lambda h: (str(h.get('completed_at') or h.get('started_at')), str(h.get('id', ''))))
    for record in completed:
        details = record.get('after') or {}
        if not isinstance(details, dict):
            details = {}
        event = {'cod': str(record['cod']), 'plant': record.get('plant'),
                 'date': _day(record.get('completed_at') or record.get('started_at')),
                 'supplier': details.get('supplier') or 'Company not recorded',
                 'type': details.get('certificate_type') or 'standard',
                 'inward_date': _day(details.get('inward_date')), 'date_source': 'history'}
        key = _cod(record['cod'])
        before = record.get('before') or {}
        if not isinstance(before, dict):
            before = {}
        before_details = before.get('details') if isinstance(before.get('details'), dict) else before
        was_issued = (_cod(before_details.get('cod_no')) == key
                      and status_value(before.get('status') or before.get('issued')) in ISSUED)
        if key in seen_history:
            revisions.append(event)
            recent.append(dict(event, action='Reissued'))
        elif record.get('previous_pdfs') or was_issued:
            # A correction to a legacy PDF is not a new certificate today.
            original_day = _day(before_details.get('cod_date'))
            if original_day:
                first_issues[key] = dict(event, date=original_day, date_source='certificate')
            else:
                undated_issues[key] = event['plant']
            revisions.append(event)
            recent.append(dict(event, action='Reissued'))
        else:
            first_issues[key] = event
            recent.append(dict(event, action='Created'))
        seen_history.add(key)
    for row in all_rows:
        key = _cod(row.get('certificate_no'))
        if not key or key in first_issues or key in undated_issues or status_value(row.get('certificate_status')) not in ISSUED:
            continue
        day = _day(row.get('certificate_date'))
        if not day:
            undated_issues[key] = row.get('plant')
            continue
        first_issues[key] = {'cod': str(row['certificate_no']), 'plant': row.get('plant'), 'date': day,
                             'supplier': row.get('supplier', ''), 'type': row.get('certificate_type') or 'standard',
                             'inward_date': _day(row.get('inward_date')), 'date_source': 'certificate'}
    folder_unknown = set()
    for pdf in saved_pdfs:
        key = _cod(pdf.get('cod'))
        if not key:
            if plant == 'all' or pdf.get('plant') == plant:
                folder_unknown.add(pdf.get('path', 'unknown'))
            continue
        if key in first_issues:
            continue
        if not pdf.get('date') or pdf.get('type') not in TYPE_NAMES:
            if plant == 'all' or pdf.get('plant') == plant:
                folder_unknown.add(key)
            continue
        undated_issues.pop(key, None)
        first_issues[key] = dict(pdf, supplier='', inward_date=None, date_source='pdf')
    events = [e for e in first_issues.values() if (plant == 'all' or e['plant'] == plant) and e['date'] <= today]
    inside = lambda day: day is not None and start <= day <= today
    previous = lambda day: day is not None and previous_start <= day <= previous_end
    selected = [e for e in events if inside(e['date'])]
    # Supplemental destruction requests reuse an inward, so never count them
    # as a second receipt. Standalone local certificates aren't Excel intakes.
    inwards = [r for r in rows if r.get('source_kind') != 'manual' and not r.get('destruction_source')]
    received = [r for r in inwards if inside(_day(r.get('inward_date')))]
    turnaround = [(e['date'] - e['inward_date']).days for e in selected
                  if e['inward_date'] and e['inward_date'] <= e['date']]
    pending = [r for r in rows if status_value(r.get('certificate_status')) not in ISSUED | {'na'}]
    ready = [r for r in pending if status_value(r.get('certificate_status')) == 'ready'
             and not r.get('issuance_blocked') and not r.get('record_identity_conflict')
             and not r.get('system_alert') and r.get('outcome') != 'needs_review'
             and not any(not re.match(r'^Excel is open or in use\. Showing its saved contents', str(w), re.I)
                         for w in r.get('warnings', []))]
    aged = [r for r in pending if not r.get('destruction_source')]
    age_counts = Counter()
    for row in aged:
        day = _day(row.get('inward_date'))
        age = (today - day).days if day else None
        key = ('unknown' if age is None or age < 0 else '0–7' if age <= 7 else '8–14' if age <= 14
               else '15–30' if age <= 30 else '31+')
        age_counts[key] += 1
    generated_by_day = Counter(e['date'] for e in selected)
    previous_by_day = Counter(e['date'] for e in events if previous(e['date']))
    received_by_day = Counter(_day(r.get('inward_date')) for r in received)
    company_key = lambda name: ' '.join(str(name or '').casefold().split())
    types = Counter(e['type'] for e in selected)
    selected_revisions = [e for e in revisions if inside(e['date']) and (plant == 'all' or e['plant'] == plant)]
    activity = [e for e in recent if inside(e['date']) and (plant == 'all' or e['plant'] == plant)]
    activity.reverse()
    current_codes = {_cod(row.get('certificate_no')) for row in rows if _cod(row.get('certificate_no'))}
    return {
        'range': {'start': start.isoformat(), 'end': today.isoformat(), 'days': span,
                  'previous_start': previous_start.isoformat(), 'previous_end': previous_end.isoformat()},
        'totals': {'generated': len(selected), 'previous_generated': sum(previous(e['date']) for e in events),
                   'received': len(received), 'previous_received': sum(previous(_day(r.get('inward_date'))) for r in inwards),
                   'clients': len({company_key(e['supplier']) for e in selected if company_key(e['supplier'])}),
                   'median_turnaround': median(turnaround) if turnaround else None, 'turnaround_sample': len(turnaround),
                   'reissues': len(selected_revisions), 'lifetime_generated': len(events)},
        'calendar': {'today': sum(e['date'] == today for e in events),
                     'week': sum(today - timedelta(days=today.weekday()) <= e['date'] <= today for e in events),
                     'month': sum(today.replace(day=1) <= e['date'] <= today for e in events)},
        'series': [{'date': (start + timedelta(days=i)).isoformat(),
                    'previous_date': (previous_start + timedelta(days=i)).isoformat(),
                    'generated': generated_by_day[start + timedelta(days=i)],
                    'received': received_by_day[start + timedelta(days=i)],
                    'previous_generated': previous_by_day[previous_start + timedelta(days=i)]} for i in range(span)],
        'types': [{'key': key, 'label': label, 'count': types[key]} for key, label in TYPE_NAMES.items()],
        'plants': [{'key': key, 'label': label, 'generated': sum(e['plant'] == key for e in selected),
                    'received': sum(r.get('plant') == key for r in received)}
                   for key, label in (('plant1', 'Plant 1'), ('plant2', 'Plant 2')) if plant in {'all', key}],
        'queue': {'ready': len(ready), 'attention': len(pending) - len(ready),
                  'to_send': sum(status_value(r.get('certificate_status')) in {'generated', 'generated_manual'} for r in rows),
                  'destruction': sum(r.get('destruction_requested') is not False and r.get('certificate_type') == 'destruction'
                                     and status_value(r.get('certificate_status')) not in {'completed', 'na'} for r in rows),
                  'overdue': age_counts['31+']},
        'ageing': [{'label': label, 'count': age_counts[label]} for label in ('0–7', '8–14', '15–30', '31+', 'unknown')],
        'recent': [dict(e, date=e['date'].isoformat(), inward_date=e['inward_date'].isoformat() if e['inward_date'] else None,
                        has_current_record=_cod(e['cod']) in current_codes)
                   for e in activity[:6]],
        'coverage': {'fallback_dates': sum(e['date_source'] == 'certificate' for e in selected),
                     'folder_certificates': sum(e['date_source'] == 'pdf' for e in selected),
                     'folder_needs_details': len(folder_unknown),
                     'no_plant': sum(e['plant'] not in {'plant1', 'plant2'} for e in selected),
                     'missing_cod': sum(status_value(r.get('certificate_status')) in ISSUED and not _cod(r.get('certificate_no')) for r in rows),
                     'unknown_generation_dates': sum(plant == 'all' or source == plant for source in undated_issues.values()),
                     'undated_inwards': sum(_day(r.get('inward_date')) is None for r in inwards),
                     'scan_complete': bool(scan_complete)},
    }
