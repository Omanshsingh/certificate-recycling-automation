"""Compare live workbook details with the details recorded for an issued PDF."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re

from certificate_core import number, parse_date
from workflow_support import save_json


REVIEW_FILE = Path('Data') / 'Issued Excel Change Reviews.json'
FIELDS = {
    'supplier': 'Company name',
    'location': 'City / site',
    'inward_date': 'Received date',
    'form6': 'Form 6',
    'reference': 'Reference number',
    'comments': 'Comments',
    'transporter_details': 'Transporter / vehicle',
    'docket_vehicle_no': 'Docket / vehicle number',
    'weight_slip_scale': 'Weight slip / scale weight',
}
ITEM_FIELDS = ('item', 'quantity', 'per_piece_weight', 'weight')
NUMERIC_FIELDS = {'quantity', 'per_piece_weight', 'weight'}


def _value(value, key=''):
    if key == 'inward_date':
        parsed = parse_date(value)
        return parsed.isoformat() if parsed else _text(value)
    if key in NUMERIC_FIELDS:
        parsed = number(value)
        if parsed is not None:
            return parsed
    return _text(value)


def _text(value):
    return ' '.join(str('' if value is None else value).split())


def _items(details):
    raw = details.get('items')
    if not isinstance(raw, list):
        raw = [details] if any(details.get(key) not in (None, '') for key in ITEM_FIELDS) else []
    return [
        {key: _value(item.get(key), key) for key in ITEM_FIELDS}
        for item in raw if isinstance(item, dict)
        and any(item.get(key) not in (None, '') for key in ITEM_FIELDS)
    ]


def normalized_source_details(details):
    details = details if isinstance(details, dict) else {}
    result = {key: _value(details.get(key), key) for key in FIELDS}
    result['items'] = _items(details)
    return result


def source_signature(details):
    encoded = json.dumps(normalized_source_details(details), ensure_ascii=False,
                         sort_keys=True, separators=(',', ':'), allow_nan=False).encode('utf-8')
    return hashlib.sha256(encoded).hexdigest()


def _display(value, key=''):
    if value in (None, ''):
        return '(blank)'
    if key == 'inward_date':
        parsed = parse_date(value)
        if parsed:
            return parsed.strftime('%d/%m/%y')
    return _text(value)


def _display_items(details):
    rendered = []
    raw = details.get('items')
    if not isinstance(raw, list):
        raw = [details] if any(details.get(key) not in (None, '') for key in ITEM_FIELDS) else []
    for item in raw:
        if not isinstance(item, dict) or not any(item.get(key) not in (None, '') for key in ITEM_FIELDS):
            continue
        parts = [_text(item.get('item')) or 'Unnamed material']
        quantity = _display(item.get('quantity'))
        per_piece = _display(item.get('per_piece_weight'))
        weight = _display(item.get('weight'))
        if quantity != '(blank)':
            parts.append(f'Qty {quantity}')
        if per_piece != '(blank)':
            parts.append(f'{per_piece} kg each')
        if weight != '(blank)':
            parts.append(f'{weight} kg')
        rendered.append(' · '.join(parts))
    return '\n'.join(rendered) or '(none)'


def compare_details(current, issued):
    """Return concise, safe-to-render before/after rows for source fields only."""
    current = current if isinstance(current, dict) else {}
    issued = issued if isinstance(issued, dict) else {}
    current_normalized = normalized_source_details(current)
    issued_normalized = normalized_source_details(issued)
    differences = []
    for key, label in FIELDS.items():
        if current_normalized[key] != issued_normalized[key]:
            differences.append({'field': label, 'issued': _display(issued.get(key), key),
                                'current': _display(current.get(key), key)})
    if current_normalized['items'] != issued_normalized['items']:
        differences.append({'field': 'Materials and weights', 'issued': _display_items(issued),
                            'current': _display_items(current)})
    return differences


def _load_reviews(root):
    path = Path(root) / REVIEW_FILE
    if not path.exists():
        return {'version': 1, 'reviews': {}}
    try:
        value = json.loads(path.read_text(encoding='utf-8-sig'))
    except (OSError, ValueError, UnicodeError):
        # Acknowledgements only hide a review notice. If this small local file
        # is damaged, fail open: show the changes again and allow a new choice.
        return {'version': 1, 'reviews': {}}
    if (not isinstance(value, dict) or value.get('version') != 1
            or not isinstance(value.get('reviews'), dict)):
        return {'version': 1, 'reviews': {}}
    for receipt_id, review in value['reviews'].items():
        if (not isinstance(receipt_id, str) or not re.fullmatch(r'[a-f0-9]{32}', receipt_id)
                or not isinstance(review, dict)
                or not re.fullmatch(r'[a-f0-9]{64}', str(review.get('source_signature', '')))):
            return {'version': 1, 'reviews': {}}
    return value


def intentional_reviews(root):
    """Load kept-value acknowledgements once per dashboard refresh."""
    return _load_reviews(root)['reviews']


def is_intentionally_kept(root, receipt, current, *, reviews=None):
    if not isinstance(receipt, dict) or not isinstance(receipt.get('id'), str):
        return False
    reviews = intentional_reviews(root) if reviews is None else reviews
    review = reviews.get(receipt['id'])
    return bool(review and review.get('source_signature') == source_signature(current))


def keep_intentional(root, receipt, plant, case_id, current):
    receipt_id = receipt.get('id') if isinstance(receipt, dict) else None
    if not isinstance(receipt_id, str) or not re.fullmatch(r'[a-f0-9]{32}', receipt_id):
        raise ValueError('The issued certificate history could not be identified. Refresh and try again.')
    value = _load_reviews(root)
    value['reviews'][receipt_id] = {
        'plant': plant,
        'case_id': str(case_id),
        'source_signature': source_signature(current),
        'accepted_at': datetime.now(timezone.utc).isoformat(),
    }
    save_json(Path(root) / REVIEW_FILE, value)


def clear_review(root, receipt_id):
    value = _load_reviews(root)
    if value['reviews'].pop(receipt_id, None) is not None:
        save_json(Path(root) / REVIEW_FILE, value)
