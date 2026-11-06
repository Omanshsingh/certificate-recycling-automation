"""SN is unique per plant; manifest identity belongs to one inward overall."""
from pathlib import Path


def identifier(value):
    return str(value or '').strip().casefold()


def conflicts(details, plant, source_kind, key, batches, manual_records=()):
    # Supplemental destruction certificates refer to the same physical inward.
    if details.get('certificate_type') == 'destruction' and details.get('destruction_source'):
        return []
    own = (source_kind, plant, identifier(key))
    candidates = []
    for batch in batches:
        identity = ('workbook', batch.get('source_id'), identifier(batch.get('case_id')))
        if identity == own:
            continue
        label = f'{str(batch.get("source_id") or "").replace("plant", "Plant ")}, SN {batch.get("case_id")}'
        if batch.get('source_row'):
            label += f', row {batch["source_row"]}'
        candidates.append((batch, label, batch.get('source_id')))
    for record in manual_records:
        if ('manual', record['plant'], identifier(record['id'])) == own:
            continue
        other = record['details']
        if other.get('certificate_type') == 'destruction' and other.get('destruction_source'):
            continue
        candidates.append((other, f'the saved certificate for {other.get("supplier") or "an unnamed company"}', record['plant']))
    result = []
    for field, title in [('case_id', 'SN (Serial Number)'), ('form6', 'Form 6 number')]:
        if field == 'case_id' and plant not in {'plant1', 'plant2'}:
            continue  # App-only records use their local ID, not an Excel SN.
        value = identifier(details.get(field))
        if not value:
            continue
        owners = [label for other, label, other_plant in candidates
                  if identifier(other.get(field)) == value and (field != 'case_id' or other_plant == plant)]
        if owners:
            result.append(f'{title} {details[field]} is already used by ' + '; '.join(dict.fromkeys(owners)) +
                          '. Use a different value; this number must be unique ' +
                          ('within this plant.' if field == 'case_id' else 'across both plants.'))
    return result


def require_unique(details, plant, source_kind, key, batches, manual_records=()):
    issues = conflicts(details, plant, source_kind, key, batches, manual_records)
    if issues:
        raise ValueError(' '.join(issues))


def shared_reference_warnings(details, plant, source_kind, key, batches, manual_records=()):
    """Compatibility API: references can be reused without a notice."""
    return []


def mark_conflicts(batches, manual_records=()):
    for batch in batches:
        issues = conflicts(batch, batch.get('source_id'), 'workbook', batch.get('case_id'), batches, manual_records)
        if issues:
            batch['blocking_issues'] = list(dict.fromkeys([*batch.get('blocking_issues', []), *issues]))
            batch['issues'] = list(dict.fromkeys([*batch.get('issues', []), *issues]))
            batch['record_identity_conflict'] = batch.get('record_identity_conflict') or ' '.join(issues)


def workspace_batches(root, overrides=None):
    """Read both configured sources; callers already hold the workspace write lock."""
    from certificate_core import read_batches
    from workflow_support import load_settings
    from io import BytesIO
    batches = []
    overrides = overrides or {}
    for index, value in enumerate(load_settings(root).get('workbook_paths', []), 1):
        if not value:
            continue
        path = Path(value).expanduser().resolve()
        source = BytesIO(overrides[str(path)]) if str(path) in overrides else path
        batches.extend(dict(batch, source_id=f'plant{index}') for batch in read_batches(source))
    return batches
