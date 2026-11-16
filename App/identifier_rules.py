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


def require_restored_identifiers(root, entries, overrides, records):
    """Check restored manifests against changes in other sources."""
    from certificate_inventory import identity as cod_identity, scan_inventory

    batches = workspace_batches(root, overrides)
    active_pdfs = None
    for entry in entries:
        source_kind, plant = entry['source_kind'], entry['plant']
        if source_kind == 'manual':
            key = entry['record_id']
            record = next((record for record in records if record['id'] == key), None)
            details = record['details'] if record else None
            others = batches + [record['details'] for record in records if record['id'] != key]
        else:
            key = entry['sn']
            details = next((batch for batch in batches if batch['source_id'] == plant
                            and identifier(batch['case_id']) == identifier(key)), None)
            others = [batch for batch in batches if batch is not details] + [record['details'] for record in records]
        if details is None:
            raise ValueError('The deleted entry could not be found in its recovery copy.')
        require_unique(details, plant, source_kind, key, batches, records)
        cod = cod_identity(details.get('cod_no'))
        if not cod:
            continue
        if active_pdfs is None:
            active_pdfs = {cod_identity(number) for pdf in scan_inventory(root, persist=False)['entries']
                           if pdf.get('present') for number in pdf.get('current_cods', pdf.get('cods', []))}
        if cod in active_pdfs or any(cod in {cod_identity(other.get('cod_no')), cod_identity(other.get('hdd_cod_no'))}
                                    for other in others):
            raise ValueError(f'COD {details["cod_no"]} now belongs to another entry or PDF. '
                             'Use the recovery copy to restore this entry with a different COD.')
