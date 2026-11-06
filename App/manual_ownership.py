"""A local certificate has a plant only through an actual Excel inward link."""


def record_plant(record):
    link = (record.get('details') or {}).get('destruction_source') or {}
    if (not record.get('detached_inward') and not link.get('record_id')
            and link.get('plant') in {'plant1', 'plant2'}):
        return link['plant']
    return None


def normalize_records(data):
    for record in data.get('records', []):
        record['plant'] = record_plant(record)
        link = (record.get('details') or {}).get('destruction_source')
        if isinstance(link, dict) and link.get('record_id'):
            link['plant'] = None
    return data
