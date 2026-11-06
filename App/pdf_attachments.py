"""Validate and retain the exact PDF selected by the operator."""
from io import BytesIO
import json
from pathlib import Path
import re
import uuid

from pypdf import PdfReader


def checked_reader(payload):
    try:
        if not payload.startswith(b'%PDF-'):
            raise ValueError('This file is not a readable PDF.')
        reader = PdfReader(BytesIO(payload))
        if reader.is_encrypted:
            raise ValueError('The attached PDF is password protected. Choose an unlocked PDF.')
        if not 0 < len(reader.pages) <= 2000:
            raise ValueError('Choose a PDF with between 1 and 2,000 pages.')
        for page in reader.pages:
            if page.mediabox.width <= 0 or page.mediabox.height <= 0:
                raise ValueError('The attached PDF has an invalid page size.')
            content = page.get_contents()
            if content is not None:
                content.get_data()
        return reader
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError('The attached PDF is damaged or unreadable. Choose another PDF.') from exc


def store_attachment(root, source, expected_revision):
    from workflow_support import run_lock, save_json
    from workspace_review import require_revision
    source = Path(source)
    if source.suffix.casefold() != '.pdf' or not 0 < source.stat().st_size <= 50_000_000:
        raise ValueError('Choose a PDF smaller than 50 MB.')
    with source.open('rb') as stream:
        payload = stream.read(50_000_001)
    if len(payload) > 50_000_000:
        raise ValueError('Choose a PDF smaller than 50 MB.')
    reader = checked_reader(payload)
    token = uuid.uuid4().hex
    attachment = {'id': token, 'name': source.name, 'pages': len(reader.pages)}
    with run_lock(root):
        require_revision(root, expected_revision)
        folder = Path(root) / 'Data' / 'Attachments'
        folder.mkdir(parents=True, exist_ok=True)
        pdf_path, metadata = folder / f'{token}.pdf', folder / f'{token}.json'
        try:
            with pdf_path.open('xb') as stream:
                stream.write(payload)
            save_json(metadata, attachment)
        except Exception:
            pdf_path.unlink(missing_ok=True)
            metadata.unlink(missing_ok=True)
            raise
    return attachment


def append_attachment(writer, root, token):
    if not token:
        return
    if not re.fullmatch(r'[a-f0-9]{32}', str(token)):
        raise ValueError('Choose the attached PDF again.')
    path = Path(root) / 'Data' / 'Attachments' / f'{token}.pdf'
    if not path.is_file():
        raise ValueError('The attached annexure is missing. Choose its PDF again before creating the certificate.')
    if not 0 < path.stat().st_size <= 50_000_000:
        raise ValueError('The attached PDF is empty or too large. Choose its PDF again.')
    reader = checked_reader(path.read_bytes())
    for page in reader.pages:
        writer.add_page(page)


def attachment_catalog(root):
    result = {}
    for path in (Path(root) / 'Data' / 'Attachments').glob('*.json'):
        if not re.fullmatch(r'[a-f0-9]{32}', path.stem) or not path.with_suffix('.pdf').is_file():
            continue
        try:
            entry = json.loads(path.read_text(encoding='utf-8'))
            if entry.get('id') == path.stem and isinstance(entry.get('name'), str):
                result[path.stem] = entry
        except (OSError, ValueError, AttributeError):
            continue
    return result
