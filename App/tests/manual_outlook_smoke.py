"""Explicit opt-in live Outlook smoke test. Never sends; removes its own draft.

Run directly, not through unittest discovery. Uses the configured mailbox but
isolates all source workbooks, outputs and COD history in a temporary directory.
"""
from datetime import date
from pathlib import Path
import shutil
import sys
import tempfile
import uuid

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from certificate_core import read_batches
from draft_service import OutlookWriter
from run_certificates import run
from workflow_support import load_settings, save_json
from app_paths import DATA_ROOT
from test_workflow import make_book


def main():
    import pythoncom
    mailbox = load_settings(DATA_ROOT)['outlook_mailbox']
    if not mailbox:
        raise RuntimeError('Select a Classic Outlook mailbox in the app first.')
    key = 'SMOKE-' + uuid.uuid4().hex
    pythoncom.CoInitialize()
    writer = None
    cleaned = 0
    try:
        writer = OutlookWriter(mailbox)
        with tempfile.TemporaryDirectory(prefix='ecoreco-smoke-') as folder:
            root = Path(folder)
            (root / 'templates').mkdir()
            for name in ('Certificate of Recycling.png', 'signature-banner.png'):
                shutil.copy2(ROOT / 'templates' / name, root / 'templates' / name)
            save_json(root / 'settings.example.json', {
                'initial_cod_confirmed': True, 'certificate_number_start': 1,
                'certificate_number_prefix': 'TEST-', 'certificate_number_width': 4,
            })
            paths = [root / 'Plant 1.xlsx', root / 'Plant 2.xlsx']
            make_book(paths[0], [{'SN': key, 'Supplier Name': 'Automation Smoke Test',
                                 'Inward Date': date.today(), 'Client Email': 'certificate-smoke@example.test'}])
            make_book(paths[1], [{'Certificate Issued': 'N/A'}])
            created, review, output = run(paths, mailbox, root, limit=1, progress=lambda _: None)
            assert (created, review) == (1, 0), (output / 'Results.txt').read_text()
            batch = read_batches(paths[0])[0]
            assert batch['cod_no'] == 'TEST-0001'
            matches = []
            for item in writer.folder.Items:
                prop = item.UserProperties.Find('EcorecoCertificateKey')
                if prop and str(prop.Value) == key:
                    matches.append(item)
            assert len(matches) == 1, 'Expected exactly one uniquely identified test draft'
            item = matches[0]
            assert not item.Sent
            assert str(item.To).lower() == 'certificate-smoke@example.test'
            pdf = next(output.glob('*.pdf'))
            attachment = next(a for a in item.Attachments if str(a.FileName) == pdf.name)
            extracted = root / 'readback.pdf'
            attachment.SaveAsFile(str(extracted))
            assert extracted.read_bytes() == pdf.read_bytes(), 'Attached PDF differs from generated certificate'
            assert not attachment.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x7FFE000B')
            banner = next(a for a in item.Attachments if str(a.FileName) == 'signature-banner.png')
            assert banner.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F') == 'ecoreco-banner'
            assert run(paths, mailbox, root, progress=lambda _: None)[0] == 0, 'Repeat created a duplicate'
            print('PASS: same-day Ready, normal plant workflow, saved unsent draft, visible byte-matching PDF, inline banner, Excel COD/date update, repeat without duplicates.')
    finally:
        try:
            if writer:
                items = writer.folder.Items
                for index in range(items.Count, 0, -1):
                    item = items.Item(index)
                    prop = item.UserProperties.Find('EcorecoCertificateKey')
                    if prop and str(prop.Value) == key:
                        if item.Sent:
                            raise RuntimeError('Unexpected sent state; cleanup refused.')
                        item.Delete()
                        cleaned += 1
                print(f'Cleanup: {cleaned} uniquely identified test draft(s) moved to Deleted Items. No email sent.')
        finally:
            writer = None
            pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
