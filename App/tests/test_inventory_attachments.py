"""Attached certificates reserve their numbers without becoming the main COD."""
from datetime import date
import hashlib
import json
from pathlib import Path
import tempfile
import unittest

from pypdf import PdfReader
from reportlab.pdfgen import canvas

from certificate_drafts import generate_draft, read_register
from certificate_inventory import scan_inventory, cod_conflicts, INSPECTION_VERSION
from destruction_certificate import request_for_inward
from pdf_attachments import store_attachment
from tests.test_certificate_templates import write_register
from workflow_support import save_json
from workspace_review import workspace_revision

AS_OF = date(2026, 10, 20)


def certificate_pdf(path, numbers):
    path.parent.mkdir(parents=True, exist_ok=True)
    pdf = canvas.Canvas(str(path))
    for number in numbers:
        pdf.drawString(30, 700, f'Certificate No. {number}')
        pdf.showPage()
    pdf.save()
    return path


class InventoryAttachmentTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.paths = [self.root / f'plant{i}.xlsx' for i in (1, 2)]
        for index, path in enumerate(self.paths, 1):
            write_register(path, case_id=f'SN{index}', cod=f'{300+index:04}', status='Completed')
        save_json(self.root / 'settings.local.json', {'workbook_paths':list(map(str, self.paths))})

    def details(self, kind='standard', token=None):
        fields = {'_annexure':token} if token else {}
        if kind == 'destruction':
            fields.update(destruction_date='2026-10-19', device_receipts='[{"date":"2026-10-01","count":"12"}]')
        if kind == 'siemens':
            fields.update(model_name='Model', model_serial='M1', tube_serial='T1', site_location='Pune')
        return {'certificate_type':kind, 'supplier':'PDF Test Client', 'location':'Pune',
                'reference':f'REF-{kind}', 'form6':f'FORM-{kind}', 'inward_date':'2026-10-01',
                'certificate_weight':'10', 'certificate_fields':fields}

    def journal(self, path, kind='destruction', fields=None, state='complete'):
        save_json(self.root / 'Data' / 'Issuance History' / 'fixture.json', {
            'version':1, 'state':state, 'pdf_path':path.relative_to(self.root).as_posix(),
            'cod': '0205',
            'pdf_sha256':hashlib.sha256(path.read_bytes()).hexdigest(),
            'after':{'certificate_type':kind,'certificate_fields':fields or {
                '_annexure':'a'*32, 'device_receipts':'[{"date":"2026-10-01","count":"12"}]'}}})

    def test_all_templates_show_own_number_and_keep_attachment_numbers_reserved(self):
        attached = certificate_pdf(self.root / 'Original.pdf', ['0189'])
        token = store_attachment(self.root, attached, workspace_revision(self.root))['id']
        before = [path.read_bytes() for path in self.paths]
        for index, kind in enumerate(('standard', 'destruction', 'sita', 'siemens'), 303):
            with self.subTest(kind=kind):
                result = generate_draft(self.root, 'plant1', self.details(kind, token), confirmed=True, as_of=AS_OF)
                self.assertEqual(result['certificate_no'], f'{index:04}')
                path = self.root / 'Certificates' / result['relative_path']
                original = path.read_bytes()
                inventory = scan_inventory(self.root, persist=True)
                entry = next(row for row in inventory['entries'] if row['path'] == path.relative_to(self.root).as_posix())
                self.assertEqual(entry['current_cods'], [f'{index:04}'])
                self.assertEqual(entry['attachment_cods'], ['0189'])
                self.assertEqual(set(entry['cods']), {'0189', f'{index:04}'})
                self.assertEqual(entry['warnings'], [])
                self.assertFalse(cod_conflicts(inventory, '0189'))
                self.assertEqual(path.read_bytes(), original)
                self.assertIn('0189', PdfReader(path).pages[-1].extract_text())
        self.assertFalse(scan_inventory(self.root)['warnings'], 'Shared attached PDFs are not duplicate main certificates.')
        original = certificate_pdf(self.root / 'Certificates' / 'Auto Generated' / '0189_Recycling.pdf', ['0189'])
        inventory = scan_inventory(self.root)
        self.assertFalse(inventory['warnings'])
        self.assertFalse(cod_conflicts(inventory, '0189', [original.relative_to(self.root).as_posix()]),
                         'An attached copy must not block correcting the original certificate.')
        self.assertEqual(before, [path.read_bytes() for path in self.paths])

    def test_existing_cached_pdf_is_rechecked_without_rewriting_or_releasing_numbers(self):
        path = certificate_pdf(self.root / 'Certificates' / 'Auto Generated' / '0205_Client_hdd.pdf', ['0205','0205','0189'])
        stale = scan_inventory(self.root, persist=True)
        self.assertTrue(stale['warnings'])
        stale.pop('warnings')
        stale['entries'][0]['inspection_version'] = 4
        for key in ('certificate_cods','attachment_cods','inspection_certificate_pages'):
            stale['entries'][0].pop(key, None)
        save_json(self.root / 'Data' / 'Certificate Inventory.json', stale)
        self.journal(path)
        original = path.read_bytes()
        result = scan_inventory(self.root, persist=True)
        entry = result['entries'][0]
        self.assertEqual(entry['current_cods'], ['0205'])
        self.assertEqual(entry['certificate_cods'], ['0205'])
        self.assertEqual(set(entry['cods']), {'0189','0205'})
        self.assertEqual(entry['inspection_version'], INSPECTION_VERSION)
        self.assertEqual(entry['warnings'], [])
        self.assertEqual(scan_inventory(self.root)['entries'][0], entry)
        self.assertEqual(path.read_bytes(), original)
        # A changed PDF must lose the attachment exemption, even with the same name.
        certificate_pdf(path, ['0205','0205','0499'])
        changed = scan_inventory(self.root)['entries'][0]
        self.assertIn('0499', changed['current_cods'])
        self.assertTrue(any('differ' in warning for warning in changed['warnings']))

    def test_main_page_mismatch_and_unproven_attachments_still_warn(self):
        path = certificate_pdf(self.root / 'Certificates' / 'Auto Generated' / '0205_Client_hdd.pdf', ['0999','0999','0189'])
        self.journal(path)
        entry = scan_inventory(self.root)['entries'][0]
        self.assertEqual(set(entry['current_cods']), {'0205','0999'})
        self.assertTrue(any('differ' in warning for warning in entry['warnings']))
        self.journal(path, state='prepared')
        entry = scan_inventory(self.root)['entries'][0]
        self.assertIn('0189', entry['current_cods'])
        self.assertTrue(entry['warnings'])

    def test_hdd_receipt_continuations_are_part_of_main_certificate(self):
        path = certificate_pdf(self.root / 'Certificates' / 'Auto Generated' / '0205_Client_hdd.pdf', ['0205','0205','0999','0189'])
        rows = [{'date':'2026-10-01','count':'1'} for _ in range(11)]
        self.journal(path, fields={'_annexure':'a'*32,'device_receipts':json.dumps(rows)})
        entry = scan_inventory(self.root)['entries'][0]
        self.assertEqual(set(entry['current_cods']), {'0205','0999'})
        self.assertEqual(entry['attachment_cods'], ['0189'])
        self.assertTrue(entry['warnings'], 'A wrong number on a receipt continuation must remain visible.')

    def test_new_hdd_requests_and_standalone_hdd_share_sequence_with_recycling(self):
        recycling = generate_draft(self.root, 'plant1', self.details(), confirmed=True, as_of=AS_OF)
        request = request_for_inward(self.root, 'plant1', 'SN1', workspace_revision(self.root))
        saved = next(record for record in read_register(self.root)['records'] if record['id'] == request)
        self.assertEqual(saved['details']['cod_no'], '', 'Opening a request must not copy its inward COD.')
        linked = dict(saved['details'])
        linked['certificate_fields'] = self.details('destruction')['certificate_fields']
        hdd = generate_draft(self.root, 'plant1', linked, request, confirmed=True, as_of=AS_OF)
        standalone = generate_draft(self.root, 'plant2', self.details('destruction'), confirmed=True, as_of=AS_OF)
        next_details = dict(self.details(), form6='FORM-next-recycling')
        next_recycling = generate_draft(self.root, 'plant2', next_details, confirmed=True, as_of=AS_OF)
        self.assertEqual([result['certificate_no'] for result in (recycling,hdd,standalone,next_recycling)], ['0303','0304','0305','0306'])
        reissued = generate_draft(self.root, 'plant1', linked, request, confirmed=True, as_of=AS_OF)
        self.assertEqual(reissued['certificate_no'], '0304', 'Correcting an issued HDD keeps its own number.')
        self.assertTrue(hdd['relative_path'].endswith('_HDD.pdf'))

    def test_numbers_seen_only_in_attachments_cannot_be_allocated_again(self):
        attached = certificate_pdf(self.root / 'Original.pdf', ['0999'])
        token = store_attachment(self.root, attached, workspace_revision(self.root))['id']
        first = generate_draft(self.root, 'plant1', self.details(token=token), confirmed=True, as_of=AS_OF)
        following = generate_draft(self.root, 'plant2', self.details('destruction'), confirmed=True, as_of=AS_OF)
        self.assertEqual(first['certificate_no'], '0303')
        self.assertEqual(following['certificate_no'], '1000')


if __name__ == '__main__':
    unittest.main()
