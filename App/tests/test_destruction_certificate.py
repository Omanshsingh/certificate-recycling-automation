"""Destruction is supplemental to an inward; all issuance tests use disposable data."""
from datetime import date
from io import BytesIO
import json
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader
from certificate_drafts import (save_draft, generate_draft, manual_rows, read_register,
                                complete_draft, reconcile_manual, prepare_certificate)
from certificate_templates import make_template_certificate, template_specs, template_id
from destruction_certificate import receipts, request_for_inward, content_issues
from tests.test_certificate_templates import write_register
from workflow_support import save_json
from workspace_review import workspace_revision
import web_app

AS_OF = date(2026, 10, 17)


class DestructionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.paths = [self.root / f'plant{i}.xlsx' for i in (1, 2)]
        for i, path in enumerate(self.paths, 1):
            write_register(path, case_id=f'CLIENT{i}', cod=f'{300+i:04}', status='Completed')
        save_json(self.root / 'settings.local.json', {'workbook_paths': list(map(str, self.paths))})

    def details(self, count=12):
        return {'certificate_type': 'destruction', 'supplier': 'Example International Ltd',
                'location': 'London, UK', 'reference':'REF-HDD-TEST', 'certificate_fields': {'destruction_date': '2026-10-16',
                'device_receipts': json.dumps([{'date': '2026-10-01', 'count': str(count)}])}}

    def test_request_is_idempotent_and_does_not_mutate_either_inward(self):
        before = [path.read_bytes() for path in self.paths]
        record = request_for_inward(self.root, 'plant1', 'CLIENT1', workspace_revision(self.root))
        self.assertEqual(record, request_for_inward(self.root, 'plant1', 'CLIENT1', workspace_revision(self.root)))
        self.assertEqual(len(read_register(self.root)['records']), 1)
        row = manual_rows(self.root)[0]
        self.assertTrue(row['destruction_pending'])
        self.assertEqual(row['destruction_source'], {'plant': 'plant1', 'case_id': 'CLIENT1'})
        self.assertFalse(row['certificate_no'])
        self.assertEqual(before, [path.read_bytes() for path in self.paths])

    def test_international_client_can_issue_without_an_inward_or_weight(self):
        result = generate_draft(self.root, 'plant2', self.details(), as_of=AS_OF)
        self.assertEqual(result['created'], 1)
        self.assertEqual(result['certificate_no'], '0303')
        self.assertEqual(result['relative_path'], '0303_Example International Ltd_London, UK_HDD.pdf')
        path = self.root / 'Certificates' / result['relative_path']
        pdf = PdfReader(path)
        self.assertEqual(len(pdf.pages), 2)
        self.assertIn('EXAMPLE INTERNATIONAL LTD', pdf.pages[0].extract_text())
        self.assertIn('16/10/2026', pdf.pages[0].extract_text())
        self.assertIn('12', pdf.pages[1].extract_text())
        self.assertEqual(pdf.metadata['/CertificateDate'],AS_OF.isoformat())
        self.assertEqual(pdf.metadata['/CertificatePlant'],'')
        self.assertTrue(manual_rows(self.root)[0]['destruction_pending'])
        complete_draft(self.root, result['record_id'])
        self.assertFalse(manual_rows(self.root)[0]['destruction_pending'])
        path.unlink()
        reconcile_manual(self.root)
        self.assertTrue(manual_rows(self.root)[0]['destruction_pending'])

    def test_request_can_be_removed_and_restored_without_losing_its_details_or_changing_excel(self):
        before = [path.read_bytes() for path in self.paths]
        record_id = request_for_inward(self.root, 'plant1', 'CLIENT1')
        details = read_register(self.root)['records'][0]['details']
        details['certificate_fields'] = self.details()['certificate_fields']
        save_draft(self.root, 'plant1', details, record_id)
        request_for_inward(self.root, 'plant1', 'CLIENT1', workspace_revision(self.root), needed=False)
        self.assertEqual(manual_rows(self.root), [])
        self.assertEqual(len(read_register(self.root)['records']), 1)
        self.assertEqual(request_for_inward(self.root, 'plant1', 'CLIENT1', workspace_revision(self.root)), record_id)
        restored = manual_rows(self.root)[0]
        self.assertTrue(restored['destruction_pending'])
        self.assertEqual(restored['certificate_fields'], details['certificate_fields'])
        self.assertEqual(before, [path.read_bytes() for path in self.paths])

    def test_issued_request_cannot_be_removed_and_keeps_pdf_number_and_history(self):
        record_id = request_for_inward(self.root, 'plant1', 'CLIENT1')
        details = read_register(self.root)['records'][0]['details']
        details['certificate_fields'] = self.details()['certificate_fields']
        result = generate_draft(self.root, 'plant1', details, record_id, as_of=AS_OF)
        pdf_path = self.root / 'Certificates' / result['relative_path']
        before = pdf_path.read_bytes()
        excel = [path.read_bytes() for path in self.paths]
        for confirmed in (False, True):
            with self.assertRaisesRegex(ValueError, 'has been issued'):
                request_for_inward(self.root, 'plant1', 'CLIENT1', workspace_revision(self.root),
                                   needed=False, confirmed=confirmed)
        row = manual_rows(self.root)[0]
        self.assertTrue(row['destruction_requested'])
        self.assertTrue(row['destruction_pending'])
        self.assertEqual(row['certificate_no'], result['certificate_no'])
        self.assertEqual(pdf_path.read_bytes(), before)
        self.assertEqual(excel, [path.read_bytes() for path in self.paths])
        self.assertEqual(row['record_id'], record_id)
        pdf_path.unlink()
        reconcile_manual(self.root)
        self.assertTrue(manual_rows(self.root)[0]['pdf_missing'])

    def test_request_removal_rejects_a_stale_review(self):
        request_for_inward(self.root, 'plant1', 'CLIENT1')
        revision = workspace_revision(self.root)
        request_for_inward(self.root, 'plant2', 'CLIENT2')
        with self.assertRaisesRegex(ValueError, 'files changed'):
            request_for_inward(self.root, 'plant1', 'CLIENT1', revision, needed=False)
        self.assertTrue(all(row['destruction_pending'] for row in manual_rows(self.root)))

    def test_linked_issuance_writes_hdd_below_recycling_and_keeps_other_plant(self):
        before = [path.read_bytes() for path in self.paths]
        recycling = self.root / 'Certificates' / 'Auto Generated' / '0301_Independent Client Ltd_Pune.pdf'
        make_template_certificate({'certificate_type':'standard', 'supplier':'Independent Client Ltd',
                                  'location':'Pune', 'inward_date':date(2026,9,1),
                                  'display_weight':10, 'form6':'FORM-1', 'certificate_fields':{}},
                                  '0301', AS_OF, recycling)
        recycling_before = recycling.read_bytes()
        record = request_for_inward(self.root, 'plant1', 'CLIENT1')
        details = read_register(self.root)['records'][0]['details']
        details['certificate_fields'] = self.details()['certificate_fields']
        result = generate_draft(self.root, 'plant1', details, record, as_of=AS_OF)
        self.assertFalse(result['requires_confirmation'])
        self.assertEqual(result['created'], 1)
        self.assertEqual(result['certificate_no'], '0303')
        self.assertNotEqual(before[0],self.paths[0].read_bytes())
        self.assertEqual(before[1],self.paths[1].read_bytes())
        from certificate_core import read_batches
        from openpyxl import load_workbook
        inward=read_batches(self.paths[0])[0]
        self.assertEqual(inward['cod_no'],'0301');self.assertEqual(inward['issued'],'Completed')
        self.assertEqual(inward['hdd_cod_no'],'0303')
        book=load_workbook(self.paths[0]);sheet=book.active
        from certificate_core import select_register
        _,_,cols=select_register(book)
        self.assertEqual(sheet.cell(3,cols['cod_no']+1).value,'0303')
        self.assertEqual(sheet.cell(3,cols['cod_date']+1).number_format,'dd/mm/yy')
        self.assertEqual(sheet.cell(3,cols['cod_no']+1).alignment.vertical,'top');book.close()
        self.assertEqual(recycling_before, recycling.read_bytes())
        # PDF-only corrections cannot detach the original inward.
        details.pop('destruction_source')
        save_draft(self.root, 'plant1', details, record)
        self.assertEqual(read_register(self.root)['records'][0]['details']['destruction_source']['case_id'], 'CLIENT1')

    def test_required_facts_cannot_be_bypassed_and_preview_is_non_mutating(self):
        details = self.details()
        details['certificate_fields']['destruction_date'] = ''
        with self.assertRaisesRegex(ValueError, 'actual date'):
            generate_draft(self.root, 'plant1', details, confirmed=True, as_of=AS_OF)
        before = workspace_revision(self.root)
        _, _, _, warnings = prepare_certificate(self.root, 'plant1', details, as_of=AS_OF)
        self.assertTrue(any('actual date' in warning for warning in warnings))
        self.assertEqual(before, workspace_revision(self.root))
        for count in ['-1', '0', '1.5', '1 Lot', 'NaN']:
            with self.subTest(count=count), self.assertRaises(ValueError):
                receipts(json.dumps([{'date':'2026-10-01', 'count':count}]))
        details = self.details()
        details['certificate_fields']['destruction_date'] = '2026-09-01'
        self.assertTrue(any('before' in issue for issue in content_issues(details, AS_OF)))

    def test_only_nine_receipts_fit_and_no_certificate_number_is_added_on_page_two(self):
        details = self.details()
        details['certificate_fields']['device_receipts'] = json.dumps([
            {'date':'2026-10-01','count':str(i)} for i in range(1,10)])
        path = self.root / 'nine-rows.pdf'
        make_template_certificate(details,'0303',AS_OF,path)
        pdf=PdfReader(path)
        self.assertEqual(len(pdf.pages),2)
        second=pdf.pages[1].extract_text()
        self.assertNotIn('Certificate No.',second)
        self.assertNotIn('0303',second)
        self.assertEqual([line.strip() for line in second.splitlines() if line.strip().isdigit()],[*[str(i) for i in range(1,10)],'45'])
        details['certificate_fields']['device_receipts'] = json.dumps([
            {'date':'2026-10-01','count':'1'} for _ in range(10)])
        with self.assertRaisesRegex(ValueError,'Only 9'):
            generate_draft(self.root,'plant1',details,confirmed=True,as_of=AS_OF)

    def test_cannot_replace_recycling_or_move_link(self):
        with self.assertRaisesRegex(ValueError, 'Review & edit.*Data Destruction needed'):
            web_app.workbook_form({}, self.details(), self.paths)
        record = request_for_inward(self.root, 'plant1', 'CLIENT1')
        details = self.details()
        details['certificate_type'] = 'standard'
        with self.assertRaisesRegex(ValueError, 'own certificate record'):
            save_draft(self.root, 'plant1', details, record)
        with self.assertRaisesRegex(ValueError, 'linked inward'):
            save_draft(self.root, 'plant2', self.details(), record)

    def test_spec_and_two_page_preview(self):
        self.assertEqual(template_id('HDD COD'), 'destruction')
        self.assertIn('destruction', {spec['id'] for spec in template_specs()})
        with patch.object(web_app, 'ROOT', self.root):
            response = web_app.preview_certificate({'source_kind':'manual', 'plant':'plant1', 'details':self.details()})
        self.assertEqual(len(response['preview_detail_images']), 1)
        self.assertFalse(read_register(self.root)['records'])

    def test_hdd_handles_a_broken_excel_style_reference(self):
        book = self.paths[1]
        original = book.read_bytes()
        damaged = BytesIO()
        with ZipFile(BytesIO(original)) as source, ZipFile(damaged, 'w') as target:
            for member in source.infolist():
                payload = source.read(member)
                if member.filename == 'xl/worksheets/sheet1.xml':
                    sheet = ET.fromstring(payload)
                    cell = sheet.find('.//{*}c[@r="A2"]')
                    cell.set('s', '65535')
                    payload = ET.tostring(sheet)
                target.writestr(member, payload)
        book.write_bytes(damaged.getvalue())

        with patch.object(web_app, 'ROOT', self.root):
            response = web_app.preview_certificate({
                'source_kind': 'manual', 'plant': 'plant1', 'details': self.details(),
            })
        self.assertEqual(len(response['preview_detail_images']), 1)
        self.assertTrue(any('Plant 2 workbook' in warning and 'invalid formatting' in warning
                            for warning in response['warnings']))
        with self.assertRaisesRegex(ValueError, 'invalid formatting or text reference'):
            request_for_inward(self.root, 'plant2', 'CLIENT2')
        from hdd_workbook import stage_hdd_tracking
        with self.assertRaisesRegex(ValueError, 'invalid formatting or text reference'):
            stage_hdd_tracking(SimpleNamespace(data=book.read_bytes()), 'CLIENT2', '0303', AS_OF)
        self.assertEqual(book.read_bytes(), damaged.getvalue())

    def test_filled_reference_uses_bold_values_slash_dates_and_padded_cod(self):
        details = self.details(272)
        details.update(supplier='Duck Creek Technologies India LLP',location='Mumbai')
        details['certificate_fields'] = {'destruction_date':'2026-09-09',
            'device_receipts':json.dumps([{'date':'2026-09-09','count':'272'}])}
        path = self.root / 'reference-preview.pdf'
        make_template_certificate(details,'0172',date(2026,10,8),path)
        pdf = PdfReader(path)
        front = pdf.pages[0].extract_text()
        self.assertIn('DUCK CREEK TECHNOLOGIES INDIA LLP',front)
        self.assertIn('MUMBAI',front)
        self.assertIn('0172',front)
        self.assertIn('09/09/2026',front)
        self.assertIn('08/10/2026',front)
        self.assertNotIn('08.10.2026',front)
        self.assertEqual(pdf.pages[1].extract_text().count('272'),2)
        fonts=[]
        def visit(text, cm, tm, font, size):
            if text.strip(): fonts.append((str(font.get('/BaseFont','')),size))
        pdf.pages[0].extract_text(visitor_text=visit)
        self.assertTrue(all('Bold' in font and size == 14 for font,size in fonts))
        self.assertFalse((self.root/'Data'/'Issuance History').exists())

    def test_receipt_dates_destruction_date_and_issue_date_remain_independent(self):
        details=self.details()
        details['certificate_fields']={'destruction_date':'2026-09-15',
            'device_receipts':json.dumps([{'date':'2026-09-01','count':'50'},{'date':'2026-09-09','count':'222'}])}
        path=self.root/'separate-dates.pdf'
        make_template_certificate(details,'0303',date(2026,10,8),path)
        pdf=PdfReader(path)
        self.assertIn('15/09/2026',pdf.pages[0].extract_text())
        self.assertIn('08/10/2026',pdf.pages[0].extract_text())
        text=pdf.pages[1].extract_text()
        for value in ('01/09/2026','09/09/2026','50','222','272'): self.assertIn(value,text)

    def test_templates_are_blank_and_use_only_the_supplied_backgrounds(self):
        from certificate_templates import TEMPLATE_ROOT
        pdf=PdfReader(TEMPLATE_ROOT/'destruction.pdf')
        self.assertEqual(len(pdf.pages),2)
        for page in pdf.pages:
            self.assertEqual(len(page.images),1)
            self.assertFalse(page.extract_text().strip())
            self.assertEqual((float(page.mediabox.width),float(page.mediabox.height)),(960,540))
        with ZipFile(TEMPLATE_ROOT/'destruction.pptx') as archive:
            for name in archive.namelist():
                if name.startswith('ppt/slides/slide') and name.endswith('.xml'):
                    root=ET.fromstring(archive.read(name))
                    self.assertFalse(any((text.text or '').strip() for text in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t')))


if __name__ == '__main__':
    unittest.main()
