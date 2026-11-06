"""Exercise actual PDF copying, preview, issuance, reopening and removal."""
from datetime import date
from io import BytesIO
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from PIL import Image, ImageChops

from certificate_drafts import generate_draft, read_register
from pdf_attachments import store_attachment, attachment_catalog
from workspace_review import workspace_revision
from workflow_support import save_json
from tests.test_certificate_templates import write_register
import web_app


AS_OF = date(2026, 10, 19)


def make_list(path, prefix='LIST A', count=2):
    pdf = canvas.Canvas(str(path), pagesize=(300, 400))
    for index in range(count):
        pdf.drawString(30, 350, f'{prefix} PAGE {index + 1}')
        pdf.showPage()
    pdf.save()


class AttachmentTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.paths = [self.root / f'Plant {i}.xlsx' for i in (1, 2)]
        for i, path in enumerate(self.paths, 1):
            write_register(path, case_id=f'SN{i}', cod=f'{300+i:04}', status='Completed')
        save_json(self.root / 'settings.local.json', {'workbook_paths':list(map(str,self.paths))})
        self.source = self.root / 'My Device List.pdf'
        make_list(self.source)

    def details(self, kind='standard', token=None):
        fields = {'_annexure':token} if token else {}
        if kind == 'siemens':
            fields.update(model_name='Model A', model_serial='M1', tube_serial='T1', site_location='Pune')
        if kind == 'destruction':
            fields.update(destruction_date='2026-10-18',device_receipts='[{"date":"2026-10-01","count":"12"}]')
        return {'certificate_type':kind,'supplier':'Attachment Test Client', 'location':'Pune',
                'inward_date':'2026-10-01','certificate_weight':'10', 'form6':f'USER-F6-{kind}',
                'reference':'MY OWN REF / 09-A', 'certificate_fields':fields}

    def test_all_four_previews_and_final_certificates_include_every_attachment_page(self):
        attachment = store_attachment(self.root, self.source, workspace_revision(self.root))
        self.source.unlink()  # A retained attachment does not depend on the original path.
        before = [path.read_bytes() for path in self.paths]
        for kind in ('standard','sita','siemens','destruction'):
            with self.subTest(kind=kind):
                details = self.details(kind, attachment['id'])
                with patch.object(web_app,'ROOT',self.root):
                    preview = web_app.preview_certificate({'plant':'plant1','source_kind':'manual','details':details})
                preview_path = self.root / 'Data' / 'Previews' / Path(preview['preview_url']).name
                pdf = PdfReader(preview_path)
                self.assertEqual(len(pdf.pages), 4 if kind == 'destruction' else 3)
                self.assertIn('LIST A PAGE 1', pdf.pages[-2].extract_text())
                self.assertIn('LIST A PAGE 2', pdf.pages[-1].extract_text())
                attached = preview['preview_attachment']
                self.assertEqual(attached['name'],'My Device List.pdf')
                self.assertEqual(len(attached['pages']),2)
                self.assertEqual(len(preview['preview_detail_images']),1 if kind=='destruction' else 0)
                identifier = preview_path.stem
                with patch.object(web_app,'ROOT',self.root):
                    images = [web_app.preview_page_image(identifier,index) for index in range(len(pdf.pages)-1,len(pdf.pages)+1)]
                with Image.open(images[0]) as first, Image.open(images[1]) as second:
                    self.assertIsNotNone(ImageChops.difference(first.convert('RGB'),Image.new('RGB',first.size,'white')).getbbox())
                    self.assertIsNotNone(ImageChops.difference(first.convert('RGB'),second.convert('RGB')).getbbox())
                result = generate_draft(self.root,'plant1',details,confirmed=True,as_of=AS_OF)
                output = PdfReader(self.root / 'Certificates' / result['relative_path'])
                self.assertEqual(len(output.pages),len(pdf.pages))
                self.assertIn('LIST A PAGE 2', output.pages[-1].extract_text())
                saved = read_register(self.root)['records'][-1]['details']
                self.assertEqual(saved['certificate_fields']['_annexure'],attachment['id'])
                self.assertEqual(saved['reference'],'MY OWN REF / 09-A')
        self.assertEqual(before,[path.read_bytes() for path in self.paths])
        self.assertEqual(attachment_catalog(self.root)[attachment['id']]['name'],'My Device List.pdf')

    def test_replace_remove_and_reopen_keep_correct_pages_and_same_cod(self):
        a = store_attachment(self.root,self.source,workspace_revision(self.root))
        details = self.details(token=a['id'])
        first = generate_draft(self.root,'plant1',details,as_of=AS_OF)
        replacement = self.root / 'Replacement.pdf'; make_list(replacement,'LIST B',1)
        b = store_attachment(self.root,replacement,workspace_revision(self.root))
        details['certificate_fields']['_annexure'] = b['id']
        second = generate_draft(self.root,'plant1',details,first['record_id'],as_of=AS_OF)
        pdf = PdfReader(self.root / 'Certificates' / second['relative_path'])
        self.assertEqual(len(pdf.pages),2)
        self.assertIn('LIST B',pdf.pages[-1].extract_text())
        self.assertNotIn('LIST A',' '.join(page.extract_text() or '' for page in pdf.pages))
        del details['certificate_fields']['_annexure']
        third = generate_draft(self.root,'plant1',details,first['record_id'],as_of=AS_OF)
        self.assertEqual(len(PdfReader(self.root/'Certificates'/third['relative_path']).pages),1)
        self.assertEqual(first['certificate_no'],third['certificate_no'])
        self.assertNotIn('_annexure',read_register(self.root)['records'][0]['details']['certificate_fields'])

    def test_invalid_encrypted_and_empty_pdfs_never_get_saved(self):
        bad = self.root / 'Broken.pdf'; bad.write_bytes(b'%PDF-1.7\nbroken')
        empty = self.root / 'Empty.pdf'; PdfWriter().write(empty)
        encrypted = self.root / 'Locked.pdf'
        writer = PdfWriter(); writer.add_blank_page(300,400); writer.encrypt('secret'); writer.write(encrypted)
        for path in (bad,empty,encrypted):
            with self.subTest(path=path.name), self.assertRaises(ValueError):
                store_attachment(self.root,path,workspace_revision(self.root))
        self.assertFalse(list((self.root/'Data'/'Attachments').glob('*')))

    def test_failed_metadata_write_and_stale_workspace_leave_no_partial_attachment(self):
        expected = workspace_revision(self.root)
        save_json(self.root/'Data'/'Custom Certificates.json',{'version':1,'records':[]})
        with self.assertRaisesRegex(ValueError,'files changed'):
            store_attachment(self.root,self.source,expected)
        with patch('workflow_support.save_json',side_effect=OSError('disk full')), self.assertRaises(OSError):
            store_attachment(self.root,self.source,workspace_revision(self.root))
        self.assertFalse(list((self.root/'Data'/'Attachments').glob('*')))

    def test_attachment_corrupted_after_selection_blocks_final_publication(self):
        a = store_attachment(self.root,self.source,workspace_revision(self.root))
        path = self.root/'Data'/'Attachments'/f'{a["id"]}.pdf';path.write_bytes(b'%PDF-1.7\nbroken')
        with self.assertRaises(ValueError):
            generate_draft(self.root,'plant1',self.details(token=a['id']),confirmed=True,as_of=AS_OF)
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))
        self.assertFalse(read_register(self.root)['records'])

    def test_attachment_pages_load_only_when_requested_and_page_bounds_are_checked(self):
        make_list(self.source,count=12)
        attachment=store_attachment(self.root,self.source,workspace_revision(self.root))
        with patch.object(web_app,'ROOT',self.root):
            preview=web_app.preview_certificate({'plant':'plant1','source_kind':'manual','details':self.details(token=attachment['id'])})
            identifier=Path(preview['preview_url']).stem
            self.assertEqual(len(preview['preview_attachment']['pages']),12)
            self.assertEqual(list((self.root/'Data'/'Previews').glob('*_page_*.png')),[])
            last=web_app.preview_page_image(identifier,13)
            self.assertTrue(last.read_bytes().startswith(b'\x89PNG'))
            before=last.stat().st_mtime_ns
            self.assertEqual(web_app.preview_page_image(identifier,13).stat().st_mtime_ns,before)
            for bad_id,page in [('../escape',1),(identifier,0),(identifier,14)]:
                with self.assertRaises(ValueError): web_app.preview_page_image(bad_id,page)
            (self.root/'Data'/'Previews'/f'{identifier}.pdf').unlink()
            with self.assertRaises(ValueError): web_app.preview_page_image(identifier,13)

    def test_nine_hdd_rows_stay_before_the_attached_pdf_preview(self):
        attachment=store_attachment(self.root,self.source,workspace_revision(self.root))
        details=self.details('destruction',attachment['id'])
        details['certificate_fields']['device_receipts']=json.dumps([{'date':'2026-10-01','count':'1'} for _ in range(9)])
        with patch.object(web_app,'ROOT',self.root):
            preview=web_app.preview_certificate({'plant':'plant1','source_kind':'manual','details':details})
        self.assertEqual(len(preview['preview_detail_images']),1)
        self.assertTrue(preview['preview_attachment']['pages'][0].endswith('/page/3.png'))
        self.assertTrue(preview['preview_attachment']['pages'][1].endswith('/page/4.png'))
        pdf=PdfReader(self.root/'Data'/'Previews'/Path(preview['preview_url']).name)
        self.assertNotIn('Certificate No.',pdf.pages[1].extract_text())
        self.assertIn('LIST A PAGE 1',pdf.pages[2].extract_text())


if __name__ == '__main__':
    unittest.main()
