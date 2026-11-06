"""Folder-only certificates feed analytics without inventing intake or dates."""
from datetime import date
import hashlib
import json
from pathlib import Path
import tempfile
from unittest import TestCase
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

from analytics_pdf import folder_events, _inspect
from certificate_inventory import scan_inventory, filename_cod, text_identifiers
from certificate_templates import make_template_certificate
from certificate_analytics import build_analytics
from run_certificates import certificate_filename, numbering_plan
from workflow_support import load_settings

TODAY = date(2026,10,19)


class FolderAnalyticsTests(TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name)
        self.folder=self.root/'Certificates'/'Auto Generated'; self.folder.mkdir(parents=True)
        _inspect.cache_clear()

    def pdf(self,name,lines):
        path=self.folder/name
        pdf=canvas.Canvas(str(path),invariant=True)
        for index,line in enumerate(lines): pdf.drawString(40,760-index*25,line)
        pdf.save(); return path

    def test_folder_only_hdd_and_sita_use_printed_issue_date_and_shared_sequence(self):
        hdd=self.pdf('0400_Client_London_hdd.pdf',['Secured Data Destruction','HDD No.: 0400','Certificate Date: 16/10/2026'])
        self.pdf('0401_Client_Mumbai.pdf',['CERTIFICATE OF E-WASTE RECYCLING','SITA INFORMATION NETWORKING','Certificate Date: 17.10.2026'])
        before={p:p.read_bytes() for p in self.folder.iterdir()}
        inventory=scan_inventory(self.root)
        events=folder_events(self.root,inventory)
        result=build_analytics([],[],saved_pdfs=events,today=TODAY)
        self.assertEqual(result['totals']['generated'],2)
        self.assertEqual(result['calendar']['today'],0)
        self.assertEqual(result['totals']['received'],0)
        self.assertEqual({item['key']:item['count'] for item in result['types']},dict(standard=0,destruction=1,siemens=0,sita=1))
        plan=numbering_plan([],load_settings(self.root),self.root/'Certificates',inventory=inventory)
        self.assertEqual(plan['next_cod'],'0402')
        self.assertEqual(before,{p:p.read_bytes() for p in self.folder.iterdir()})
        self.assertFalse((self.root/'Data').exists())
        with patch('analytics_pdf._inspect',side_effect=AssertionError('known certificate was parsed')):
            self.assertEqual(folder_events(self.root,inventory,['400','401']),[])

    def test_legacy_hdd_table_uses_issue_date_not_destruction_or_receipt_date(self):
        path=self.folder/'0172_Client_Mumbai_hdd.pdf'
        make_template_certificate({'certificate_type':'destruction','supplier':'Client','location':'Mumbai',
            'certificate_fields':{'destruction_date':'2026-09-15','device_receipts':json.dumps([{'date':'2026-08-01','count':'12'}])}},
            '0172',date(2026,10,8),path)
        reader=PdfReader(path); writer=PdfWriter(); writer.append_pages_from_reader(reader)
        with path.open('wb') as stream: writer.write(stream)
        events=folder_events(self.root,scan_inventory(self.root))
        self.assertEqual(events[0]['date'],date(2026,10,8))
        self.assertEqual(events[0]['type'],'destruction')

    def test_only_receipt_or_destruction_date_does_not_become_issue_date(self):
        self.pdf('0204_Client_Mumbai_hdd.pdf',['HDD No.: 0204','Date of destruction: 19/10/2026','Date received: 01/10/2026'])
        result=build_analytics([],[],saved_pdfs=folder_events(self.root,scan_inventory(self.root)),today=TODAY)
        self.assertEqual(result['totals']['generated'],0)
        self.assertEqual(result['coverage']['folder_needs_details'],1)

    def test_normal_recycling_footer_does_not_misclassify_as_destruction(self):
        self.pdf('0205_Client_Mumbai.pdf',['CERTIFICATE OF E-WASTE RECYCLING','Certificate No. & Date 0205 17.10.2026','Services: Secured Data Destruction'])
        event=folder_events(self.root,scan_inventory(self.root))[0]
        self.assertEqual(event['type'],'standard'); self.assertEqual(event['date'],date(2026,10,17))

    def test_replaced_pdf_is_not_read_using_stale_inventory(self):
        path=self.pdf('0205_Client_Mumbai_hdd.pdf',['Certificate Date: 17/10/2026'])
        digest=hashlib.sha256(path.read_bytes()).hexdigest()
        self.pdf(path.name,['Certificate Date: 18/10/2026'])
        self.assertIsNone(_inspect(str(path),digest))

    def test_hdd_labels_and_filename_prefixes_reserve_the_numeric_sequence(self):
        self.assertEqual(filename_cod('HDD_0204_Client_Mumbai_hdd.pdf'),'0204')
        self.assertEqual(text_identifiers('HDD No.: 0204')[0],['0204'])
        name=certificate_filename({'certificate_type':'destruction','supplier':'A'*500,'location':'B'*500},cod='0204')
        self.assertTrue(name.endswith('_HDD.pdf')); self.assertLessEqual(len(Path(name).stem),240)
