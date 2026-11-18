"""Settings deletion tests use disposable workbooks, local records and PDFs."""
from datetime import date
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import uuid

from certificate_core import read_batches
from certificate_drafts import read_register, save_draft, generate_draft, complete_draft, current_record_pdf
from entry_deletion import delete_entries
from issuance_safety import latest_issuance, audit_records, recover_issuances
from workbook_store import WorkbookSession
from workflow_support import run_lock, load_settings
from workspace_review import workspace_revision
import web_app
from tests import test_simple_certificate_flow as fixtures


class EntryDeletionTests(unittest.TestCase):
    setUp = fixtures.SimpleCertificateFlowTests.setUp
    source = fixtures.SimpleCertificateFlowTests.source
    body = fixtures.SimpleCertificateFlowTests.body
    create = fixtures.SimpleCertificateFlowTests.create

    def issue_completed(self, hdd=False):
        payload = self.body('standard')
        payload.update(review_edits=True, edit_scope='certificate_only', data_destruction_needed=hdd)
        self.assertEqual(self.create(payload)['created'], 1)
        source = self.source()
        with WorkbookSession(self.paths[0], self.root/'Data'/'Workbook Backups') as session:
            session.set_issued_status(source['case_id'], source['source_sheet'], source['source_row'],
                                      source['cod_no'], 'Generated', 'Completed')
        if hdd:
            record = read_register(self.root)['records'][0]
            details = dict(record['details'], certificate_fields={
                'destruction_date':'2026-10-29','device_receipts':'[{"date":"2026-09-01","count":"2"}]'})
            self.assertEqual(generate_draft(self.root,'plant1',details,record_id=record['id'],
                as_of=date(2026,10,30),expected_revision=workspace_revision(self.root))['created'],1)
        return latest_issuance(self.root,'workbook','plant1','1001')

    def payload(self, *, selected=None):
        if selected is None:
            source = self.source()
            selected = [dict(source_kind='workbook',plant='plant1',sn=source['case_id'],
                             sheet=source['source_sheet'],row=source['source_row'])]
        return dict(selected=selected, request_id=uuid.uuid4().hex,
                    expected_revision=workspace_revision(self.root))

    def linked_choices(self):
        identifier = self.pending_hdd()
        return [*self.payload()['selected'], {'source_kind':'manual', 'plant':'plant1', 'record_id':identifier}]

    def reviewed(self, **options):
        payload = self.payload(**options)
        review = delete_entries(self.root,payload)
        return dict(payload,checked=True,review_token=review['review_token']), review['review']

    def snapshot(self):
        return {str(path.relative_to(self.root)):path.read_bytes() for path in self.root.rglob('*') if path.is_file()}

    def pending_hdd(self):
        from destruction_certificate import request_for_inward
        return request_for_inward(self.root,'plant1','1001',workspace_revision(self.root))

    def test_completed_entry_cannot_be_reviewed_or_deleted(self):
        self.issue_completed()
        payload=self.payload();before=self.snapshot()
        for commit in (False,True):
            with self.subTest(commit=commit),self.assertRaisesRegex(ValueError,'generated.*cannot be deleted'):
                delete_entries(self.root,payload,commit=commit)
            self.assertEqual(self.snapshot(),before)

    def test_missing_pdf_and_cleared_status_do_not_hide_issuance(self):
        receipt=self.issue_completed()
        (self.root/receipt['pdf_path']).unlink()
        from openpyxl import load_workbook
        book=load_workbook(self.paths[0]);book.active['M2']='Ready';book.save(self.paths[0]);book.close()
        before=self.snapshot()
        with self.assertRaisesRegex(ValueError,'cannot be deleted'):
            delete_entries(self.root,self.payload())
        self.assertEqual(self.snapshot(),before)

    def test_generated_local_certificate_cannot_be_deleted(self):
        details=self.body('sita')['details'];details.update(case_id='',form6='')
        result=generate_draft(self.root,'plant1',details,starting_cod='0001')
        payload=self.payload(selected=[dict(source_kind='manual',plant='plant1',record_id=result['record_id'])])
        before=self.snapshot()
        with self.assertRaisesRegex(ValueError,'cannot be deleted'):
            delete_entries(self.root,payload)
        self.assertEqual(self.snapshot(),before)

    def test_pending_hdd_can_be_deleted_with_its_unissued_inward(self):
        payload,review=self.reviewed(selected=self.linked_choices())
        self.assertEqual(len(review['entries']),2);self.assertEqual(review['pdfs'],[])
        self.assertTrue(delete_entries(self.root,payload,commit=True)['deleted'])
        self.assertEqual(read_batches(self.paths[0]),[])
        self.assertEqual(read_register(self.root)['records'],[])

    def test_linked_and_standalone_hdd_can_be_selected_without_deleting_excel(self):
        linked = self.pending_hdd()
        standalone = save_draft(self.root, None, {'certificate_type':'destruction', 'supplier':'Standalone HDD'})
        before = self.paths[0].read_bytes()
        for identifier, plant in ((linked, 'plant1'), (standalone, None)):
            with self.subTest(identifier=identifier):
                selected = [{'source_kind':'manual', 'plant':plant, 'record_id':identifier}]
                payload, review = self.reviewed(selected=selected)
                self.assertEqual([entry['record_id'] for entry in review['entries']], [identifier])
                self.assertTrue(delete_entries(self.root, payload, commit=True)['deleted'])
                self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertEqual(read_register(self.root)['records'], [])

    def test_pending_hdd_can_be_kept_standalone(self):
        identifier=self.pending_hdd()
        payload,review=self.reviewed()
        self.assertEqual(review['kept_hdd'][0]['record_id'],identifier)
        delete_entries(self.root,payload,commit=True)
        record=read_register(self.root)['records'][0]
        self.assertNotIn('destruction_source',record['details'])
        self.assertEqual(record['details']['case_id'],'')

    def test_generated_hdd_protects_its_unissued_inward(self):
        identifier=self.pending_hdd();record=read_register(self.root)['records'][0]
        details=dict(record['details'],certificate_fields={'destruction_date':'2026-10-29',
            'device_receipts':'[{"date":"2026-09-01","count":"2"}]'})
        generate_draft(self.root,'plant1',details,identifier,starting_cod='0001')
        before=self.snapshot()
        with self.assertRaisesRegex(ValueError,'cannot be deleted'):
            delete_entries(self.root,self.payload())
        self.assertEqual(self.snapshot(),before)

    def test_deletion_requires_checked_review_and_rejects_legacy_hdd_option(self):
        payload,_=self.reviewed();before=self.snapshot()
        for change in ({'checked':False},{'checked':'yes'},{'review_token':'forged'},{'include_linked_hdd':True}):
            with self.subTest(change=change),self.assertRaises(ValueError):
                delete_entries(self.root,dict(payload,**change),commit=True)
        self.assertEqual(self.snapshot(),before)

    def test_old_automatic_hdd_option_cannot_expand_a_new_review(self):
        self.pending_hdd()
        before = self.snapshot()
        with self.assertRaisesRegex(ValueError, 'Select each Data Destruction entry'):
            delete_entries(self.root, dict(self.payload(), include_linked_hdd=True))
        self.assertEqual(self.snapshot(), before)

    def test_external_change_invalidates_review(self):
        payload,_=self.reviewed()
        from openpyxl import load_workbook
        book=load_workbook(self.paths[1]);book.active['C2']='Changed outside app';book.save(self.paths[1]);book.close()
        before=self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError,'files changed'):
            delete_entries(self.root,payload,commit=True)
        self.assertEqual(self.paths[0].read_bytes(),before)

    def test_multiple_deletions_close_row_and_serial_gaps(self):
        from tests.test_certificate_workflow import make_workbook
        make_workbook(self.paths[0],[dict(sn='1',form6='F6-1'),dict(sn='2',form6='F6-2'),dict(sn='3',form6='F6-3',status='N/A')])
        rows=read_batches(self.paths[0]);other=self.paths[1].read_bytes()
        selected=[dict(source_kind='workbook',plant='plant1',sn=row['case_id'],sheet=row['source_sheet'],row=row['source_row']) for row in rows[:2]]
        payload,review=self.reviewed(selected=selected)
        self.assertEqual(len(review['entries']),2)
        delete_entries(self.root,payload,commit=True)
        remaining=read_batches(self.paths[0]);self.assertEqual(len(remaining),1)
        self.assertEqual((remaining[0]['case_id'],remaining[0]['source_row'],remaining[0]['form6']),('1',2,'F6-3'))
        self.assertEqual(self.paths[1].read_bytes(),other)

    def test_deletion_renumbers_a_certified_neighbour_without_losing_its_pdf_or_hdd_link(self):
        from tests.test_certificate_workflow import make_workbook
        from run_certificates import run, find_certificate
        from destruction_certificate import request_for_inward
        make_workbook(self.paths[0],[dict(sn='1',form6='F6-1'),dict(sn='2',form6='F6-2')])
        self.assertEqual(run(self.paths,self.root,selected=('plant1','2'),starting_cod='0001',progress=None)[0],1)
        receipt=latest_issuance(self.root,'workbook','plant1','2');pdf=self.root/receipt['pdf_path'];pdf_before=pdf.read_bytes()
        identifier=request_for_inward(self.root,'plant1','2',workspace_revision(self.root))
        payload,_=self.reviewed();delete_entries(self.root,payload,commit=True)
        kept=read_batches(self.paths[0])[0]
        self.assertEqual((kept['case_id'],kept['source_row'],kept['cod_no']),('1',2,receipt['cod']))
        self.assertEqual(latest_issuance(self.root,'workbook','plant1','1')['id'],receipt['id'])
        self.assertEqual(read_register(self.root)['records'][0]['details']['destruction_source']['case_id'],'1')
        self.assertEqual(pdf.read_bytes(),pdf_before)

    def test_failed_workbook_save_keeps_entry_and_local_request(self):
        payload,_=self.reviewed(selected=self.linked_choices());before=self.snapshot()
        with patch.object(WorkbookSession,'_persist',side_effect=OSError('Excel is busy')):
            with self.assertRaisesRegex(OSError,'Excel is busy'):
                delete_entries(self.root,payload,commit=True)
        self.assertEqual(self.paths[0].read_bytes(),before[str(self.paths[0].relative_to(self.root))])
        self.assertEqual((self.root/'Data'/'Custom Certificates.json').read_bytes(),before['Data\\Custom Certificates.json'])

    def test_interrupted_multi_source_delete_is_restored_on_recovery(self):
        class Interrupted(BaseException): pass
        payload,_=self.reviewed(selected=self.linked_choices())
        before=self.paths[0].read_bytes(),(self.root/'Data'/'Custom Certificates.json').read_bytes()
        persist=WorkbookSession._persist
        def interrupt(session,data,**options):
            result=persist(session,data,**options)
            raise Interrupted()
        with patch.object(WorkbookSession,'_persist',interrupt):
            with self.assertRaises(Interrupted):delete_entries(self.root,payload,commit=True)
        with run_lock(self.root):recover_issuances(self.root)
        self.assertEqual(self.paths[0].read_bytes(),before[0])
        self.assertEqual((self.root/'Data'/'Custom Certificates.json').read_bytes(),before[1])

    def test_lost_response_retry_does_not_delete_a_replacement_inward(self):
        payload,_=self.reviewed();delete_entries(self.root,payload,commit=True)
        from tests.test_certificate_workflow import make_workbook
        make_workbook(self.paths[0],[dict(sn='1001',supplier='Replacement inward')])
        before=self.paths[0].read_bytes()
        self.assertTrue(delete_entries(self.root,payload,commit=True)['deleted'])
        self.assertEqual(self.paths[0].read_bytes(),before)

    def test_assigned_but_ungenerated_cod_can_be_deleted_and_stays_reserved(self):
        from openpyxl import load_workbook
        book=load_workbook(self.paths[0]);book.active['K2']='0900';book.active['M2']='Ready';book.save(self.paths[0]);book.close()
        payload,review=self.reviewed();self.assertEqual(review['pdfs'],[])
        delete_entries(self.root,payload,commit=True)
        from run_certificates import numbering_plan
        self.assertGreaterEqual(numbering_plan([],load_settings(self.root),self.root/'Certificates')['sequence'],901)


if __name__=='__main__':unittest.main()
