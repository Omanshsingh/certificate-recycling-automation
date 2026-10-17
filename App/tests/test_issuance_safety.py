"""Crash and stale-review tests; every file belongs to a temporary workspace."""
from datetime import date
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from reportlab.pdfgen.canvas import Canvas
from issuance_safety import IssuanceTransaction, recover_issuances, valid_pdf, audit_records
from run_certificates import find_certificate
from workflow_support import save_json
from workspace_review import workspace_revision, reviewed_operation


def pdf(path, text='Certificate No. 0001'):
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = Canvas(str(path))
    doc.drawString(40, 100, text)
    doc.save()


class SafetyTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.source = self.root / 'Data' / 'Custom Certificates.json'
        save_json(self.source, {'version': 1, 'records': []})
        self.destination = self.root / 'Certificates' / '0001_New_Mumbai_1.pdf'
        self.temporary = self.root / 'work.pdf'
        pdf(self.temporary)

    def tx(self, **values):
        return IssuanceTransaction(self.root, self.destination, self.temporary, self.source,
            self.source.read_bytes(), source_kind='manual', plant='plant1', key='CLIENT',
            cod='0001', issued_date=date(2026, 10, 1), before={'supplier': 'Old'},
            after={'supplier': 'New'}, warnings=['Missing manifest'], override_reason='Approved exception', **values)

    def test_nonempty_garbage_or_encrypted_file_is_not_a_certificate(self):
        self.destination.parent.mkdir(parents=True)
        self.destination.write_bytes(b'not a real PDF')
        self.assertFalse(valid_pdf(self.destination))
        self.assertIsNone(find_certificate(self.destination.parent, {'source_id': 'plant1', 'case_id': 'CLIENT', 'cod_no': '0001'}))

    def test_interruption_after_pdf_before_source_rolls_back_and_keeps_recovery_copy(self):
        original = self.source.read_bytes()
        tx = self.tx()
        tx.publish()
        self.assertTrue(self.destination.is_file())
        messages = recover_issuances(self.root)
        self.assertFalse(self.destination.exists())
        self.assertEqual(self.source.read_bytes(), original)
        self.assertTrue(messages)
        self.assertEqual(audit_records(self.root)[0]['state'], 'rolled_back')
        self.assertEqual(len(list((self.root / 'Data' / 'Archive' / 'Uncommitted Certificates').rglob('*.pdf'))), 1)
        self.assertEqual(recover_issuances(self.root), [])

    def test_interruption_after_source_commit_finishes_and_archives_old_pdf(self):
        old = self.root / 'Certificates' / '0001_Old_Mumbai_1.pdf'
        pdf(old, 'Old approved certificate')
        tx = self.tx(old_paths=[old])
        tx.publish()
        new_source = b'{"version": 1, "records": []}'
        tx.expect_source(new_source)
        self.source.write_bytes(new_source)
        recover_issuances(self.root)
        self.assertTrue(self.destination.is_file())
        self.assertFalse(old.exists())
        record = audit_records(self.root)[0]
        self.assertEqual(record['state'], 'complete')
        self.assertEqual(record['override_reason'], 'Approved exception')
        self.assertEqual(record['before']['supplier'], 'Old')
        self.assertTrue((self.root / record['previous_pdfs'][0]['archive']).exists())
        self.assertEqual(recover_issuances(self.root), [])

    def test_same_filename_reissue_retains_previous_version(self):
        pdf(self.destination, 'Old certificate contents')
        before = self.destination.read_bytes()
        tx = self.tx()
        tx.publish()
        tx.expect_source(b'updated')
        self.source.write_bytes(b'updated')
        tx.complete()
        self.assertEqual((self.root / tx.data['previous_pdfs'][0]['archive']).read_bytes(), before)

    def test_committed_recovery_retries_partial_or_complete_staged_pdf(self):
        for partial in (False, True):
            with self.subTest(partial=partial):
                tx = self.tx()
                tx.publish()
                tx.expect_source(b'updated source')
                self.source.write_bytes(b'updated source')
                self.destination.unlink()
                stage = self.destination.with_name(f'.{tx.id}.recover.pdf')
                stage.write_bytes(b'partial PDF' if partial else self.temporary.read_bytes())
                recover_issuances(self.root)
                self.assertEqual(self.destination.read_bytes(), self.temporary.read_bytes())
                self.assertEqual(audit_records(self.root)[-1]['state'], 'complete')
                self.assertFalse(stage.exists())

    def test_rollback_retries_partial_staged_previous_pdf(self):
        pdf(self.destination, 'Original certificate')
        original = self.destination.read_bytes()
        tx = self.tx()
        tx.publish()
        stage = self.destination.with_name(f'.{tx.id}.restore.pdf')
        stage.write_bytes(b'partial previous PDF')
        recover_issuances(self.root)
        self.assertEqual(self.destination.read_bytes(), original)
        self.assertEqual(audit_records(self.root)[-1]['state'], 'rolled_back')
        self.assertFalse(stage.exists())

    def test_unknown_source_or_pdf_edits_stop_recovery_without_overwriting(self):
        tx = self.tx()
        tx.publish()
        self.source.write_bytes(b'outside edit')
        with self.assertRaisesRegex(ValueError, 'source changed'):
            recover_issuances(self.root)
        self.assertEqual(self.source.read_bytes(), b'outside edit')
        self.assertTrue(self.destination.is_file())

    def test_last_version_missing_cannot_be_masked_by_old_same_cod_copy(self):
        tx = self.tx()
        tx.publish()
        tx.expect_source(b'updated')
        self.source.write_bytes(b'updated')
        tx.complete()
        # Workbook receipt for the same stable source, with an old file re-added.
        data = tx.data
        data['source_kind'] = 'workbook'
        save_json(tx.path, data)
        self.destination.unlink()
        pdf(self.root / 'Certificates' / '0001_Old_Mumbai_1.pdf')
        self.assertIsNone(find_certificate(self.root / 'Certificates', {'source_id':'plant1', 'case_id':'CLIENT', 'cod_no':'0001'}))

    def test_stale_revision_blocks_before_action(self):
        revision = workspace_revision(self.root)
        save_json(self.source, {'version': 1, 'records': [], 'changed': True})
        calls = []
        with self.assertRaisesRegex(ValueError, 'files changed'):
            reviewed_operation(self.root, '/api/create', {'expected_revision': revision}, lambda x: calls.append(x))
        self.assertFalse(calls)

    def test_confirmation_is_bound_to_exact_payload_and_warnings(self):
        revision = workspace_revision(self.root)
        calls = []
        def action(values):
            calls.append(values)
            return {'created':1} if values.get('confirmed') else {'requires_confirmation': True, 'warnings':['Missing weight']}
        payload = {'expected_revision':revision, 'details':{'supplier':'Acme'}}
        review = reviewed_operation(self.root, '/api/create', payload, action)
        confirmed = dict(payload, confirmed=True, review_token=review['review_token'], override_reason='Approved by manager')
        self.assertEqual(reviewed_operation(self.root, '/api/create', confirmed, action)['created'], 1)
        changed = dict(confirmed, details={'supplier':'Someone else'})
        self.assertTrue(reviewed_operation(self.root, '/api/create', changed, action)['requires_confirmation'])
        self.assertFalse(calls[-1]['confirmed'])

    def test_bare_confirmation_cannot_bypass_review_but_reviewed_acceptance_needs_no_reason(self):
        calls = []
        def action(values):
            calls.append(values)
            return {'created': 1} if values.get('confirmed') else {'requires_confirmation': True, 'warnings':['Missing weight']}
        payload = {'expected_revision':workspace_revision(self.root), 'confirmed':True}
        review = reviewed_operation(self.root, '/api/create', payload, action)
        self.assertTrue(review['requires_confirmation'])
        self.assertFalse(calls[-1]['confirmed'])
        confirmed = dict(payload, review_token=review['review_token'])
        self.assertEqual(reviewed_operation(self.root, '/api/create', confirmed, action)['created'], 1)
        self.assertEqual(calls[-1]['override_reason'], 'Warnings reviewed; user chose Continue anyway.')
        with self.assertRaisesRegex(ValueError, '2,000'):
            reviewed_operation(self.root, '/api/create', dict(confirmed, override_reason='x' * 2001), action)


if __name__ == '__main__':
    unittest.main()
