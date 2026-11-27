"""Alert lifecycle, scheduler contention and bounded read caches on disposable data."""
from datetime import date
from pathlib import Path
from threading import Event, Thread
import json
import os
import tempfile
import time
import unittest
from unittest.mock import Mock, patch

import certificate_drafts as drafts
import file_fingerprints
import issuance_safety as safety
import preview_store
import run_certificates as runner
import scheduled_run
import web_app
import workflow_support as support
from workbook_store import LockBusyError
from workspace_review import workspace_revision, require_revision
from tests import test_missing_certificate_recovery as recovery_tests
from openpyxl import load_workbook


class AlertReliabilityTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory()))

    def record(self, **changes):
        row = {'id': 'a' * 32, 'plant': 'plant1', 'status': 'completed', 'warnings': [],
               'details': {'certificate_type': 'standard', 'supplier': 'Audit client',
                           'inward_date': '2026-10-01', 'cod_no': '0001',
                           'certificate_weight': 10, 'certificate_fields': {}}}
        row.update(changes)
        support.save_json(drafts.register_path(self.root), {'version': 1, 'records': [row]})
        return row

    def test_local_missing_pdf_alert_survives_reconciliation_and_repeated_reads(self):
        self.record()
        with patch.object(drafts, 'current_record_pdf', return_value=None):
            drafts.reconcile_manual(self.root)
            for _ in range(2):
                row = drafts.manual_rows(self.root, as_of=date(2026, 10, 12))[0]
                self.assertEqual(row['certificate_status'], 'ready')
                self.assertTrue(row['system_alert'])
                self.assertTrue(row['pdf_missing'])
                self.assertIn('PDF missing', row['detail'])
                self.assertIn(row, support._actionable_rows({'rows': [row]}))

    def test_workbook_missing_pdf_alert_survives_next_checks_without_rewriting_status(self):
        case = recovery_tests.MissingCertificateRecoveryTests()
        case.setUp()
        self.addCleanup(case.doCleanups)
        case.source(status='Generated')
        book = load_workbook(case.paths[0])
        book.active['B2'] = date(2026, 9, 20)
        book.save(case.paths[0]); book.close()
        for _ in range(3):
            report = {}
            runner.run(case.paths, case.root, as_of=date(2026, 10, 1), check_only=True,
                       progress=None, report_sink=report)
            row = case.report_row(report)
            self.assertTrue(row['system_alert'])
            self.assertTrue(row['pdf_missing'])
            self.assertFalse(row['reminder_due'])
            self.assertEqual(row['certificate_status'], 'ready')

    def test_midnight_invalidates_unchanged_file_stamps(self):
        with patch.object(web_app, 'ROOT', self.root), patch.object(web_app, 'date') as day, \
                patch.object(web_app, 'WORKBOOK_CHECK_STAMPS', {}), \
                patch.object(web_app, 'LINKED_STATUS_STAMPS', {}), \
                patch.object(web_app, 'configured_paths', return_value=[]), \
                patch.object(web_app, 'load_settings', return_value={'workbook_paths': ['a', 'b']}), \
                patch.object(web_app, 'latest_result', return_value={'rows': [], 'scan_complete': True}), \
                patch.object(web_app, 'refresh', return_value={'rows': []}) as refresh:
            day.today.return_value = date(2026, 10, 25)
            stamp = web_app.workbook_check_stamp([])
            web_app.WORKBOOK_CHECK_STAMPS[str(self.root)] = stamp
            web_app.LINKED_STATUS_STAMPS[str(self.root)] = stamp
            self.assertFalse(web_app.check_workbook_changes()['changed'])
            day.today.return_value = date(2026, 10, 26)
            self.assertTrue(web_app.check_workbook_changes()['changed'])
            refresh.assert_called_once_with([])

    def test_disconnected_dashboard_detects_new_local_records_and_next_day(self):
        with patch.object(web_app, 'ROOT', self.root), patch.object(web_app, 'date') as day, \
                patch.object(web_app, 'WORKBOOK_CHECK_STAMPS', {}), \
                patch.object(web_app, 'load_settings', return_value={'workbook_paths': ['', '']}), \
                patch.object(web_app, 'refresh', return_value={'rows': []}) as refresh:
            day.today.return_value = date(2026, 10, 25)
            web_app.WORKBOOK_CHECK_STAMPS[str(self.root)] = ('disconnected', web_app.workbook_check_stamp([]))
            self.assertFalse(web_app.check_workbook_changes()['changed'])
            self.record()
            self.assertTrue(web_app.check_workbook_changes()['changed'])
            web_app.WORKBOOK_CHECK_STAMPS[str(self.root)] = ('disconnected', web_app.workbook_check_stamp([]))
            day.today.return_value = date(2026, 10, 26)
            self.assertTrue(web_app.check_workbook_changes()['changed'])
            self.assertEqual(refresh.call_count, 2)

    def test_new_issue_notifies_but_reordered_duplicate_warnings_do_not(self):
        row = {'plant': 'plant1', 'sn': '1', 'certificate_status': 'ready',
               'age_days': 12, 'system_alert': True, 'warnings': ['First issue']}
        notifier = Mock()
        options = {'today': date(2026, 10, 26)}
        self.assertTrue(support.notify_once(self.root, {'rows': [row]}, notifier, **options))
        row['warnings'] = ['Second issue', 'First issue']
        self.assertTrue(support.notify_once(self.root, {'rows': [row]}, notifier, **options))
        row['warnings'] = ['First issue', 'Second issue', 'First issue']
        self.assertFalse(support.notify_once(self.root, {'rows': [row]}, notifier, **options))
        self.assertEqual(notifier.call_count, 2)

    def test_retired_references_are_hidden_for_every_local_certificate_type(self):
        old = 'Reference SHARED is also used by another entry. You can keep the same reference and continue.'
        for kind in ('standard', 'sita', 'siemens', 'destruction'):
            row = self.record(status='ready', warnings=[old, 'An actual issue'])
            row['details']['certificate_type'] = kind
            if kind == 'destruction':
                row['details']['certificate_fields'] = {
                    'destruction_date':'2026-10-16',
                    'device_receipts':json.dumps([{'date':'2026-10-14', 'count':'1'}]),
                }
            support.save_json(drafts.register_path(self.root), {'version': 1, 'records': [row]})
            self.assertEqual(drafts.manual_rows(self.root)[0]['warnings'], ['An actual issue'])
            self.assertEqual(drafts.read_register(self.root)['records'][0]['warnings'], [old, 'An actual issue'])

    def test_scheduler_covers_local_hdd_without_workbooks(self):
        pending = {'source_kind': 'manual', 'record_id': 'hdd', 'plant': 'plant1',
                   'destruction_pending': True, 'reminder_due': False}
        with patch.object(scheduled_run, 'load_settings', return_value={
                'workbook_paths': ['', ''], 'reminders_enabled': True, 'reminder_time': '00:00'}), \
                patch.object(drafts, 'reconcile_manual'), patch.object(drafts, 'manual_rows', return_value=[pending]), \
                patch.object(runner, 'run') as run, \
                patch.object(scheduled_run, 'notify_once', return_value=True) as notify:
            self.assertEqual(scheduled_run.main(self.root), 0)
            run.assert_not_called()
            rows = notify.call_args.args[1]['rows']
            self.assertEqual(support._actionable_rows({'rows': rows}), [pending])
            self.assertEqual(support.alert_workbook_targets({'rows': rows}, ['a.xlsx', 'b.xlsx']), [])

    def test_local_only_reminder_setup_can_be_enabled(self):
        from validate_daily_setup import validate_daily_setup
        support.save_json(self.root / 'settings.local.json', {'workbook_paths': ['', '']})
        self.assertEqual(validate_daily_setup(self.root), ())

    def test_busy_scheduler_waits_past_four_seconds_then_succeeds(self):
        with patch.object(scheduled_run, 'load_settings', return_value={
                'workbook_paths': ['a', 'b'], 'reminders_enabled': True, 'reminder_time': '00:00'}), \
                patch.object(runner, 'run', side_effect=[LockBusyError('busy')] * 5 + [None]) as run, \
                patch.object(drafts, 'reconcile_manual'), patch.object(drafts, 'manual_rows', return_value=[]), \
                patch.object(scheduled_run.time, 'sleep') as sleep, \
                patch.object(scheduled_run, 'notify_once', return_value=False):
            self.assertEqual(scheduled_run.main(self.root), 0)
            self.assertEqual(run.call_count, 6)
            self.assertEqual(sum(call.args[0] for call in sleep.call_args_list), 50)

    def test_ordinary_dismissal_starts_no_powershell_process(self):
        with patch.object(support, '_powershell') as powershell:
            support.dismiss_reminders_today(self.root)
        powershell.assert_not_called()
        saved = json.loads((self.root / 'Data/reminder-state.json').read_text())
        self.assertEqual(saved['dismissed_day'], date.today().isoformat())

    def test_snooze_marker_only_exists_after_success_and_cancel_clears_it(self):
        marker = self.root / 'Data/Snoozed Reminder.json'
        with patch.object(support, '_powershell', return_value=Mock(returncode=0)) as powershell:
            support.schedule_one_hour_reminder(self.root)
            self.assertTrue(marker.is_file())
            self.assertTrue(support.cancel_snoozed_reminder(self.root))
            self.assertFalse(marker.exists())
            self.assertEqual(powershell.call_count, 2)
            self.assertFalse(support.cancel_snoozed_reminder(self.root))
            self.assertEqual(powershell.call_count, 2)
        with patch.object(support, '_powershell', return_value=Mock(returncode=1, stderr='failed')):
            with self.assertRaises(RuntimeError):
                support.schedule_one_hour_reminder(self.root)
        self.assertFalse(marker.exists())

    def test_dismissed_snooze_does_not_scan_workbooks_or_launch_cleanup_without_marker(self):
        support.save_json(self.root / 'Data/reminder-state.json', {
            'dismissed_day': date.today().isoformat(), 'dismissed_time': '00:00'})
        with patch.object(scheduled_run, 'load_settings', return_value={
                'workbook_paths': ['a', 'b'], 'reminders_enabled': True, 'reminder_time': '00:00'}), \
                patch.object(runner, 'run') as run, patch.object(support, '_powershell') as powershell:
            self.assertEqual(scheduled_run.main(self.root, reminder_only=True), 0)
        run.assert_not_called(); powershell.assert_not_called()

    def test_local_certificate_reminders_stop_on_completion(self):
        for status, expected in [('ready', True), ('generated_manual', True), ('completed', False)]:
            self.record(status=status)
            with patch.object(drafts, 'current_record_pdf', return_value=self.root / 'current.pdf'):
                row = drafts.manual_rows(self.root, as_of=date(2026, 10, 26))[0]
            self.assertEqual(row['age_days'], 25)
            self.assertEqual(row['reminder_due'], expected)

    def test_desktop_alert_destroys_window_before_slow_action_and_ignores_second_click(self):
        closed, started, release = Event(), Event(), Event()
        window = Mock()
        window.winfo_screenwidth.return_value = 1920
        window.winfo_screenheight.return_value = 1080
        window.winfo_width.return_value = 500
        window.winfo_height.return_value = 220
        window.destroy.side_effect = closed.set
        order = []
        def slow_action():
            order.append(closed.is_set())
            started.set()
            release.wait(3)
        with patch.object(support.tk, 'Tk', return_value=window), \
                patch.object(support.tk, 'Frame'), patch.object(support.tk, 'Label'), \
                patch.object(support.tk, 'Button') as button:
            window.mainloop.side_effect = lambda: button.call_args_list[2].kwargs['command']()
            worker = Thread(target=support.show_centered_reminder,
                            args=('Certificates', 'One record needs a check'),
                            kwargs={'on_dismiss': slow_action, 'error_root': self.root})
            worker.start()
            try:
                self.assertTrue(closed.wait(2))
                self.assertTrue(started.wait(2))
                self.assertTrue(worker.is_alive(), 'Background cleanup should still be blocked after the window closed')
                button.call_args_list[2].kwargs['command']()
                self.assertEqual(order, [True])
            finally:
                release.set(); worker.join(3)
            self.assertFalse(worker.is_alive())
            window.destroy.assert_called_once()
            window.focus_force.assert_not_called()
            self.assertNotIn((('-topmost', True), {}), [(call.args, call.kwargs) for call in window.attributes.call_args_list])

    def test_receipt_index_reads_history_once_and_invalidates_after_a_commit(self):
        folder = self.root / 'Data/Issuance History'
        for index in range(10):
            support.save_json(folder / f'{index}.json', {'version': 1, 'id': str(index),
                'state': 'complete', 'source_kind': 'workbook', 'plant': 'plant1',
                'key': str(index), 'cod': str(index), 'started_at': f'{index:02}'})
        original = Path.read_text
        reads = []
        def counted(path, *args, **kwargs):
            if path.parent == folder:
                reads.append(path)
            return original(path, *args, **kwargs)
        with patch.object(Path, 'read_text', counted):
            for index in range(10):
                receipt = safety.latest_issuance(self.root, 'workbook', 'plant1', str(index))
                self.assertEqual(receipt['cod'], str(index))
                receipt['cod'] = 'caller edit'
            self.assertEqual(len(reads), 10)
        saved = json.loads((folder / '9.json').read_text())
        saved['cod'] = '99'
        support.save_json(folder / '9.json', saved)
        self.assertEqual(safety.latest_issuance(self.root, 'workbook', 'plant1', '9')['cod'], '99')

    def test_publication_revision_forces_bytes_even_when_display_fingerprint_is_cached(self):
        path = self.root / 'Certificates/current.pdf'
        path.parent.mkdir()
        path.write_bytes(b'first bytes')
        stamp = file_fingerprints.file_signature(path)
        with patch.object(file_fingerprints, 'file_signature', return_value=stamp):
            revision = workspace_revision(self.root)
            path.write_bytes(b'other bytes')
            self.assertEqual(workspace_revision(self.root), revision)
            with self.assertRaisesRegex(ValueError, 'files changed'):
                require_revision(self.root, revision)

    def test_unavailable_change_metadata_disables_digest_reuse(self):
        path = self.root / 'file.pdf'
        path.write_bytes(b'first bytes')
        signature = (*file_fingerprints.file_signature(path)[:-1], None)
        with patch.object(file_fingerprints, 'file_signature', return_value=signature):
            first = file_fingerprints.fingerprint(path)[0]
            path.write_bytes(b'other bytes')
            second = file_fingerprints.fingerprint(path)[0]
        self.assertNotEqual(first, second)

    def test_preview_retention_keeps_active_group_and_removes_only_disposable_groups(self):
        folder = self.root / 'Data/Previews'
        folder.mkdir(parents=True)
        old, active = 'a' * 32, 'b' * 32
        for name in [f'{old}.pdf', f'{old}_page_2.png', f'{active}.pdf', f'{active}_page_2.png',
                     '.unfinished.tmp.pdf', 'important.pdf']:
            path = folder / name
            path.write_bytes(b'preview')
            os.utime(path, (time.time() - 3600, time.time() - 3600))
        preview_store.retain_preview(self.root, old, 'editor')
        preview_store.retain_preview(self.root, active, 'editor')
        preview_store.cleanup_previews(self.root, max_bytes=0, force=True)
        self.assertFalse((folder / f'{old}.pdf').exists())
        self.assertFalse((folder / f'{old}_page_2.png').exists())
        self.assertTrue((folder / f'{active}.pdf').exists())
        self.assertTrue((folder / f'{active}_page_2.png').exists())
        self.assertTrue((folder / '.unfinished.tmp.pdf').exists())
        self.assertTrue((folder / 'important.pdf').exists())

    def test_live_editor_presence_keeps_preview_protected_until_panel_closes(self):
        folder = self.root / 'Data/Previews'
        folder.mkdir(parents=True)
        identifier = 'c' * 32
        path = folder / f'{identifier}.pdf'
        path.write_bytes(b'preview')
        os.utime(path, (time.time() - 100000, time.time() - 100000))
        preview_store.retain_preview(self.root, identifier, 'editor', 'client')
        preview_store.preview_presence(self.root, 'client', 'editor')
        preview_store.cleanup_previews(self.root, max_bytes=0, force=True)
        self.assertTrue(path.exists())
        preview_store.preview_presence(self.root, 'client')
        preview_store.cleanup_previews(self.root, max_bytes=0, force=True)
        self.assertFalse(path.exists())

    def test_redirected_preview_folder_is_never_cleaned(self):
        data = self.root / 'Data'
        data.mkdir()
        outside = self.root / 'Other files'
        outside.mkdir()
        victim = outside / ('d' * 32 + '.pdf')
        victim.write_bytes(b'keep this')
        os.utime(victim, (time.time() - 100000, time.time() - 100000))
        original = Path.resolve
        def redirected(path, *args, **kwargs):
            return outside if path == data / 'Previews' else original(path, *args, **kwargs)
        with patch.object(Path, 'resolve', redirected):
            preview_store.cleanup_previews(self.root, max_bytes=0, force=True)
        self.assertEqual(victim.read_bytes(), b'keep this')


if __name__ == '__main__':
    unittest.main()
