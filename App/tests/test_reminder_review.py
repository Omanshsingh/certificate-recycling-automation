"""Reminder navigation, dismissal and summaries; no Windows tasks are changed."""
from datetime import date, timedelta
from pathlib import Path
import json
import tempfile
import unittest
from unittest.mock import Mock, patch

import web_app
import workflow_support as support


class ReminderReviewTests(unittest.TestCase):
    def test_open_excel_targets_only_affected_plants_and_deduplicates_rows(self):
        with tempfile.TemporaryDirectory() as folder:
            paths = [Path(folder) / 'Plant 1.xlsx', Path(folder) / 'Plant 2.xlsx']
            for path in paths:
                path.write_bytes(b'test file; native opening is mocked')
            for plants in (['plant1'], ['plant2'], ['plant1', 'plant2']):
                with self.subTest(plants=plants), patch.object(support.os, 'startfile', create=True) as opener:
                    rows = [{'plant':plant, 'reminder_due':True} for plant in plants for _ in range(2)]
                    rows.append({'plant':'plant2' if plants == ['plant1'] else 'plant1', 'reminder_due':False})
                    result = {'rows':rows}
                    opened = support.open_alert_workbooks(result, [str(path) for path in paths])
                    expected = [paths[int(plant[-1]) - 1] for plant in plants]
                    self.assertEqual(opened, expected)
                    self.assertEqual([call.args[0] for call in opener.call_args_list], [str(path) for path in expected])
            self.assertEqual([path.read_bytes() for path in paths], [b'test file; native opening is mocked'] * 2)

    def test_unknown_alert_files_cannot_open_unconfigured_files_and_missing_files_fail_before_opening(self):
        with tempfile.TemporaryDirectory() as folder:
            first = Path(folder) / 'Plant 1.xlsx'; first.write_bytes(b'test')
            paths = [str(first), str(Path(folder) / 'Missing Plant 2.xlsx')]
            with patch.object(support.os, 'startfile', create=True) as opener:
                with self.assertRaisesRegex(FileNotFoundError, 'No plant Excel file'):
                    support.open_alert_workbooks({'rows':[{'system_alert':True,'file':str(Path(folder) / 'unrelated.xlsx')}]}, paths)
                with self.assertRaisesRegex(FileNotFoundError, 'Plant 2 Excel file'):
                    support.open_alert_workbooks({'rows':[{'plant':'plant1','system_alert':True}, {'plant':'plant2','system_alert':True}]}, paths)
                opener.assert_not_called()
            targets = support.alert_workbook_targets({'rows':[{'system_alert':True,'file':str(first)}]}, paths)
            self.assertEqual(targets, [('plant1',first)])

    def test_excel_button_label_matches_affected_plants_without_a_dropdown(self):
        for plants, label in [(['plant1'],'Open Plant 1 Excel'), (['plant2'],'Open Plant 2 Excel'),
                              (['plant1','plant2'],'Open both Excel files')]:
            with self.subTest(plants=plants):
                window = Mock()
                window.winfo_screenwidth.return_value = 1920; window.winfo_screenheight.return_value = 1080
                window.winfo_width.return_value = 560; window.winfo_height.return_value = 320
                open_excel = Mock()
                with patch.object(support.tk, 'Tk', return_value=window), patch.object(support.tk, 'Frame'), \
                        patch.object(support.tk, 'Label'), patch.object(support.tk, 'Button') as button, \
                        patch.object(support.tk, 'OptionMenu') as dropdown:
                    support.show_centered_reminder('Needs attention','Review now',
                                                   on_open_workbook=open_excel, workbook_plants=plants)
                excel_buttons = [call for call in button.call_args_list if call.kwargs['text'] == label]
                self.assertEqual(len(excel_buttons), 1)
                self.assertEqual(button.call_count, 4)
                excel_buttons[0].kwargs['command']()
                open_excel.assert_called_once()
                dropdown.assert_not_called()

    def test_new_time_rearms_same_day_but_same_time_still_suppresses_duplicates(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            today = date(2026, 10, 12)
            result = {'rows': [{'reminder_due': True, 'sn': '1', 'age_days': 25}]}
            notices = Mock()
            self.assertTrue(support.notify_once(root, result, notices, today=today, reminder_time='11:00'))
            self.assertFalse(support.notify_once(root, result, notices, today=today, reminder_time='11:00'))
            self.assertTrue(support.notify_once(root, result, notices, today=today, reminder_time='13:01'))
            self.assertFalse(support.notify_once(root, result, notices, today=today, reminder_time='13:01'))
            self.assertEqual(notices.call_count, 2)
            self.assertTrue(support.notify_once(root, result, notices, today=today,
                                              reminder_time='13:01', repeat_pending=True))

    def test_old_reminder_state_without_time_does_not_suppress_the_new_schedule(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            today = date(2026, 10, 12)
            row = {'reminder_due': True, 'sn': '1', 'plant': 'plant1', 'age_days': 25}
            support.save_json(root / 'Data/reminder-state.json', {
                'last_day': today.isoformat(), 'signature': 'plant1|1|None|25'})
            notices = Mock()
            self.assertTrue(support.notify_once(root, {'rows':[row]}, notices, today=today,
                                              reminder_time='13:01'))
            self.assertFalse(support.notify_once(root, {'rows':[row]}, notices, today=today,
                                               reminder_time='13:01'))

    def test_summary_has_counts_and_review_action_without_confusing_day_examples(self):
        text = support._notice_text([
            {'reminder_due': True, 'age_days': 100},
            {'system_alert': True, 'age_days': -8, 'supplier': 'Future company'},
        ])
        self.assertIn('1 certificate pending.', text)
        self.assertIn('1 record has details to check.', text)
        self.assertIn('Review now', text)
        self.assertNotIn('day -8', text)
        self.assertNotIn('Future company', text)

    def test_dismiss_for_today_suppresses_changed_and_snoozed_alerts_only_today(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            today = date(2026, 10, 12)
            result = {'rows': [{'reminder_due': True, 'sn': '1', 'age_days': 25}]}
            notices = Mock()
            with patch.object(support, 'cancel_snoozed_reminder') as cancel:
                support.dismiss_reminders_today(root, today=today)
                cancel.assert_called_once_with(root)
            for repeat in (False, True):
                self.assertFalse(support.notify_once(root, result, notices, today=today, repeat_pending=repeat))
            result['rows'].append({'system_alert': True, 'sn': '2'})
            self.assertFalse(support.notify_once(root, result, notices, today=today))
            notices.assert_not_called()
            self.assertTrue(support.notify_once(root, result, notices, today=today, reminder_time='13:01'))
            self.assertTrue(support.notify_once(root, result, notices, today=today + timedelta(days=1)))
            self.assertEqual(json.loads((root / 'Data/reminder-state.json').read_text())['dismissed_day'], today.isoformat())

    def test_review_launches_html_dashboard_with_attention_flag(self):
        with patch.object(support.subprocess, 'Popen') as start:
            support.launch_review_app(Path('C:/test-workspace'))
        args = start.call_args.args[0]
        self.assertEqual(Path(args[1]).name, 'web_app.py')
        self.assertEqual(args[2:], ['--attention'])

    def test_review_reuses_existing_dashboard_at_its_actual_port(self):
        response = Mock(status=200)
        response.getheader.return_value = 'application/json'
        connection = Mock()
        connection.getresponse.return_value = response
        with patch.object(web_app, 'HTTPConnection', return_value=connection), \
                patch.object(web_app, 'ensure_certificate_folders'), \
                patch.object(web_app.webbrowser, 'open') as open_browser, \
                patch.object(web_app, 'reuse_dashboard', return_value=False):
            web_app.main(attention=True)
        open_browser.assert_called_once_with('http://127.0.0.1:8765/?view=attention')

    def test_small_reminder_has_three_actions_and_no_excel_dropdown(self):
        window = Mock()
        window.winfo_screenwidth.return_value = 1920
        window.winfo_screenheight.return_value = 1080
        window.winfo_width.return_value = 560
        window.winfo_height.return_value = 265
        review, snooze, dismiss = Mock(), Mock(), Mock()
        with patch.object(support.tk, 'Tk', return_value=window), \
                patch.object(support.tk, 'Frame'), patch.object(support.tk, 'Label'), \
                patch.object(support.tk, 'Button') as button, patch.object(support.tk, 'OptionMenu') as dropdown:
            support.show_centered_reminder('Needs attention', 'Review now', on_open_automation=review,
                                           on_snooze=snooze, on_dismiss=dismiss)
        self.assertEqual([call.kwargs['text'] for call in button.call_args_list],
                         ['Review now', 'Remind me in 1 hour', 'Dismiss for today'])
        for call in button.call_args_list:
            call.kwargs['command']()
        review.assert_called_once(); snooze.assert_called_once(); dismiss.assert_called_once()
        dropdown.assert_not_called()
        window.grab_set.assert_not_called()


if __name__ == '__main__':
    unittest.main()
