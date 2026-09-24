"""Regression tests for refreshing the operator's review list after reminders."""
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from desktop_app import App
import scheduled_run


class ReviewRefreshTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.latest = self.base / "Output" / "Latest Results.json"
        self.latest.parent.mkdir()
        self.app = object.__new__(App)
        self.app.base = self.base
        self.app._latest_results_signature = None
        self.app._next_latest_results_check = 0.0
        self.app.show_results = Mock()

    def write_latest(self, value):
        self.latest.write_text(value, encoding="utf-8")
        time.sleep(0.002)

    def test_changed_latest_results_refresh_the_open_review_list(self):
        self.write_latest('{"rows":[]}')

        self.assertTrue(self.app.refresh_latest_results_if_changed())
        self.app.show_results.assert_called_once_with()
        self.assertEqual(self.app._latest_results_signature,
                         (self.latest.stat().st_mtime_ns, self.latest.stat().st_size))

    def test_unchanged_results_do_not_redraw_the_review_list(self):
        self.write_latest('{"rows":[]}')
        signature = (self.latest.stat().st_mtime_ns, self.latest.stat().st_size)
        self.app._latest_results_signature = signature

        self.assertFalse(self.app.refresh_latest_results_if_changed())
        self.app.show_results.assert_not_called()

    def test_result_poll_is_rate_limited_to_once_per_second(self):
        self.write_latest('{"rows":[]}')
        with patch("desktop_app.time.monotonic", return_value=10.0):
            self.assertTrue(self.app.refresh_latest_results_if_changed())
            self.assertFalse(self.app.refresh_latest_results_if_changed())
        self.app.show_results.assert_called_once_with()

    def test_failed_snooze_run_clears_the_consumed_one_shot(self):
        with patch("workflow_support.load_settings", side_effect=ValueError("broken settings")), \
             patch("workflow_support.cancel_snoozed_reminder") as cancel, \
             patch("scheduled_run.diagnostic") as diagnostic:
            self.assertEqual(scheduled_run.main(self.base, reminder_only=True), 1)
        cancel.assert_called_once_with(self.base, force=True)
        diagnostic.assert_called_once()
        self.assertIn("broken settings", diagnostic.call_args.args[1])


if __name__ == "__main__":
    unittest.main()
