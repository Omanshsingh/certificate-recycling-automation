"""In-place reminder review, closed-tab fallback and request dispatch."""
from io import BytesIO
import json
import unittest
from unittest.mock import Mock, patch

from dashboard_review import DashboardReview
import web_app


class DashboardReviewTests(unittest.TestCase):
    def test_review_goes_to_one_visible_tab_and_is_acknowledged(self):
        coordinator = DashboardReview()
        coordinator.presence('visible', True)
        coordinator.presence('hidden', False)
        request = coordinator.request()
        self.assertTrue(request['reused'])
        self.assertIsNone(coordinator.presence('hidden', False))
        self.assertEqual(coordinator.presence('visible', True)['id'], request['request_id'])
        self.assertFalse(coordinator.status(request['request_id'])['acknowledged'])
        coordinator.presence('visible', True, acknowledged=request['request_id'])
        self.assertTrue(coordinator.status(request['request_id'])['acknowledged'])
        self.assertIsNone(coordinator.presence('visible', True))

    def test_closed_tabs_expire_and_cancel_prevents_a_delayed_tab_from_handling_fallback(self):
        coordinator = DashboardReview()
        with patch('dashboard_review.time.monotonic', return_value=0):
            coordinator.presence('tab', True)
            request = coordinator.request()
        coordinator.status(request['request_id'], cancel=True)
        self.assertIsNone(coordinator.presence('tab', True))
        with patch('dashboard_review.time.monotonic', return_value=10**12):
            self.assertFalse(coordinator.request()['reused'])

    def test_cancel_racing_with_receipt_does_not_open_a_duplicate_page(self):
        coordinator = DashboardReview()
        coordinator.presence('tab')
        request = coordinator.request()
        coordinator.presence('tab', acknowledged=request['request_id'])
        self.assertTrue(coordinator.status(request['request_id'], cancel=True)['acknowledged'])

    def test_review_and_presence_do_not_wait_for_certificate_publication(self):
        coordinator = DashboardReview()
        coordinator.presence('tab')
        handler = object.__new__(web_app.Handler)
        handler.path = '/api/dashboard-review'
        handler.headers = {'Host':'127.0.0.1','Content-Length':'2'}
        handler.rfile = BytesIO(b'{}')
        handler.send_json = Mock()
        with patch.object(web_app, 'DASHBOARD_REVIEW', coordinator), \
                patch.object(web_app, 'STOPPING', False), patch.object(web_app, 'OPERATION_LOCK') as lock:
            handler.do_POST()
        lock.__enter__.assert_not_called()
        self.assertTrue(handler.send_json.call_args.args[0]['reused'])
        self.assertFalse(handler.send_json.call_args.args[0]['changed'])

    def test_startup_does_not_open_a_new_page_when_existing_tab_receives_review(self):
        connection = Mock()
        response = Mock(status=200)
        response.getheader.return_value = 'application/json'
        response.read.return_value = b'{"application":"ecoreco-certificates","changed":false}'
        connection.getresponse.return_value = response
        with patch.object(web_app, 'HTTPConnection', return_value=connection), \
                patch.object(web_app, 'ensure_certificate_folders'), \
                patch.object(web_app, 'reuse_dashboard', return_value=True) as reuse, \
                patch.object(web_app.webbrowser, 'open') as open_browser:
            web_app.main(attention=True)
        reuse.assert_called_once_with(8765)
        open_browser.assert_not_called()

    def test_reuse_waits_for_receipt_and_focus_does_not_launch_a_browser(self):
        connection = Mock()
        request = Mock(status=200); request.read.return_value = json.dumps({'reused':True,'request_id':'review'}).encode()
        receipt = Mock(status=200); receipt.read.return_value = b'{"acknowledged":true}'
        connection.getresponse.side_effect = [request,receipt]
        with patch.object(web_app, 'HTTPConnection', return_value=connection), \
                patch.object(web_app, 'focus_dashboard_window') as focus:
            self.assertTrue(web_app.reuse_dashboard(8765))
        focus.assert_called_once()
        self.assertEqual(json.loads(connection.request.call_args.args[2]), {'request_id':'review'})


if __name__ == '__main__':
    unittest.main()
