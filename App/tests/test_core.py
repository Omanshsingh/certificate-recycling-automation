import tempfile
import unittest
from datetime import date
from pathlib import Path

from certificate_core import certificate_reminder_due, effective_certificate_status, eligible, read_batches
from draft_service import draft_html
from run_certificates import run

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "sample" / "Dummy_Inward_5_Clients.xlsx"


class CertificateTests(unittest.TestCase):
    def test_reads_five_batches_and_calculates_formula_weight(self):
        batches = read_batches(SAMPLE)
        self.assertEqual([b["case_id"] for b in batches], ["2001", "2002", "2003", "2004", "2005"])
        self.assertAlmostEqual(batches[0]["total_weight"], 135.64, places=2)
        self.assertEqual(batches[0]["email"], "plant2.client1@example.test")
        self.assertFalse(batches[0]["issues"])

    def test_eligibility_is_thirty_days(self):
        batches = read_batches(SAMPLE)
        self.assertTrue(eligible(batches[0], date(2026, 6, 1), 30))
        self.assertFalse(eligible(batches[2], date(2026, 6, 1), 30))

    def test_alert_cutoff_is_exactly_30_days_and_excludes_completed_or_na(self):
        today = date(2026, 9, 23)
        for status in (None, "", "Blank", "Ready", "Drafted"):
            with self.subTest(status=status):
                batch = {"case_id": "100", "inward_date": date(2026, 8, 24), "issued": status}
                self.assertTrue(certificate_reminder_due(batch, today))
        for status in ("N/A", "Completed", "Sent"):
            with self.subTest(status=status):
                batch = {"case_id": "100", "inward_date": date(2026, 8, 24), "issued": status}
                self.assertFalse(certificate_reminder_due(batch, today))
        self.assertFalse(certificate_reminder_due(
            {"case_id": "100", "inward_date": date(2026, 8, 25), "issued": "Blank"}, today))
        self.assertFalse(certificate_reminder_due(
            {"case_id": "100", "inward_date": None, "issued": "Blank"}, today))
        self.assertFalse(certificate_reminder_due(
            {"case_id": "", "inward_date": date(2026, 8, 1), "issued": "Blank"}, today))

    def test_legacy_completed_is_still_drafted_until_history_confirms_sent(self):
        batch = {"issued": "Completed"}
        record = {"status": "saved", "certificate_no": "0042"}
        self.assertEqual(effective_certificate_status(batch, record), "drafted")
        self.assertEqual(effective_certificate_status(
            batch, record, {"certificate_no": "0042"}), "completed")
        self.assertEqual(effective_certificate_status({"issued": "Sent"}), "completed")

    def test_certificate_email_uses_salesforce_format_and_banner(self):
        html = draft_html({
            "supplier": "Dummy & Sons",
            "location": "Mumbai",
            "items": [{"item": "Laptop & monitor", "quantity": 2, "weight": 12.5}],
            "total_weight": 12.5,
        })

        self.assertIn('font-family:"Calibri",sans-serif', html)
        self.assertIn('font-size:11.0pt', html)
        self.assertIn('width="540"', html)
        self.assertIn("table-layout:fixed", html)
        self.assertIn("border:1px solid #000", html)
        self.assertIn("padding:4pt 8pt", html)
        self.assertIn("DUMMY &amp; SONS, MUMBAI", html)
        self.assertIn("Laptop &amp; monitor", html)
        self.assertIn("Kindly review the attached certificate", html)
        self.assertIn('width="530" height="95" src="cid:ecoreco-banner"', html)
        self.assertNotIn("Arial,sans-serif;font-size:11.0pt", html)

if __name__ == "__main__":
    unittest.main()
