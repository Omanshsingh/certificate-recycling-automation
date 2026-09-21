import tempfile
import unittest
from datetime import date
from pathlib import Path

from certificate_core import eligible, read_batches
from run_certificates import run

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "sample" / "Dummy_Inward_5_Clients.xlsx"


class CertificateTests(unittest.TestCase):
    def test_reads_five_batches_and_calculates_formula_weight(self):
        batches = read_batches(SAMPLE)
        self.assertEqual([b["case_id"] for b in batches], ["1001", "1002", "1003", "1004", "1005"])
        self.assertAlmostEqual(batches[0]["total_weight"], 135.64, places=2)
        self.assertEqual(batches[0]["email"], "client1@example.test")
        self.assertFalse(batches[0]["issues"])

    def test_eligibility_is_thirty_days(self):
        batches = read_batches(SAMPLE)
        self.assertTrue(eligible(batches[0], date(2026, 6, 1), 30))
        self.assertFalse(eligible(batches[2], date(2026, 6, 1), 30))

    def test_pdf_only_run_creates_one_and_does_not_duplicate(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            for name in ("settings.example.json",):
                (root / name).write_text((ROOT / name).read_text(), encoding="utf-8")
            (root / "templates").mkdir()
            (root / "templates" / "Certificate of Recycling.png").write_bytes((ROOT / "templates" / "Certificate of Recycling.png").read_bytes())
            first = run(SAMPLE, "", root, limit=1, create_drafts=False, as_of=date(2026, 6, 1))
            self.assertEqual(first[0], 1)
            second = run(SAMPLE, "", root, limit=1, create_drafts=False, as_of=date(2026, 6, 1))
            self.assertEqual(second[0], 1)
            self.assertEqual(len(list((root / "Output" / "2026-06-01").glob("*.pdf"))), 2)


if __name__ == "__main__":
    unittest.main()
