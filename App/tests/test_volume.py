"""Full-size dummy runs: real PDFs and Excel writes, simulated Outlook only."""
from collections import Counter
from datetime import timedelta
import json
from pathlib import Path
import shutil
import tempfile
import time
import unittest
from unittest.mock import patch
from pypdf import PdfReader

from certificate_core import read_batches
from run_certificates import run
from workflow_support import notify_once, save_json
from test_workflow import FakeOutlook, ROOT, TODAY, make_book


class VolumeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.paths = [self.root / "Plant 1.xlsx", self.root / "Plant 2.xlsx"]
        save_json(self.root / "settings.example.json", {"workbook_paths": ["", ""], "initial_cod_confirmed": True})
        (self.root / "templates").mkdir()
        shutil.copy2(ROOT / "templates" / "Certificate of Recycling.png", self.root / "templates" / "Certificate of Recycling.png")
        FakeOutlook.saved = []
        self.fake = patch("run_certificates.OutlookWriter", FakeOutlook)
        self.fake.start()

    def tearDown(self):
        self.fake.stop()
        self.temp.cleanup()

    def execute(self):
        result = run(self.paths, "dummy@example.test", self.root, as_of=TODAY, progress=lambda _: None)
        return json.loads((result[2] / "Results.json").read_text())

    def test_100_ready_batches_use_one_sequence_and_repeat_without_duplicates(self):
        for path in self.paths:
            make_book(path, [{"SN": str(i + 1)} for i in range(50)])
        started = time.perf_counter()
        result = self.execute()
        elapsed = time.perf_counter() - started
        self.assertEqual((result["created"], result["review"]), (100, 0), result)
        self.assertTrue(result["scan_complete"])
        self.assertEqual(len(FakeOutlook.saved), 100)
        self.assertEqual(len(list(Path(result["output"]).glob("*.pdf"))), 100)
        batches = [batch for path in self.paths for batch in read_batches(path)]
        self.assertEqual(sorted(int(b["cod_no"]) for b in batches), list(range(1, 101)))
        self.assertTrue(all(b["cod_date"].date() == TODAY for b in batches))
        self.assertTrue(all(b["total_weight"] == 7 for b in batches))
        indexed = {(f"plant{index}", b["case_id"]): b for index, path in enumerate(self.paths, 1) for b in read_batches(path)}
        for index in (0, 49, 50, 99):
            plant, sn, subject, filename = FakeOutlook.saved[index]
            batch = indexed[plant, sn]
            pdf = PdfReader(filename)
            self.assertEqual(len(pdf.pages), 1)
            self.assertEqual(tuple(float(v) for v in pdf.pages[0].mediabox[2:]), (960, 540))
            overlay = pdf.pages[0].extract_text()
            for field in (batch["cod_no"], batch["reference"], batch["form6"], "7.0 KG", batch["supplier"].upper()):
                self.assertIn(field, overlay)
        self.assertEqual(len(list((self.root / "Data" / "Workbook Backups").glob("*.xlsx"))), 2)
        repeat = self.execute()
        self.assertEqual((repeat["created"], repeat["review"]), (0, 0), repeat)
        self.assertEqual(len(FakeOutlook.saved), 100)
        self.assertEqual(Counter(r["outcome"] for r in repeat["rows"]), {"already_prepared": 100})
        self.assertTrue(all(r["certificate_status"] == "drafted" and r.get("awaiting_send")
                            for r in repeat["rows"]))
        print(f"\n100-dummy-certificate run: {elapsed:.2f}s; real PDF generation and Excel writes; Outlook simulated.", flush=True)

    def test_100_mixed_batches_have_complete_results_and_one_quiet_alert(self):
        variants = ({}, {"Certificate Issued": ""}, {"Certificate Issued": "N/A", "Client Email": None}, {"Reference No.": None}, {"Inward Date": TODAY - timedelta(days=29)})
        for path in self.paths:
            make_book(path, [dict(variants[i % 5], SN=str(i + 1)) for i in range(50)])
        result = self.execute()
        self.assertEqual((result["created"], result["review"]), (40, 40), result)
        self.assertEqual(len(result["rows"]), 100)
        self.assertEqual(Counter(r["outcome"] for r in result["rows"]), {"draft_created": 40, "needs_review": 40, "not_applicable": 20})
        messages = []
        notifier = lambda title, body: messages.append(body) or True
        notify_once(self.root, result, notifier, TODAY, repeat_pending=True)
        self.assertEqual(len(messages), 1)
        self.assertIn("60 inward records are 30+ days old", messages[0])
        self.assertNotIn("draft", messages[0].casefold())
        repeat = self.execute()
        notify_once(self.root, repeat, notifier, TODAY + timedelta(days=1), repeat_pending=True)
        self.assertEqual(len(FakeOutlook.saved), 40)
        self.assertEqual(len(messages), 2)
        self.assertIn("60 inward records are 30+ days old", messages[1])


if __name__ == "__main__":
    unittest.main()
