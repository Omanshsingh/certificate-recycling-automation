"""Find matching certificates without opening every PDF in the folder."""
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from reportlab.pdfgen.canvas import Canvas

import run_certificates as runner
from issuance_safety import valid_pdf


class CertificateLookupTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.folder = self.root / 'Certificates' / 'Auto Generated'
        self.folder.mkdir(parents=True)
        self.batch = dict(cod_no='0193', source_id='plant1', case_id='5',
                          supplier='Client', location='Mumbai')

    def pdf(self, name):
        path = self.folder / name
        canvas = Canvas(str(path))
        canvas.drawString(20, 20, 'Certificate')
        canvas.save()
        return path

    def test_only_matching_files_are_opened_and_validated(self):
        expected = self.pdf('0193_Client_Mumbai.pdf')
        for number in range(200, 225):
            self.pdf(f'{number:04d}_Other_Mumbai.pdf')
        with patch('issuance_safety.valid_pdf', wraps=valid_pdf) as validate:
            self.assertEqual(runner.find_certificate(self.root / 'Certificates', self.batch), expected)
        self.assertEqual([call.args[0] for call in validate.call_args_list], [expected])

    def test_no_match_opens_no_pdf(self):
        self.pdf('0200_Other_Mumbai.pdf')
        with patch('issuance_safety.valid_pdf', wraps=valid_pdf) as validate:
            self.assertIsNone(runner.find_certificate(self.root / 'Certificates', self.batch))
        validate.assert_not_called()

    def test_duplicate_valid_candidates_remain_ambiguous_and_corrupt_matches_are_rejected(self):
        expected = self.pdf('0193_Client_Mumbai.pdf')
        duplicate = self.pdf('0193_Legacy_Mumbai_1.pdf')
        self.assertIsNone(runner.find_certificate(self.root / 'Certificates', self.batch))
        duplicate.write_bytes(b'not a PDF')
        self.assertEqual(runner.find_certificate(self.root / 'Certificates', self.batch), expected)
        expected.write_bytes(b'not a PDF')
        self.assertIsNone(runner.find_certificate(self.root / 'Certificates', self.batch))

    def test_legacy_names_and_leading_zero_numbers_still_resolve(self):
        for filename in ['0193_Legacy_Mumbai_1.pdf', '0193.pdf', '193.pdf', 'COD_0193.pdf', 'certificate_0193.pdf']:
            with self.subTest(filename=filename):
                path = self.pdf(filename)
                self.assertEqual(runner.find_certificate(self.root / 'Certificates', self.batch), path)
                path.unlink()
