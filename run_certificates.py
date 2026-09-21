"""Command-line runner used by the desktop app and optional daily task."""
from __future__ import annotations

from datetime import date
from datetime import timedelta
import json
from pathlib import Path

from certificate_core import certificate_number, eligible, make_certificate, read_batches
from draft_service import History, OutlookWriter, draft_html


def run(workbook_path, mailbox, root, limit=None, create_drafts=True, as_of=None, progress=print):
    root = Path(root)
    config = root / "settings.local.json"
    if not config.exists(): config = root / "settings.example.json"
    settings = json.loads(config.read_text(encoding="utf-8"))
    as_of = as_of or date.today()
    batches = read_batches(workbook_path)
    output = root / "Output" / as_of.isoformat()
    output.mkdir(parents=True, exist_ok=True)
    history = History(root / "Data" / "certificate-history.sqlite")
    writer = OutlookWriter(mailbox) if create_drafts else None
    created = 0
    review = 0
    try:
        for batch in batches:
            if limit is not None and created >= limit:
                break
            if not eligible(batch, as_of, int(settings.get("eligibility_days", 30))):
                progress(f"{batch['case_id']}: waiting until {batch['inward_date']} + {settings.get('eligibility_days', 30)} days")
                continue
            if batch["issues"]:
                review += 1
                progress(f"{batch['case_id']}: review required - {', '.join(batch['issues'])}")
                continue
            fingerprint = history.fingerprint(batch)
            existing = history.inspect(batch["case_id"], fingerprint)
            if existing:
                progress(f"{batch['case_id']}: {existing}")
                continue
            sequence = history.next_sequence(int(settings.get("certificate_number_start", 1)))
            cert_no = certificate_number(sequence, settings)
            if history.reserve(batch, cert_no):
                continue
            pdf_path = output / f"Certificate_{batch['case_id']}_{cert_no}.pdf"
            certificate_date = batch["inward_date"] + timedelta(days=int(settings.get("eligibility_days", 30)))
            make_certificate(batch, cert_no, certificate_date, root / "templates" / "Certificate of Recycling.png", pdf_path)
            if writer:
                subject = f"Certificate of E-Waste Recycling - {batch['supplier']} - {cert_no}"
                entry = writer.save_draft(batch, subject, draft_html(batch), pdf_path)
            else:
                entry = "PDF_ONLY"
            history.complete(batch, pdf_path, entry)
            created += 1
            progress(f"{batch['case_id']}: certificate created and {'Outlook draft saved' if writer else 'PDF saved'}")
        (output / "Results.txt").write_text(f"Created: {created}\nReview required: {review}\n", encoding="utf-8")
        return created, review, output
    finally:
        history.close()
