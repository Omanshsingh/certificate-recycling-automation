from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from draft_service import draft_html


root = Path(__file__).resolve().parent
out = root / "Output" / "preview"
out.mkdir(parents=True, exist_ok=True)
pdf_path = root / "Output" / "2026-09-21" / "Certificate_1002_0002.pdf"
if not pdf_path.exists():
    pdf_path = root / "Output" / "2026-06-01" / "Certificate_1001_0001.pdf"
banner_path = root / "templates" / "signature-banner.png"

batch = {
    "supplier": "Dummy Client 2 Pvt. Ltd.",
    "location": "Pune",
    "email": "dummy.client2@example.test",
    "items": [
        {"item": "Laptop computers", "quantity": 4, "weight": 18.6},
        {"item": "Mobile phones", "quantity": 8, "weight": 1.2},
    ],
    "total_weight": 19.8,
}

message = MIMEMultipart("mixed")
message["From"] = "crm@example.test"
message["To"] = batch["email"]
message["Subject"] = "Certificate of E-Waste Recycling - Dummy Client 2 - 0002"

related = MIMEMultipart("related")
alternative = MIMEMultipart("alternative")
alternative.attach(MIMEText("Please find attached your Certificate of E-Waste Recycling.", "plain", "utf-8"))
alternative.attach(MIMEText(draft_html(batch), "html", "utf-8"))
related.attach(alternative)

with banner_path.open("rb") as stream:
    image = MIMEImage(stream.read(), _subtype="png")
image.add_header("Content-ID", "<ecoreco-banner>")
image.add_header("Content-Disposition", "inline", filename="signature-banner.png")
related.attach(image)
message.attach(related)

with pdf_path.open("rb") as stream:
    attachment = MIMEApplication(stream.read(), _subtype="pdf")
attachment.add_header("Content-Disposition", "attachment", filename="Certificate_Dummy_Client_2_0002.pdf")
message.attach(attachment)

target = out / "Certificate_Draft_Preview.eml"
target.write_bytes(message.as_bytes())
print(target)
