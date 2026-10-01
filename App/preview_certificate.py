"""Create a local dummy layout preview without reserving a certificate number."""
from datetime import date
from certificate_core import make_certificate, read_batches
from app_paths import DATA_ROOT, PROGRAM_ROOT

ROOT = DATA_ROOT


if __name__ == "__main__":
    batch = read_batches(PROGRAM_ROOT / "sample" / "Dummy_Inward_10_Clients.xlsx")[0]
    batch["supplier"] = "Dummy Advanced Electronics and Resource Recovery Solutions Private Limited"
    batch["location"] = "Mumbai"
    destination = ROOT / "Data" / "Reports" / "Vendor alignment preview.pdf"
    make_certificate(batch, "PREVIEW", date.today(), PROGRAM_ROOT / "templates" / "Certificate of Recycling.png", destination)
    print(destination)
