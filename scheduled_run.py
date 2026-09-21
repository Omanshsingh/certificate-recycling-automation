"""Run the configured draft-only job from Windows Task Scheduler."""
from pathlib import Path
import json
from run_certificates import run

ROOT = Path(__file__).resolve().parent
config = ROOT / "settings.local.json"
if not config.exists(): config = ROOT / "settings.example.json"
settings = json.loads(config.read_text(encoding="utf-8"))
workbook = settings.get("workbook_path", "")
mailbox = settings.get("outlook_mailbox", "")
if not workbook or not Path(workbook).exists():
    raise SystemExit("Set the workbook path in the desktop app before scheduling the daily job.")
if not mailbox:
    raise SystemExit("Select the Outlook mailbox in the desktop app before scheduling the daily job.")
run(workbook, mailbox, ROOT, create_drafts=True)
