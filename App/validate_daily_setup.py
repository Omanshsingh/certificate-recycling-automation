"""Confirm the two saved plant workbooks before enabling reminders."""
from pathlib import Path

from workflow_support import load_settings


def validate_daily_setup(root):
    root = Path(root)
    settings_path = root / "settings.local.json"
    if not settings_path.is_file():
        raise ValueError("Open Prepare Certificates.bat and save both plant workbook locations first.")
    paths = load_settings(root)["workbook_paths"]
    if not all(paths):
        raise ValueError("Save both plant Excel file locations in the app before enabling reminders.")
    files = tuple(Path(item).expanduser().resolve() for item in paths)
    if not all(item.is_file() for item in files):
        raise ValueError("A saved plant workbook cannot be found. Choose its current location in the app.")
    if files[0] == files[1] or files[0].samefile(files[1]):
        raise ValueError("Choose two different Excel files, one for each plant.")
    return files


if __name__ == "__main__":
    try:
        validate_daily_setup(Path.cwd())
    except (OSError, ValueError, KeyError) as exc:
        print(str(exc))
        raise SystemExit(1)
    print("Reminder setup is ready: both plant workbooks are saved.")
