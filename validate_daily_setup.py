"""Validate the saved plant workbooks and mailbox before installing a task."""
from __future__ import annotations

from pathlib import Path

from workflow_support import load_settings


def validate_daily_setup(root: Path) -> tuple[Path, Path]:
    root = Path(root)
    if not (root / "settings.local.json").is_file():
        raise ValueError("Open Prepare Certificates.bat and save both plant locations and the Outlook mailbox first.")

    settings = load_settings(root)
    paths = settings["workbook_paths"]
    if len(paths) != 2 or not all(paths):
        raise ValueError("Save both plant Excel locations in the app before creating the daily task.")

    files = tuple(Path(path).resolve() for path in paths)
    if not all(path.is_file() for path in files):
        raise ValueError("A saved plant workbook is unavailable. Reconnect its drive or select its new location in the app.")
    if files[0] == files[1] or files[0].samefile(files[1]):
        raise ValueError("Choose a different Excel file for each plant.")
    if not settings.get("outlook_mailbox", "").strip():
        raise ValueError("Load Outlook accounts and select the mailbox in the app first.")
    return files


def main() -> int:
    try:
        validate_daily_setup(Path.cwd())
    except (OSError, ValueError, KeyError) as exc:
        print(str(exc))
        return 1
    print("Preflight passed: two distinct plant workbooks and an Outlook mailbox are saved.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
