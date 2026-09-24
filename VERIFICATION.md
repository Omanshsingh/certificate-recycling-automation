# Verification — 24 September 2026

## Current behavior

- Ready and complete: eligible on the next draft run, including the receipt day. No 30-day waiting period. Future receipt dates require correction.
- Blank Certificate Issued: remains in Needs attention until the operator chooses Ready or N/A. Blank does not prove recycling is complete or an email was never sent.
- N/A: skipped, with no COD allocation.
- One local installation coordinates numbering across both selected plant workbooks. COD/date write-back follows a confirmed Outlook save, with backups and interruption recovery.
- The work list filters pending decisions, Ready rows, prepared drafts, manual sent confirmations and skipped rows. Open selected plant identifies the sheet, row and SN to edit.
- Mark sent records the operator's confirmation locally with a UTC timestamp. It never sends, verifies delivery, or changes the plant workbook. Undo is supported.
- Installed daily task: one silent summary per day while unresolved issues remain, with a mute control. It does not show a popup for every client. Without the task, use Check both files manually.

## Live Outlook verification

The opt-in `tests/manual_outlook_smoke.py` ran through the normal plant processing function with two isolated temporary workbooks, an independent TEST-prefixed COD register and a non-deliverable example.test recipient. It verified:

- A Ready row received today produced exactly one unsent draft.
- Its visible PDF attachment was byte-for-byte identical to the generated certificate.
- The signature banner had the expected inline content identifier.
- The plant COD was updated and a repeat run did not create another draft.
- The uniquely identified test draft was moved to Outlook Deleted Items afterward. No email was sent. This is recoverable deletion, not permanent mailbox cleanup.

No live plant workbook or live COD sequence was used by that test. Attachment read-back is not a claim of pixel-perfect Outlook rendering.

## Automated checks

Final local run: **179 tests passed**, no skips, in 102.033 seconds. The 100-certificate volume run completed in 54.07 seconds with real PDF generation and Excel updates and a simulated Outlook writer. An export of the staged release also passed its three core/template/sample tests from its own directory, without local settings or history.

Run `python -m unittest discover -s tests -v` after installing requirements-dev.txt. Tests cover two-plant numbering, same-day Ready eligibility, future dates, missing fields, wrong statuses, protected/locked files, formatting preservation, duplicate prevention, interrupted saves, workbook recovery, hidden/missing attachments, UI filters, sent-confirmation cancellation/undo, damaged history, quiet alerts, mute, partial scans and volume runs.

## Deployment limits

Keep Data, settings.local.json and outputs when updating. Install/run under the Outlook user's Windows account. Code updates do not install or update the Windows scheduled task; run Create Daily Draft Task.bat separately after confirming the selected files and mailbox.

The operator's actual network files, official COD sequence and Office/security policies still need a one-case acceptance check. Two independent installations cannot safely share numbering through separate local registers. Uncertain saves require review rather than automatic retry. This version makes recycling certificates only and never sends mail automatically. Marked sent is a manual record, not mailbox delivery tracking.
