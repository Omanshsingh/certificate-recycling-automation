# Verification — 24 September 2026

## Current behavior

- Ready and complete: eligible on the next draft run, including the receipt day. No 30-day waiting period. Future receipt dates require correction.
- Blank Certificate Issued: remains in Needs attention until the operator chooses Ready or N/A. Blank does not prove recycling is complete or an email was never sent.
- N/A: skipped, with no COD allocation.
- One local installation coordinates numbering across both selected plant workbooks. COD/date write-back follows a confirmed Outlook save, with backups and interruption recovery.
- The work list filters pending decisions, Ready rows, prepared drafts, manual sent confirmations and skipped rows. Open selected plant identifies the sheet, row and SN to edit.
- Mark sent updates only the selected batch's Certificate Issued cell to Sent, makes a workbook backup, and records operator confirmation locally. It never sends or verifies delivery. Undo returns the status to Completed without unsending the email.
- Installed daily task: each eligible run can show one centered summary for inward records at least 30 days old whose status is not Sent or N/A, plus a grouped count of missing/invalid/future Inward Dates and automation problems that reach the notification stage. Separate later runs, including a one-hour snooze, may remind again the same day; a duplicate popup for the same run is blocked. Blank, Ready, and Completed aged rows qualify; ordinary blank fields and status decisions remain visible in the results but do not create per-row popups. Invalid dates are separately summarized because they cannot be age-classified. Failures before the runner/dependencies load go to a diagnostic file instead of a popup. Remind me later snoozes one hour and performs a read-only refresh. Yes/open review refreshes the already-open app after its latest results change. Default time is 09:00 local; Settings can change and apply it. Saving the reminder time clears an earlier same-day dismissal. This also changes when the scheduled check prepares Ready drafts. A sign-in catch-up may prepare drafts earlier, but never displays a reminder before the configured time. Without the task, use Check both files manually.

## Live Outlook verification

The opt-in `tests/manual_outlook_smoke.py` ran through the normal plant processing function with two isolated temporary workbooks, an independent TEST-prefixed COD register and a non-deliverable example.test recipient. It verified:

- A Ready row received today produced exactly one unsent draft.
- Its visible PDF attachment was byte-for-byte identical to the generated certificate.
- The signature banner had the expected inline content identifier.
- The plant COD was updated and a repeat run did not create another draft.
- The uniquely identified test draft was moved to Outlook Deleted Items afterward. No email was sent. This is recoverable deletion, not permanent mailbox cleanup.

No live plant workbook or live COD sequence was used by that test. Attachment read-back is not a claim of pixel-perfect Outlook rendering.

## Automated checks

Final local run: **224 tests passed**, no skips. The suite includes reminder age/status boundaries, missing/invalid/future date handling, daily aggregation, partial and resolved scans, snooze recovery, review-list refresh, workbook status updates and undo, and the certificate workflow. Tests use simulated Outlook and do not send email or create live drafts.

Run `python -m unittest discover -s tests -v` after installing requirements-dev.txt. Tests cover two-plant numbering, same-day Ready eligibility, future dates, missing fields, wrong statuses, protected/locked files, formatting preservation, duplicate prevention, interrupted saves, workbook recovery, hidden/missing attachments, UI filters, sent-confirmation cancellation/undo, damaged history, 30-day alert boundaries, missing-date digest and migration, mute/snooze, partial scans, one hundred mixed cases, and supported Windows task-reschedule parameters.

## Deployment limits

Keep Data, settings.local.json and Plant Workbooks when updating. Output contains generated certificates and reports. Install/run under the Outlook user's Windows account. Code updates do not install or update the Windows scheduled task; run Setup/Create Daily Draft Task.bat separately after confirming the selected files and mailbox.

The operator's actual network files, official COD sequence and Office/security policies still need a one-case acceptance check. Two independent installations cannot safely share numbering through separate local registers. Uncertain saves require review rather than automatic retry. This version makes recycling certificates only and never sends mail automatically. Marked sent is a manual record, not mailbox delivery tracking.

The existing Windows task was rescheduled to the saved **11:54** local time and its sign-in catch-up was preserved. Its next planned daily run is **25 September 2026 at 11:54**. The previous run's result remains `0x800710E0`; Task Scheduler's Operational log is disabled, so the underlying cause is unknown. The task was not manually started during this review, so no live Outlook draft was created. Plant 1 had an Excel lock at review time; neither plant workbook was modified.
