# Verification — 26 September 2026

## Current behavior

- Ready and complete: eligible on the next draft run, including the receipt day. No 30-day waiting period. Future receipt dates require correction.
- Blank Certificate Issued: remains in Needs attention until the operator chooses Ready or N/A. Blank does not prove recycling is complete or an email was never sent.
- N/A: skipped, with no COD allocation.
- One local installation coordinates numbering across both selected plant workbooks. COD/date write-back follows a confirmed Outlook save, with backups and interruption recovery.
- Draft creation writes **Drafted** after Outlook confirms the unsent draft. After the person sends it, **Mark completed / Undo** writes **Completed** and records that confirmation. The app never sends mail or verifies delivery. Undo restores Drafted and does not unsend the email.
- The work list filters pending decisions, Ready rows, Drafted certificates, Completed certificates and skipped rows. Open selected plant identifies the sheet, row and SN to edit.
- A Drafted but unsent certificate can be explicitly corrected from the app. The same Outlook draft is updated with the corrected subject/PDF, the workbook and shared history are updated together, and the earlier PDF plus reason/old values are retained. Direct Excel edits alone do not alter the certificate. Completed certificates are protected and require the company's separate replacement process.
- Installed daily task: each eligible run can show one centered summary for inward records at least 30 days old whose status is not Completed or N/A, plus a grouped count of missing/invalid/future Inward Dates and automation problems that reach the notification stage. Separate later runs, including a one-hour snooze, may remind again the same day; a duplicate popup for the same run is blocked. Blank, Ready, and Drafted aged rows qualify; Completed rows do not. Ordinary blank fields and status decisions remain visible in the results but do not create per-row popups. Invalid dates are separately summarized because they cannot be age-classified. Failures before the runner/dependencies load go to a diagnostic file instead of a popup. Remind me later snoozes one hour and performs a read-only refresh. Yes/open review refreshes the already-open app after its latest results change. Default time is 09:00 local; Settings can change and apply it. Saving the reminder time clears an earlier same-day dismissal. This also changes when the scheduled check prepares Ready drafts. A sign-in catch-up may prepare drafts earlier, but never displays a reminder before the configured time. Without the task, use Check both files manually.
- Legacy handling: old Sent rows display as Completed. Old Completed rows without a matching completion confirmation in history are treated as Drafted and remain overdue; this prevents an old unsent draft from being incorrectly cleared from the reminder list.

## Live Outlook verification

The opt-in `tests/manual_outlook_smoke.py` ran through the normal plant processing function with two isolated temporary workbooks, an independent TEST-prefixed COD register and a non-deliverable example.test recipient. It verified:

- A Ready row received today produced exactly one unsent draft.
- Its visible PDF attachment was byte-for-byte identical to the generated certificate.
- The signature banner had the expected inline content identifier.
- The plant COD was updated and a repeat run did not create another draft.
- The uniquely identified test draft was moved to Outlook Deleted Items afterward. No email was sent. This is recoverable deletion, not permanent mailbox cleanup.

No live plant workbook or live COD sequence was used by that test. Attachment read-back is not a claim of pixel-perfect Outlook rendering.

## Automated checks

The current local regression suite is run with `python -m unittest discover -s tests -v`. It covers reminder age/status boundaries, old-status migration, data validation and color rules, missing/invalid/future date handling, daily aggregation, partial and resolved scans, snooze recovery, workbook status updates and undo, and the certificate workflow. Correction tests cover mismatched Excel COD/date values, duplicate and previously used numbers, Completed protection, interrupted recovery, and updating the existing Outlook draft. Tests use simulated Outlook; no email is sent.

Run `python -m unittest discover -s tests -v` after installing requirements-dev.txt. Tests cover two-plant numbering, same-day Ready eligibility, future dates, missing fields, wrong statuses, protected/locked files, formatting preservation, duplicate prevention, interrupted saves, workbook recovery, hidden/missing attachments, UI filters, sent-confirmation cancellation/undo, audited certificate correction and recovery, damaged history, 30-day alert boundaries, missing-date digest and migration, mute/snooze, partial scans, one hundred mixed cases, and supported Windows task-reschedule parameters.

## Deployment limits

Keep Data, settings.local.json and Plant Workbooks when updating. Output contains generated certificates and reports. Install/run under the Outlook user's Windows account. Code updates do not install or update the Windows scheduled task; run Setup/Create Daily Draft Task.bat separately after confirming the selected files and mailbox.

The operator's actual network files, official COD sequence and Office/security policies still need a one-case acceptance check. Two independent installations cannot safely share numbering through separate local registers. Uncertain saves require review rather than automatic retry. This version makes recycling certificates only and never sends mail automatically. Completed is a manual record, not mailbox delivery tracking.

The existing Windows task was rescheduled to the saved **11:54** local time and its sign-in catch-up was preserved. Its next planned daily run is **25 September 2026 at 11:54**. The previous run's result remains `0x800710E0`; Task Scheduler's Operational log is disabled, so the underlying cause is unknown. The task was not manually started during this review, so no live Outlook draft was created. Plant 1 had an Excel lock at review time; neither plant workbook was modified.
