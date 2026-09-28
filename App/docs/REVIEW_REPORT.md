# Certificate automation review

Review date: 23 September 2026; reminder/status follow-up verified 24 September 2026. Local project: `C:\certificate-recycling-automation`.

## Outcome and scope

The app uses one plant workflow for every selected workbook, regardless of filename or recipient domain. This turn added a guard against reissuing a previously saved batch if its SN is changed and its COD/date are cleared; the run now blocks until that saved row is restored or an administrator reconciles an archived register. The renamed Plant 1 and Plant 2 workbooks contain 10 and 5 register batches plus 24 documented scenarios each on a separate, non-processed **Test Cases** tab. A real check-only scan found both files, created zero drafts, and reported 15 batches: 3 already prepared, 10 ready, 1 N/A, and 1 intentional blank-status case for review.

A live Outlook smoke test used temporary copies of those workbooks and a unique non-deliverable `@example.test` address. The app created one unsent draft in the configured Classic Outlook Drafts folder, read it back and verified its PDF, then removed that exact test draft. The original workbooks and shared history were not changed by the smoke test; no email was sent. No scheduled task was registered and nothing was pushed to GitHub.

This is a tested local implementation, not a claim that every possible failure is covered or that the company's desktop environment has been certified. Outlook and Windows notification behaviour must be checked on the actual operator's machine.

At the time of the 23 September review, **170 tests passed**. The 24 September follow-up regression run passed **207 tests**; details are recorded below. The earlier 100-certificate volume run took 117.91 seconds. The renamed workbooks retained their original register rows, formulas, validation and formatting. The shared next COD remained **0305** through the file-path migration.

## Existing setup found on this laptop

The read-only desktop check found a pre-existing, enabled task named `Certificate Recycling Drafts`. It still launches the batch file directly, has only the daily 09:00 trigger, has missed-run catch-up disabled, and blocks battery starts. Its last recorded run on 23 September returned `0x800710E0`; this review does not infer the cause from that code alone. Its configuration was not changed.

The saved application settings now point to `Plant 1.xlsx` and `Plant 2.xlsx` in the local `outputs` folder, in that order. The existing shared COD history was safely backed up and rebound to those files; its records and next number (**0305**) were preserved. The original input workbooks remain untouched. Confirm the approved source workbooks and official COD sequence before production use. After the live setup, rerun the updated daily-task installer as the operator to replace the older task with the tested hidden-process/sign-in/catch-up configuration. Do not assume the code changes alone updated the registered Windows task.

## What was fixed

| Area | Risk found | Current behaviour |
|---|---|---|
| Alerts | Reprocessing one result could count its drafts again | A durable run identifier prevents repeat counting. |
| Alerts | Failed/interrupted notifications could lose pending work | Pending counts are recorded before notifying; an ambiguous attempt is deferred until a later day. |
| Alerts | A second run could finish while the first notification was showing | Its report is queued on disk and processed without another popup storm. |
| Alerts | Missing files could make earlier issues appear resolved | Partial scans preserve unresolved issues from unchecked plants. |
| Alerts | Muting could leave a surprise backlog | Muted runs acknowledge current issues/events; new issues after unmuting can notify. |
| Alerts | Corrupt alert history could stop notifications indefinitely | The damaged file is retained, diagnostics are written, and notification state can recover. Unknown historical counts are not guessed. |
| Data | Unsupported weight formulas were silently substituted | Only an explicit same-row quantity/per-piece multiplication is calculated; unsupported formulas require review. |
| Data | Number-like text, ambiguous headers and orphan rows could slip through | Strict validation identifies these problems before issuance. |
| Data | Long reference/manifest text could overrun the template | Field fitting is checked before a number is reserved. |
| Excel | Merged/protected/formula COD cells and failed backups needed safeguards | Those cells are not overwritten; backups must succeed before replacement. |
| Numbering | Workbook numbers could conflict with another owner's reservation | Both workbooks and the history register are checked together. |
| Numbering | Swapping plant slots could attach history to the wrong register | Existing history is bound to the selected paths and plant order. |
| Numbering | Changing an issued SN and clearing its COD could make it look new | Missing saved rows now block issuance until the original row is restored or an administrator reconciles an archived register. |
| Recovery | N/A or a manually filled COD could conceal an interrupted save | Pending saves remain visible; unsafe re-drafting is blocked. |
| Recovery | Global errors could leave eligible rows out of the report | Unprocessed candidates are listed as waiting instead of disappearing. |
| Outlook | An EntryID alone was treated as confirmation | The saved item is read back from its store and checked for its identity, folder, unsent status and attachments. |
| Startup | Corrupt settings, missing packages and unwritable folders could cause silent failures | The UI retains recovery controls; scheduled errors have reports/fallback diagnostics. |
| Permissions | Access-denied errors could be labelled "already running" | Actual lock contention is distinguished from permissions failure. |

The read-back implementation follows Microsoft's documented [GetItemFromID interface](https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getitemfromid), using both the item ID and the selected store ID. API documentation is not a substitute for testing the operator's Outlook provider.

## Alert rules

- Current policy: each eligible reminder run may show one centered summary for inward batches at least 30 days old whose Certificate Issued status is not Completed or N/A. Blank, Ready, and Drafted all qualify once aged; a saved draft is not proof of sending. Distinct later runs may alert again the same day; duplicate popups for the same run are blocked.
- Blank decisions and missing required details appear in the app results immediately but do not cause a popup before the inward date reaches 30 days. A missing inward date remains visible but cannot be aged.
- Automation failures that reach the notification stage (for example, a source workbook unavailable or an uncertain Outlook save) are included in the next scheduled summary; a partial scan is not treated as a clean full scan. Failures before the runner or its dependencies load go to a diagnostic file instead of a popup.
- The reminder is aggregated, not one popup per case, and unresolved overdue cases repeat on later eligible runs. “No, remind tomorrow” pauses for the rest of today; saving a reminder time clears that same-day pause. “Remind me later” snoozes for one hour and runs a read-only refresh. Marking a row Completed or N/A removes it from the overdue reminder set.
- The app does not send mail. After sending manually, Mark completed records the operator's confirmation and updates the status cell in Excel; it does not verify delivery. Undo restores Drafted, not an unsent message.
- The daily task does one pass at its configured time; it is not an all-day watcher. A shut-down or signed-out desktop cannot display a reminder until a later eligible run.
- The OS accepting a notification does not prove the user saw it. Focus Assist, Windows settings and company policy can suppress it.

## Tests performed

The follow-up regression run on 24 September passed **207 tests in 99.463 seconds**. Tests include the exact day-29/day-30 boundary, Blank/Ready/Completed versus N/A/Sent statuses, Excel Sent update/undo, alert aggregation and simulated Outlook. No emails were sent by this suite.

The latest local regression run on 26 September passed **248 tests in 128.353 seconds** after the Drafted/Completed lifecycle and audited COD/date correction changes. It includes simulated Outlook only; no emails were sent.

The automated suite uses temporary Excel workbooks/PDF rendering, simulated Outlook objects and simulated Windows notifications. Separately, one normal workflow draft was verified in the selected Classic Outlook profile with a non-deliverable recipient, then removed from Drafts. No tests sent email or registered a Windows task.

| Test group | Cases covered |
|---|---|
| Normal workflow | Ready/N/A/blank, required fields, Form 6, preview-only mode, repeated runs, saved paths, same-day Ready eligibility (no 30-day gate) |
| Data and workbook preservation | Duplicate/alias headings, continuation conflicts, orphan rows, malformed numeric values, unsupported formulas, long fields, protected/merged COD cells, absent XML cells, 1904 date epoch, formulas/styles/validation and untouched ZIP package parts |
| Numbering and recovery | Global duplicates, registry ownership, swapped/moved sources, removed/rekeyed saved rows, newly added rows, missing pending rows, changed source details, uncertain saves, workbook-write failure, history-commit failure, concurrent allocations and atomic save claims |
| Alerts | Empty runs, 100 issues in one summary, retries, day boundaries, clock rollback, muting/unmuting, resolution/reappearance, corrupt state/queue, notifier exceptions, interruptions, concurrent passes, access denied and state-write failure |
| Outlook contracts | Correct mailbox/folder, PDF/banner attachment, account assignment failure, empty save ID, read-back failure/mismatch, retrying reads without another save, no Send or Display calls; one live non-deliverable-recipient draft read back successfully |
| Desktop and scheduler | Damaged settings/results, failed saves, missing packages, worker errors, close during processing, hidden task action, sign-in delay, catch-up setting, overlap protection, same-user execution and explicit-only package installation |
| Volume | 100 Ready batches across two workbooks: 100 real PDFs, one COD sequence, COD/date writeback, two original-file backups and zero duplicate drafts on repeat |
| Mixed volume | 100 batches: 20 drafted, 40 needing review, 20 N/A, 20 waiting; one summary; repeated run creates no duplicate drafts or unchanged-issue alerts |

The 100-certificate local run took approximately 118 seconds in the final measured run. Outlook was simulated for the volume test, so this is not a prediction of live Outlook speed. Processing is sequential; it does not open 100 message windows.

## What is still needed before live use

1. Select the two actual plant files once, in a fixed Plant 1/Plant 2 order. Confirm both are writable by the Outlook operator and that SN remains unique within each register. Test on copies first.
2. Ask Chahna to add/populate Reference No. Confirm Form 6 No, Client Email, Certificate Issued, COD No and COD Date columns exist. Ready must mean the recycling work and certificate details are approved. There is no age-based approval or waiting period.
3. Confirm the official COD prefix, current sequence, whether numbering resets by financial year, and the approval/voiding procedure. Do not infer an official starting value from dummy data.
4. Before production, create one draft from an approved plant row after the official COD sequence and certificate fields are confirmed. Verify its PDF and matching Excel number/date before preparing more drafts.
5. Install the daily task under the same Windows user only after that test. Test one notification, one normal run, a missed scheduled start, and an unavailable/locked workbook. Confirm company policy permits the task and that Focus Assist does not hide required alerts.
6. Keep and back up the Data folder, saved settings and certificate outputs when updating. Choose an IT-approved backup/retention policy; the app does not automatically delete old backups or PDFs.

## Final alert and scheduler pass — 24 September 2026

The final local suite passed **224 tests**. Added checks cover the grouped data-quality reminder for missing/invalid/future inward dates, all three centered-reminder buttons and hidden-window cleanup, stale review-list refresh, failed snooze cleanup, and a case changed to N/A after a saved certificate. Automated Outlook tests are mocked; no live Outlook draft or email was created during this final pass.

The daily reminder includes: (1) inward records at least 30 days old whose status is not Completed or N/A—including Blank, Ready, and Drafted, because Drafted means an unsent draft exists; (2) one aggregate count for inward dates that cannot be age-classified; and (3) automation failures that need action. Ordinary blank fields and recent blank statuses remain in the review list without a popup per row. The user marks Completed only after sending manually; the app never sends or verifies delivery.

The installed `Certificate Recycling Drafts` task was found at 09:00 while saved Settings specified 11:54. The first reschedule attempt exposed an unsupported `Set-ScheduledTask -Description` parameter. The installer now updates the retrieved task definition instead, and the actual Windows PowerShell reschedule succeeded. The task remains interactive under the existing user, retains its Python action and sign-in trigger, and is now scheduled for 11:54. Next daily trigger: 25 September 2026 at 11:54. The previous result remains `0x800710E0`; Task Scheduler Operational history is disabled, so its cause was not inferred and the task was not manually run. Plant 1 was locked in Excel; both plant workbooks were left untouched.

The centered prompt uses a local interactive scheduled task rather than a Windows toast so it can offer **Yes, open review**, **No, remind tomorrow**, and **Remind me in 1 hour**. The one-hour action re-reads both files read-only; when an existing app is focused, its list refreshes from the new report. The reminder task cannot run while the computer is off or the user is signed out; Windows catch-up is best-effort, not a guaranteed exact-time delivery.

Design references: [Task Scheduler StartWhenAvailable](https://learn.microsoft.com/en-us/windows/win32/taskschd/tasksettings-startwhenavailable), [logon triggers](https://learn.microsoft.com/en-us/windows/win32/taskschd/logon-triggers), [Set-ScheduledTask](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/set-scheduledtask?view=windowsserver2025-ps), and [Windows notification UX guidance](https://learn.microsoft.com/en-us/windows/apps/design/shell/tiles-and-notifications/toast-ux-guidance?view=powershell-4.0).

## Remaining product limitations / highest-value follow-ups

- **Drafts are not sent automatically.** The 30+ day reminder includes a Drafted batch until the operator sends it manually and marks it Completed in the app. This is not mailbox-wide monitoring: if the operator sends outside the app and does not mark the row Completed, reminders continue.
- **One installation owns the numbering register.** Two independent desktops do not safely coordinate through separate Data folders. If both plants must issue independently, a shared transactional service/database and deployment plan are needed; do not copy the SQLite file between live machines.
- **Ambiguous Outlook saves require reconciliation.** The app deliberately stops rather than guessing. There is not yet an operator-friendly recovery screen for linking an existing draft or approving a reviewed void. Do not delete history to clear this condition.
- **Recycling only.** This build has a recycling certificate template. A distinct data-destruction certificate, its fields and approval rules are not implemented.
- **File locations are fixed after issuance.** Moving/renaming a numbered register or rotating to a new workbook requires an explicit checked history migration. Reselecting an unrelated file cannot silently reuse old history.
- **No endless background watcher.** The task checks at 09:00 and shortly after sign-in. If files remain locked at that point, retry manually after closing Excel or wait for the next scheduled pass. A hung Outlook/provider call cannot be made reliable purely by mocks.
- **One client email per batch.** Multiple recipients are flagged; company-specific To/CC rules need agreement before adding multi-recipient certificate delivery.
- **Desktop/network validation remains necessary.** Actual mapped drives/SMB behaviour, macro-enabled company templates, antivirus controls, Office prompts, sleep/resume and account-provider behaviour have not all been exercised here.
- **Reminder delivery depends on the desktop task.** It runs at the selected local time and after sign-in catch-up; it is not an always-on service and cannot alert while the PC is off or the user is signed out.

## Reproduce the checks

Use the development dependencies, then run `python -m unittest discover -s tests -v` from this project. `requirements-dev.txt` includes the PDF reader used by the tests. The test fixtures and outputs are temporary and contain dummy email addresses only.
