# Ecoreco recycling certificate automation

The desktop app reads two plant inward workbooks, generates certificates for complete batches marked **Ready**, and attaches each certificate to a Classic Outlook draft. The person reviews and sends the drafts.

## Start here

1. Run `Setup Once.bat` once as the Windows user who will use Outlook.
2. Open `Prepare Certificates.bat`.
3. Choose the two plant Excel files and the Outlook mailbox once. The app remembers them.
4. Click **Check both files**. Fix any rows marked **Needs attention** in Excel. Set approved rows to **Ready** and intentional skips to **N/A**.
5. Save and close both workbooks. Click **Create one draft** to prepare one eligible certificate, or **Create all Ready drafts** after checking the results.
6. The app uses the same workflow for any selected plant files: the certificate template, Outlook Drafts folder, shared COD sequence, and COD/date write-back to Excel. It never sends email. Review unsent drafts in Classic Outlook > Drafts.

The app remembers the two plant file locations. Keep the same plant order and file locations after issuing certificates; moving a register requires a checked history migration. The first COD number is needed only when neither plant workbook has an earlier number. Set it under **Setup options**.

The selected source files are updated at their saved locations. You do not need to choose them or copy them into this folder each day.

## Certificate Issued column

| Cell value | Suggested display | Behaviour |
|---|---|---|
| Ready | Green | Prepare on the next draft run once required details are complete. No 30-day wait. |
| N/A or Not Applicable | White | Skip the batch; do not allocate a number or send an overdue reminder. |
| Blank or empty | Red | List the batch as needing a decision; do not issue a number. |
| Completed | Yellow | A draft was saved in Outlook; it still needs a human review and send. |
| Sent | Blue | The operator confirmed the email was sent; stop overdue reminders. |

The text is the instruction. A cell painted green with no Ready text remains blank and is reported. Status colors are linked to the cell text by the status formatting in the plant workbooks.

The app lists blank decisions and missing required details right away, but does not pop up for every new blank. The daily reminder includes inward batches whose Inward Date is at least 30 days old and whose status is not **Sent** or **N/A**. That includes Blank, Ready, and Completed rows: a saved draft is not proof that the certificate was sent. It also includes one grouped data-quality count for rows with a missing, invalid, or future Inward Date, because their age cannot be safely calculated. Automation failures that reach the notification stage are summarized separately in the same reminder. A failure before startup/dependencies load is written to the daily-check diagnostic instead; no popup can be guaranteed in that case. There is no popup for every case or every blank field.

## Required source fields

Each workbook must contain one inward register with headings in row 1. Items on subsequent rows with blank SN belong to the preceding batch.

- SN, Inward Date, Supplier Name, Item. Location is included when provided.
- Weight (KG), or Quantity in Pcs and Per Pcs Weight in (KG).
- Form 6 No: printed as Manifest No.
- Client Email: one complete email address. Address lists and display-name text are flagged for review.
- Reference No.: ask Chahna to add and populate this field in both files.
- Certificate Issued, COD No and COD Date.

Missing information appears in the results list. No official number is issued to an incomplete batch. A blank weight can be calculated from quantity multiplied by per-piece weight. Formula weights are accepted only for that exact same-row multiplication; other formulas are flagged rather than guessed. Ambiguous headings, orphan rows, conflicting continuation data, invalid numbers and text too long for the certificate are reported.

Unrelated/reference tabs are allowed, but each workbook must have exactly one register tab with the required headings in row 1. The supplied Plant 1 and Plant 2 workbooks each contain only the **Inward Register** tab.

## One COD sequence for both plants

The app reads existing COD numbers from both files and uses one local register in `Data/certificate-history.sqlite`. For example, if the largest existing number is 0064, the next certificates are 0065, 0066, etc., whichever plant they belong to.

If neither file nor the register contains an earlier number, enter the approved **First COD number** once in the app. A fresh installation does not silently assume the sequence starts at 0001. Keep the Data folder when updating the application.

A number is reserved only for a complete Ready batch during drafting, never during a check. The certificate date is the day the number is first reserved; a recovery keeps that same date. Previews do not allocate numbers. Reserved numbers are never recycled following an interrupted run. Old eligibility_days settings are ignored.

## Your work list

Use the list filter to see **Needs attention**, **Ready to draft**, **Drafts to review**, **Marked sent**, or **Skipped**. Select a pending case and click **Open selected plant**; the message identifies its SN, sheet and row. Update Certificate Issued to Ready or N/A, save and close Excel, then click Check both files.

Blank means the approval decision is pending, not proof that recycling is complete or that the client has not received a certificate. Ready is a human approval; the app does not infer it from age or colour.

After actually sending an email yourself in Outlook, select its prepared case and click **Mark sent / Undo**. Confirming updates only that batch's Certificate Issued cell to **Sent** in the plant workbook and records your confirmation in the local Data register. It does not send mail or verify delivery. A backup is made before the workbook change. Undo changes the status back to **Completed**; it does not unsend the email. Set it to **Ready** only if another draft is needed; the existing COD and date are reused.

After the draft save is confirmed, the same COD number and certificate date are written to that batch's COD No and COD Date cells. Existing numbers are not overwritten. Duplicate or unrecognised number formats are reported before further numbering.

Use one installation to process both plant files. Two independent installations with separate Data folders do not share a number register.

## Workbook updates and interruptions

Both workbooks must be available and closed in Excel before the app issues numbers. A locked, read-only or unavailable file is reported rather than overwritten.

Draft creation writes the COD No and COD Date cells only after Outlook confirms the draft save. **Mark sent / Undo** changes only the selected Certificate Issued status cell. Both operations back up the workbook before changing it and preserve its other sheets, formulas, merged cells, formatting, validation, images and package parts.

The app reads back a newly saved Outlook item from the selected mailbox and checks its batch identifier, subject, draft folder and attachments. It never opens a draft window or sends the message. If a draft is saved but the workbook update fails, the next run retries the workbook update with the same number and does not create a second draft. If Outlook's save result is uncertain, issuance pauses for review rather than automatically creating another draft. A live smoke test with the configured Classic Outlook profile verified an unsent draft and its PDF attachment; the unique test draft was removed afterward. This used temporary workbook copies and a non-deliverable test address.

## Quiet daily check

After choosing both files and the Outlook mailbox, run `Create Daily Draft Task.bat` once as the same Windows user who uses Outlook.

The task runs at the time saved in Settings (default 09:00) and two minutes after that user's Windows sign-in. Use **Save & update time** in Settings to reschedule an installed task; if it has not been installed yet, the saved time is used during setup. Windows Task Scheduler is configured to catch up after a missed start. It uses a hidden Python process, performs one pass and exits. Overlapping runs are blocked. Installing the files alone does not enable the daily task. The scheduled check prepares Ready drafts; a sign-in catch-up can prepare them earlier, but the overdue reminder waits until the configured time.

Each eligible reminder run shows one centered summary of overdue inward records (30+ days old and not Sent/N/A), rows whose Inward Date cannot be age-classified, and urgent automation problems that reached the notification stage. It does not show one prompt per client, and the same run cannot create duplicate popups. A separate later run—including a one-hour snooze—may show another summary the same day. **Yes, open review** opens or focuses the app and refreshes its results list; **No, remind tomorrow** closes the reminder for today; **Remind me in 1 hour** schedules one read-only refresh of both plant files and then shows an updated summary. Dismissing the reminder does not mark any case Sent. Saving the reminder time clears a prior same-day dismissal and re-arms alerts. Unresolved overdue records repeat on later eligible runs until the operator marks them Sent or N/A. The app always retains the complete results list. The checkbox can turn reminders off. Reminders require the daily task to be installed; the checkbox alone does not install it. Ordinary missing fields and blank status decisions are visible in the work list immediately, but only become popup reminders once that batch is 30+ days old. Missing, invalid, or future inward dates are an exception: one combined data-quality count is included so an un-ageable row is not silently excluded. Failures before Python packages or the reminder runner start are recorded in a diagnostic file rather than shown in a popup.

Notification retries do not double-count drafts. Failed notification attempts retain pending information; an interrupted or uncertain notification attempt is not repeated the same day. Muted runs acknowledge current information without building a backlog. A missing workbook cannot clear its earlier unresolved issues. Corrupt notification history is backed up and reported, and resumes on a later day. Windows can suppress or delay alerts through Focus Assist, account policy, or an unavailable desktop; a submitted notification is not proof it was read. The results list remains the authoritative record.

If the Outlook connection is unavailable, the daily job still checks readable sheets for missing information. If a required Python package prevents even reading the sheets, it writes an error report where possible. If the project folder is unwritable, a fallback diagnostic is placed under the user's Local AppData `Ecoreco Certificates` folder. Run the setup under the Windows user who will operate Outlook. The launchers do not bypass company restrictions or install packages automatically during a normal run.

The desktop must be on, the user signed in, and the Classic Outlook profile available to create drafts. Excel locks may require saving/closing the workbooks and running again. Automatic sending is not enabled.

## Files and testing

- `Output/Latest Results.txt`: the latest readable report.
- `Output/Latest Results.json`: the latest app results.
- `Output/<date>/<run>/`: PDF files and the report for each run.
- `Output/Vendor alignment preview.pdf`: static layout preview of the certificate template.
- `settings.local.json`: saved file paths, mailbox and preferences.
- `Data/`: the shared number register, alert state and workbook backups.
- `Data/certificate-history.sqlite`: shared COD register for both selected plant workbooks.

Plant Workbooks, Output, history, backups and local settings are excluded from Git. Keep real client files out of the repository.

Developer checks: install `requirements-dev.txt`, then run `python -m unittest discover -s tests -v`. The integration tests use temporary dummy workbooks and a fake Outlook writer; they never send mail.

See `VERIFICATION.md` for verified behavior, remaining limitations and deployment checks. This project currently generates recycling certificates; there is no separate data-destruction certificate template or unattended email-sending workflow.
