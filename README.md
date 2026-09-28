# Ecoreco certificates — start here

Double-click **Prepare Certificates.bat**.

1. Save and close the two plant Excel files.
2. Click **Create all Ready drafts**. Both files are checked automatically; you do not need to click Check first.
3. Review and send the prepared emails from Classic Outlook > Drafts.

For a listed case, select it and **Open selected plant**. Choose Ready, N/A, or Blank in Excel, then save and close it. After the app saves an Outlook draft, it sets the row to **Drafted**. Review and send the email yourself in Outlook, then select the case in the app and click **Mark completed / Undo**. This updates Excel and keeps a recoverable backup; it never sends email for you.

In **Certificate Issued**, Ready (green) means approved; N/A (white) means skip; **Blank** is red; **Drafted** (yellow) means an unsent draft is waiting for review; and **Completed** (blue) means the person confirmed it was sent. A truly empty cell also needs a decision. Drafted, Ready, and Blank cases stay in the 30+ day reminder list until the person marks them Completed or N/A. The reminder also groups inward dates that cannot be safely age-checked and urgent automation problems. It does not pop up for every blank field or every client. Older rows marked Sent remain treated as Completed; older Completed rows without a saved completion confirmation remain Drafted.

Before the first draft, the user may enter a unique COD number and certificate date in Excel; if COD No is blank, the app assigns the next shared number. After a draft is saved, use **Correct certificate** in the app if its COD number or date needs a genuine correction: select the unsent Drafted case, enter the corrected values and a reason, and confirm. Even if the cells were already edited in Excel, the app shows both the saved certificate values and current Excel values, then updates the same unsent Outlook draft, replaces its PDF, updates Excel, and keeps the old PDF and correction history. Editing those cells in Excel alone does not change the certificate. Completed certificates cannot be changed in the app; follow the company's correction/replacement process. If a correction is interrupted, the app pauses new drafts and offers to resume the saved correction. It never sends email.

To deliberately create another draft for an unchanged certificate, change Drafted or Completed back to Ready: it reuses that batch's current COD number and date, leaving the earlier draft untouched for review. If an earlier Outlook save was confirmed but the Excel update was interrupted, **Create all Ready drafts** finishes the Excel update without creating a duplicate message.

**Settings** contains the saved plant files, mailbox, one-draft test and daily reminder controls. The daily check and reminder default to 09:00 local time. Change it and click **Save & update time** to move the installed daily schedule. The scheduled check also prepares Ready drafts; a sign-in catch-up may prepare drafts earlier, but will not show a reminder before the selected time. These settings stay out of the way after setup. A new installation opens Settings automatically.

Each eligible reminder run shows at most one centered summary of overdue cases, inward dates that cannot be age-checked, and automation problems. A separate later run—including the one-hour read-only snooze—can remind again the same day; the same run will not create duplicate popups. **Yes, open review** opens/focuses the app and refreshes its results; **No, remind tomorrow** pauses reminders for today; **Remind me in 1 hour** schedules one extra read-only check. Saving the reminder time re-arms alerts after a prior “No” or window close. A local scheduled task needs this Windows account to be signed in; it cannot check the files or show a reminder while the computer is off. Click **Review alerts** any time to refresh both workbooks and see the actual overdue cases, date problems, and urgent system issues in the list; this does not create drafts. Completed is the person's confirmation, not automatic mail tracking.

## Folders

- **Setup**: one-time installation, daily reminder setup and troubleshooting. Not needed every day.
- **App**: program, templates, sample files, tests and technical documentation. Leave this folder in place.
- **Data**: certificate numbering, backups and reminders. Never delete it to reset a case.
- **Plant Workbooks**: the two selected plant Excel registers.
- **Output**: generated certificates and reports.
- **App/sample/Old test inputs**: archived older dummy input sheets; not used by the app.

When updating, keep Data, settings.local.json and Plant Workbooks. The older dummy sheets were moved out of Input and archived under App/sample/Old test inputs. Plant Workbooks is excluded from Git. The hidden scheduled_run.py and Run Daily Drafts.bat are compatibility entry points for existing daily tasks.

First-time setup: run **Setup/Setup Once.bat**, open the app, select the two files and mailbox, then close Settings. To enable daily reminders, use **Setup/Create Daily Draft Task.bat** once under the same Windows account as Outlook. The task checks at the chosen time and shortly after sign-in; sign-in catch-up will not show a reminder before the chosen time.

Developer tests: from App, run `python -m unittest discover -s tests -v`. Technical notes are in App/docs.
