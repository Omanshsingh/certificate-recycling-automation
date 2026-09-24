# Current certificate automation

Implemented locally:

- Two plant workbook selections remembered between runs.
- Ready / N/A / blank status handling and required-field checks.
- Form 6 mapped to Manifest No.; missing Reference No. reported.
- Centred vendor text, using Calibri Bold when available, with balanced wrapping for long names.
- One sequence across both selected workbooks and a persistent local register.
- COD No and COD Date writeback, workbook backups, and recovery of failed writebacks.
- Duplicate prevention and a stop for uncertain Outlook saves.
- Results shown in the app without completion popups.
- Optional silent daily summaries with duplicate suppression.
- Daily task installer with sign-in and missed-run handling.

Before using live data, select the two real files, provide the Reference No. and Certificate Issued columns, select the Classic Outlook mailbox, and check one real Outlook draft. Local automated tests use dummy files and a fake Outlook writer. This version prepares drafts; it does not send mail automatically.
