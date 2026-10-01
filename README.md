# Ecoreco Certificate Automation

This local browser app reads the two plant Excel workbooks and generates certificate PDFs. It runs only on this computer at `127.0.0.1`; it has no Outlook integration and does not create drafts or send messages.

## Everyday use

1. Open **Prepare Certificates.bat**. The local dashboard opens in your default browser.
2. Use **Settings** to choose the Plant 1 and Plant 2 workbooks. In Excel, mark approved rows **Ready**, use **N/A** to skip a certificate, and save and close Excel.
3. Find the company and choose **Review & create**. Check the details, optionally use **Preview PDF**, then create the PDF. For bulk creation, open **Create several at once**. Warnings always need confirmation before proceeding.
4. Find PDFs in the top-level **Certificates** folder. Names use `{COD}_Company name_city_1.pdf` or `{COD}_Company name_city_2.pdf`; the COD matches the number written to Excel, and the final number identifies the plant.
5. Send the PDF through the normal company process. Then choose **Mark Completed**; Excel records **Completed** and reminders stop.

The app writes a newly assigned COD number and date back to Excel. A COD or date already entered by the user is respected. Reference No. is optional; Form 6 No. remains the required manifest identifier. If a row is deliberately changed back to **Ready**, the current row values control the next PDF generation. If COD is blank, the next number is chosen from both plant workbooks and the existing certificate PDFs so the sequence is shared.

When editing an existing inward's certificate, choose **The PDF only — keep the Excel details** or **The PDF and Excel**. Both choices record the final COD, certificate date and Generated status in Excel for tracking. Previewing, checking warnings and cancelling do not save form edits. Edits are staged in memory and committed only after the PDF is successfully created; failed workbook updates roll back the PDF publication. The selected inward ID and plant stay fixed. New inward rows use **The PDF and Excel** and are added only on successful creation. Normal bulk creation still skips Blank and N/A rows.

The main screen shows company, plant, received date, status and next step, with ten companies per page and simple **Previous / Next** controls. Click a summary card to filter the list, or search by company, COD, Form 6 or reference. **More** shows concise certificate facts and warnings; full messages and backup filenames stay under **Technical details**. **Saved certificates** stays collapsed and lists only PDFs currently in the folder, with readable names and ten files per page. **File details** opens the filename and references; **Past numbers** opens a separate searchable view of previously used numbers. Inspection warnings remain visible in the section summary, and automatic number protection is unchanged. Optional material calculations, number/date changes and PDF attachments open only when needed. All original fields and safety checks remain available. The status names are the same in the dashboard and Excel: **Ready**, **Alert**, **Generated**, **Generated Manual**, **Completed**, **Blank** and **N/A**. On refresh, older status labels are converted to the matching current name; empty cells remain empty and N/A stays separate. Overdue appears separately and never replaces a certificate's true status. Floating-point tails are cleaned up on screen; untouched numeric fields retain their original values when submitted.

Use **New certificate** when the details were not entered in Excel first. Select **Excel** under **Use details from**, choose the plant or its workbook through Windows Explorer, and enter the details. The app adds/updates the SN row only when PDF creation succeeds. **Weight to print (kg)** is optional for Excel rows; when populated it replaces the summed inward weight and can contain the customer's net weight.

## Certificate types and custom certificates

Choose **Standard COD**, **Sita International**, or **Siemens Medical** in the certificate editor. The template is an explicit choice, independent of the client name. The supplied PowerPoint templates provide the original artwork, fonts and fixed wording. Standard uses the original five-row table ending at Manifest No., with no disposal-date row or warning. Sita has its own receipt/reference table; Siemens has model, model serial, X-ray tube serial and site-location fields.

For Sita and Siemens without an inward entry, choose **Enter myself (no Excel row)** under **Use details from**. This saves to a separate local register under `Data/Custom Certificates.json`. It does not invent a plant inward or require either international workbook. Pick the plant, enter the certificate details, **Preview PDF**, then **Save draft** or **Create PDF**. Saved drafts appear in the same list and can be edited later. **Copy as new**, available under **More**, reuses an earlier custom certificate's details while clearing its COD and dates for a new certificate; review the copied weights and serial numbers. Custom certificates share the COD sequence with both plant workbooks and existing PDFs, and keep their assigned number if the PDF is deleted. International forms do not show Form 6.

Use **Company line on certificate** to change only the printed name, such as Sita Laptops or Sita Metals, without replacing the supplier name in Excel. Designs, standard wording, table labels, signatory labels and footer are fixed; staff cannot edit them or move/resize the layout. Older saved wording/layout overrides are ignored when creating new PDFs, while existing PDFs are left untouched. **Weight as printed** supports exact units or weight descriptions, while the numeric certificate/net weight remains available for the register. Decimal weights retain their precision.

Missing dates, manifests, weights and model identifiers produce warnings with their consequences. Staff can explicitly choose **Create anyway**. An invalid template selection, an unreadable file, a protected workbook cell or an unavailable attachment still needs correction because the app cannot safely produce/save the requested result. If a workbook is unavailable, a custom certificate can still be created after accepting the warning to verify its COD against that workbook before sending.

Use **Attach PDF list** for an existing annexure, such as an item/serial-number list. The app copies the selected PDF into its local data folder and appends its pages to both preview and final PDF. It never reuses the example annexures automatically.

When these clients eventually enter the inward register, choose **Plant workbook** and the same template. The app persists `Certificate Type` and `Certificate Fields` columns as needed. Operators choose the type in the dashboard; the saved additional details travel with the inward row. No client-name matching rules are needed.

On workbook refresh or when the dashboard checks its saved work list, a row marked **Generated** or **Completed** with no matching PDF in `Certificates` returns to **Ready** in Excel, retaining its COD and date for recreation. A missing COD does not prevent recovery; a new number is assigned only when the PDF is created. Matching uses COD and plant, so changing a company name does not incorrectly mark an existing PDF missing. Archive copies do not replace the required current PDF. Refresh never generates a PDF. If Excel is locked or protected, the app reports the failed status update and asks you to close or correct the workbook before retrying. Custom records follow the same missing-PDF rule and track their current PDF filename; when a custom certificate is regenerated with a different name, the superseded PDF moves to `Data/Archive/Superseded Certificates` and remains recoverable. **Completed** confirms sending only while the current PDF remains in `Certificates`.

Every source or layout issue is shown as a warning. Before generation, the app lists the warnings and asks for confirmation. Continuing can produce a PDF with a blank, fallback, or manually corrected value, and the row is still marked Generated, so verify the PDF and workbook before sending it.

## Manually created PDFs and used numbers

Staff may create a certificate outside this app and put its PDF in `Certificates`. Use the same `{COD}_Company_city_1.pdf` or `{COD}_Company_city_2.pdf` naming format, with the actual COD at the start. The app checks this folder before each creation, even when that certificate has no Excel row. It reads CODs from filenames and labelled PDF text, and references from readable PDF text. Automatic CODs start above the highest observed COD; manually entered used CODs require confirmation to switch to a new automatic number. Confirmation never authorizes overwriting someone else's certificate.

The dashboard's **Saved certificates** section shows PDFs currently in the folder and inspection warnings. Click **File details** for a PDF's filename and references. **Past numbers** is a separate searchable list, so removed files do not clutter the saved-PDF list. Filename/printed-COD mismatches reserve both numbers and need review. Scanned, encrypted, oversized or unreadable PDFs may only have a usable filename COD; their references need manual verification. This is not OCR and does not automatically import unverified PDF details into Excel.

References remain optional and are not automatically generated or incremented. A reference already seen on another certificate produces a warning about a possible duplicate or shared inward. Staff may deliberately keep it after confirmation.

`Data/Certificate Inventory.json` remembers observed identifiers after a PDF is moved, replaced or deleted; backups are under `Data/Certificate Inventory Backups`. Deleting a PDF does not free its observed COD. Previewing does not reserve a new COD. The folder is checked again immediately before publishing, so an external file added during rendering cannot silently be overwritten.

## Reminders

The default reminder time is **11:00 AM** and can be changed using the time picker. Click **Save & update reminder** to enable the daily Windows reminder. It checks both saved workbooks at that time and after sign-in catch-up; the reminder is shown only in the signed-in Windows session while the computer is awake.

At inward age 25, only the **Supplier Name** cell turns red in Excel and the app shows the row for follow-up. The SN and Reference No. remain unchanged. The reminder continues through day 30 and becomes **Overdue** after day 30. It stops when the row is **Completed** or **N/A**. “Completed” is the user’s confirmation that the certificate was sent; the app cannot detect delivery. Missing or invalid inward dates and certificate data are shown as attention items because their age or PDF cannot be safely determined.

Daily reminders only check and notify. They never generate PDFs on their own. Generating PDFs is always a user action in the app.

## Setup

- Run **Setup\Setup Once.bat** to install Python packages for the current Windows account.
- Run **Prepare Certificates.bat** to start the local server and open the dashboard. Use **Close dashboard** at the top when finished.
- Save both workbook paths in Settings. Save and close Excel before checking or generating because Excel may lock an open workbook.
- To enable reminders, use **Save & update reminder** or run **Setup\Create Daily Reminder Task.bat**.

Dependencies: Python 3.11+, `openpyxl`, `reportlab`, `pypdf`, `fonttools`, `pypdfium2`, Pillow, and Tkinter (used for the Windows file picker). PowerPoint is not required to create certificates: the three supplied layouts are included as PDF templates. The local dashboard uses Python's standard library; no web framework or internet connection is needed. Outlook and pywin32 are not needed.

Setup uses `App/requirements.lock.txt` for reproducible package versions. `pypdfium2` and Pillow render the first page in the editor, so preview does not depend on the browser's built-in PDF viewer. Open the full PDF to check every attached annexure page. A normal installation needs internet once to obtain packages, not to use the app afterward.

## Safe creation and recovery

The editor has three steps: check details, review exact before/after changes, then preview. A neutral, readable interface keeps optional fields, file history and bulk operations out of the everyday path. The success screen shows the actual issued number and filename, with **Open PDF**, **Open folder** and **Create another** actions.

Working drafts are saved in this browser on this computer. Closing an edited form offers to keep or discard the draft; reopening offers to restore it. A working draft does not change Excel or reserve a COD. Clear browser data only after finishing saved drafts. A separate **Save to local register** action persists custom certificate drafts in application data.

Changes from another browser tab, Excel, a newly added PDF or a changed template invalidate the open editor's reviewed snapshot. The app keeps the entered form and asks for a refresh rather than saving stale details. Warning acceptance is tied to the exact reviewed details and warnings; staff must enter a short reason before continuing. Neither **Create anyway** nor a custom COD allows overwriting an unrelated PDF.

New PDFs are structurally checked before publication. A file with the right name but damaged, encrypted or unreadable PDF content does not count as an available certificate. For certificates issued by this version, the current file must also match its saved issuance receipt and content hash; an older same-number copy cannot hide a missing latest version. Blank and N/A rows remain excluded from automatic missing-PDF recovery and bulk creation.

Both Excel-backed and custom reissues archive superseded PDFs after the source update succeeds. Durable save journals and recovery snapshots under `Data/Issuance History` and `Data/Issuance Recovery` cover interrupted saves. On the next check, matching source hashes prove whether to complete or roll back the interrupted operation. Unknown external changes block further creation and preserve copies for IT review. Bulk creation stops if history completion is pending, so a later workbook write cannot invalidate recovery evidence. A dashboard refresh failure after successful saving is reported as a refresh problem, not as a failed certificate.

**Activity history** shows issued numbers, before/after details, template fingerprints, the Windows account, accepted warnings and reasons. It is a local operational history, not independently authenticated approval or tamper-proof compliance evidence. Manager approval of certificate wording and operator acceptance remain required before rollout.

## Data and backups

- `Plant Workbooks` and `settings.local.json` are local user data and are not part of this repository.
- Workbook backups are saved under `Data\Workbook Backups` before the app updates COD/date/status fields or alert formatting.
- The latest readable work list is under `Data\Reports`. Certificate PDFs are kept together in the top-level `Certificates` folder. Old run reports and preview artifacts are archived under `Data\Archive`.
- The three current designs and their original PowerPoint sources are in `App\templates\certificate-types`. The older `App\templates\Certificate of Recycling.png` remains compatible with older installs.
- Custom register backups, PDF previews and copied annexures remain under `Data`. Custom draft records retain their CODs after a PDF is removed, so later certificates do not reuse those numbers.

In **Settings → Backups**, choose a separate backup destination and create a verified workspace backup. The default local destination is convenient but does not protect against drive failure. Backups include configured workbooks, application data, certificate PDFs, templates and issuance recovery state; the app checks hashes and its trusted catalog before restoring. Close Excel first. Restoration requires explicit confirmation and creates a pre-restore backup.

Restoring returns active PDFs and issuance receipts to the selected point in time. Newer files absent from that snapshot are moved recoverably to `Data/Archive/Restore Extras`; previously observed CODs stay reserved to prevent reuse. Restore does not erase unrelated files or permit adopting an unrelated PDF as the restored source's certificate. Backups are not encrypted: restrict access to their destination. Preserve the ZIPs together with `Data/Backups/trust.key` and the latest `Data/Backups/catalog.json` separately. Restores are bound to this workspace and its configured workbook paths; moving to a replacement PC requires an IT-assisted migration.

Complete [the production acceptance checklist](App/docs/PRODUCTION_CHECKLIST.md) before using real certificates without supervision. The Windows reminder, backup destination on a separate drive, certificate wording and staff trial need verification on the operator's actual computer.
