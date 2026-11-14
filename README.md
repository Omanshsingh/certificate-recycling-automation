# Ecoreco Certificate Automation

This local browser app reads the two plant Excel workbooks and generates certificate PDFs. It runs only on this computer at `127.0.0.1`; it has no Outlook integration and does not create drafts or send messages.

Manual Excel editing remains supported. Refresh and preview can read saved data
while Excel holds a workbook open; the app explains when an Excel update is still
pending. Actual saves require write access. Duplicate Serial Numbers are reported
with their worksheet rows and block certificate creation, while **Open in Excel**
remains available for repair. Company lists use ascending inward dates. Supported
workbook layouts are reordered as complete inward/asset blocks on app saves and
writable refreshes, preserving Serial Numbers and existing CODs.

Before issuance, you can enter or correct inward details directly in Excel. Keep
each SN unique within its plant and each Form 6 number unique across both plants.
Leave workbook headings, weight formulas and certificate tracking to the app.
Save and close Excel, then refresh to pick up your edits.

Issued certificate details stay protected in Excel, including their material and
linked HDD rows. The four follow-up fields remain editable. Use **Review & edit**
for certificate corrections and **Mark Completed** after sending. To update an
inward marked **N/A**, use **Review & edit → Save to Excel**. The inward keeps its
**N/A** status. In **Companies**, **Mark N/A** skips an unissued Excel inward.
**Undo N/A** checks the inward details and restores **Ready** or **Alert**. For an
unissued app-only certificate, **Undo N/A** restores its previous status.

The [stability plan](App/docs/STABILITY_PLAN.md) describes the longer-term local
record store and reconciliation design that supports continued manual Excel edits.

## Everyday use

**Settings → Next COD number** sets the starting point for the next new certificate.
The app still checks used numbers across the workbooks and certificate history.
Leave the field blank to continue automatic numbering.

Use **Upload PDF** beside **Refresh** for a certificate created outside the app.
Choose an Excel inward or save the certificate in the app only. The form fills in
the printed COD when it can read it. Otherwise it uses the inward's existing COD
or suggests the next available number. If the PDF already has a number, enter
the printed COD. Confirm the type and issue date. For app-only records, also
enter the company details.
The app keeps the original PDF in **Certificates → Manually Added** and records
its status as **Generated Externally**. An uploaded Data Destruction certificate
uses its own record and COD. PDFs already in **Manually Added** can still be
linked from **Saved certificates → Details**. Mark the record **Completed** after
sending it. **Generated Externally** is recorded by the app when you attach a PDF.

In **Review & edit**, uncheck **Recycling certificate needed** to mark only the
Recycling certificate N/A. The equivalent checkbox uses the selected certificate
type for Medical or International inwards. Use **Open Data Destruction (HDD)** to
create the linked HDD certificate first or on its own. Its COD and date are
recorded separately in Excel.

Use the trash icon beside an inward in **Companies** to delete an entry entered by mistake, up to the point when a certificate is generated. An assigned number alone does not hide the action. Generated or Completed entries, including entries with an issued HDD certificate or missing issued PDF, cannot be deleted. **Settings → Advanced → Search entries to delete** follows the same rule. Pending, ungenerated HDD requests can be included in the reviewed deletion or retained as standalone records.

Review the company, plant and materials, tick the confirmation, then remove the inward. The app keeps a workbook backup, physically removes the entire inward block, shifts later rows up, and closes numeric SN gaps. Text SNs and existing certificate numbers remain unchanged; receipt and HDD links follow renumbered inwards. Cancel does not delete anything, and repeated requests cannot remove a replacement entry.

Refresh also removes wholly empty row boxes left by older deletions, with a workbook backup. Blank rows within populated merged client blocks remain intact. Workbooks with unsupported formulas or objects retain their rows and show a maintenance warning for review in Excel.

Dates throughout the app and the Excel register display as **DD/MM/YY**. Use the calendar icon next to editable dates, including Data Destruction and each device receipt. Excel stores real dates, and register details including SN use top-left alignment. App saves and writable refreshes wrap register text and make rows taller where needed, including merged client cells. You can also resize rows directly in Excel. New-client entry stays open when clicking outside the dialog. Add client and Review & edit use the same materials calculation controls: per-piece weight × quantity = line weight, with direct weight entry supported for quantities such as **1 Lot**.

When an inward's HDD certificate is generated, its number and date are written on the next line beneath the recycling COD and COD Date in the same inward. The tracking lines use the light internal asset borders. Existing recycling numbers, dates and statuses stay unchanged. The workbook, HDD PDF and local tracking are saved with recovery snapshots; close Excel before this save. Refresh also fills this tracking line for older linked HDD certificates.

COD No and COD Date remain separate cells on each material row so the light internal lines stay visible. Refresh repairs older vertical merges in these two columns; new clients, material edits and certificate saves apply the same formatting.

The **Dashboard** shows certificates created for the selected dates, today's
count and this week's count, inwards received, and typical days from inward to
first PDF. A single daily line shows the selected dates, including in the
90-day view. The top card separately labels the previous period's total. Certificate
types are **Recycling, Data destruction, Siemens and Sita**, with a separate
plant comparison. Plant and date filters apply to these charts and totals.
Ready to create, Check details, PDFs to send, pending HDD certificates and
entries waiting over 30 days open the matching work list; these pending counts
cover all dates. Large numbers and labels of at least 16px keep the wide desktop
view readable. Loading errors offer Try again and hide stale counts.
Hover over the chart to see one date and the number of certificates created that
day. Arrow keys, Home and End read the chart when it has keyboard focus. Plant
cards and certificate-type bars show their counts and shares on hover or focus.
Click a plant card to filter the dashboard; click the selected plant again to
show both plants. Escape dismisses an open popup.

Saved PDFs with no inward or app record are included using their certificate
issue dates and type. The `_HDD.pdf` ending identifies data destruction (older lowercase names are also recognised). A PDF's
copy/modified date is never treated as its certificate date. The app counts each
COD once across history, records and folder copies. Unknown dates/types produce
a short message. PDFs with no known plant contribute to **Both plants** and are
listed as having no plant, rather than being assigned to a plant by guesswork.
Imported PDFs never create an inward or an assumed sending status. Dashboard
reads do not change Excel, create PDFs or allocate numbers. There is no export
feature, counting-instructions panel or help popup.

**Open in Excel** beside a client opens its configured plant workbook and selects the entire client section, including material rows. The app checks the live Serial Number rather than relying on a saved row number. If the Serial Number is missing or repeated, it opens the workbook for repair and explains that a unique record was not selected. The same button appears beside saved and newly generated PDFs linked to an Excel client. App-only and independently added PDFs have no Excel button. Selection does not save Excel or permanently color cells; save Excel after making edits, then refresh the dashboard. Close Excel when the app needs to write changes.

Use **Add client** to prepare a new Excel inward without generating a PDF. Choose the plant, enter the company details, and click **Add to Excel**. Choose the certificate type on its generation preview. The app inserts a formatted inward block in date order, including all its assets, material weight formulas, an existing total column, vertical merges and dropdowns. New entry cells use top-left alignment. Automatic numeric SNs follow that plant's date order: a backdated entry takes its place in the sequence and later entries move up. Existing CODs stay the same. Saved certificate, data destruction and interrupted-add links follow the changed SNs through metadata saved in the same workbook transaction; original issuance history stays unchanged. Text SNs are kept. The workbook is backed up before saving. Refresh marks incomplete or invalid source details Alert and complete details Ready. Complete Siemens identifiers on the certificate preview. Repeating an interrupted add request uses its original operation record to prevent duplicate entries. Save and close Excel before adding. Unsupported table layouts or cross-record formula dependencies produce an explanation instead of a partial save.

Keep each plant's existing register layout. The app keeps **Comments**, **Transporter/Vehicle Details**, **Docket No./Vehicle No** and **Weight Slip/Weight Scale** directly before the final **COD No**, **COD Date** and **Status / Certificate Issued** columns. Missing columns in that final group are added, with existing values, formulas, merges, dropdowns and status colours retained. It does not add per-piece weight, reference, certificate type, JSON fields or a certificate weight column. Plant 1 can use its existing line-weight column, while Plant 2 uses quantity × per-piece weight where those columns exist. New clients retain the original fonts, row heights, merged client sections and border lines, with totals calculated in the existing total column. Explicit quantities such as **1 Lot** are supported. Optional certificate settings without an existing column are saved in the workbook's document properties and remain available in the app. Existing optional columns remain supported.

1. Open **Prepare Certificates.bat**. The local dashboard opens in your default browser.
2. Use **Settings** to choose the Plant 1 and Plant 2 workbooks. In Excel, mark approved rows **Ready**, use **N/A** to skip a certificate, and save and close Excel.

To start a plant register from scratch, choose **Create template** for that plant in Settings. The app makes a blank `.xlsx` with the standard headings, readable client-block formatting, a Certificate Issued dropdown and colors. The total-weight formula is added when a client is entered through **Add client**. Choose a new filename; an existing workbook is never overwritten. The new file becomes active only after you click **Save** in Settings. **Choose file** continues to connect an existing workbook without replacing its layout.
3. Search for the company and choose **Review & edit**. Change the company details, received date, materials or certificate weight and check the PDF preview beside them. Choose Recycling, Medical or International. **Data Destruction needed** adds a separate pending HDD task when you save; it does not create another PDF or assign an HDD number. Before the first PDF exists, **Save to Excel** lets you save corrected inward details without creating a certificate. Excel changes require **Yes, update Excel**. When creating a PDF, **No, keep Excel details** uses the edits only on the certificate. COD is assigned when the PDF is created. **Create several** shows a searchable checklist of Ready clients, with nothing checked initially. Search by company, SN, reference or city. Selections stay checked across searches; the footer shows the total and any selected clients hidden by the search. **Select shown** changes only matching clients. Only checked clients receive PDFs, and warnings still need review. Creating the chosen PDFs keeps the other clients visible in the company list.

   **Details to finish** lists inwards with missing Comments, Transporter / vehicle details, Docket no. / vehicle no., or Weight slip / scale weight. All four fields are visible in Add client and directly in Review & edit. Enter them when known; any blanks appear in this separate view, including after Completed. They do not block certificate creation or join urgent certificate reminders. Saving only these four fields asks **Yes, update Excel** and keeps the PDF and certificate status unchanged. Status buttons and the additional company filters stay visible.
4. Find PDFs in the top-level **Certificates** folder. Names use `{COD}_Company name_City, Country.pdf`. Underscores separate sections; spaces inside company and location names are preserved. The COD matches the number written to Excel; the plant is tracked in the source record.
5. Choose **Open PDF** to check the saved certificate and send it through the normal company process. Then choose **Mark Completed** beside that company; Excel records **Completed** and reminders stop.

The app writes a newly assigned COD number and date back to Excel. A COD or date already entered by the user is respected. Reference No. may be blank on a draft but is required before issuing a Recycling, Sita or Siemens certificate. Form 6 No. is the required manifest identifier for Standard Recycling certificates. If a row is deliberately changed back to **Ready**, the current row values control the next PDF generation. If COD is blank, the next number is chosen from both plant workbooks and the existing certificate PDFs so the sequence is shared.

An inward can be reviewed and edited before or after its first certificate. The plant and SN stay fixed. The editor asks before saving changed details to Excel, and the save endpoint checks that consent and the reviewed file version. Saving inward details alone creates no PDF and reserves no COD. PDF creation records the number, date and Generated status; the chosen template and printed details are retained in the app's issuance history.

After generation, **Review & edit** opens the certificate details and preview directly. Choose **Save corrected certificate** when finished. The prompt **Also update these details in Excel?** shows the changes. **Yes, update Excel** saves them in the selected inward too; **No, keep Excel details** uses them only on the corrected PDF. Both record final COD, date and Generated status. Reopening uses the last successful certificate snapshot, including earlier PDF-only corrections. An issued inward can be updated together with its corrected PDF. Previewing, warnings and cancelling do not save Excel changes. Failed workbook updates roll back PDF publication. Bulk creation skips empty statuses and N/A.

When the selected certificate type differs from the Excel inward's saved type,
the save prompt names both types and asks what should happen to the original inward.
**Yes, update Excel** saves the reviewed changes only to the selected inward and
creates its PDF together. Switching to Medical or International clears the old
Form 6 value because those certificates do not use it; the change appears in the review.
**Delete Excel inward; keep PDF** asks for explicit confirmation of the plant,
SN and company, then removes that inward and all its materials from Excel while
keeping the PDF as a local certificate. Other entries keep their Excel positions.
A workbook backup is saved first. The current COD, issuance history and linked
HDD requests are retained; HDD switching continues through the local certificate.
Failed saves restore the previous PDF and local register, and interrupted saves
recover from the workbook commit. Cancel leaves the original entry untouched.
Same-type corrections retain the existing PDF-only choice. Standalone Medical
and International certificates still need no Excel inward. No register columns are added.

Open **Review & edit** for any change to an issued certificate. Saving a correction keeps its current COD and archives the previous PDF in issuance history. Certificate numbers cannot be changed through the editor.

Three tabs keep everyday work together: **Dashboard** for totals and pending work, **Companies** for Excel/local records, and **Saved certificates** for files in the certificate folder. Companies shows company, plant, received date, status and next step, with ten companies per page. Search by company, COD, Form 6 or reference. **Review & edit** opens Excel facts and preview directly; **Review & edit** opens issued details directly. Alerts appear on the row and in the preview. Saved certificates uses compact, spaced rows with direct **Open PDF** links, search by company/number/reference, plant and certificate-type selection and sorting. **Details** opens the original filename and references. **Past numbers** stays separate from current files; inspection alerts and automatic number protection remain available. Certificate type, manual identifiers, corrections and attachments are on the preview page. Optional quantity calculations remain available. Status names match Excel: **Ready**, **Alert**, **Generated**, **Generated Externally**, **Completed** and **N/A**. An unselected status is an empty white box, with no Blank dropdown option. On refresh, older labels are converted to the current name; an old literal Blank is cleared. Refresh checks source details and sets Ready or Alert for unissued records. N/A and issued/completed statuses remain separate from automatic readiness. Overdue never replaces a certificate's true status. Floating-point tails are cleaned up on screen; untouched numeric fields retain their original values when submitted.

While the dashboard is open, it checks for saved Excel changes every 15 seconds and when you return to the page. Save and close Excel to allow the app to write Ready. Background checks wait while a dialog or operation is active; Refresh also checks immediately. Required source details include the SN, company, a valid received date and material descriptions with valid positive total weight. Form 6 may be left empty when adding an inward and is required before creating a Recycling certificate. International (Sita) and Medical (Siemens) certificates do not require or print it. Reference may also be entered later, but is required before creating Recycling, Sita or Siemens certificates. HDD certificates do not use references. COD/date and Siemens manual identifiers do not block automatic Ready. Formula, merged or locked protected status cells are preserved. Checking never creates a PDF.

If Excel changes during a review, the rejected action refreshes the company list automatically. An outdated deletion review closes so it can be reopened against the current rows. Entered certificate details remain in the editor; close and reopen the entry to review the latest Excel facts before saving. The app never automatically repeats the rejected save or deletion.

The everyday screen keeps search, plant selection and quick **All / Ready / Alert** choices visible. **More filters** reveals other statuses and certificate types; active selections remain visible in its heading. Clearing company search/filters keeps the selected plant. **Summary & analytics** reveals the status totals. The Alert view includes day-25 follow-ups without approving unfinished inwards or N/A rows. **Create several** lists eligible Excel rows in the selected plant, regardless of search, status or type filters. Use **Review & edit** to fill missing details for any company that is absent from the checklist. Close the batch dialog to cancel. If a batch only partly succeeds, the result shows what saved and which companies still need a certificate. Resolve the listed issues, refresh and select the unfinished companies for retry.

Use **New certificate** when the details were not entered in Excel first. **Where details are saved** contains **Save details in**, the plant/workbook selection and SN. A new numeric ID is suggested from the selected plant when all its IDs are numeric; you can change it. A new Excel record is saved only when PDF creation succeeds, and an existing ID is rejected rather than replaced. Under the materials section, enter a material and its weight, then use **+ Add another material** for additional lines. The app calculates the total and previews it automatically; optional quantity-based calculations remain available. Adding materials in Review & edit inserts the required rows before the next client, moves that client's complete block, and extends the selected inward's merges and supported total formulas.

**Certificate Weight** can be changed before or after the first PDF. Enter the weight to print, or clear it to use the Excel weight or material total. The Yes/No prompt controls whether changed details update Excel. Cancel asks **Discard these changes?** only when there are unsaved edits, including an unsaved Data Destruction choice. Edits autosave in this browser for recovery; custom records also have optional **Save for later** tools.

## Certificate types and custom certificates

Choose **Standard COD**, **Sita International**, or **Siemens Medical** on the first preview page, independently of the supplier name. Selecting Sita/Siemens applies its official printed company wording without renaming the Excel supplier; the printed line can be adjusted here. Their printed designs omit Form 6; Excel inwards still require the source manifest number. Siemens opens Model name, Model serial number and X-ray tube serial number immediately. Site location fills from Excel Location and stays locked for first generation when present; a missing location can be entered manually. Corrections may change it later. The original PowerPoint designs, artwork, fonts and fixed wording remain in use.

For certificates without an inward entry, choose **This app only (no Excel row)** under **Where details are saved → Save details in**. This uses a separate local register under `Data/Custom Certificates.json` and has no plant assignment. Enter the details, check the automatic preview and choose **Create PDF**. Optional **Save for later** tools save the custom record without generating. Companies and Saved certificates label these records **App only**, with a separate filter; they are excluded from Plant 1 and Plant 2 totals. Custom certificates share the global COD sequence and retain their assigned number if the PDF is deleted. International forms do not show Form 6. HDD certificates explicitly linked to an Excel inward retain its real plant. Startup backs up and clears legacy plant assignments on standalone local records while preserving existing PDFs, CODs and immutable issuance journals.

Use **Company line on certificate** to change only the printed name, such as Sita Laptops or Sita Metals, without replacing the supplier name in Excel. Designs, standard wording, table labels, signatory labels and footer are fixed; staff cannot edit them or move/resize the layout. Older saved wording/layout overrides are ignored when creating new PDFs, while existing PDFs are left untouched. **Weight as printed** supports exact units or weight descriptions, while the numeric certificate/net weight remains available for the register. Decimal weights retain their precision.

Missing dates, manifests, weights and model identifiers produce warnings with their consequences. Staff can explicitly choose **Create anyway**. An invalid template selection, an unreadable file, a protected workbook cell or an unavailable attachment still needs correction because the app cannot safely produce/save the requested result. If a workbook is unavailable, a custom certificate can still be created after accepting the warning to verify its COD against that workbook before sending.

Use **Attach PDF list** for an existing annexure, such as an item/serial-number list. The app copies the selected PDF into its local data folder and appends its pages to both preview and final PDF. It never reuses the example annexures automatically.

For PDFs verified against the app's saved issuance history, attached certificate numbers stay reserved but are not shown as the main certificate's number or reported as conflicting numbers. The PDF is not rewritten. New HDD certificates take a new number from the shared COD sequence; correcting an issued HDD keeps its existing number.

When these clients enter the inward register, select that inward and choose the template directly on its preview page. First issuance retains the choice in the app's history. A later correction saved with **Yes, update Excel** can persist `Certificate Type` and `Certificate Fields` columns. No supplier-name matching rules are needed.

On workbook refresh or when the dashboard checks its saved work list, a row marked **Generated** or **Completed** with no matching PDF in `Certificates` returns to **Ready** in Excel, retaining its COD and date for recreation. A missing COD does not prevent recovery; a new number is assigned only when the PDF is created. Matching uses COD and plant, so changing a company name does not incorrectly mark an existing PDF missing. Archive copies do not replace the required current PDF. Refresh never generates a PDF. If Excel is locked or protected, the app reports the failed status update and asks you to close or correct the workbook before retrying. Custom records follow the same missing-PDF rule and track their current PDF filename; when a custom certificate is regenerated with a different name, the superseded PDF moves to `Data/Archive/Superseded Certificates` and remains recoverable. **Completed** confirms sending only while the current PDF remains in `Certificates`.

Every source or layout issue is shown as a warning. Before generation, the app lists the warnings and asks for confirmation. Continuing can produce a PDF with a blank, fallback, or manually corrected value, and the row is still marked Generated, so verify the PDF and workbook before sending it.

## Requested data destruction certificates

To remove a linked request before creating its PDF, clear **Keep this Data Destruction request active** in the HDD editor and choose **Save request**; it leaves pending work while retaining its saved details. If an HDD PDF was already issued, confirm removal; the PDF and number are preserved.

Open **Review & edit**, check **Data Destruction needed**, and save. This marks an additional HDD certificate as needed without generating it or allocating a number. **Review Data Destruction** appears inside the same panel to open the saved task later. The app copies the company/location and links it to the plant and SN. Enter the actual destruction date, received dates and whole device counts when preparing its PDF. Device counts are never inferred from material weights. Reopening reuses the existing task. The HDD gets its own number from the shared sequence when its PDF is created.

For an international client without an inward, choose **New certificate → Secured Data Destruction** and enter its company/location and device receipts. These certificates use the local register and the shared certificate number sequence. They do not require Form 6 or weight. Incomplete device/date details must be corrected before final creation.

COD and HDD use **one shared number sequence**, for example COD 0204 followed by
HDD 0205. Newly created HDD PDFs are saved alongside automatic certificates in
`Certificates/Auto Generated`, named `{number}_Company_City_HDD.pdf`. This applies
to standalone clients as well as inward-linked requests. Existing PDFs retain
their paths until corrected or recreated. HDD numbers found in saved files also
reserve that number in the shared sequence. Independently created HDD PDFs can
be placed in this folder and counted by the dashboard without an inward entry.

The HDD layout follows the filled reference: bold company/location and values,
dates printed as **DD/MM/YYYY**, and the supplied signature on the details page.
The destruction date, receipt dates and certificate issue date remain separate.
The receipt editor shows a running device total and flags incomplete entries.
The details page holds at most nine device receipts. The editor prevents a tenth row, and final creation rejects excess rows without dropping entries. No certificate number or extra heading is added above the supplied table.

The supplied `HDD COD.pptx` artwork is included as a two-page template. Both pages appear in the preview. At most nine receipt rows can be entered. Optional PDF lists follow the certificate pages.

A banner shows outstanding data destruction requests, including PDFs awaiting sending. Click it to show those requests. Send the PDF through the usual company process, then **Mark Completed** on that request. Completing the recycling certificate does not complete the data destruction request. A missing or damaged current PDF returns the request to pending. The payment note is a follow-up reminder; the app does not change invoices or payment records.

## Manually created PDFs and used numbers

Staff may create a certificate outside this app and put its PDF in `Certificates/Manually Added`. Use the `{COD}_Company name_City, Country.pdf` naming format, with the actual COD at the start. Underscores divide the sections; spaces remain inside names. The app checks this folder before each creation, even when that certificate has no Excel row. It reads CODs from filenames and labelled PDF text, and references from readable PDF text. Automatic CODs start above the highest observed COD; manually entered used CODs require confirmation to switch to a new automatic number. Confirmation never authorizes overwriting someone else's certificate.

The **Saved certificates** tab shows current folder files, including manually added PDFs, with search, plant filters and direct Open PDF links. There are **two folders** in `Certificates`: **Auto Generated** contains every PDF created by this app, whether from Excel, New certificate, or a data destruction request; **Manually Added** contains PDFs made outside the app. Startup safely moves files from the former third folder into Auto Generated and removes that empty folder. Existing links and issuance checks keep working through a recorded move; original issuance history remains unchanged. Copies with the same filename are kept under distinct names, and a move backup is stored in Data. Put outside-app PDFs in `Manually Added`; existing PDFs left directly in `Certificates` continue to be checked. Click **Details** for the original filename and references. **Past numbers** is separate, so removed files do not clutter the saved list. Filename/printed-COD mismatches reserve both numbers and need review. Scanned, encrypted, oversized or unreadable PDFs may only have a usable filename COD; their references need manual verification. This is not OCR and does not automatically import unverified PDF details into Excel.

References are entered manually and required before creating Recycling, Sita or Siemens certificates. HDD / data destruction certificates do not use references. The same reference may be reused on other certificates without an alert. References are never automatically generated or incremented. They may be left blank when adding an inward or saving a draft, and entered later when creating the certificate.

`Data/Certificate Inventory.json` remembers observed identifiers after a PDF is moved, replaced or deleted; backups are under `Data/Certificate Inventory Backups`. Deleting a PDF does not free its observed COD. Previewing does not reserve a new COD. The folder is checked again immediately before publishing, so an external file added during rendering cannot silently be overwritten.

## Reminders

The default reminder time is **11:00 AM**. In Settings, turn on the daily reminder, choose its time and click **Save**. Changing reminder preferences updates the Windows task as part of that save; unchanged preferences do not reinstall it. Optional **Reminder options → Save & repair reminder** retries a task update. If Windows cannot confirm the update, the app clearly reports that Settings were saved but the reminder needs repair. The task checks both workbooks at that time and after sign-in catch-up; it runs only in the signed-in Windows session while the computer is awake.

Daily reminders also include pending local HDD requests, missing local PDFs, and local certificates at least 25 days after receipt. A workspace with no Excel files connected can enable these reminders. Missing-PDF alerts stay active until a valid replacement exists. The dashboard recalculates ages on the next check after midnight. A new issue on the same record can notify again; **Dismiss for today** still suppresses that day's reminders.

The reminder window is smaller and does not force keyboard focus or stay above other apps. Closing it removes the window immediately; Windows task cleanup finishes in the background. If another operation is busy at reminder time, the check retries for up to ten minutes. Installed tasks also support failure retries after **Save & repair reminder**.

Automatic PDF previews are disposable: old inactive previews expire after a day, and a 256 MiB cleanup budget removes older superseded groups when needed. Active editor previews are retained, with a short grace period for recently rendered files. Certificate PDFs and issuance recovery files are outside this cleanup.

At inward age 25, the **Supplier Name** cell turns red in Excel and the app shows the row for follow-up. A missing, invalid or future inward date is highlighted amber in the date cell instead, so date problems are not confused with an overdue reminder. The SN and Reference No. remain unchanged. The reminder continues through day 30 and becomes **Overdue** after day 30. It stops when the row is **Completed** or **N/A**. “Completed” is the user’s confirmation that the certificate was sent; the app cannot detect delivery. Missing or invalid inward dates and certificate data are shown as attention items because their age or PDF cannot be safely determined.

Daily reminders only check and notify. They never generate PDFs on their own. Generating PDFs is always a user action in the app.

## Setup

- Run **Setup\Setup Once.bat**. It finds supported 64-bit Python even when it is not on PATH; if none is available, it downloads Python 3.13.17 from python.org, verifies the installer SHA-256, and installs it for the current Windows account. It then installs the app's locked packages.
- Run **Prepare Certificates.bat** to start the local server and open the dashboard. Use **Close dashboard** at the top when finished.
- Save both workbook paths in Settings. Save and close Excel before checking or generating because Excel may lock an open workbook.
- To enable reminders, turn on the reminder in Settings and click **Save**, or run **Setup\Create Daily Reminder Task.bat**.

Dependencies: 64-bit Python 3.11+, `openpyxl`, `reportlab`, `pypdf`, `fonttools`, `pypdfium2`, Pillow, and Tkinter (used for the Windows file picker). PowerPoint is not required to create certificates: the supplied layouts are included as PDF templates. The local dashboard uses Python's standard library; no web framework is needed. Internet is required once if Setup must download Python and install packages. Outlook and pywin32 are not needed.

Setup uses `App/requirements.lock.txt` for reproducible package versions. `pypdfium2` and Pillow render certificate pages in the editor, so preview does not depend on the browser's built-in PDF viewer. The **Attached PDF** section shows the selected list with **Previous / Next** buttons and a page count. Attachment images load only as requested, using the same PDF snapshot as the full preview. Failed page loads offer a retry; changing or removing the attachment clears the previous images. The full PDF remains available. A normal installation needs internet once to download Python when missing and install packages, not to use the app afterward.

## Safe creation and recovery

The everyday flow is **find company → check preview → create PDF**. Editing is optional; previews update automatically and never reserve a number or write an Excel row. Issues appear as readable cards: a short heading, what will happen and **What to do**. Duplicate COD warnings show the old and replacement numbers. Company rows show short issue labels. Preview, confirmation, file alerts and accepted history warnings use the same wording. **Original details** retains the complete technical message and filenames; unfamiliar issues remain visible in full. Exact before/after changes appear beside the preview and in the Excel save prompt. The success screen shows the actual issued number, with **Open PDF**, **Open folder**, **Create another** and optional filename details.

Working edits autosave in this browser on this computer. Cancel asks only whether to discard changes; No returns to editing and Yes discards them. Reopening after an interruption offers recovery. These browser edits do not change Excel or reserve a COD. Clear browser data only after finishing them. Custom certificates can also be saved to the local register without creating a PDF.

Success feedback appears after saving. Errors stay visible until dismissed. After a lost connection, **Check connection** rereads saved state without repeating creation or saving again; the open form is preserved. Check Saved certificates before retrying an uncertain creation. Settings recovery rereads committed preferences while keeping pending inputs.

Changes from another browser tab, Excel, a newly added PDF or a changed template invalidate the open editor's reviewed snapshot. The app keeps the entered form and asks for a refresh rather than saving stale details. Warning acceptance is tied to the exact reviewed details and warnings; staff must enter a short reason before continuing. Neither **Create anyway** nor a custom COD allows overwriting an unrelated PDF.

New PDFs are structurally checked before publication. A file with the right name but damaged, encrypted or unreadable PDF content does not count as an available certificate. For certificates issued by this version, the current file must also match its saved issuance receipt and content hash; an older same-number copy cannot hide a missing latest version. Inwards with empty statuses and N/A remain excluded from missing-PDF recovery and bulk creation. Complete empty statuses first become Ready on the next workbook check.

Both Excel-backed and custom reissues archive superseded PDFs after the source update succeeds. Durable save journals and recovery snapshots under `Data/Issuance History` and `Data/Issuance Recovery` cover interrupted saves. On the next check, matching source hashes prove whether to complete or roll back the interrupted operation. Unknown external changes block further creation and preserve copies for IT review. Bulk creation stops if history completion is pending, so a later workbook write cannot invalidate recovery evidence. A dashboard refresh failure after successful saving is reported as a refresh problem, not as a failed certificate.

**Activity history** shows issued numbers, before/after details, template fingerprints, the Windows account, accepted warnings and reasons. It is a local operational history, not independently authenticated approval or tamper-proof compliance evidence. Manager approval of certificate wording and operator acceptance remain required before rollout.

## Data and backups

- `Plant Workbooks` and `settings.local.json` are local user data and are not part of this repository.
- Workbook backups are saved under `Data\Workbook Backups` before the app updates COD/date/status fields or alert formatting.
- The latest readable work list is under `Data\Reports`. Certificate PDFs are kept together in the top-level `Certificates` folder. Old run reports and preview artifacts are archived under `Data\Archive`.
- The four current designs and their original PowerPoint sources are in `App\templates\certificate-types`. The older `App\templates\Certificate of Recycling.png` remains compatible with older installs.
- Custom register backups, PDF previews and copied annexures remain under `Data`. Custom draft records retain their CODs after a PDF is removed, so later certificates do not reuse those numbers.

In **Settings → Backups**, choose a separate backup destination and create a verified workspace backup. The default local destination is convenient but does not protect against drive failure. Backups include configured workbooks, application data, certificate PDFs, templates and issuance recovery state; the app checks hashes and its trusted catalog before restoring. Close Excel first. Restoration requires explicit confirmation and creates a pre-restore backup.

Restoring returns active PDFs and issuance receipts to the selected point in time. Newer files absent from that snapshot are moved recoverably to `Data/Archive/Restore Extras`; previously observed CODs stay reserved to prevent reuse. Restore does not erase unrelated files or permit adopting an unrelated PDF as the restored source's certificate. Backups are not encrypted: restrict access to their destination. Preserve the ZIPs together with `Data/Backups/trust.key` and the latest `Data/Backups/catalog.json` separately. Restores are bound to this workspace and its configured workbook paths; moving to a replacement PC requires an IT-assisted migration.

Complete [the production acceptance checklist](App/docs/PRODUCTION_CHECKLIST.md) before using real certificates without supervision. The Windows reminder, backup destination on a separate drive, certificate wording and staff trial need verification on the operator's actual computer.

## Manual references, unique identifiers and attachments

Reference No. is required for Recycling, Sita and Siemens certificates and is always entered by the user. HDD / data destruction certificates do not use a reference number. The app never creates, increments or replaces references. SN numbering is separate for each plant, so the same SN may occur in different plants. An SN must be unique within its plant. COD and Form 6 numbers must be unique across both plant workbooks and the local certificate register. Duplicate Form 6 or assigned COD values block creation even when warnings are accepted; correct those values in Excel before retrying. A linked data destruction request refers to its original inward and does not create another inward SN or manifest.

Attach PDF list retains a validated local copy and appends every page after the certificate pages in preview and final output. Reopening retains the filename and page count. Replacement uses only the newly selected list; Remove list omits it on the next save. Invalid, empty or password-protected PDFs are rejected, and unavailable or damaged retained files block publication.

The Companies table shows SN beside Company. Find SN and Received date narrow the list; Sort by supports numeric SN, received date in either direction and company A–Z. HDD records open Excel when their inward link or a single exact company/location match identifies the correct entry. Add client assigns SN automatically, places Weight after Per Pcs × Quantity, and displays save errors at the top of the form.
