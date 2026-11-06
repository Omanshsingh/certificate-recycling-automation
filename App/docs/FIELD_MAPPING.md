# Certificate field mapping

| Certificate field | Source or rule |
|---|---|
| Certificate No. | User-entered COD No., or next number after values in both workbooks and existing certificate PDFs. Written back to COD No. |
| Certificate Date | Assigned automatically when the PDF is generated and written back to COD Date. Existing certificates keep their original date when corrected or recreated. |
| COD allocation | Uses the current values in both plant workbooks and CODs found in every PDF under `Certificates`. If a PDF has no identifiable COD, creation pauses with instructions to rename it with its COD or move it outside the certificate folders. |
| Date of Material Receipt | Inward Date. |
| Weight | Certificate Weight (KG) / Net Weight (KG) when populated; otherwise the sum of the inward batch's item weights. Displayed to one decimal place in KG. |
| Reference No. | Entered manually and required before creating Recycling, Sita or Siemens certificates. HDD certificates do not use it. References may be reused without an alert. May be left blank when adding an inward. |
| Manifest No. | Form 6 No. (optional when adding an inward; required for Recycling. Medical and International do not use it; changing to either clears the old Form 6 in the reviewed Excel update.) |
| Supplier / vendor | Supplier Name plus Location, uppercase and centered in the certificate template. |
| SN | Assigned separately for each plant. Add client places automatic numeric SNs in received-date order and shifts later SNs, keeping saved certificate and data destruction links with the correct company. Supplier Name and SN identify the row in Excel and the app work list. |
| Status | Ready is eligible for PDF creation. Alert needs review and is never an automatic creation trigger. Generated means the app created the PDF from an inward row; Generated Externally identifies a PDF prepared outside the app. Completed is the user's confirmation that it was sent. N/A is excluded; Blank remains blank and needs a decision. Warnings can be explicitly accepted before generation and remain flagged in the report. |

PDF names use `{COD}_Company name_city.pdf`. The leading COD matches the number written to Excel; the plant remains tracked in the source record.

Inward age 25+ through day 30 is a follow-up reminder; after day 30 it stays overdue until Completed or N/A. The app does not infer recycling completion or email delivery.
