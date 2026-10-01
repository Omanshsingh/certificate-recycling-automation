# Certificate field mapping

| Certificate field | Source or rule |
|---|---|
| Certificate No. | User-entered COD No., or next number after values in both workbooks and existing certificate PDFs. Written back to COD No. |
| Certificate Date | User-entered COD Date, or the date the PDF is generated. Written back to COD Date. |
| Date of Material Receipt | Inward Date. |
| Weight | Certificate Weight (KG) / Net Weight (KG) when populated; otherwise the sum of the inward batch's item weights. Displayed to one decimal place in KG. |
| Reference No. | Optional Reference No. column. Blank is allowed. |
| Manifest No. | Form 6 No. (the required manifest identifier; missing values remain a warning that must be explicitly accepted) |
| Supplier / vendor | Supplier Name plus Location, uppercase and centered in the certificate template. |
| Client / inward ID | Supplier Name and SN identify the row in Excel and the app work list. |
| Status | Ready is eligible for PDF creation. Alert needs review and is never an automatic creation trigger. Generated means the app created the PDF from an inward row; Generated Manual identifies a manually prepared/custom PDF. Completed is the user's confirmation that it was sent. N/A is excluded; Blank remains blank and needs a decision. Warnings can be explicitly accepted before generation and remain flagged in the report. |

PDF names use `{COD}_Company name_city_1.pdf` or `{COD}_Company name_city_2.pdf`. The leading COD matches the number written to Excel; Plant 1 and Plant 2 are identified only by the final number.

Inward age 25+ through day 30 is a follow-up reminder; after day 30 it stays overdue until Completed or N/A. The app does not infer recycling completion or email delivery.
