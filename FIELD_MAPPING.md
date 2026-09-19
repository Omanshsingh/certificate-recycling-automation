# Certificate field mapping

This mapping is based on the supplied `COD_0061.pdf` reference certificate and the dummy inward workbook.

| Certificate field | Observed source or rule | Status |
|---|---|---|
| Certificate No. | `0061` in the reference PDF | Source column not present in dummy workbook. Needs an approved numbering rule. |
| Certificate Date | `01.06.2026` in the reference PDF | Needs an approved rule. It is not simply the same as the inward date. |
| Date of Material Receipt | `Inward Date` | Confirmed. |
| Weight | Sum of the batch item weights, rounded to one decimal place | Confirmed from `217.15` to displayed `217.2 KG`. |
| Reference No. | `2026040047` in the reference PDF | Source column not identified in dummy workbook. Needs mapping. |
| Manifest No. | `Form 6 No.`; reference shows `10164` | Confirmed for the sample. |
| Supplier and city | `Supplier Name` plus `Location`; reference displays uppercase supplier and city | Confirmed formatting rule. |

## Key consequence

The automation must not generate a final certificate until the missing Reference No. source and the certificate numbering/date rule are confirmed. It can generate a clearly labelled test preview from the dummy workbook, but production output must stop and request review when those values are unavailable.
