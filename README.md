# Certificate of Recycling Automation

Private project for preparing certificates of e-waste recycling from the facility inward workbook.

## Proposed workflow

1. The responsible person copies the completed facility inward workbook into the input folder.
2. The automation groups each inward batch by its `SN` and keeps all item rows belonging to that batch.
3. A batch becomes eligible when its `Inward Date` is at least 30 days old.
4. The automation validates the required certificate fields and flags missing or conflicting data for review.
5. It generates a certificate using the supplied certificate design and creates one Outlook draft per supplier/batch.
6. The responsible person reviews the certificate and draft, then sends it manually.
7. The automation records the certificate number, generated date, draft status, and any review reason so a batch is not processed twice.

## Important safeguards

- The first release will create drafts only. It will not send certificates automatically.
- One certificate is created per inward batch, not per item row.
- A missing client email address, missing inward date, missing supplier, invalid weight, or incomplete reference data stops that batch for review.
- The automation must not treat a blank `SN` on an item row as a new batch; it belongs to the most recent populated `SN` above it.
- The project contains only the supplied dummy workbook and template image. Do not add live client data to this repository.

## Workbook observations

The supplied workbook uses merged cells. The batch-level values appear on the first row of each batch, while item, quantity, per-piece weight, and weight values continue on following rows. The certificate needs a clear mapping for:

- Certificate number and certificate date
- Date of material receipt
- Total weight
- Supplier name
- Reference number
- Manifest number
- Client email address

The current sample does not contain a client email-address column or an explicit certificate-number column. Those should be added to the live process, or maintained in a separate controlled mapping file before automated sending is enabled.

## Recommended rollout

Start with a two-case test that generates PDFs and Outlook drafts without sending. After the responsible person confirms the values and visual layout, test a full day of dummy batches. Only after that should automatic sending be considered, with an approval/status column and an audit log.
