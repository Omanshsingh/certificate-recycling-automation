# Project plan

## Phase 1: understand and validate the source

- Read the merged-cell workbook without changing the original.
- Normalize each inward batch into one record with child item rows.
- Calculate and validate total weight from the item rows.
- Identify the official source for certificate number, reference number, manifest number, and recipient email.

## Phase 2: generate the certificate

- Reproduce the supplied certificate design as a controlled template.
- Fill batch-level values in the certificate.
- Export one PDF per eligible batch.
- Keep the source workbook unchanged.

## Phase 3: prepare communication

- Create one Outlook draft per eligible batch.
- Attach the matching PDF certificate.
- Use a standard subject and body approved by the business.
- Do not open many compose windows and do not send automatically.

## Phase 4: scheduling and audit

- Run daily or on demand from the responsible person’s desktop.
- Select the latest copied workbook.
- Process only batches at least 30 days old.
- Record generated, drafted, sent, skipped, and review-required states.
- Make reruns idempotent so a batch is not duplicated.

## Decisions needed before implementation

1. Is the delay exactly 30 calendar days from `Inward Date`, or one month by calendar date?
2. Where does the client email address come from?
3. How are certificate numbers assigned?
4. Which source fields populate Reference No. and Manifest No.?
5. Should the person approve every draft, or should approved rows be sent automatically later?
