# Keeping manual Excel editing reliable

## Recommended architecture

Use a local SQLite register as the app's durable record store, with Excel as an
editable import/export surface. People can continue making manual Excel changes.
Saved changes pass through validation and reconciliation before the app uses them
to issue a certificate. A workbook row number is a location, not an identity.

This is a proposed migration. The current patch keeps the workbook-based register
and adds the safeguards below; it does not claim to have migrated existing data.

## Safeguards implemented now

- Duplicate Serial Numbers are checked within each plant, ignoring surrounding
  spaces and letter case. The app explains conflicting worksheet rows before a
  save. Manually introduced duplicates block issuance for those records rather
  than choosing the first matching company. Unrelated inwards remain usable:
  date sorting moves and verifies whole physical blocks independently of their
  Serial Numbers.
- Reusing a Serial Number does not by itself authorize restoring an earlier
  certificate's details. History must match the current COD, or, when COD is
  blank, the stored source facts. Mismatches are visible conflicts and block
  generation and completion until the identifiers are reconciled.
- Opening a workbook for repair remains available when a unique record cannot be
  selected. Preview uses saved workbook contents without taking a write lock.
- A refresh can show saved contents when Excel prevents exclusive write access.
  Computed Ready and missing-PDF recovery states are labelled as pending Excel
  updates. A successful preview or read never implies that an Excel write worked.
- App saves and writable refreshes put supported inward blocks in date order.
  Asset rows, merges, block-local formulas, and record numbers travel together.
  When an Excel feature prevents safe reordering, saves retain the current row
  order and show one workbook note. Data verification failures still stop a save.
- Generation preflight reads saved workbooks and inventory without changing
  source facts, recovering an interrupted save, or replacing the full work list.
  Only the final save or an explicit refresh updates those sources.
- A history conflict stops its own record; it does not stop other complete
  inwards in that plant from becoming Ready.
- Asset edits made for a PDF do not depend on spare Excel rows. Excel updates
  verify the resulting material values, including formula-backed weights,
  before reporting success. Adding another material preserves entered weights.
- Add-client retries retain their request identity after a lost response. An
  unchanged retry can reach the server's recovery check even if its saved row
  has already appeared in the company list.
- Reviewed content revisions, issuance receipts, backups, and publication recovery
  continue to protect actual saves and certificate numbering.

## Migration sequence

1. **Create and verify a baseline.** Back up both workbooks, certificate files,
   app register, and issuance history. Produce a reconciliation report for
   duplicate Serial Numbers, reused CODs, unreadable rows, and missing files.
   Resolve ambiguous ownership explicitly; do not guess which record is correct.
2. **Introduce permanent identity.** Give each inward a UUID independent of its
   plant-specific Serial Number, worksheet position, date, and company name.
   Retain the UUID in an app-managed workbook column so it moves with a complete
   inward. Existing certificates retain their COD and history. Legacy matching
   must be unique and verified before assigning an identity.
3. **Build a validated local register.** Store inwards, assets, issued certificate
   versions, workbook imports, conflicts, and before/after changes transactionally.
   Serial Number uniqueness is a database constraint scoped to a plant. Certificate
   number allocation is a separate, globally unique operation.
4. **Reconcile saved Excel changes.** Compare each incoming record with its last
   imported version and the current database version. Apply valid, non-conflicting
   changes. Present duplicate IDs, copied UUIDs, simultaneous edits, unexplained
   deletions, or changed issued identifiers for review. Keep other readable records
   accessible. Never interpret a missing row as permission to delete history.
5. **Export safely.** Compare the reviewed workbook revision immediately before
   writing, acquire write access, back up the file, and publish atomically. If Excel
   has changed again, preserve the user's input and ask for reconciliation. Label
   pending exports accurately while the file is open.
6. **Switch only after parity checks.** Compare old and new imports, weights,
   numbering floors, and certificate outputs against the same workbook copies.
   Exercise duplicates, backdating, moved blocks, direct Excel edits, concurrent
   edits, open files, interrupted saves, and restored backups before switching the
   authoritative record store. Retain a recovery route to the verified baseline.

## What this resolves

Reordering rows or correcting a Serial Number would no longer change the app's
record identity. Invalid Excel edits would become specific import conflicts,
rather than failures scattered across opening, previewing, numbering, and saving.
This reduces whole classes of bugs; it does not remove the need for validation,
recovery, and checks when the app changes.
