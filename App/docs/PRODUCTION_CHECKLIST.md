# Production acceptance and backup operations

Status: human acceptance is pending. Automated tests do not approve the artwork,
business rules, field accuracy, Windows installation, or operational recovery.
Complete this checklist on a disposable acceptance workspace using representative
client data. Record the release revision, tester, date, evidence, and manager
decision. Do not run fault injection against the live plant workbooks.

## Release record

| Item | Evidence / owner / decision |
| --- | --- |
| Release revision and installer identity | Pending |
| Windows version, Excel version, machine and standard user account | Pending |
| Operations manager and certificate reviewer | Pending |
| IT installation and recovery owner | Pending |
| Acceptance workspace and representative workbook copies | Pending |
| Backup device, retention period and protected trust-store copy | Pending |
| Blocking defects and retest evidence | Pending |
| Manager approval, IT approval and approval date | Pending |

## Certificate and workbook acceptance

- [ ] Review printed and on-screen PDFs for all three types: Standard, Siemens,
  and SITA. Approve artwork, company/site names, addresses, wording, stamps,
  signatures, dates, units, table alignment and every type-specific field.
- [ ] Test long company names, city/site names, addresses, references, item
  descriptions and serial/model values. Verify wrapping and readable fonts;
  check the warning and explicit acceptance flow when text does not fit.
- [ ] Reconcile net weights, quantity × unit weight, continuation rows,
  per-certificate weight overrides, rounding, thousands separators, decimals,
  zero, blank, invalid and negative input. Confirm the intended weight appears
  in the PDF, workbook and saved details. Review differences among the templates.
- [ ] Verify certificate-only changes and changes written back to Excel. Check
  original inward values, formulas, other sheets, validation and formatting.
  Test both XLSX and XLSM; open the restored XLSM in Excel and verify macros.
- [ ] Review field overrides, warning explanations and required override reasons.
  Check that a changed file or stale confirmation requires a fresh review.
- [ ] Attach multi-page PDF annexures, replace/remove the selected annexure,
  and test missing, corrupt, encrypted and oversized attachments. Review page
  order, orientation, text, attachment metadata and the final combined PDF.
- [ ] Delete a certificate in the acceptance copy and test missing-PDF recovery.
  Corrupt a PDF and verify visible review guidance. Confirm retained CODs,
  filenames, printed CODs, references and inventory remain consistent.
- [ ] Test duplicate/manual CODs, manually added PDFs, corrupt inventory and
  invalid custom registers. No reused historical number may slip through.
- [ ] Reissue with the same COD, changed details, company/site name or template.
  Confirm the previous PDF is archived, a new receipt is recorded, and failure
  preserves the previous source/PDF and history.
- [ ] Create, edit, issue and complete custom certificates without an inward
  workbook row. Verify working drafts and the custom register survive restart.
- [ ] Open Excel during save, issuance, backup and restore. Verify a clear lock
  error, no false success, no lost changes and successful retry after closing it.
- [ ] Open two browser tabs. Review in both, save in one, then save/confirm in
  the stale tab. Verify optimistic SHA checks protect the first save. Repeat
  after a restore, attachment change, external workbook edit and reissue.

## Backend integration contract

These functions are provided by `App/backup_service.py`; wiring the web controls,
native folder picker, authentication and startup recovery belongs to the backend.

| Call | Result / required behavior |
| --- | --- |
| `create_backup(root, destination=None, *, locked=False)` | Returns `backup_id`, absolute ZIP `path`, UTC `created_at`, `purpose`, `file_count`, uncompressed `total_bytes`, `original_workbook_paths`, `trust_files`, and `warnings`. |
| `list_backups(root)` | Returns registered backups newest first. Healthy entries have the same metadata plus `valid=True`; damaged/missing entries have `backup_id`, `valid=False`, and `error`. Listing hashes every payload and performs no writes. |
| `restore_backup(root, backup_id, confirmed=False, *, locked=False)` | Requires the literal boolean `True`. Returns `backup_id`, `restored_files`, `archived_files`, absolute `extras_path`, `reserved_cods` (count of newly retained distinct COD reservations), `pre_restore_backup_id`, `pre_restore_path`, `recovery_path`, and `warnings`. |
| `check_restore_health(root)` | Read-only guard returning `{"healthy": True}` or raising `BackupError` for an interrupted/unusable restore journal. Call before startup issuance recovery and inside each writer's `run_lock`. |

`BackupError` means an unsafe, incompatible or invalid operation was refused.
Filesystem/Excel/run-lock errors may propagate as `OSError`/`PermissionError`.
`RestoreError` exposes `recovery_path` and `rolled_back`; the backend must show
the failure and whether supervised recovery is required, without claiming success.

The default destination is `root/Data/Backups`. An explicit native-picker folder
overrides the persisted `settings.local.json` field `backup_destination` for one
call. This module never changes that preference. It reads the field locally
for compatibility with settings loaders that do not preserve it. A blank persisted
value selects the default folder. The backend
must retain the field when saving other settings and validate picker results on
the server. Restore accepts an ID from the signed local catalog, never a client
archive path, manifest path, target path or upload.

Creation and restore acquire `run_lock`. Set `locked=True` only in trusted server
code already holding that lock; never take a nested lock or expose this argument
as a browser-controlled parameter. Close active Excel/database users first.
Exclusive source handles protect capture on Windows; they are released immediately
before an atomic replacement because Windows requires it. An intervening Excel
lock or filesystem error causes rollback. Local processes must use the same
workflow lock; this is not protection against an administrator changing files.

Restore verifies the entire archive, stages and verifies every payload, checks
current and saved workbook mappings, locks existing targets, preserves their
original bytes, and creates a registered pre-restore ZIP before replacement.
It writes a durable transaction journal and verifies replacement checksums.
Exceptions trigger rollback of all attempted replacements. Originally absent
files are moved to recoverable `.restore-recovery-<transaction>` siblings during
rollback. Files belonging to current issuance state but absent from the snapshot
are recoverably moved into `Data/Archive/Restore Extras/<transaction>/`, preserving
their original workspace-relative paths. This includes active certificate PDFs,
legacy certificate PDFs, later issuance receipts and recovery snapshots, and a
custom register or working drafts created after a snapshot that lacked them.
Unrelated notes, previews, reports, reusable attachments and existing archives
remain in place. The selected snapshot supplies active PDF/receipt ownership;
later receipts do not remain active and invalidate a restored same-COD PDF.
The restored inventory unions previously observed COD reservations (including
deleted-file history and receipt-only CODs), without copying later managed owners
or PDF hashes. Reservation-only entries cannot establish issuance ownership.
If the current reservation ledger or receipts cannot be read safely, restore
requires review instead of silently discarding their historical reservations.
Exceptions roll back both replacement files and extras archive moves. The journal
records every planned move, and the pre-restore ZIP retains the full prior state.
Keep recovery originals and pre-restore snapshots until acceptance is complete.

The archive restores only these approved locations:

- Exact configured XLSX/XLSM paths, including a workbook outside the application
  folder only when its original absolute path still matches the current plant
  mapping. Plant order must match. Missing files may be recreated; missing,
  swapped or changed mappings are refused. A different workspace root is refused.
- `settings.local.json`, existing default settings, `Certificates`, legacy
  `Output/Certificates`, critical `Data` files/folders, attachments, audit,
  inventory, custom register, working drafts and issuance history/recovery.
- Template PDF/PPTX/image/geometry/font assets in `templates` or `App/templates`.
  Program executables and source code are not restored.
- Historic PDFs in `Data/Archive`, including Superseded and Uncommitted
  Certificates. Restore Extras archives also include their retired receipts,
  recovery binary snapshots, registers and working drafts, so subsequent backups
  preserve the complete recovery material. Other non-PDF archive contents are excluded.

Preview/report/cache folders, backup trees, transient/lock files and SQLite
sidecars are excluded. Active SQLite WAL/SHM/journal sidecars make backup/restore
fail: close the users and checkpoint the database with qualified IT assistance.
Do not delete sidecars to force success. Quiescent SQLite files, including corrupt
original bytes needed for a pre-restore snapshot, are preserved byte for byte.
Links, junctions, hardlinks, traversal, Windows device names, alternate data stream
paths, duplicate targets/members and expansion beyond safety limits are refused.

`Data/Issuance History/*.json` and all `Data/Issuance Recovery/<id>` contents are
included, including `new.pdf`, `previous-*.pdf` and `source-before.bin`. Complete
and rolled-back journals remain audit evidence. Pending restored journals must
reference the exact configured source or local custom register, approved PDF
paths and present snapshots with matching hashes. After restore, the backend
must run its issuance recovery under `run_lock`, refresh the workspace revision,
and require new reviews in every browser tab before further issuance. This module
does not invoke issuance recovery or preserve old browser approval tokens.

## Backup trust and portability

ZIP manifests and the catalog are authenticated with HMAC-SHA256 using the
workspace's random `Data/Backups/trust.key`. The catalog pins whole-archive SHA256;
each manifest records payload SHA256, sizes and original workbook paths. This
protects against archive tampering, not an attacker who controls the workstation
and its trust key. Backups are not encrypted; secure the files with appropriate
Windows permissions and approved encrypted/removable storage.

For a portable disaster-recovery copy, preserve all of the following on a
separate approved device after every backup:

1. The ZIP archives at the chosen backup destination.
2. `Data/Backups/trust.key` and the latest `Data/Backups/catalog.json` together,
   in a protected recovery folder outside the ZIPs. `trust_files` in the API
   result identifies these files. They are deliberately excluded from snapshots.
3. The original workspace path, plant workbook absolute paths, plant order,
   chosen backup directory paths and installer identity in the IT recovery record.
4. Any retained restore transaction folder needed for an interrupted restore.

On a rebuilt Windows machine, install the approved app, recreate the SAME workspace,
workbook and backup directory paths (including drive letters), place the preserved
key/catalog into `Data/Backups`, and reconnect the exact plant mappings before
selecting a registered backup ID. The restored workbook files themselves may be
missing before restore. Never generate a new key to replace a lost key. A ZIP alone
cannot be restored through this API if its trust key/catalog is lost. Moving the
workspace, changing drive letters or importing archives is deliberately unsupported
by the ordinary restore controls; those require a separately reviewed IT migration
tool/process. Do not edit signed catalog contents to bypass these checks.

A backup on the same physical disk is local recovery, not disaster protection.
Agree frequency, separate-device copies, encryption, retention, capacity and a
responsible owner. No backup is automatically purged by this module. Large stores
need a capacity/retention procedure; current limits are 100,000 files, 2 GiB per
file, 20 GiB expanded total and 32 MiB manifest. Listing does full verification
and may take time; run it away from the web request/UI thread when needed.

## Restore, crash and clean-install acceptance

- [ ] Use the native folder picker to choose separate storage. Cancel the picker,
  restart the app and edit reminder settings; verify the destination survives.
  Reject destinations within included data/program folders or containing a plant
  workbook. Verify disconnected, full, read-only and non-directory destinations.
- [ ] Create and list a backup. Inspect both workbook formats, macro bytes,
  settings, inventory, custom drafts/register, attachments, archive PDFs,
  issuance history and every issuance recovery snapshot. Verify exclusions.
- [ ] Require visible, explicit restore confirmation for the selected ID. Test
  cancellation, missing ZIP, altered ZIP/manifest, unknown ID, arbitrary paths,
  changed plant order, different workspace, lost trust key and linked targets.
- [ ] Change representative files and restore. Verify staged checks happen before
  overwrite, a pre-restore snapshot exists, later active PDFs/receipts/recovery
  snapshots move into Restore Extras, unrelated files remain in place, the
  restored settings point to the same sources and all types still render correctly.
  Reconcile COD history and missing/deleted-number reservations against the
  pre-restore snapshot before issuing again.
- [ ] Back up an issued certificate, reissue the same COD with changed PDF bytes
  and/or filename, then restore. Verify the original receipt matches the restored
  PDF hash, the later receipt/recovery material is archived, and only the original
  source ownership remains active. Create a new certificate after a backup and
  restore; verify it is archived, its COD remains reserved, and no ownership of
  that later issuance is adopted. Repeat restore and make another backup to verify
  reservations persist and Restore Extras receipts/binary snapshots are included.
- [ ] Inject a write error after the first replacement in a disposable copy.
  Verify every attempted target returns to its exact original bytes, newly added
  files are recoverably moved, originals are retained, and failure is visible.
  Repeat after extras have moved, including a failure immediately after a rename;
  verify extras return to their original locations and the prior ledger is exact.
- [ ] Force a rollback failure. Verify `rolled_back=False`, preserved originals,
  a readable recovery path and a persistent block on backup/restore mutations.
- [ ] Terminate a disposable process during restore, during issuance PDF publish,
  source commit, archive creation and journal update. Restart and verify no false
  success, no silently reused COD and no unreviewed overwrite. Record real process
  termination evidence; an exception test alone does not prove crash recovery.
- [ ] Rehearse IT recovery of an interrupted restore. `in_progress`, malformed
  and `rollback_failed` journals block further backup/restore; there is no automatic
  multi-file crash rollback in this module. Stop issuance and scheduled jobs, close
  file users, preserve the transaction/trust store, validate the journal targets
  against the original workspace and plant mapping, and use preserved originals
  or the pre-restore snapshot in a supervised recovery. Do not clear the journal
  until restored bytes and COD continuity are verified. Atomic file replacement
  and fsync do not promise a multi-file transaction through sudden power loss.
- [ ] Ensure backend startup calls `check_restore_health(root)` before issuance
  recovery and each writer calls it while holding `run_lock`. Verify the public
  guard blocks interrupted/malformed journals without writing any files. The backup
  service does not itself stop unrelated writers; integration must enforce that
  maintenance state. Do not automatically replay unknown external journal paths.
- [ ] Install on a clean supported Windows machine as a standard user. Verify
  dependency/package availability, launch shortcut, native picker, local-browser
  access, offline operation, permissions, Unicode/space-containing paths and Excel.
- [ ] Install the daily reminder task, restart/sign in, and verify enabled/disabled
  state, configured time, day-25–30 follow-up, overdue reminders, snooze, dismissal,
  opening Excel/app, retry after failure and lack of duplicate task registrations.
- [ ] Rehearse the separate-device recovery procedure above on the clean machine.
  Have the manager approve actual output and IT approve recovery before release.

Run automated backup checks from `App` with
`python -m unittest tests.test_backup_service -v`. They use temporary directories
and byte fixtures rather than live Excel. A privileged symlink test may be skipped
on Windows without developer mode; the Windows junction test covers reparse-point
rejection independently. Record actual test output and any skips; leave the human
acceptance boxes unchecked until the relevant reviewer has completed them.
