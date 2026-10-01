# Verification

The regression suite is run with:

```powershell
python -m unittest discover -s tests -v
```

Tests use temporary workbooks and temporary output folders. They exercise the certificate PDF renderer, the shared COD sequence across both plant files, custom and blank CODs, repeated generation after a deliberate Ready selection, workbook backups and write verification, PDF rollback when Excel writes fail, Windows-safe COD filenames, flat number-named output, alert status exclusions and retry after notification failure, day-25/day-30 boundaries, and Excel conditional formatting/data validation. Scheduled checks are verified as alert-only. No live plant workbook is modified, no Windows task is registered, and no mail client is opened.

Before production use, select the real Plant 1 and Plant 2 workbooks in the app and review one generated PDF and its written-back COD/date/status. The scheduled reminder requires the Windows user to be signed in and the computer awake. It only checks and alerts; certificate generation and sending remain deliberate user actions.

## Local verification, 1 October 2026

- Full Python suite: 230 tests; 228 passed, one privileged symlink test skipped,
  one ReportLab font-subsetting `MemoryError` in Siemens company alignment.
  The entire eight-test company alignment suite passed on an isolated rerun.
  This is not recorded as a clean full-suite pass; investigate/repeat under the
  operator's target environment before production sign-off.
- Latest focused interrupted-save/review tests: 24 passed, including recovery
  retries after partial staged files. Backup service: 44 tests, 43 passed and
  one privileged symlink test skipped. UI behavior tests: 44 passed.
- Installer PowerShell syntax, package imports and Python compilation checked.
  No installer, Windows scheduled task or real reminder was registered/run.
- Browser checks used a disposable workspace at `127.0.0.1:8766`, never the
  live plant files. Created Standard, Sita and Siemens PDFs; rendered their first
  pages for visual inspection. Verified certificate-only company edits retain
  the original Excel company while tracking COD/date/Generated status; raster
  preview; warning reason validation; visible saved filenames; backup creation
  and restore; activity reasons; unsaved draft protection; graceful app shutdown.
- The temporary QA app was stopped afterward. No Git commit or push was made.

The [production acceptance checklist](PRODUCTION_CHECKLIST.md) remains pending
manager and IT approval. Automated checks cannot approve business wording,
actual Windows reminder operation, XLSM macro preservation or separate-device
backup recovery. The local activity journal is not tamper-proof authentication.
