# V1 certificate creation check — 16 October 2026

Scope: existing saved plant workbooks → preview → generate PDF → save COD,
certificate date and Generated status. Verification uses disposable workbooks,
PDFs, history and browser data. The operator's workbooks and certificates were
not changed. The running operator app was not restarted.

## Defects fixed

- A Form 6 number supplied during first creation could be dropped when the
  form was projected back onto the Excel record. Preserve it in preview and
  save it with COD/date/status when issuance commits. Company/material edits
  made for the PDF remain separate from that required Excel field.
- The generation runner enforced missing Form 6 only for requests without form
  details. Enforce it for both paths, including confirmed warnings.
- Source-control refresh forced a filled inward into edit mode, bypassing the
  intended review summary. Preserve the selected review/edit mode; a deliberate
  certificate-type change still exposes the new template's editable fields.
- Low-memory exceptions can have no message. The browser exposed a blank
  generation error on this computer. Give resource failures a short recovery
  message in generation reports and API errors.

## Evidence

- Created Standard, Sita and Siemens certificates through the actual creation
  API with separate Serial Numbers, the same company, and a shared reference.
  Verified distinct CODs 0001, 0002 and 0003, correct dates, 7 kg weights,
  Generated status, and matching issuance records. Rendered and visually
  reviewed each resulting page.
- Used the actual browser against a disposable server. Created Acme Recycling
  Ltd's Standard certificate and verified its PDF against the workbook:
  COD 0002, 12.5 kg, Generated status. Verified the corrected review summary
  in the browser for another filled inward.
- A later browser generation failed while Windows was short on available
  committed memory. Its inward retained Ready status and a blank COD; no
  partial PDF was published. A focused injected-memory-failure check verifies
  a useful message, unchanged source, no published certificate, and successful
  retry with the same next COD.
- Deferred Form 6 checks cover preview without writing Excel, final save and
  printed value, preservation of PDF-only company/material edits, failed save
  rollback, and successful retry.
- Updated older assertions to the user's current requirements: Form 6 optional
  at intake, required at issuance; visible calculation/additional fields; red
  alerts; fixed issued CODs; nested recovery archives; preserved entered unit
  weights. Tests still assert actual PDF/source values and rollback outcomes.

### Automated results

Final focused Python regression: **103 tests passed in 400.176 seconds**, exit
code 0. The saved output is `.artifact_work/v1-creation-review/regression.log`;
the recorded exit code is in `regression-exit.txt`. This is the scoped V1 run
listed below, not a claim that the complete repository test suite passed.
UI checks: 167 passed with a constrained Node heap. An unconstrained attempt
hit this workstation's memory limit; this is not counted as a clean run.
An earlier Python regression run found obsolete expectations. A subsequent
unconstrained run exited without a complete result during the same resource
pressure. Neither is represented as a passing run.

Commands for the final checks (run the Python command from `App`):

```powershell
$env:OPENBLAS_NUM_THREADS='1'
python -m unittest tests.test_simple_certificate_flow tests.test_excel_edit_resilience tests.test_register_layout tests.test_issuance_safety tests.test_certificate_workflow tests.test_certificate_templates.WebGenerationTests.test_selected_generation_warning_keeps_workspace_unchanged_until_confirmation tests.test_certificate_edit_scope.CertificateEditScopeTests.test_duplicate_cod_on_an_issued_inward_cannot_be_confirmed_into_a_new_number -q
```

Run UI checks from the repository root:

```powershell
node --max-semi-space-size=4 --max-old-space-size=192 --test --test-reporter=dot App/tests/test_simple_ui.cjs
```

## Boundaries

- This checks the ordinary register layout and creation workflow, rather than
  certifying that every manually edited Excel structure is supported.
- Generation still needs both numbering workbooks readable and write access
  to their files. Save and close Excel before issuance. Unreadable COD history
  and duplicate Serial Numbers still require correction. References may be
  reused without an alert.
- Hidden duplicate registers and worksheet renames affecting saved optional
  settings remain documented edge cases in the audit. They were not migrated
  or repaired in this creation pass.
- Visual examples are representative checks, not approval of all possible
  long names, wording or template customizations. Installation, physical power
  loss and live Excel COM selection were not part of this pass.
- Low available system memory can still prevent a certificate from being
  created. These changes explain the failure and protect the save; they do not
  change the user's Windows memory settings or close their applications.

Disposable evidence is retained under `.artifact_work/v1-creation-review/`.
The temporary browser tab and server were closed after verification.
