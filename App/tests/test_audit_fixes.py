"""Production-audit regressions using disposable workbooks and real PDFs."""
from datetime import date
from io import BytesIO
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils.datetime import CALENDAR_MAC_1904, to_excel
from pypdf import PdfReader

from certificate_core import inward_details_complete, number, parse_date, read_batches
from excel_values import formula_number
from tests.test_certificate_workflow import make_workbook, AS_OF
from workbook_store import WorkbookSession, exclusive_file, stage_certificate_details
from run_certificates import run
from workflow_support import save_json
import web_app

APP = Path(__file__).resolve().parents[1]


class AuditFixTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.paths = [self.root/'plant1.xlsx', self.root/'plant2.xlsx']
        make_workbook(self.paths[0], [dict(sn='A', status='')])
        make_workbook(self.paths[1], [dict(sn='B', status='N/A')])
        shutil.copytree(APP/'templates', self.root/'App/templates')
        save_json(self.root/'settings.local.json', dict(workbook_paths=[str(p) for p in self.paths], reminders_enabled=False))
        for name, value in (('ROOT', self.root), ('BACKUP_DIR', self.root/'Data/Workbook Backups'),
                            ('RESULTS_PATH', self.root/'Data/Reports/Latest Results.json')):
            self.enterContext(patch.object(web_app, name, value))

    def mutate(self, action, path=None):
        path = path or self.paths[0]
        book = load_workbook(path)
        try:
            action(book)
            book.save(path)
        finally:
            book.close()

    def execute(self, **options):
        # These disposable workspaces start without historical CODs; the
        # test operator explicitly chooses the first number, as the UI does.
        options.setdefault('starting_cod', '0001')
        return run(self.paths, self.root, as_of=AS_OF, progress=None, **options)

    def report(self):
        return json.loads(web_app.RESULTS_PATH.read_text(encoding='utf-8'))

    def pdf_text(self):
        pdfs = list((self.root/'Certificates').rglob('*.pdf'))
        self.assertEqual(len(pdfs), 1)
        return '\n'.join(page.extract_text() or '' for page in PdfReader(pdfs[0]).pages)

    @staticmethod
    def total(book, value=13):
        book.active.append([None,None,None,None,'Monitors',2,3,6])
        book.active.append([None,None,None,None,'Total',None,None,value])

    def test_numeric_and_formula_totals_are_reconciled_in_preview_and_final_pdf(self):
        for total in (13, '=SUM(H2:H3)'):
            with self.subTest(total=total):
                make_workbook(self.paths[0], [dict(sn='A', status='')])
                self.mutate(lambda book: self.total(book, total))
                source = read_batches(self.paths[0])[0]
                self.assertEqual(source['total_weight'], 13)
                self.assertEqual(len(source['items']), 2)
                self.assertEqual(source['summary_rows'], [4])
                self.assertTrue(inward_details_complete(source, AS_OF))
                before = self.paths[0].read_bytes()
                details = web_app.editor_details(source)
                body = dict(plant='plant1', source_kind='workbook', source_sn='A', source_plant='plant1',
                            details=details, starting_cod='0001')
                preview = web_app.preview_certificate(body)
                self.assertEqual(preview['printed_details']['certificate_weight'].strip(), '13 KG')
                self.assertEqual(self.paths[0].read_bytes(), before)
                self.assertEqual(self.execute()[0], 1)
                self.assertIn('13 KG', self.pdf_text())
                self.assertNotIn('26 KG', self.pdf_text())
                self.assertFalse(self.report()['rows'][0]['warnings'])
                # Tracking writes leave the source summary, including formulas, intact.
                book = load_workbook(self.paths[0]); self.assertEqual(book.active['H4'].value, total); book.close()
                shutil.rmtree(self.root/'Certificates')
                shutil.rmtree(self.root/'Data')

    def test_mismatched_total_cannot_be_accepted_to_generate(self):
        self.mutate(lambda book: self.total(book, 99))
        self.mutate(lambda book: setattr(book.active['M2'], 'value', 'Ready'))
        before = self.paths[0].read_bytes()
        self.assertEqual(self.execute(allow_warnings=True)[0], 0)
        row = next(row for row in self.report()['rows'] if row['sn'] == 'A')
        self.assertTrue(row['issuance_blocked'])
        self.assertIn('row 4', row['detail'])
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))
        # Formatting is allowed on a check, but no issuance tracking is changed.
        self.assertEqual(read_batches(self.paths[0])[0]['cod_no'], '')

    def test_total_with_material_inputs_is_not_silently_discarded(self):
        self.mutate(lambda book: book.active.append([None,None,None,None,'Total',2,3,6]))
        batch = read_batches(self.paths[0])[0]
        self.assertFalse(inward_details_complete(batch, AS_OF))
        self.assertTrue(batch['blocking_issues'])

    def test_common_heading_aliases_reach_the_final_pdf(self):
        def aliases(book):
            for coordinate, label in {'A1':'Inward ID','C1':'Customer','D1':'Site Location','E1':'Description',
                                       'I1':'Manifest No.','J1':'Ref No.','M1':'Status'}.items():
                book.active[coordinate] = label
        self.mutate(aliases)
        source = read_batches(self.paths[0])[0]
        self.assertEqual((source['reference'],source['location'],source['form6']), ('REF-2','Mumbai','FORM6-2'))
        self.assertTrue(inward_details_complete(source, AS_OF))
        self.assertEqual(self.execute()[0],1)
        text = self.pdf_text()
        for value in ('REF-2','MUMBAI','FORM6-2'): self.assertIn(value,text)

    def test_shifted_headers_are_used_by_refresh_generation_and_corrections(self):
        def shift(book):
            book.active.insert_rows(1,4)
            book.active['A1'] = 'Operator inward register'
        self.mutate(shift)
        batch = read_batches(self.paths[0])[0]
        self.assertEqual((batch['source_header_row'],batch['source_row']), (5,6))
        self.assertEqual(self.execute()[0],1)
        source = read_batches(self.paths[0])[0]
        details = web_app.editor_details(source); details['supplier']='Corrected company'
        with WorkbookSession(self.paths[0],self.root/'Data/Workbook Backups') as session:
            staged, proposed = stage_certificate_details(session,details)
        self.assertEqual(proposed['supplier'],'Corrected company')
        book = load_workbook(BytesIO(staged))
        self.assertEqual(book.active['A1'].value,'Operator inward register')
        self.assertEqual(book.active['C6'].value,'Corrected company')
        self.assertEqual(book.active['C5'].value,'Supplier Name')
        book.close()
        with WorkbookSession(self.paths[0],self.root/'Data/Workbook Backups') as session:
            session.set_issued_status('A',source['source_sheet'],6,source['cod_no'],'Generated','Completed')
        self.assertEqual(read_batches(self.paths[0])[0]['issued'],'Completed')

    def test_hidden_identical_register_is_accepted_but_different_archive_is_not_ignored(self):
        def copy(book):
            archive=book.copy_worksheet(book.active);archive.title='Old copy';archive.sheet_state='hidden'
        self.mutate(copy)
        self.assertEqual(len(read_batches(self.paths[0])),1)
        self.mutate(lambda book:setattr(book['Old copy']['K2'],'value','9999'))
        with self.assertRaisesRegex(ValueError,'Hidden register.*different records'):
            read_batches(self.paths[0])

    def test_multiple_visible_registers_and_duplicate_mapped_columns_remain_explicit(self):
        self.mutate(lambda book:book.copy_worksheet(book.active))
        with self.assertRaisesRegex(ValueError,'More than one inward register'):
            read_batches(self.paths[0])
        make_workbook(self.paths[0],[dict(sn='A',status='')])
        self.mutate(lambda book:setattr(book.active['N1'],'value','Company'))
        with self.assertRaisesRegex(ValueError,'Ambiguous duplicate headings'):
            read_batches(self.paths[0])

    def test_unknown_populated_site_heading_and_merged_heading_block_generation(self):
        for change in (lambda book:setattr(book.active['D1'],'value','Site region'),
                       lambda book:book.active.merge_cells('C1:D1')):
            make_workbook(self.paths[0],[dict(sn='A',status='Ready')])
            self.mutate(change)
            self.assertTrue(read_batches(self.paths[0])[0]['blocking_issues'])
            self.assertEqual(self.execute(allow_warnings=True)[0],0)

    def test_populated_form6_mail_tracking_is_ignored_and_preserved(self):
        def add_tracking(book):
            book.active['O1'] = 'Form 6 POG Copies Mail sent date'
            book.active['O2'] = date(2026, 9, 20)
        self.mutate(add_tracking)
        source = read_batches(self.paths[0])[0]
        self.assertFalse(source['blocking_issues'])
        self.assertEqual(source['form6'], 'FORM6-2')
        self.assertTrue(inward_details_complete(source, AS_OF))
        before = self.paths[0].read_bytes()
        with patch.object(web_app, 'make_template_certificate', wraps=web_app.make_template_certificate) as render:
            preview = web_app.preview_certificate(dict(plant='plant1', source_sn='A', source_plant='plant1',
                                                       details=web_app.editor_details(source), starting_cod='0001'))
        self.assertEqual(render.call_args.args[0]['form6'], 'FORM6-2')
        self.assertTrue(preview['preview_url'])
        self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertEqual(self.execute()[0], 1)
        self.assertIn('FORM6-2', self.pdf_text())
        book = load_workbook(self.paths[0])
        try:
            tracking_column = next(cell.column for cell in book.active[1]
                                   if cell.value == 'Form 6 POG Copies Mail sent date')
            self.assertEqual(book.active.cell(2, tracking_column).value.date(), date(2026, 9, 20))
        finally:
            book.close()

    def test_mail_tracking_heading_variants_do_not_become_certificate_fields(self):
        for heading in ('FORM 6 - POG COPIES - E-MAIL SENT DATE', 'Form 6 Mail Sent Date',
                        'Manifest Copies Email Received Date', 'Reference Email Sent Status'):
            with self.subTest(heading=heading):
                self.mutate(lambda book:setattr(book.active['O1'], 'value', heading))
                self.mutate(lambda book:setattr(book.active['O2'], 'value', '2026-09-20'))
                source = read_batches(self.paths[0])[0]
                self.assertFalse(source['blocking_issues'])
                self.assertEqual((source['form6'], source['reference']), ('FORM6-2', 'REF-2'))

    def test_mail_tracking_does_not_supply_missing_manifest_or_hide_ambiguous_fields(self):
        self.mutate(lambda book:(setattr(book.active['O1'], 'value', 'Form 6 POG Copies Mail sent date'),
                                 setattr(book.active['O2'], 'value', date(2026, 9, 20)),
                                 setattr(book.active['I2'], 'value', None)))
        source = read_batches(self.paths[0])[0]
        # Form 6 is collected on the create-certificate step. A mail timestamp
        # must never be mistaken for it, but an incomplete inward can be saved.
        self.assertEqual(source['blocking_issues'], [])
        self.assertEqual(source['form6'], '')
        self.assertTrue(inward_details_complete(source, AS_OF))
        for heading in ('Form 6 alternate number', 'Site region', 'Reference Email Sent Date / Reference No'):
            with self.subTest(heading=heading):
                self.mutate(lambda book:setattr(book.active['P1'], 'value', heading))
                self.mutate(lambda book:setattr(book.active['P2'], 'value', 'Unmapped value'))
                self.assertTrue(any('not recognised' in issue
                                    for issue in read_batches(self.paths[0])[0]['blocking_issues']))

    def test_note_footer_is_not_a_fake_inward(self):
        self.mutate(lambda book:book.active.append(['Note: checked by supervisor']))
        self.assertEqual(len(read_batches(self.paths[0])),1)

    def test_supported_formulas_are_recalculated_without_cached_values(self):
        for formula, expected in [('=ROUND(F2*G2,2)',7),('=SUM(3,4)',7),('=ROUND(2.675,2)',2.68),
                                  ('=SUM(F2,G2)',5.5),('=($F$2 * $G$2)',7)]:
            with self.subTest(formula=formula):
                self.mutate(lambda book:setattr(book.active['H2'],'value',formula))
                batch=read_batches(self.paths[0])[0]
                self.assertEqual(batch['total_weight'],expected)
                self.assertTrue(inward_details_complete(batch,AS_OF))

    def test_unknown_external_circular_and_cross_row_formulas_are_not_executed(self):
        for formula in ('=SUM(H2:H3)','=H2','=F3*G3',"='[other.xlsx]Sheet1'!F2",'=TODAY()',
                        '=1/0','=1E999','=__import__("os")','=ROUND(1,100)', '='+'('*30+'1'+')'*30):
            with self.subTest(formula=formula):
                self.assertIsNone(formula_number(formula,{'F2':2.0,'G2':3.5}))

    def test_named_unambiguous_us_and_serial_dates_use_the_workbook_epoch(self):
        for value in ('29 Sep 2026','September 29, 2026','09/29/2026',to_excel(date(2026,9,29))):
            with self.subTest(value=value):
                self.mutate(lambda book:setattr(book.active['B2'],'value',value))
                self.assertEqual(read_batches(self.paths[0])[0]['inward_date'],date(2026,9,29))
        self.mutate(lambda book:(setattr(book,'epoch',CALENDAR_MAC_1904),setattr(book.active['B2'],'value',
                                                           to_excel(date(2026,9,29),CALENDAR_MAC_1904))))
        self.assertEqual(read_batches(self.paths[0])[0]['inward_date'],date(2026,9,29))
        self.assertEqual(parse_date('09/10/2026'),date(2026,10,9))
        self.assertIsNone(parse_date('31/02/2026'))

    def test_explicit_kg_values_are_read_without_guessing_other_units(self):
        for coordinate in ('G2','H2'):
            self.mutate(lambda book:setattr(book.active[coordinate],'value','3.5 kg' if coordinate=='G2' else '7 kg'))
        batch=read_batches(self.paths[0])[0]
        self.assertEqual(batch['total_weight'],7)
        self.assertTrue(inward_details_complete(batch,AS_OF))
        self.mutate(lambda book:setattr(book.active['H2'],'value','7000 g'))
        self.assertFalse(inward_details_complete(read_batches(self.paths[0])[0],AS_OF))

    def test_empty_status_selected_sita_and_siemens_generate_with_required_form6(self):
        for kind in ('sita','siemens'):
            with self.subTest(kind=kind):
                make_workbook(self.paths[0],[dict(sn='A',status='',form6='FORM-A')])
                source=read_batches(self.paths[0])[0]
                details=web_app.editor_details(source);details['certificate_type']=kind
                if kind=='siemens': details['certificate_fields']=dict(model_name='Scanner',model_serial='M1',tube_serial='T1')
                body=dict(plant='plant1',source_kind='workbook',source_sn='A',source_plant='plant1',
                          details=details,starting_cod='0001')
                before=self.paths[0].read_bytes()
                preview=web_app.preview_certificate(body)
                self.assertFalse([message for message in preview['warnings'] if 'Form 6' in message])
                self.assertEqual(self.paths[0].read_bytes(),before)
                with self.assertRaisesRegex(ValueError,'certificate type differs from Excel'):
                    web_app.create_certificate(body)
                self.assertEqual(self.paths[0].read_bytes(),before)
                body.update(edit_scope='update_excel',confirmed_excel=True)
                result=web_app.create_certificate(body)
                self.assertEqual(result['created'],1)
                after=read_batches(self.paths[0])[0]
                self.assertEqual((after['supplier'],after['location'],after['form6']),
                                 (source['supplier'],source['location'],source['form6']))
                self.assertEqual(after['issued'],'Generated')
                self.assertEqual(after['certificate_type'],kind)
                text=self.pdf_text()
                if kind=='siemens':
                    for value in ('M1','T1','Mumbai'): self.assertIn(value,text)
                shutil.rmtree(self.root/'Certificates');shutil.rmtree(self.root/'Data')

    def test_incomplete_common_details_and_explicit_na_cannot_use_type_selection_to_issue(self):
        for changes in (dict(supplier=''),dict(weight=-7),dict(status='N/A')):
            make_workbook(self.paths[0],[dict(dict(sn='A',status='',form6=''),**changes)])
            details=web_app.editor_details(read_batches(self.paths[0])[0]);details['certificate_type']='sita'
            with self.assertRaises(ValueError):
                web_app.workbook_form(dict(plant='plant1',source_sn='A',source_plant='plant1',details=details),details,self.paths)

    def test_empty_status_report_names_the_material_problem(self):
        self.mutate(lambda book:setattr(book.active['H2'],'value',-7))
        self.execute(check_only=True)
        row=next(row for row in self.report()['rows'] if row['sn']=='A')
        self.assertTrue(any('weight' in warning for warning in row['warnings']))
        self.assertIn('weight', row['detail'])
        self.assertEqual(row['certificate_status'],'alert')
        self.assertEqual(read_batches(self.paths[0])[0]['issued'],'Alert')

    def test_future_cod_date_requires_review_for_ready_and_empty_status(self):
        for status in ('Ready', ''):
            with self.subTest(status=status):
                make_workbook(self.paths[0], [dict(sn='A', status=status, cod_date=date(2050,1,1))])
                self.assertEqual(self.execute()[0],0)
                row=next(row for row in self.report()['rows'] if row['sn']=='A')
                self.assertEqual(row['certificate_status'], 'alert')
                self.assertTrue(any('COD Date is in the future' in warning for warning in row['warnings']))
                self.assertFalse(inward_details_complete(read_batches(self.paths[0])[0],AS_OF))
                self.assertEqual(self.execute(allow_warnings=True,override_reason='Future date agreed')[0],0)
                self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))
                self.assertEqual(read_batches(self.paths[0])[0]['cod_no'], '')

                # Correcting the source date makes this automatic Alert Ready
                # on the next review; confirmation cannot approve a bad date.
                # COD Date is in column P after the register layout migration.
                self.mutate(lambda book:setattr(book.active['P2'],'value',AS_OF))
                self.execute(check_only=True)
                reviewed = next(row for row in self.report()['rows'] if row['sn']=='A')
                self.assertEqual(reviewed['certificate_status'], 'ready')
                self.assertFalse(any('COD Date is in the future' in warning for warning in reviewed['warnings']))
        self.assertEqual(self.execute()[0],1)
        self.assertIn(AS_OF.strftime('%d.%m.%Y'), self.pdf_text())

    def test_shared_validation_and_colour_rules_keep_non_status_ranges(self):
        def rules(book):
            sheet=book.active
            for refs in ('L2:M20','M25:O30'):
                rule=DataValidation(type='custom',formula1='L2>0');sheet.add_data_validation(rule);rule.add(refs)
            for refs in ('J2:M20','M25:O30'):
                sheet.conditional_formatting.add(refs,FormulaRule(formula=['$M2="Ready"'],
                    fill=PatternFill(fill_type='solid',fgColor='FFABCDEF')))
        self.mutate(rules)
        with WorkbookSession(self.paths[0],self.root/'Data/Workbook Backups') as session:
            session.ensure_alert_formatting('Inward Register',as_of=AS_OF)
        book=load_workbook(self.paths[0]);sheet=book.active
        validations={str(rule.sqref):rule.formula1 for rule in sheet.data_validations.dataValidation}
        self.assertEqual(validations['P2:P20'],'P2>0')
        self.assertEqual(validations['K25:L30'],'P2>0')
        self.assertIn('Q2:Q10000',validations)
        colours={str(ref.sqref):[rule.formula for rule in rules] for ref,rules in sheet.conditional_formatting._cf_rules.items()}
        self.assertEqual(colours['J2:J20 O2:P20'],[['$Q2="Ready"']])
        self.assertEqual(colours['K25:L30'],[['$Q2="Ready"']])
        book.close()

    def test_bad_other_plant_keeps_readable_companies_visible_and_blocks_issuance(self):
        self.paths[1].write_bytes(b'not an xlsx')
        before=self.paths[0].read_bytes()
        self.assertEqual(self.execute()[0],0)
        report=self.report();self.assertFalse(report['scan_complete'])
        company=next(row for row in report['rows'] if row['sn']=='A')
        self.assertTrue(company['issuance_blocked'])
        for row in report['rows']:
            if row['source_kind']=='system':
                self.assertFalse(any('inward date' in warning.lower() for warning in row['warnings']))
        self.assertEqual(self.paths[0].read_bytes(),before)
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))

    def test_actual_windows_lock_keeps_other_plant_visible_without_writes(self):
        make_workbook(self.paths[1],[dict(sn='B',status='Ready')])
        before=[path.read_bytes() for path in self.paths]
        with exclusive_file(self.paths[0]): self.execute(check_only=True)
        report=self.report()
        self.assertFalse(report['scan_complete'])
        self.assertTrue(any(row['sn']=='B' for row in report['rows']))
        self.assertEqual([path.read_bytes() for path in self.paths],before)

    def test_file_disappearing_after_scan_returns_a_readable_uncached_state(self):
        real_run=web_app.run
        def vanish(*args,**kwargs):
            result=real_run(*args,**kwargs);self.paths[0].unlink();return result
        with patch.object(web_app,'run',side_effect=vanish): state=web_app.refresh(self.paths)
        self.assertFalse(state['scan_complete'])
        self.assertIn('became unavailable',state['setup_warning'])
        self.assertNotIn(str(self.root),web_app.WORKBOOK_CHECK_STAMPS)
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))

    def test_correction_keeps_summary_formula_and_updates_numeric_summary(self):
        for total in (13,'=SUM(H2:H3)'):
            make_workbook(self.paths[0],[dict(sn='A',status='Ready')])
            self.mutate(lambda book:self.total(book,total))
            source=read_batches(self.paths[0])[0]
            details=web_app.editor_details(source)
            details['items'][1]['weight']='8'
            with WorkbookSession(self.paths[0],self.root/'Data/Workbook Backups') as session:
                staged,proposed=stage_certificate_details(session,details)
            self.assertEqual(proposed['total_weight'],15)
            self.assertFalse(proposed['blocking_issues'])
            book=load_workbook(BytesIO(staged))
            self.assertEqual(book.active['E4'].value,'Total')
            self.assertEqual(book.active['H4'].value,15 if total==13 else total)
            book.close()


if __name__=='__main__': unittest.main()
