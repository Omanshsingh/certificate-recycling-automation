"""Production dates, corrections, queue totals and drilldown source semantics."""
from copy import deepcopy
from datetime import date, datetime
from unittest import TestCase
from unittest.mock import Mock, patch

from certificate_analytics import build_analytics
import web_app


TODAY = date(2026, 10, 19)


def row(sn='1', **values):
    return dict(source_kind='workbook', plant='plant1', sn=sn, supplier='Client',
                inward_date='2026-10-10', certificate_status='ready', **values)


def event(cod, day='2026-10-19', **values):
    result = dict(id=str(cod)+day, state='complete', cod=str(cod), plant='plant1',
                  completed_at=day+'T12:00:00', after={'supplier':'Client','certificate_type':'standard','inward_date':'2026-10-10'})
    result.update(values)
    return result


class AnalyticsTests(TestCase):
    def build(self, rows=(), history=(), **kwargs):
        return build_analytics(list(rows), list(history), today=TODAY, **kwargs)

    def test_first_issue_deduplicates_zero_padded_cods_and_ignores_unfinished_publications(self):
        result = self.build(history=[event('0001','2026-09-25'), event('1'), event('0002'), event('3',state='published')])
        self.assertEqual(result['totals']['generated'],1)
        self.assertEqual(result['totals']['reissues'],1)
        self.assertEqual(result['totals']['lifetime_generated'],2)
        self.assertCountEqual([e['action'] for e in result['recent']],['Created','Reissued'])

    def test_receipts_exclude_local_and_supplemental_requests_and_keep_repeat_inwards(self):
        rows = [row('1'),row('2'),dict(row('3'),source_kind='manual',destruction_source={'sn':'1'}),dict(row('4'),source_kind='manual')]
        history = [event('1'),event('2',after={'supplier':' CLIENT  ','inward_date':'2026-10-11','certificate_type':'destruction'})]
        result = self.build(rows,history)
        self.assertEqual(result['totals']['received'],2)
        self.assertEqual(result['totals']['clients'],1)
        self.assertEqual(sum(x['count'] for x in result['types']),2)
        self.assertEqual(result['totals']['median_turnaround'],8.5)

    def test_calendar_week_is_monday_but_rolling_week_has_seven_days(self):
        result = self.build(history=[event('1','2026-10-18'),event('2'),event('3','2026-10-12')],period='week')
        self.assertEqual(result['calendar'],{'today':1,'week':1,'month':3})
        self.assertEqual(result['totals']['generated'],2)
        self.assertEqual(result['totals']['previous_generated'],1)
        self.assertEqual(len(result['series']),7)
        self.assertEqual(result['range']['previous_end'],'2026-10-12')

    def test_legacy_rows_use_certificate_date_and_no_date_is_not_invented(self):
        rows = [dict(row(),certificate_status='generated',certificate_no='0001',certificate_date='2026-10-03'),
                dict(row('2'),certificate_status='completed',certificate_no='2'),
                dict(row('3'),certificate_status='ready',certificate_no='3',certificate_date='2026-10-03')]
        result = self.build(rows)
        self.assertEqual(result['totals']['generated'],1)
        self.assertEqual(result['coverage']['fallback_dates'],1)
        self.assertEqual(result['coverage']['unknown_generation_dates'],1)
        self.assertEqual(result['recent'],[])

    def test_first_recorded_reissue_of_legacy_pdf_does_not_inflate_today(self):
        history = [event('1',previous_pdfs=[{'path':'old.pdf'}],before={'cod_no':'1','cod_date':'2026-09-01','issued':'Generated'}),
                   event('2',previous_pdfs=[{'path':'old2.pdf'}])]
        result = self.build(history=history)
        self.assertEqual(result['totals']['generated'],0)
        self.assertEqual(result['totals']['reissues'],2)
        self.assertEqual(result['calendar']['today'],0)
        self.assertEqual(result['coverage']['unknown_generation_dates'],1)

    def test_queue_and_ageing_do_not_double_count_linked_destruction_intakes(self):
        rows = [row(),dict(row('2'),certificate_status='generated'),dict(row('3'),certificate_status='completed'),
                dict(row('4'),certificate_status='na'),dict(row('5'),certificate_status='blank',inward_date='2026-08-01'),
                dict(row('6'),source_kind='manual',certificate_type='destruction',destruction_source={'sn':'1'},certificate_status='generated_manual'),
                dict(row('7'),warnings=['Check weight']),dict(row('8'),inward_date=''),
                dict(row('9'),inward_date='2026-10-20')]
        result = self.build(rows)
        self.assertEqual(result['queue'],{'ready':3,'attention':2,'to_send':2,'destruction':1,'overdue':1})
        self.assertEqual(next(x['count'] for x in result['ageing'] if x['label']=='unknown'),2)

    def test_plant_scope_applies_to_calendar_types_queue_and_history(self):
        result = self.build([row(),dict(row('2'),plant='plant2')], [event('1'),event('2',plant='plant2')],plant='plant2')
        self.assertEqual(result['calendar']['today'],1)
        self.assertEqual(result['totals']['received'],1)
        self.assertEqual(result['queue']['ready'],1)
        self.assertEqual(len(result['plants']),1)
        self.assertEqual(result['plants'][0]['key'],'plant2')
        self.assertEqual(result['recent'][0]['cod'],'2')

    def test_removed_hdd_request_is_not_a_pending_task_but_its_issued_certificate_still_counts(self):
        hdd = dict(row('HDD'), source_kind='manual', certificate_type='destruction',
                   certificate_status='generated_manual', destruction_requested=False)
        result = self.build([hdd], [event('0303', after={'supplier':'Client','certificate_type':'destruction'})])
        self.assertEqual(result['queue']['destruction'], 0)
        self.assertEqual(result['totals']['generated'], 1)

    def test_empty_period_has_real_zeros_and_no_turnaround_estimate(self):
        result = self.build(scan_complete=False)
        self.assertEqual(result['totals']['generated'],0)
        self.assertIsNone(result['totals']['median_turnaround'])
        self.assertFalse(result['coverage']['scan_complete'])
        self.assertEqual(len(result['series']),19)

    def test_input_is_immutable_and_future_dates_are_not_counted_as_activity(self):
        rows,history = [row()],[event('1','2026-10-20')]
        original = deepcopy((rows,history))
        result = self.build(rows,history)
        self.assertEqual((rows,history),original)
        self.assertEqual(result['totals']['generated'],0)

    def test_invalid_filters_are_rejected(self):
        for filters in ({'period':'invalid'},{'plant':'plant3'}):
            with self.assertRaises(ValueError): self.build(**filters)

    def test_event_dates_follow_the_computers_local_timezone(self):
        stamp = '2026-10-19T00:15:00+14:00'
        local = datetime.fromisoformat(stamp).astimezone().date()
        result = self.build(history=[event('1', completed_at=stamp)])
        self.assertEqual(result['recent'][0]['date'], local.isoformat())

    def test_issued_records_without_cods_are_disclosed_in_data_notes(self):
        result = self.build([dict(row(),certificate_status='completed',certificate_date='2026-10-01')])
        self.assertEqual(result['totals']['generated'],0)
        self.assertEqual(result['coverage']['missing_cod'],1)

    def test_recent_activity_only_links_to_a_current_company_record(self):
        result=self.build(history=[event('0001')])
        self.assertFalse(result['recent'][0]['has_current_record'])
        result=self.build([dict(row(),certificate_no='1')],[event('0001')])
        self.assertTrue(result['recent'][0]['has_current_record'])

    def test_endpoint_is_read_only_and_loads_all_history_not_the_recent_fifty(self):
        handler = object.__new__(web_app.Handler)
        handler.path = '/api/analytics?period=week&plant=plant2'
        handler.send_json,handler.send_error_json = Mock(),Mock()
        with patch.object(web_app,'latest_result',return_value={'rows':[], 'scan_complete':True}), \
                patch.object(web_app,'manual_rows',return_value=[]), \
                patch.object(web_app,'audit_records',return_value=[event(str(i),plant='plant2') for i in range(65)]), \
                patch.object(web_app,'scan_inventory',return_value={'entries':[]}), \
                patch.object(web_app,'refresh') as refresh, patch.object(web_app,'save_json') as save:
            handler._do_get()
        refresh.assert_not_called(); save.assert_not_called()
        handler.send_error_json.assert_not_called()
        self.assertEqual(handler.send_json.call_args.args[0]['totals']['lifetime_generated'],65)

    def test_saved_hdd_without_inward_counts_once_with_plant_unknown(self):
        saved = [{'cod':'0204','type':'destruction','date':TODAY,'plant':None}]*2
        result = self.build(saved_pdfs=saved)
        self.assertEqual(result['totals']['generated'],1)
        self.assertEqual(result['totals']['received'],0)
        self.assertEqual(result['queue']['to_send'],0)
        self.assertEqual(result['coverage']['no_plant'],1)
        self.assertEqual(next(item['count'] for item in result['types'] if item['key']=='destruction'),1)
        self.assertEqual(self.build(saved_pdfs=saved,plant='plant1')['totals']['generated'],0)

    def test_saved_pdf_does_not_duplicate_history_and_unknown_dates_are_not_today(self):
        saved = [{'cod':'0001','type':'destruction','date':TODAY,'plant':None},
                 {'cod':'2','type':'destruction','date':None,'plant':None}]
        result = self.build(history=[event('1')],saved_pdfs=saved)
        self.assertEqual(result['totals']['generated'],1)
        self.assertEqual(result['coverage']['folder_needs_details'],1)
        self.assertEqual(result['coverage']['folder_certificates'],0)
        self.assertEqual([item['label'] for item in result['types']],['Recycling','Data Destruction','Medical','International'])

    def test_comparison_line_matches_previous_total_and_includes_imports(self):
        saved = [{'cod':'3','type':'destruction','date':date(2026,9,20),'plant':'plant2'}]
        result = self.build(history=[event('1','2026-09-19'),event('2')],saved_pdfs=saved)
        self.assertEqual(result['totals']['previous_generated'],2)
        self.assertEqual(sum(item['previous_generated'] for item in result['series']),2)
