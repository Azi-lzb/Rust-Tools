"""Assertions on retained independent before/after evidence and saved outputs."""
import copy
import hashlib
import json
import sys
import unittest
from pathlib import Path
import openpyxl

BASE=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(BASE))
from capabilities import snapshot
from report import formula_parser_note, umya_validation_note

ROOT=BASE/'results/root-cause-20261114'


class RootCauseRegression(unittest.TestCase):
    def test_report_uses_current_evidence_instead_of_stale_failure_text(self):
        before=json.loads((BASE/'results/capabilities-windows-20261114-verified/report.json').read_text('utf-8'))
        after=json.loads((ROOT/'capabilities-postfix/report.json').read_text('utf-8'))
        self.assertIn('UNEXPECTED_PARSER_DIFFERENCE',formula_parser_note(before))
        self.assertIn('MATCH',formula_parser_note(after))
        self.assertNotIn('UNEXPECTED_PARSER_DIFFERENCE',formula_parser_note(after))
        self.assertIn('上下限丢失',umya_validation_note(before))
        self.assertIn('未发现验证上下限丢失',umya_validation_note(after))

    def test_two_xls_generators_reproduce_and_patch_matches_xlsx_control(self):
        before=json.loads((ROOT/'upstream-report.json').read_text('utf-8'))
        after=json.loads((ROOT/'patched-report.json').read_text('utf-8'))
        self.assertEqual([len(x['differences']) for x in before['formulas']],[41,40,0])
        self.assertEqual([x['cases'] for x in after['formulas']],[44,44,44])
        self.assertTrue(all(not x['differences'] for x in after['formulas']))
        self.assertTrue(all(not x['differences'] for x in after['independent_reader']))
        self.assertEqual(before['input_hashes'],after['input_hashes'])

    def test_validation_memory_was_correct_but_save_lost_bounds_now_preserved(self):
        before=json.loads((ROOT/'upstream-report.json').read_text('utf-8'))['validation']
        after=json.loads((ROOT/'patched-report.json').read_text('utf-8'))['validation']
        self.assertEqual(before['in_memory_before_save'][0]['formula1'],'0')
        self.assertEqual(before['in_memory_before_save'][0]['formula2'],'1000')
        self.assertEqual(before['after_reload'][0]['formula1'],'')
        self.assertEqual(before['after_reload'][0]['formula2'],'')
        self.assertTrue(after['raw_xml_equal'])
        self.assertEqual(after['raw_xml_before'],after['raw_xml_after'])
        self.assertEqual(after['in_memory_before_save'],after['after_reload'])

    def test_postfix_edits_keep_intended_values_styles_merges_and_validation_bounds(self):
        root=ROOT/'capabilities-postfix'
        report=json.loads((root/'report.json').read_text('utf-8'))
        source=snapshot(root/'edit-input.xlsx')
        self.assertEqual(len(report['edits']),12)
        for case in report['edits']:
            prefix='ooxml' if case['engine']=='ooxml-patch' else case['engine']
            actual=snapshot(root/f'{prefix}-{case["operation"]}.xlsx')
            with self.subTest(engine=case['engine'],op=case['operation']):
                for sheet in ('Data','Calc'):
                    expected={k:(v['value'],v['type']) for k,v in source[sheet]['cells'].items()}
                    if sheet=='Data' and case['operation']=='value': expected['B12']=(999,'n')
                    self.assertEqual({k:(v['value'],v['type']) for k,v in actual[sheet]['cells'].items()},expected)
                    self.assertEqual(actual[sheet]['merges'],source[sheet]['merges'])
                if case['operation']=='style': self.assertTrue(actual['Data']['cells']['B12']['bold'])
                if case['operation']=='header': self.assertEqual(actual['Data']['header'],source['Data']['header'].replace('synthetic','edited'))
                if case['engine']=='umya': self.assertEqual(actual['Data']['validations'],source['Data']['validations'])

    def test_calculation_is_explicit_and_rust_stays_unsupported_without_output(self):
        root=ROOT/'capabilities-postfix'; report=json.loads((root/'report.json').read_text('utf-8'))
        self.assertEqual(report['formula']['aspose_actual'],[8,68,6,5])
        self.assertEqual(report['formula']['formula_text_diagnostic']['status'],'MATCH')
        self.assertEqual(report['formula']['rust']['status'],'UNSUPPORTED')
        self.assertFalse((root/'rust-degraded.xlsx').exists())
        wb=openpyxl.load_workbook(root/'aspose-calculated.xlsx',data_only=True)
        self.assertEqual([wb['Calc'].cell(i,1).value for i in range(1,5)],[8,68,6,5]); wb.close()
        raw=[json.loads(s) for s in (root/'aspose.stdout.jsonl').read_text('utf-8').splitlines()]
        self.assertEqual(raw[-1]['calculation'],'explicit Workbook.CalculateFormula')

    def test_prior_workbooks_and_protected_config_and_dist_remain_identical(self):
        for entry in json.loads((BASE/'ARTIFACTS.json').read_text('utf-8')):
            self.assertEqual(hashlib.sha256((BASE/entry['path']).read_bytes()).hexdigest(),entry['sha256'],entry['path'])
        for entry in json.loads((BASE/'protected-after.json').read_text('utf-8')):
            self.assertEqual(hashlib.sha256(Path(entry['path']).read_bytes()).hexdigest(),entry['before_sha256'],entry['path'])


if __name__=='__main__': unittest.main()
