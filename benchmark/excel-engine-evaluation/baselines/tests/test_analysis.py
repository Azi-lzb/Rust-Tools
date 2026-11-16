import importlib.util
import json
import sys
import tempfile
import unittest
import zipfile
import subprocess
import xml.etree.ElementTree as ET
from pathlib import Path
import openpyxl

BASE=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(BASE))
# Explicit file load avoids the fixtures/ and capabilities module name collision.
spec=importlib.util.spec_from_file_location('baseline_capabilities',BASE/'capabilities.py')
cap=importlib.util.module_from_spec(spec); spec.loader.exec_module(cap)
from compare.package_diff import package_diff
from runner import summarize


class AnalysisTests(unittest.TestCase):
    def test_exclusive_trial_sst_filtered_but_user_shared_string_retained(self):
        root=BASE/'results/capabilities-windows-20261114-verified'
        source=root/'edit-input.xlsx'; target=root/'aspose-noop.xlsx'
        before=target.read_bytes(); view=cap.eval_view(source,target)
        removed=view['excluded_exclusive_evaluation_strings']
        self.assertTrue(any('Aspose' in x['text'] and 'Evaluation' in x['text'] for x in removed))
        self.assertEqual(target.read_bytes(),before)
        marker=next(x['original_index'] for x in removed if 'Aspose' in x['text'])
        package=cap.parts(target)
        xml=ET.fromstring(package['xl/worksheets/sheet1.xml'])
        cell=next(c for c in xml.iter(f'{{{cap.MAIN}}}c') if c.get('r')=='A1')
        cell.set('t','s'); cell.find(f'{{{cap.MAIN}}}v').text=str(marker)
        package['xl/worksheets/sheet1.xml']=ET.tostring(xml,encoding='utf-8')
        with tempfile.TemporaryDirectory() as tmp:
            clone=Path(tmp)/'shared-marker-copy.xlsx'
            with zipfile.ZipFile(clone,'w',zipfile.ZIP_DEFLATED) as z:
                for n,b in package.items(): z.writestr(n,b)
            shared=cap.eval_view(source,clone)
            self.assertNotIn(marker,[x['original_index'] for x in shared['excluded_exclusive_evaluation_strings']])

    def test_saved_capability_outputs_have_intended_edit_and_preserve_other_values(self):
        root=BASE/'results/capabilities-windows-20261114-verified'
        report=json.loads((root/'report.json').read_text('utf-8'))
        source=cap.snapshot(root/'edit-input.xlsx')
        for case in report['edits']:
            with self.subTest(engine=case['engine'],operation=case['operation']):
                prefix='ooxml' if case['engine']=='ooxml-patch' else case['engine']
                actual=cap.snapshot(root/f'{prefix}-{case["operation"]}.xlsx')
                for sheet in ('Data','Calc'):
                    expected_values={k:(v['value'],v['type']) for k,v in source[sheet]['cells'].items()}
                    if sheet=='Data' and case['operation']=='value': expected_values['B12']=(999,'n')
                    self.assertEqual({k:(v['value'],v['type']) for k,v in actual[sheet]['cells'].items()},expected_values)
                    self.assertEqual(actual[sheet]['merges'],source[sheet]['merges'])
                if case['operation']=='style': self.assertTrue(actual['Data']['cells']['B12']['bold'])
                if case['operation']=='header': self.assertEqual(actual['Data']['header'],source['Data']['header'].replace('synthetic','edited'))
                if case['engine']=='ooxml-patch':
                    self.assertEqual([r['part'] for r in case['raw_package']['changed']],['xl/worksheets/sheet1.xml'])
                    self.assertFalse(case['raw_package']['added'] or case['raw_package']['removed'])

    def test_validation_bounds_loss_is_never_default_normalized(self):
        base=[['B2:B16','whole','between','0','1000']]
        self.assertEqual(cap.classify(dict(path='/Data/validations',before=base,
            after=[['B2:B16','whole',None,'0','1000']])), 'explicit_default_validation_operator')
        self.assertEqual(cap.classify(dict(path='/Data/validations',before=base,
            after=[['B2:B16','whole','between',None,None]])), 'DATA_VALIDATION_BOUNDS_LOST')

    def test_rust_formula_refusal_preserves_input_and_has_no_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'input.xls'; cap.make_formula(source)
            before=source.read_bytes(); target=root/'rejected.xlsx'
            requests=root/'requests.json'
            requests.write_text(json.dumps([dict(op='formula',input=str(source),output=str(target))]),'utf-8')
            p=subprocess.run(['F:/RustBuild/pytools/release/capabilities.exe',str(requests)],capture_output=True,text=True)
            self.assertEqual(p.returncode,0,p.stderr)
            response=json.loads(p.stdout)
            self.assertEqual(response['status'],'UNSUPPORTED')
            self.assertEqual(len(response['cells']),4)
            self.assertFalse(target.exists())
            self.assertEqual(source.read_bytes(),before)

    def test_part_diff_ignores_container_compression_but_detects_part_change(self):
        with tempfile.TemporaryDirectory() as tmp:
            a,b,c=[Path(tmp)/n for n in ('a.xlsx','b.xlsx','c.xlsx')]
            for path,compress,body in [(a,zipfile.ZIP_STORED,b'<a/>'),(b,zipfile.ZIP_DEFLATED,b'<a/>'),(c,zipfile.ZIP_DEFLATED,b'<b/>')]:
                with zipfile.ZipFile(path,'w',compress) as z: z.writestr('x.xml',body)
            self.assertFalse(package_diff(a,b)['changed'])
            self.assertEqual(package_diff(a,c)['changed'][0]['part'],'x.xml')

    def test_vendor_view_never_modifies_raw_or_filters_original_sheet(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); source=root/'in.xlsx'; target=root/'out.xlsx'
            wb=openpyxl.Workbook(); wb.active.title='Data'; wb.save(source)
            wb.create_sheet('Evaluation Warning')['A1']='Evaluation Only. Created with Aspose.Cells for .NET.'; wb.save(target)
            before=target.read_bytes(); view=cap.eval_view(source,target)
            self.assertEqual(view['excluded_sheets'],['Evaluation Warning'])
            self.assertEqual(before,target.read_bytes())
            self.assertEqual(cap.eval_view(target,target)['excluded_sheets'],[])

    def test_biff_fixture_proves_empty_formula_caches(self):
        with tempfile.TemporaryDirectory() as tmp:
            result=cap.make_formula(Path(tmp)/'input.xls')
            self.assertEqual(len(result['biff_formula_caches']),4)
            self.assertTrue(all(x['empty_cache'] for x in result['biff_formula_caches']))

    def test_failed_or_missing_round_not_reported_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            records=[dict(id='tiny',op='read',engine='rust',phase='measured',processing_ms=i+1,
                          ok=True,validation=dict(ok=i!=9)) for i in range(10)]
            rows=summarize(records,Path(tmp))
            self.assertEqual(rows[0]['n'],9)
            self.assertEqual(rows[0]['status'],'INCOMPLETE_OR_FAILED')


if __name__=='__main__': unittest.main()
