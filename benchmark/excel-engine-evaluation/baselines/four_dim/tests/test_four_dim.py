import copy
import json
import math
import sys
import tempfile
import unittest
from pathlib import Path

HERE=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(HERE))
from fixtures import calc_spec,write_value
from run import calc_validation,summary
from structure import fingerprint,package_audit,custom_xml_stores,parts,xml_semantics
from chart_audit import resolved_bindings


class HarnessTests(unittest.TestCase):
    def test_controlled_profile_is_complete_and_uses_corrected_source(self):
        root=HERE.parent/'results/four-dim-20261114-controlled'
        env=json.loads((root/'run/environment.json').read_text('utf-8'))
        self.assertEqual(env['dotnet_profile'],'TieredCompilation=0; TieredPGO=0')
        stats=json.loads((root/'run/summary.json').read_text('utf-8'))
        self.assertEqual(len(stats),17)
        self.assertTrue(all(s['status']=='PASS' and s['n']==10 for s in stats))
        rows=[json.loads(s) for s in (root/'run/rounds.jsonl').read_text('utf-8').splitlines()]
        self.assertEqual(len(rows),221)
        self.assertEqual([r['workbook_count'] for r in rows if r['engine']=='aspose'],list(range(1,92)))
        s=json.loads((root/'run/structure-summary.json').read_text('utf-8'))
        self.assertEqual(len(s),6)
        self.assertTrue(all(x['preserved']==13 for x in s if x['engine']=='aspose'))
        native=next(x for x in s if x['condition']=='modify-header-narrow' and x['engine']=='rust')
        self.assertEqual(native['precise_patch'],13)

    def test_corrected_cohort_and_preservation_are_reported_separately(self):
        root=HERE.parent/'results/four-dim-20261114/modify-corrected'
        rows=[json.loads(s) for s in (root/'rounds.jsonl').read_text('utf-8').splitlines()]
        self.assertEqual(len(rows),78);self.assertTrue(all(r['validation']['ok'] for r in rows))
        self.assertEqual([r['workbook_count'] for r in rows if r['engine']=='aspose'],list(range(1,27)))
        s=json.loads((root/'structure-summary.json').read_text('utf-8'))
        self.assertTrue(all(x['outputs']==13 for x in s))
        self.assertTrue(all(x['preserved']==13 for x in s if x['engine']=='aspose'))
        native=next(x for x in s if x['condition']=='modify-header-narrow' and x['engine']=='rust')
        self.assertEqual(native['precise_patch'],13)
        charts=json.loads((root/'chart-audit.json').read_text('utf-8'))['outputs']
        self.assertTrue(all(x['match'] for x in charts if x['engine']=='aspose'))

    def test_chart_binding_normalization_keeps_raw_semantics_for_unknown_sheets(self):
        a=[dict(formulas=['Data!B1','Data!$B$2:$B$11'])]
        b=[dict(formulas=['Data!$B$1','Data!B2:B11'])]
        self.assertEqual(resolved_bindings(a),resolved_bindings(b))
        self.assertEqual(a[0]['formulas'][0],'Data!B1')
        other=[dict(formulas=["'Sheet$Name'!$B$1"])];self.assertEqual(resolved_bindings(other),other)

    def test_corrected_customxml_has_workbook_ownership_and_unchanged_sheet_parts(self):
        root=HERE.parent/'results/four-dim-20261114/fixtures'
        a=parts(root/'modify-complex.xlsx');b=parts(root/'modify-complex-workbook-customxml.xlsx')
        self.assertEqual(len(custom_xml_stores(a)),0)
        self.assertEqual(len(custom_xml_stores(b)),1)
        for name in a:
            if name.startswith(('xl/worksheets/','xl/charts/','xl/styles','xl/tables/')):self.assertEqual(a[name],b[name],name)
        self.assertNotEqual(xml_semantics(b'<a>pre<b/>tail</a>'),xml_semantics(b'<a>pre<b/>changed</a>'))

    def test_formula_fixture_caches_are_stale_in_both_formats(self):
        import openpyxl,xlrd
        root=HERE.parent/'results/four-dim-20261114/fixtures'
        wb=openpyxl.load_workbook(root/'calc-stale.xlsx',data_only=True)
        values=[v for row in wb['Calc'].iter_rows(values_only=True) for v in row];wb.close()
        self.assertEqual(values,[-9999]*2048)
        ws=xlrd.open_workbook(root/'calc-stale.xls').sheet_by_name('Calc')
        self.assertEqual([ws.cell_value(r,c) for r in range(256) for c in range(8)],values)

    def test_fixed_write_schema(self):
        self.assertEqual([write_value(0,i) for i in range(5)],[0,'row-000000-col-01',True,None,-4])
        self.assertEqual(write_value(49999,19),-50018)

    def test_calc_golden_is_not_derived_from_any_engine_output(self):
        s=calc_spec();self.assertEqual(s['formula_count'],2048)
        self.assertEqual(s['expected'][:7],[32897,16384,156,1,369,5,0])
        self.assertNotIn(-9999,s['expected'])
        self.assertTrue(calc_validation(s['expected'],s['expected'])['ok'])

    def test_calculation_rejects_boolean_nan_wrong_count_and_wrong_results(self):
        for values,expected in [([False],[0]),([float('nan')],[0]),([float('inf')],[0]),([],[0]),([1],[0])]:
            self.assertFalse(calc_validation(values,expected)['ok'])

    def test_invalid_round_cannot_win_or_get_pass_status(self):
        with tempfile.TemporaryDirectory() as tmp:
            records=[dict(condition='calc',engine='test',phase='measured',ok=True,processing_ms=i+1,validation=dict(ok=i!=9)) for i in range(10)]
            result=summary(records,Path(tmp),'test')[0]
            self.assertEqual(result['n'],9);self.assertEqual(result['total'],10)
            self.assertEqual(result['status'],'FAILED_OR_INCOMPLETE')

    def test_part_audit_does_not_hide_metadata_or_unknown_parts(self):
        a={'docProps/core.xml':b'<a time="1"/>','customXml/item1.xml':b'<keep/>','xl/sheet.xml':b'<a x="1" y="2"/>'}
        b={'docProps/core.xml':b'<a time="2"/>','xl/sheet.xml':b'<a y="2" x="1"/>'}
        result=package_audit(a,b,{})
        self.assertEqual(result['removed'],['customXml/item1.xml'])
        self.assertEqual(len(result['changed']),2)
        self.assertTrue(next(x for x in result['changed'] if x['part']=='xl/sheet.xml')['canonical_xml_equal'])
        c=dict(a);c['docProps/core.xml']=b'<a time="3"/>'
        self.assertEqual(fingerprint(a),fingerprint(c))
        c['customXml/item1.xml']=b'<changed/>';self.assertNotEqual(fingerprint(a),fingerprint(c))

    def test_current_measurements_have_all_rounds_and_single_aspose_lifecycle(self):
        root=HERE.parent/'results/four-dim-20261114/run'
        rows=[json.loads(s) for s in (root/'rounds.jsonl').read_text('utf-8').splitlines()]
        self.assertEqual(len(rows),221)
        self.assertEqual(sum(r['phase']=='measured' for r in rows),170)
        self.assertTrue(all(r['ok'] and r['validation']['ok'] for r in rows))
        self.assertEqual([r['workbook_count'] for r in rows if r['engine']=='aspose'],list(range(1,92)))
        for r in rows:self.assertGreaterEqual(r['request_e2e_ms']+1,r['processing_ms'])


if __name__=='__main__':unittest.main()
