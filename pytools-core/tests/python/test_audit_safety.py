import __future__, ast, json, subprocess, sys, tempfile, unittest
from pathlib import Path
import openpyxl
ROOT = Path(__file__).resolve().parents[3]
class AuditSafety(unittest.TestCase):
    def test_print_same_sheet_replaces_merges_and_dimensions(self):
        rust=(ROOT/'pytools-core/src/static_print_engines.rs').read_text(encoding='utf-8')
        script=rust.split('    let script = r#"',1)[1].split('"#;',1)[0]
        with tempfile.TemporaryDirectory() as d:
            d=Path(d); source=d/'source.xlsx'; target=d/'target.xlsx'
            w=openpyxl.Workbook(); a=w.active; a.title='first'; a['A1']='merged'; a.merge_cells('A1:B1'); a.row_dimensions[1].height=70; a.column_dimensions['A'].width=50
            b=w.create_sheet('second'); b['A1']='left'; b['B1']='right'; w.save(source)
            tasks=[]
            for i,name in enumerate(['first','second']):
                tasks.append(dict(row_no=i+2,mode=1,source_wb=str(source),source_ws=name,source_ranges='A1:B1',target_wb=str(target),target_ws='same',fit_wide=1,fit_tall=0,orientation='landscape',center_h=True,center_v=False,drop_comments=True,hide_zero=False))
            plan=d/'plan.json'; plan.write_text(json.dumps(dict(tasks=tasks)),encoding='utf-8')
            result=subprocess.run([sys.executable,'-I','-c',script,str(plan)],capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stderr)
            report=json.loads(result.stdout); self.assertEqual(report['task_failed'],0,report)
            ws=openpyxl.load_workbook(target).active
            self.assertEqual([ws['A1'].value,ws['B1'].value],['left','right'])
            self.assertFalse(list(ws.merged_cells.ranges)); self.assertNotEqual(ws.row_dimensions[1].height,70); self.assertNotEqual(ws.column_dimensions['A'].width,50)
            self.assertTrue(ws.print_area); self.assertTrue(ws.oddFooter.center.text)
    def test_worker_xls_append_refuses_before_any_io(self):
        text=(ROOT/'pytools-core/python/pywin32_excel_conversion_worker.py').read_text(encoding='utf-8')
        tree=ast.parse(text)
        method=next(n for n in ast.walk(tree) if isinstance(n,ast.FunctionDef) and n.name=='append_rows_com')
        method.decorator_list=[]
        module=ast.Module(body=[method],type_ignores=[]); env={'Any':object}
        exec(compile(ast.fix_missing_locations(module),'isolated-method','exec',flags=__future__.annotations.compiler_flag),env)
        with tempfile.TemporaryDirectory() as d:
            path=Path(d)/'synthetic.xls'; original=b'synthetic-placeholder-not-a-COM-fixture'; path.write_bytes(original)
            values=['text',12.5,True,None,{'date_time':{'serial':'45000','number_format':'yyyy-mm-dd'}},{'error':'#N/A'},{'formula':'=A12','cached':3}]
            for start in [1,12,100,0,-1]:
                for value in values:
                    with self.assertRaisesRegex(ValueError,'disabled'):
                        env['append_rows_com'](object(),str(path),'Sheet',start,[[value]*27])
                    self.assertEqual(path.read_bytes(),original)
            self.assertEqual(list(Path(d).iterdir()),[path])
    def test_glibc_ceiling_accepts_older_and_rejects_unknown(self):
        bash=Path('C:/Program Files/Git/bin/bash.exe')
        if not bash.exists(): bash=Path('/bin/bash')
        text=(ROOT/'pytools-core/packaging/build_deb.sh').read_text(encoding='utf-8')
        code=text.split('max_glibc="',1)[1].split('\necho "== staging',1)[0]
        code='max_glibc="'+code
        for version,good in [('2.17',True),('2.28',True),('2.28.0',True),('2.28.1',False),('2.9',True),('2.29',False),('',False)]:
            # Mock objdump only; never execute packaging, cargo or dpkg.
            payload='set -euo pipefail\nbinary=synthetic\nobjdump() { echo "GLIBC_'+version+'"; }\n'+code
            result=subprocess.run([str(bash)],input=payload,text=True,capture_output=True)
            self.assertEqual(result.returncode==0,good,(version,result.stdout,result.stderr))
    def test_release_guard_requires_main_and_explicit_acceptance(self):
        bash=Path('C:/Program Files/Git/bin/bash.exe')
        if not bash.exists(): bash=Path('/bin/bash')
        text=(ROOT/'pytools-core/packaging/build_deb.sh').read_text(encoding='utf-8')
        guard=text.split('case "${PACKAGE_KIND:-test}"',1)[1].split('esac',1)[0]
        guard='case "${PACKAGE_KIND:-test}"'+guard+'esac'
        for branch,kind,accepted,good in [('main','release','yes',True),('main','release','',False),('feature/business-development','release','yes',False),('test/business-coverage','test','',True)]:
            script='set -eu\nVERSION=synthetic\nbranch="'+branch+'"\nPACKAGE_KIND="'+kind+'"\nRELEASE_ACCEPTED="'+accepted+'"\n'+guard
            result=subprocess.run([str(bash)],input=script,text=True,capture_output=True)
            self.assertEqual(result.returncode==0,good,(branch,kind,result.stderr))
    def test_packaging_has_no_host_install_side_effects(self):
        text=(ROOT/'pytools-core/packaging/build_deb.sh').read_text(encoding='utf-8')
        self.assertNotIn('dpkg -i',text); self.assertNotIn('dpkg -P',text); self.assertNotIn('Z99',text)
        self.assertIn('PACKAGE_KIND:-test',text); self.assertIn('RELEASE_ACCEPTED',text)
        self.assertIn('sort -V',text); self.assertNotIn('"$max_glibc" != "2.28"',text)
if __name__=='__main__': unittest.main()
