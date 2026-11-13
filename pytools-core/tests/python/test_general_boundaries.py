"""Synthetic boundary cases. Invoked with the current Cargo-built CLI, never dist EXE."""
import hashlib,json,os,subprocess,sys,tempfile,unittest
from pathlib import Path
import openpyxl
from openpyxl.styles import Border,Side
ROOT=Path(__file__).resolve().parents[3]
def digest(p): return hashlib.sha256(p.read_bytes()).hexdigest()
class Boundaries(unittest.TestCase):
 def capture(self,name,out):
  p=Path(os.environ.get('RUSTTOOLS_BOUNDARY_CAPTURE_DIR',str(ROOT/'test-runs/general-boundary-20261113')))/name;p.mkdir(parents=True,exist_ok=True)
  (p/'stdout.txt').write_text(out.stdout,encoding='utf-8');(p/'stderr.txt').write_text(out.stderr,encoding='utf-8');(p/'exit.txt').write_text(str(out.returncode))
 def cli(self,config,commands,name):
  env=os.environ.copy();env['PYTOOLS_OPENPYXL_PYTHON']=sys.executable;env['PYTOOLS_OPENPYXL_WORKER']=str(ROOT/'crates/excel-editor/python/openpyxl_worker.py')
  out=subprocess.run([os.environ['RUSTTOOLS_BOUNDARY_CLI'],'menu','--config',str(config),'--log-view','user'],input=commands,capture_output=True,text=True,encoding='utf-8',env=env,timeout=90)
  self.capture(name,out);self.assertEqual(out.returncode,0,out.stdout+out.stderr);return out.stdout+out.stderr
 def config(self,d,target=None,prefix=False):
  p=d/'config.xlsx';w=openpyxl.Workbook();g=w.active;g.title='全局配置';g.append(['配置项','值'])
  for key,value in [('输入目录',str(d)),('输出目录',str(d/'output')),('日志目录',str(d/'logs')),('错误策略','fail_fast')]:g.append([key,value])
  if target is not None:
   s=w.create_sheet('去重追加数据配置');s.append(['是否启用','任务名','源数据工作簿','源数据工作表','标识列序号','目标工作簿','目标工作表','执行模式']);s.append(['Y','Synthetic',str(d/'source.xlsx'),'Data','1',str(target),'Combined','追加'])
  if prefix:
   s=w.create_sheet('重命名配置');s.cell(2,15,'Y');s.cell(2,16,'Data@A1');s=w.create_sheet('文件排序配置');s.append(['是否启用','排序前缀','匹配关键词','备注']);s.append(['Y','001','original','synthetic'])
  w.save(p);return p
 def print_mode(self,mode):
  rust=(ROOT/'pytools-core/src/static_print_engines.rs').read_text(encoding='utf-8');script=rust.split('    let script = r#"',1)[1].split('"#;',1)[0]
  with tempfile.TemporaryDirectory() as folder:
   d=Path(folder);src=d/'source.xlsx';target=d/'target.xlsx';w=openpyxl.Workbook();a=w.active;a.title='first';a['A1']='old';a.merge_cells('A1:B1');a.row_dimensions[1].height=70;a.column_dimensions['A'].width=50
   b=w.create_sheet('second');b['A1']='left';b['B1']='right';b['A2']=0;b['B2']='0';b['AA10']='00123';b['A100']='12345678901234567890';b.row_dimensions[1].height=22;b.column_dimensions['A'].width=17;b['B1'].border=Border(left=Side(style='thin'));w.save(src);before=digest(src)
   tasks=[]
   for i,(sheet,area) in enumerate([('first','A1:B1'),('second','A1:AA100')]):tasks.append(dict(row_no=i+2,mode=mode,source_wb=str(src),source_ws=sheet,source_ranges=area,target_wb=str(target),target_ws='same',fit_wide=1,fit_tall=2,orientation='横向',center_h=True,center_v=False,drop_comments=True,hide_zero=True))
   plan=d/'plan.json';plan.write_text(json.dumps(dict(tasks=tasks)),encoding='utf-8');out=subprocess.run([sys.executable,'-I','-c',script,str(plan)],capture_output=True,text=True,encoding='utf-8',timeout=60);self.capture('print-mode-'+str(mode),out);self.assertEqual(out.returncode,0,out.stderr);report=json.loads(out.stdout);self.assertEqual((report['task_ok'],report['task_failed']),(2,0),report)
   config=self.config(d);settings=openpyxl.load_workbook(config);sheet=settings.create_sheet('打印配置');sheet.append(['是否启用','打印模式','源工作簿','源工作表','源工作表打印区域','目标工作簿','目标工作表','FitToPagesWide','FitToPagesTall','打印方向','水平居中','垂直居中','不输出批注','零值不输出']);cli_target=d/'cli-print.xlsx'
   for name,area in [('first','A1:B1'),('second','A1:AA100')]:sheet.append(['Y',mode,str(src),name,area,str(cli_target),'same',1,2,'横向','Y','N','Y','Y'])
   settings.save(config);text=self.cli(config,'3\n1\n0\n0\n','print-menu-mode-'+str(mode));self.assertNotIn('执行失败',text)
   ws=openpyxl.load_workbook(cli_target).active;self.assertEqual((ws['A1'].value,ws['B1'].value),('left','right'));self.assertFalse(list(ws.merged_cells.ranges));self.assertIsNone(ws['A2'].value);self.assertEqual(ws['B2'].value,'0');self.assertEqual(ws['AA10'].value,'00123');self.assertEqual(ws['A100'].value,'12345678901234567890')
   self.assertEqual(ws.row_dimensions[1].height,22 if mode==1 else None);self.assertEqual(ws.column_dimensions['A'].width,17 if mode in (1,2) else 13);self.assertEqual(ws['B1'].border.left.style if ws['B1'].border.left else None,'thin' if mode in (1,2) else None)
   self.assertEqual(str(ws.print_area),"'same'!$A$1:$AA$101");self.assertEqual(ws.page_setup.orientation,'landscape');self.assertEqual((ws.page_setup.fitToWidth,ws.page_setup.fitToHeight),(1,2));self.assertTrue(ws.print_options.horizontalCentered);self.assertFalse(ws.print_options.verticalCentered);self.assertEqual(ws.oddFooter.center.text,'第 1 页 / 共 1 页');self.assertEqual(digest(src),before)
 def test_print_mode_1(self):self.print_mode(1)
 def test_print_mode_2(self):self.print_mode(2)
 def test_print_mode_3(self):self.print_mode(3)
 def test_xlsx_append_real_menu_preserves_old_rows_and_repeat_is_noop(self):
  with tempfile.TemporaryDirectory() as folder:
   d=Path(folder);s=d/'source.xlsx';t=d/'target.xlsx';w=openpyxl.Workbook();w.active.title='Data';w.active.append(['ID','Amount']);w.active.append(['001',10]);w.active.append(['002',20]);w.save(s);w=openpyxl.Workbook();w.active.title='Combined';w.active.append(['ID','Amount']);w.active.append(['001',99]);w.save(t);sb=digest(s);config=self.config(d,t);cb=digest(config)
   text=self.cli(config,'2\n3\n0\n','xlsx-append');self.assertNotIn('执行失败',text);rows=list(openpyxl.load_workbook(t,data_only=False).active.values);self.assertEqual(rows,[('ID','Amount'),('001',99),('002',20)])
   text=self.cli(config,'2\n3\n0\n','xlsx-append-repeat');self.assertNotIn('执行失败',text);self.assertEqual(list(openpyxl.load_workbook(t).active.values),rows);self.assertEqual(digest(s),sb);self.assertEqual(digest(config),cb)
 @unittest.skipUnless(os.environ.get('RUSTTOOLS_TEST_COM')=='1','requires explicit COM synthetic-fixture authorization')
 def test_existing_xls_rejected_in_real_menu_with_real_biff_fixture(self):
  import win32com.client
  with tempfile.TemporaryDirectory() as folder:
   d=Path(folder);t=d/'target.xls';s=d/'source.xlsx';w=openpyxl.Workbook();w.active.title='Data';w.active.append(['ID','Amount']);w.active.append(['002',20]);w.save(s)
   app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;book=None
   try:
    book=app.Workbooks.Add();sheet=book.Worksheets(1);sheet.Name='Combined';sheet.Cells(1,1).Value='ID';sheet.Cells(1,2).Value='Amount';sheet.Cells(2,1).Value='001';sheet.Cells(2,2).Value=99;book.SaveAs(str(t),FileFormat=56);book.Close(False);book=None
   finally:
    if book is not None:book.Close(False)
    app.Quit()
   self.assertEqual(t.read_bytes()[:8],bytes.fromhex('d0cf11e0a1b11ae1'));tb=digest(t);sb=digest(s);config=self.config(d,t);cb=digest(config)
   text=self.cli(config,'2\n3\n0\n','xls-refusal');self.assertIn('既有 XLS 去重追加已安全停用',text);self.assertEqual(digest(t),tb);self.assertEqual(digest(s),sb);self.assertEqual(digest(config),cb);self.assertEqual(list(d.glob('*.xls')),[t]);self.assertFalse(any('candidate' in p.name.lower() and p.suffix.lower()=='.xls' for p in d.rglob('*')))
 def test_delete_refuses_cross_sheet_formula_and_merged_duplicate_without_changes(self):
  from openpyxl.workbook.defined_name import DefinedName
  with tempfile.TemporaryDirectory() as folder:
   d=Path(folder);source=d/'source.xlsx';w=openpyxl.Workbook();a=w.active;a.title='Data';a.append(['ID','Amount','note','extra']);a.append(['001',10,'first','safe']);a.append(['001',20,'duplicate',None]);a.merge_cells('C3:D3');a.append(['002',30,'keep','tail']);w.create_sheet('References')['A1']='=Data!B4';w.defined_names.add(DefinedName('KeepRows',attr_text="'Data'!$A$1:$D$4"));w.save(source);before=digest(source);config=self.config(d,d/'unused.xlsx');cb=digest(config)
   text=self.cli(config,'2\n2\n0\n','delete-structural-refusal');self.assertIn('执行失败',text);self.assertNotIn('删除重复行数：1',text);self.assertEqual(digest(source),before);self.assertEqual(digest(config),cb);self.assertFalse(any('pytools-dedup-' in p.name for p in d.rglob('*')))
 @unittest.skipUnless(sys.platform=='win32','Windows share-lock semantics')
 def test_delete_locked_source_fails_without_counting_or_losing_rows(self):
  import win32file,win32con
  with tempfile.TemporaryDirectory() as folder:
   d=Path(folder);source=d/'source.xlsx';w=openpyxl.Workbook();w.active.title='Data';w.active.append(['ID','Amount']);w.active.append(['001',10]);w.active.append(['001',20]);w.save(source);before=digest(source);config=self.config(d,d/'unused.xlsx')
   handle=win32file.CreateFile(str(source),win32con.GENERIC_READ,win32con.FILE_SHARE_READ,None,win32con.OPEN_EXISTING,0,None)
   try:
    text=self.cli(config,'2\n2\n0\n','delete-locked-source');self.assertIn('执行失败',text);self.assertNotIn('删除重复行数：1',text);self.assertEqual(digest(source),before)
   finally:handle.Close()
   self.assertFalse(any('pytools-dedup-' in p.name for p in d.rglob('*')))
 def test_prefix_real_menu_middle_text_sorted_undo_and_collision(self):
  with tempfile.TemporaryDirectory() as folder:
   d=Path(folder);s=d/'memo_A_file.xlsx';w=openpyxl.Workbook();w.active.title='Data';w.active['A1']='A';w.save(s);before=digest(s);config=self.config(d,prefix=True)
   text=self.cli(config,'4\n3\n1\n1\n'+str(s)+'\n0\n0\n','prefix-add');renamed=d/'A_memo_A_file.xlsx';self.assertTrue(renamed.exists(),text);self.assertEqual(digest(renamed),before)
   self.cli(config,'4\n3\n1\n1\n'+str(renamed)+'\n0\n0\n','prefix-repeat');self.assertTrue(renamed.exists());self.assertFalse((d/'A_A_memo_A_file.xlsx').exists())
   s.write_bytes(b'preexisting conflict');self.cli(config,'4\n3\n2\n1\n'+str(renamed)+'\n0\n0\n','prefix-undo-conflict');self.assertEqual(s.read_bytes(),b'preexisting conflict');self.assertEqual(digest(d/'memo_A_file_1.xlsx'),before)
   sorted_name=d/'001_A_original.xlsx';sorted_name.write_bytes((d/'memo_A_file_1.xlsx').read_bytes());self.cli(config,'4\n3\n2\n1\n'+str(sorted_name)+'\n0\n0\n','prefix-sorted-undo');self.assertEqual(digest(d/'001_original.xlsx'),before)
if __name__=='__main__':unittest.main(verbosity=2)
