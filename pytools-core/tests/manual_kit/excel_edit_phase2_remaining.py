"""Run additional existing business menus with copied production configuration."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import openpyxl
from excel_edit_phase2 import digest, com_open

repo=Path(__file__).resolve().parents[3]
evidence=repo/'artifacts/excel-edit-phase2'
main=Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Tools')
original_config=main/'config.xlsx'
config_hash=digest(original_config)
stage=json.loads((evidence/'special-staging.json').read_text(encoding='utf-8'))
special=Path(stage['scratch'])/'xlsx'/'2026.09 外发用.xlsx'
scratch=Path(tempfile.mkdtemp(prefix='rusttools-phase2-remaining-'))
exe=scratch/'RustTools-test.exe'
shutil.copy2(Path(r'F:\RustBuild\pytools-feature\debug\RustTools.exe'),exe)
faults='--fault-cases' in sys.argv
results=[]
cases=[('5-2',m) for m in ['OpenpyxlOnly','RustNativeOnly','UMYAOnly','BetterOfficeOnly','Auto']]
cases += [('5-3',m) for m in ['OpenpyxlOnly','UMYAOnly','BetterOfficeOnly','Auto']]
cases += [('7-1',m) for m in ['OpenpyxlOnly','Auto']]
cases += [('2-3',m) for m in ['OpenpyxlOnly','Auto']]
if faults:
    cases=[x for x in cases if x[0] in ['7-1','2-3']]
for index,(menu,mode) in enumerate(cases):
    folder=scratch/f'{index}-{menu}-{mode}'
    folder.mkdir()
    config_path=folder/'config.xlsx'
    b=openpyxl.load_workbook(original_config)
    g=b['全局配置']
    for row in g:
        if row[0].value in ['输入目录','输出目录','日志目录']:
            row[1].value=str(folder/str(row[0].value))
    b.save(config_path)
    b.close()
    source=folder/'source.xlsx'
    shutil.copy2(special,source)
    env=os.environ.copy()
    env.update(PYTOOLS_EXCEL_EDIT_ROUTE=mode,PYTOOLS_OPENPYXL_PYTHON=sys.executable,
               PYTOOLS_OPENPYXL_WORKER=str(repo/'crates/excel-editor/python/openpyxl_worker.py'),
               PYTOOLS_COM_PYTHON=sys.executable,
               PYTOOLS_COM_WORKER=str(repo/'pytools-core/python/pywin32_excel_conversion_worker.py'))
    def invoke(args, transcript=None, tag='main'):
        p=subprocess.run([str(exe),*args],input=transcript,env=env,text=True,
                         capture_output=True,encoding='utf-8',errors='replace',timeout=180)
        (evidence/f'{"fault-" if faults else ""}remaining-{index}-{menu}-{mode}-{tag}.log').write_text(
            p.stdout+'\n[STDERR]\n'+p.stderr,encoding='utf-8')
        return p
    if menu=='5-3':
        env['PYTOOLS_EXCEL_EDIT_ROUTE']='OpenpyxlOnly'
        prep=invoke(['institution-normalize','--config',str(config_path),'--apply',
                     '--in-place','openpyxl','--view','developer',str(source)],tag='normalization')
        env['PYTOOLS_EXCEL_EDIT_ROUTE']=mode
    if menu=='2-3':
        shutil.copy2(main/'测试用例/2/检查重复去重用例.xlsx',source)
        target=folder/'target.xlsx'
        shutil.copy2(main/'测试用例/2/目标用例.xlsx',target)
        b=openpyxl.load_workbook(config_path)
        del b['去重追加数据配置']
        sheet=b.create_sheet('去重追加数据配置')
        sheet.append(['是否启用','任务名','源数据工作簿','源数据工作表','标识列序号','目标工作簿','目标工作表','执行模式'])
        sheet.append(['是','real-append',str(source),'五篇大文章分机构','1;2;3;4;5',str(target),'五篇大文章分机构','追加'])
        b.save(config_path);b.close()
    preparation=[]
    if faults and menu=='7-1':
        b=openpyxl.load_workbook(config_path)
        found=False
        for ws in b:
            for row in ws:
                for cell in row:
                    if cell.comment is not None:
                        preparation.append({'kind':'remove_one_comment_for_repair','sheet':ws.title,'cell':cell.coordinate})
                        cell.comment=None
                        found=True
                        break
                if found:break
            if found:break
        assert found, 'no existing comment available for repair fixture'
        b.save(config_path);b.close()
    if faults and menu=='2-3':
        import datetime
        b=openpyxl.load_workbook(source)
        ws=b['五篇大文章分机构']
        values=[c.value for c in ws[ws.max_row]]
        values[2]=datetime.datetime(2026,10,31)
        ws.append(values)
        preparation.append({'kind':'add_one_unique_record','sheet':ws.title,'row':ws.max_row,'date':'2026-10-31'})
        b.save(source);b.close()
    before={str(p):digest(p) for p in folder.rglob('*.xlsx')}
    if menu=='5-2':
        args=['institution-normalize','--config',str(config_path),'--apply','--in-place',
              'openpyxl','--view','developer',str(source)]
        proc=invoke(args)
    elif menu=='5-3':
        args=['remove-foreign-bank','--config',str(config_path),'--create-copy',
              '--output-dir',str(folder/'output'),'--log-dir',str(folder/'logs'),
              '--view','developer',str(source)]
        proc=invoke(args)
    else:
        transcript='7\n1\n1\n0\n0\n' if menu=='7-1' else '2\n3\n0\n0\n'
        args=['menu','--config',str(config_path),'--log-view','developer']
        proc=invoke(args,transcript)
    after={str(p):digest(p) for p in folder.rglob('*.xlsx')}
    changed=[p for p,h in after.items() if before.get(p)!=h]
    result={'menu':menu,'mode':mode,'exit':proc.returncode,'case_directory':str(folder),
            'command':args,'preparation':preparation,'changed_or_created_files':changed,
            'com':{p:com_open(Path(p)) for p in changed},
            'has_error_line':any(t in proc.stdout+proc.stderr for t in ['[ERROR]','配置操作失败','状态 failed','"status": "failed"']),
            'engine_lines':[l for l in (proc.stdout+proc.stderr).splitlines() if 'ExcelEditPlan 路由' in l],
            'production_config_unchanged':digest(original_config)==config_hash}
    results.append(result)
    (evidence/('remaining-fault-results.json' if faults else 'remaining-results.json')).write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
    print(menu,mode,'exit',proc.returncode,'changed',len(changed),'error',result['has_error_line'],flush=True)
assert digest(original_config)==config_hash
