"""Exercise existing special-report business plans on staged real XLSX copies."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from excel_edit_phase2 import digest, package, com_open

repo = Path(__file__).resolve().parents[3]
evidence = repo/'artifacts/excel-edit-phase2'
staged = json.loads((evidence/'special-staging.json').read_text(encoding='utf-8'))
data = Path(staged['scratch'])/'xlsx'
scratch = Path(tempfile.mkdtemp(prefix='rusttools-phase2-special-runs-'))
exe = scratch/'RustTools-test.exe'
shutil.copy2(Path(r'F:\RustBuild\pytools-feature\debug\RustTools.exe'),exe)
cases=[]
for mode in ['OpenpyxlOnly','UMYAOnly','BetterOfficeOnly','Auto']:
    for name in ['2026.09 本外币.xlsx','2026.09 外发用.xlsx']:
        cases.append(('5-1',mode,name))
for mode in ['OpenpyxlOnly','RustNativeOnly','UMYAOnly','Auto']:
    cases.append(('5-4',mode,'2026.09 外发用.xlsx'))
for mode in ['OpenpyxlOnly','RustNativeOnly','UMYAOnly','BetterOfficeOnly','Auto']:
    cases.append(('5-7',mode,'2026.09. 专项二批次.xlsx'))
only_menus=set(filter(None,os.environ.get('PHASE2_ONLY_MENUS','').split(',')))
if only_menus:
    cases=[case for case in cases if case[0] in only_menus]
tag=os.environ.get('PHASE2_RUN_TAG','')
results=[]
for index,(menu,mode,name) in enumerate(cases):
    folder=scratch/f'{index}-{menu}-{mode}'
    folder.mkdir()
    source=data/name
    target=folder/name
    shutil.copy2(source,target)
    env=os.environ.copy()
    env.update(PYTOOLS_EXCEL_EDIT_ROUTE=mode,PYTOOLS_OPENPYXL_PYTHON=sys.executable,
               PYTOOLS_OPENPYXL_WORKER=str(repo/'crates/excel-editor/python/openpyxl_worker.py'),
               PYTOOLS_COM_PYTHON=sys.executable,
               PYTOOLS_COM_WORKER=str(repo/'pytools-core/python/pywin32_excel_conversion_worker.py'))
    common=['--output-dir',str(folder/'output'),'--log-dir',str(folder/'logs')]
    if menu=='5-1':
        args=['village-split','--current',str(data/'2026.09 商业银行共享一批.xlsx'),
              '--previous',str(data/'2025.09 商业银行共享一批.xlsx'),
              '--in-place','openpyxl','--log-view','developer',*common,str(target)]
    elif menu=='5-4':
        args=['fx-header','--in-place','openpyxl','--view','developer',*common,str(target)]
    else:
        args=['rural-loan-plan','--current',str(target),
              '--previous',str(data/'2026.08. 专项二批次.xlsx'),
              '--output',str(folder/'candidate.xlsx'),'--plan-output',str(folder/'plan.json'),
              '--in-place','openpyxl']
    proc=subprocess.run([str(exe),*args],env=env,capture_output=True,text=True,
                        encoding='utf-8',errors='replace',timeout=180)
    log=proc.stdout+'\n[STDERR]\n'+proc.stderr
    prefix=f'{tag}-' if tag else ''
    log_path=evidence/f'{prefix}{index}-{menu}-{mode}.log'
    log_path.write_text(log,encoding='utf-8')
    old_parts,old_formulas=package(source)
    new_parts,new_formulas=package(target)
    changed=digest(source)!=digest(target)
    result={'menu':menu,'mode':mode,'input':str(source),'output_copy':str(target),
            'command':args,'exit':proc.returncode,'changed':changed,
            'staged_input_unchanged':digest(source)==next(x['sha256'] for x in staged['derived_xlsx'] if Path(x['path'])==source),
            'changed_parts':sorted(n for n in old_parts.keys()|new_parts.keys() if old_parts.get(n)!=new_parts.get(n)),
            'formula_and_cache_equal':old_formulas==new_formulas,
            'com':com_open(target),'terminal_log':str(log_path),
            'engine_lines':[line for line in log.splitlines() if 'ExcelEditPlan 路由' in line],
            'plan_files':[str(p) for p in folder.rglob('*.json') if 'plan' in p.name]}
    results.append(result)
    result_name=f'special-results-{tag}.json' if tag else 'special-results.json'
    (evidence/result_name).write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
    print(menu,mode,name.encode('unicode_escape').decode(),'exit',proc.returncode,
          'changed',changed,'COM',result['com']['status'],flush=True)
assert all(digest(Path(p))==h for p,h in staged['source_hashes'].items())
