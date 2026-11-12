"""Real 4-5 business acceptance using the formula-bearing user sample."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import openpyxl
from excel_edit_phase2 import digest, package, com_open

root = Path(__file__).resolve().parents[3]
source = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Tools\测试用例\2\目标用例.xlsx')
evidence = root / 'artifacts/excel-edit-phase2'
tag=os.environ.get('PHASE2_RUN_TAG','')
scratch = Path(tempfile.mkdtemp(prefix='rusttools-phase2-rename-'))
before_hash = digest(source)
before_parts, before_formulas = package(source)
original = openpyxl.load_workbook(source)
old_name = max(original, key=lambda s:s.max_row*s.max_column).title
original.close()
new_name = '验收改名'
runs = []
for mode in ['LegacyDefault', 'OpenpyxlOnly', 'UMYAOnly', 'Auto']:
    folder = scratch/mode
    folder.mkdir()
    target = folder/'source.xlsx'
    shutil.copy2(source,target)
    target_before_hash = digest(target)
    config = openpyxl.Workbook()
    sheet = config.active
    sheet.title = '重命名配置'
    sheet.append(['原表名','新表名'])
    sheet.append([old_name,new_name])
    config.save(folder/'config.xlsx')
    env = os.environ.copy()
    env.pop('PYTOOLS_EXCEL_EDIT_ROUTE',None)
    if mode != 'LegacyDefault':
        env['PYTOOLS_EXCEL_EDIT_ROUTE']=mode
    env['PYTOOLS_COM_PYTHON']=sys.executable
    env['PYTOOLS_COM_WORKER']=str(root/'pytools-core/python/pywin32_excel_conversion_worker.py')
    env['PYTOOLS_OPENPYXL_PYTHON']=sys.executable
    env['PYTOOLS_OPENPYXL_WORKER']=str(root/'crates/excel-editor/python/openpyxl_worker.py')
    proc = subprocess.run([r'F:\RustBuild\pytools-feature\debug\RustTools.exe',
                           'rename-sheets','--config',str(folder/'config.xlsx'),str(target)],
                          env=env,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=120)
    log=proc.stdout+'\n[STDERR]\n'+proc.stderr
    prefix=f'{tag}-' if tag else ''
    (evidence/(prefix+'4-5-'+mode+'.log')).write_text(log,encoding='utf-8')
    after_parts,after_formulas=package(target)
    expected={part:[(cell,formula.replace(old_name,new_name),cache) for cell,formula,cache in cells]
              for part,cells in before_formulas.items()}
    book=openpyxl.load_workbook(target)
    renamed=new_name in book.sheetnames and old_name not in book.sheetnames
    book.close()
    record={'menu':'4-5','mode':mode,'input':str(source),'output_copy':str(target),
            'exit':proc.returncode,'renamed':renamed,
            'formula_and_cache_equal_original':before_formulas==after_formulas,
            'formula_and_cache_expected_after_rename':expected==after_formulas,
            'output_copy_unchanged':digest(target)==target_before_hash,
            'dynamic_reference_rejected':('INDIRECT selector' in log or
                'unsupported_dynamic_sheet_reference' in log or
                'incomplete rewrite' in log),
            'changed_parts':sorted(n for n in before_parts.keys()|after_parts.keys()
                                   if before_parts.get(n)!=after_parts.get(n)),
            'com':com_open(target),'stdout':proc.stdout,'stderr':proc.stderr,
            'original_unchanged':digest(source)==before_hash}
    runs.append(record)
    result_name=f'rename-results-{tag}.json' if tag else 'rename-results.json'
    (evidence/result_name).write_text(json.dumps(runs,ensure_ascii=False,indent=2),encoding='utf-8')
    print(mode,'renamed',renamed,'formulas/cache original',record['formula_and_cache_equal_original'],
          'copy_unchanged',record['output_copy_unchanged'],
          'dynamic_rejected',record['dynamic_reference_rejected'],
          'COM',record['com']['status'],flush=True)
assert digest(source)==before_hash
