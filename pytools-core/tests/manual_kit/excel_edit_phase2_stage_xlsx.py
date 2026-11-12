"""Stage user XLS samples through the existing conversion CLI, on copies only."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import openpyxl
from excel_edit_phase2 import digest

repo = Path(__file__).resolve().parents[3]
samples = Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Tools\测试用例\5')
scratch = Path(tempfile.mkdtemp(prefix='rusttools-phase2-special-'))
inputs = scratch/'xls-copies'
inputs.mkdir()
sources = list(samples.rglob('*.xls'))
hashes = {str(p):digest(p) for p in sources}
copies = []
for i, source in enumerate(sources):
    folder = inputs/str(i)
    folder.mkdir()
    target = folder/source.name
    shutil.copy2(source,target)
    copies.append(target)
config = openpyxl.Workbook()
sheet = config.active
sheet.title = '全局配置'
sheet.append(['配置项','值'])
for key,val in [('输入目录',str(inputs)),('输出目录',str(scratch/'xlsx')),
                ('日志目录',str(scratch/'logs'))]:
    sheet.append([key,val])
config_path = scratch/'config.xlsx'
config.save(config_path)
exe = scratch/'RustTools-test.exe'
shutil.copy2(Path(r'F:\RustBuild\pytools-feature\debug\RustTools.exe'),exe)
env = os.environ.copy()
env['PYTOOLS_COM_PYTHON'] = sys.executable
env['PYTOOLS_COM_WORKER'] = str(repo/'pytools-core/python/pywin32_excel_conversion_worker.py')
proc = subprocess.run([str(exe),'convert-excel','--config',str(config_path),
                       '--to','xlsx','--engine','pywin32','--output-dir',str(scratch/'xlsx'),
                       '--log-dir',str(scratch/'logs'),'--log-view','developer',
                       *map(str,copies)],capture_output=True,text=True,encoding='utf-8',
                      errors='replace',env=env,timeout=300)
evidence = repo/'artifacts/excel-edit-phase2'
(evidence/'special-staging.log').write_text(proc.stdout+'\n[STDERR]\n'+proc.stderr,encoding='utf-8')
result={'scratch':str(scratch),'exit':proc.returncode,'source_hashes':hashes,
        'originals_unchanged':all(digest(Path(p))==h for p,h in hashes.items()),
        'derived_xlsx':[{'path':str(p),'sha256':digest(p)} for p in (scratch/'xlsx').rglob('*.xlsx')]}
(evidence/'special-staging.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=True,indent=2))
assert result['originals_unchanged']
