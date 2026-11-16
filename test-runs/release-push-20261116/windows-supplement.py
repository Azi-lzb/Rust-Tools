"""Explicitly run affected gated tests and Python discovery on the current CLI."""
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys

ROOT=Path(__file__).resolve().parents[2]
RUN=Path('F:/RustBuild/pytools/release-push-20261116')
log=(RUN/'windows-workspace.log').read_text(encoding='utf-8-sig',errors='replace')
assert (RUN/'windows-workspace.exit').exists(), 'wait for the full Windows run'
executables=list(dict.fromkeys(re.findall(r'\((F:[^)\r\n]+\.exe)\)',log)))
directory=RUN/'bin'
directory.mkdir(exist_ok=False)
cli=directory/'RustTools.exe'
shutil.copy2('F:/RustBuild/pytools/debug/RustTools.exe',cli)
fixtures=RUN/'office-fixtures'
fixtures.mkdir(exist_ok=True)
for name in ['source.xlsx','target.xls']:
    if (fixtures/name).exists():
        assert (fixtures/name).read_bytes() == (ROOT/'test-runs/platform-validation-20261113/excel-final-fixtures'/name).read_bytes()
    shutil.copy2(ROOT/'test-runs/platform-validation-20261113/excel-final-fixtures'/name,fixtures/name)
env=os.environ.copy()
python='F:/RustBuild/pytools/debug/python/python.exe'
env.update(PYTOOLS_OPENPYXL_PYTHON=python,RUST_WORKSPACE_OPENPYXL_PYTHON=python,
    PYTOOLS_OPENPYXL_WORKER=str(ROOT/'crates/excel-editor/python/openpyxl_worker.py'),
    PYTOOLS_COM_PYTHON=python,PYTOOLS_COM_WORKER=str(ROOT/'pytools-core/python/pywin32_excel_conversion_worker.py'),
    RUSTTOOLS_TEST_EXE=str(cli),RUSTTOOLS_LOGGING_CLI=str(cli),RUSTTOOLS_BOUNDARY_CLI=str(cli),
    RUSTTOOLS_BOUNDARY_CAPTURE_DIR=str(RUN/'python-cli'),
    RUSTTOOLS_OFFICE_FIXTURE_DIR=str(RUN/'office-fixtures'),
    RUSTTOOLS_ASPOSE_DOTNET='F:/RustBuild/pytools/aspose-worker-sdk-10.0.404/dotnet.exe',
    RUSTTOOLS_ASPOSE_WORKER_DLL='F:/RustBuild/pytools/aspose-worker-selfcontained/publish/AsposeWorker.dll',
    DOTNET_ROOT='F:/RustBuild/pytools/aspose-worker-sdk-10.0.404',
    ASPOSE_TEST_PYTHON='C:/Users/AZI/AppData/Local/Programs/Python/Python312/python.exe')
env.pop('PYTOOLS_EXCEL_EDIT_ROUTE',None)
for key in ['RUSTTOOLS_TEST_COM','RUSTTOOLS_OFFICE_PROGID','RUSTTOOLS_WPS_READONLY']:
    env.pop(key,None)
results=[]
jobs=[('aspose_menu_cli',None),('aspose_modify_cli',None),('openpyxl_worker',None),('config_vba_workbook_parity','real_vba_workbook_loader_parity'),
      ('conversion_pywin32_windows',None),
      ('conversion_menu_cli','conversion_menu_batch_uses_pywin32_and_preserves_sources'),
      ('sheet_rename_execution','legacy_xls_sheet_rename_uses_dedicated_com_route'),
      ('conversion_menu_cli','conversion_cli_converts_verified_biff_with_et_suffix_without_claiming_native_et'),
      ('conversion_menu_cli','conversion_cli_does_not_publish_biff_with_et_suffix_without_content_evidence')]
for index,(name,pattern) in enumerate(jobs):
    found=[p for p in executables if Path(p).name.startswith(name+'-')]
    assert len(found)==1,(name,found)
    command=[found[0],'--ignored','--test-threads=1']
    if pattern: command.extend([pattern,'--exact'])
    path=RUN/f'windows-gate-{index}-{name}.log'
    with path.open('w',encoding='utf-8') as output:
        process=subprocess.run(command,cwd=ROOT,env=env,stdout=output,stderr=subprocess.STDOUT,timeout=600)
    results.append({'name':name,'filter':pattern,'exit':process.returncode,'log':path.name})
    (RUN/'windows-gates.json').write_text(json.dumps(results,indent=2),encoding='utf-8')
    print(name,pattern,process.returncode,flush=True)
with (RUN/'windows-python.log').open('w',encoding='utf-8') as output:
    process=subprocess.run([python,'-X','utf8','-I','-m','unittest','discover','-s',
        'pytools-core/tests/python','-p','test_*.py','-v'],cwd=ROOT,env=env,
        stdout=output,stderr=subprocess.STDOUT,timeout=1200)
(RUN/'windows-python.exit').write_text(str(process.returncode),encoding='ascii')
print('Python discovery',process.returncode,flush=True)
raise SystemExit(1 if process.returncode or any(r['exit'] for r in results) else 0)
