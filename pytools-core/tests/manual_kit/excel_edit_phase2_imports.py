"""Real-data 1-1/5-6 imports into an existing formula-bearing target workbook."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import openpyxl
from excel_edit_phase2 import digest, package, com_open

repo=Path(__file__).resolve().parents[3]
samples=Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Tools\测试用例\2')
evidence=repo/'artifacts/excel-edit-phase2'
scratch=Path(tempfile.mkdtemp(prefix='rusttools-phase2-imports-'))
exe=scratch/'RustTools-test.exe'
shutil.copy2(Path(r'F:\RustBuild\pytools-feature\debug\RustTools.exe'),exe)
originals={str(p):digest(p) for p in samples.glob('*.xlsx')}
results=[]
for menu in ['1-1','5-6']:
    for mode in ['OpenpyxlOnly','UMYAOnly','Auto']:
        folder=scratch/(menu+'-'+mode)
        folder.mkdir()
        source=folder/'source.xlsx'
        target=folder/'target.xlsx'
        shutil.copy2(samples/'检查重复去重用例.xlsx',source)
        shutil.copy2(samples/'目标用例.xlsx',target)
        before_hash=digest(target)
        old_parts,old_formulas=package(target)
        config=openpyxl.Workbook()
        sheet=config.active
        sheet.title='全局配置'
        sheet.append(['配置项','值'])
        for key,val in [('输入目录',str(folder)),('输出目录',str(folder/'output')),
                        ('日志目录',str(folder/'logs')),('错误策略','fail_fast')]:
            sheet.append([key,val])
        if menu=='1-1':
            task=config.create_sheet('时序提取规则')
            task.append(['是否启用','规则名称','工作簿关键字','工作表关键字','行头列',
                         '列表头行','必含列头','必含行头','数据起始行','数据结束行',
                         '数据起始列','数据结束列','跳过关键字','目标工作簿路径',
                         '目标工作表','启用目标写入','set区域','目标去重列','是否输出行序号'])
            task.append(['是','真实数据导入','source','五篇大文章分机构','A','1','','',
                         '2','463','F','J','',str(target),'业务汇总',1,'日期@C2','','是'])
            config.create_sheet('路径标准化映射').append(['是否启用','映射名称','作用对象'])
        else:
            task=config.create_sheet('工作表提取')
            task.append(['是否启用','工作簿关键字','工作表关键字','排除关键字','提取区域',
                         '多区域拼接方式','输出文件夹','文件名','目标工作表','备注'])
            task.append(['是','source','五篇大文章分机构','','','',str(folder),
                         'target.xlsx','业务提取','真实整表导入既有目标'])
        config_path=folder/'config.xlsx'
        config.save(config_path)
        env=os.environ.copy()
        env.update(PYTOOLS_EXCEL_EDIT_ROUTE=mode,PYTOOLS_OPENPYXL_PYTHON=sys.executable,
                   PYTOOLS_OPENPYXL_WORKER=str(repo/'crates/excel-editor/python/openpyxl_worker.py'))
        command=['wide-summary','--config',str(config_path),'--log-view','developer',str(source)] if menu=='1-1' else [
            'extract-sheets','--config',str(config_path),'--append-existing',
            '--log-view','developer',str(source)]
        proc=subprocess.run([str(exe),*command],env=env,capture_output=True,text=True,
                            encoding='utf-8',errors='replace',timeout=180)
        log=proc.stdout+'\n[STDERR]\n'+proc.stderr
        logpath=evidence/(menu+'-'+mode+'-import.log')
        logpath.write_text(log,encoding='utf-8')
        new_parts,new_formulas=package(target)
        book=openpyxl.load_workbook(target)
        result={'menu':menu,'mode':mode,'command':command,'exit':proc.returncode,
                'source':str(source),'target':str(target),'changed':digest(target)!=before_hash,
                'sheet_names':book.sheetnames,'expected_new_sheet_present':
                ('业务汇总' if menu=='1-1' else '业务提取') in book.sheetnames,
                'existing_formulas_and_caches_equal':all(new_formulas.get(k)==v for k,v in old_formulas.items()),
                'changed_parts':sorted(n for n in old_parts.keys()|new_parts.keys() if old_parts.get(n)!=new_parts.get(n)),
                'com':com_open(target),'terminal_log':str(logpath),
                'engine_lines':[l for l in log.splitlines() if 'ExcelEditPlan 路由' in l],
                'originals_unchanged':all(digest(Path(p))==h for p,h in originals.items())}
        book.close()
        results.append(result)
        (evidence/'import-results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
        print(menu,mode,'exit',proc.returncode,'changed',result['changed'],
              'new_sheet',result['expected_new_sheet_present'],flush=True)
