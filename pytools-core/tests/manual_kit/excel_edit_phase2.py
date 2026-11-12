"""Run real dedup menus on untouched sample copies and retain acceptance evidence.

This is test support, not a production Excel route. No sample is saved in place.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

import openpyxl


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def package(path):
    with zipfile.ZipFile(path) as archive:
        hashes = {name: hashlib.sha256(archive.read(name)).hexdigest()
                  for name in archive.namelist()}
        formulas = {}
        for name in archive.namelist():
            if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                root = ET.fromstring(archive.read(name))
                ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                formulas[name] = [(c.get('r'), c.findtext('s:f', namespaces=ns),
                                   c.findtext('s:v', namespaces=ns))
                                  for c in root.findall('.//s:c', ns)
                                  if c.find('s:f', ns) is not None]
        return hashes, formulas


def com_open(path):
    app = book = None
    try:
        import win32com.client
        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        app.AutomationSecurity = 3
        book = app.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True,
                                  IgnoreReadOnlyRecommended=True, CorruptLoad=0)
        return {'status': 'opened_read_only', 'sheets': book.Worksheets.Count,
                'read_only': bool(book.ReadOnly), 'saved': False}
    except Exception as exc:
        return {'status': 'failed_or_unavailable', 'error': str(exc)}
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if app is not None:
            app.Quit()


def run(args):
    source = args.samples / '2' / '检查重复去重用例.xlsx'
    before_hash = digest(source)
    scratch = Path(tempfile.mkdtemp(prefix='rusttools-phase2-'))
    args.evidence.mkdir(parents=True, exist_ok=True)
    results = []
    cases = [('2-1', mode) for mode in ('OpenpyxlOnly', 'UMYAOnly', 'Auto')]
    cases += [('2-2', mode) for mode in
              ('OpenpyxlOnly', 'UMYAOnly', 'BetterOfficeOnly', 'Auto')]
    original_book = openpyxl.load_workbook(source)
    original_sheet = max(original_book, key=lambda s: s.max_row * s.max_column)
    sheet_name = original_sheet.title
    original_rows = list(original_sheet.values)
    original_book.close()
    seen, duplicate_rows = set(), []
    for index, row in enumerate(original_rows[1:], 2):
        value = row[7]
        if value is None or str(value).strip() == '':
            continue
        key = str(value)
        if key in seen:
            duplicate_rows.append(index)
        seen.add(key)
    old_parts, old_formulas = package(source)
    for menu, mode in cases:
        folder = scratch / (menu + '-' + mode)
        folder.mkdir()
        target = folder / 'source.xlsx'
        shutil.copy2(source, target)
        config = openpyxl.Workbook()
        global_sheet = config.active
        global_sheet.title = '全局配置'
        global_sheet.append(['配置项', '值'])
        for key, val in [('输入目录', str(folder)), ('输出目录', str(folder/'output')),
                         ('日志目录', str(folder/'logs')), ('错误策略', 'fail_fast')]:
            global_sheet.append([key, val])
        task = config.create_sheet('去重追加数据配置')
        task.append(['是否启用','任务名','源数据工作簿','源数据工作表','标识列序号',
                     '目标工作簿','目标工作表','执行模式'])
        task.append(['是', menu, str(target), sheet_name, '8', str(folder/'unused.xlsx'),
                     sheet_name, '追加'])
        config_path = folder/'config.xlsx'
        config.save(config_path)
        env = os.environ.copy()
        env['PYTOOLS_EXCEL_EDIT_ROUTE'] = mode
        env['PYTOOLS_OPENPYXL_PYTHON'] = sys.executable
        env['PYTOOLS_OPENPYXL_WORKER'] = str(args.worker.resolve())
        proc = subprocess.run([str(args.exe), 'menu', '--config', str(config_path),
                               '--log-view', 'developer'],
                              input=f'2\n{menu[-1]}\n0\n0\n', text=True,
                              encoding='utf-8', errors='replace', capture_output=True,
                              env=env, timeout=120)
        log = proc.stdout + '\n[STDERR]\n' + proc.stderr
        log_path = args.evidence / (menu + '-' + mode + '.log')
        log_path.write_text(log, encoding='utf-8')
        new_parts, new_formulas = package(target)
        book = openpyxl.load_workbook(target)
        sheet = book[sheet_name]
        if menu == '2-1':
            actual = [r for r in range(2, sheet.max_row+1)
                      if all(sheet.cell(r,c).fill.fgColor.rgb == 'FFFF0000'
                             for c in range(1,sheet.max_column+1))]
            values_match = list(sheet.values) == original_rows
            correct = actual == duplicate_rows and values_match
            differences = {'marked_rows': actual, 'values_unchanged': values_match}
        else:
            expected = [row for i,row in enumerate(original_rows,1)
                        if i not in duplicate_rows]
            correct = list(sheet.values) == expected
            differences = {'rows_before': len(original_rows),
                           'rows_after': sheet.max_row, 'surviving_values_match': correct}
        book.close()
        record = {'menu': menu, 'mode': mode, 'input': str(source),
                  'output_copy': str(target), 'exit': proc.returncode,
                  'business_correct': correct, 'differences': differences,
                  'changed_parts': sorted(n for n in old_parts.keys() | new_parts.keys()
                                          if old_parts.get(n) != new_parts.get(n)),
                  'formula_and_cache_equal': old_formulas == new_formulas,
                  'com': com_open(target), 'terminal_log': str(log_path),
                  'engine_lines': [line for line in log.splitlines() if 'ExcelEditPlan 路由' in line],
                  'has_error': '[ERROR]' in log,
                  'output_sha256': digest(target)}
        results.append(record)
        print(menu, mode, 'correct=', correct, 'COM=', record['com']['status'], flush=True)
        (args.evidence/'dedup-results.json').write_text(json.dumps(
            {'scratch':str(scratch),'original_sha256':before_hash,
             'original_unchanged':digest(source)==before_hash,'runs':results},
            ensure_ascii=False,indent=2),encoding='utf-8')
    assert digest(source) == before_hash, 'original sample changed'


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--samples', type=Path, required=True)
    parser.add_argument('--exe', type=Path, required=True)
    parser.add_argument('--worker', type=Path, required=True)
    parser.add_argument('--evidence', type=Path, required=True)
    run(parser.parse_args())
