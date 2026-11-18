"""Windows parity for static printing from one isolated XLS/XLSX fixture pair.

Run --baseline with the pre-change executable in a new evidence directory,
then run with the updated executable and the same directory.
"""

import argparse
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import struct
import subprocess
import sys
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font
import win32com.client


HEADERS = [
    "是否启用", "打印模式", "源工作簿", "源工作表", "源工作表打印区域",
    "目标工作簿", "目标工作表", "FitToPagesWide", "FitToPagesTall", "打印方向",
    "水平居中", "垂直居中", "不输出批注", "零值不输出",
]
MENUS = {"openpyxl": "3\n1\n1\n0\n0\n", "rust": "3\n1\n2\n0\n0\n",
         "precheck-menu": "3\n3\n0\n0\n"}


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def patch_xlsx_cache(path):
    with ZipFile(path) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    part = "xl/worksheets/sheet1.xml"
    xml = parts[part]
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    for address, expected in [("E2", "901"), ("F2", None)]:
        tree = ElementTree.fromstring(xml)
        value = tree.find(f".//s:c[@r='{address}']/s:v", ns)
        assert value is not None
        value.text = expected
        ElementTree.register_namespace("", ns["s"])
        xml = ElementTree.tostring(tree)
    parts[part] = xml
    with ZipFile(path, "w", ZIP_DEFLATED) as target:
        for name, data in parts.items():
            target.writestr(name, data)


def patch_biff_cache(path):
    data = path.read_bytes()
    assert data.startswith(bytes.fromhex("d0cf11e0a1b11ae1"))
    for column, original, wanted in [
        (4, 107.0, struct.pack("<d", 901.0)),
        (5, 108.0, bytes.fromhex("030000000000ffff")),
    ]:
        # Fixed result in the FORMULA record. This controlled fixture must have
        # exactly one matching header, position and numeric result.
        prefix = rb"\x06\x00..\x01\x00" + re.escape(struct.pack("<H", column)) + b".."
        record = re.compile(prefix + re.escape(struct.pack("<d", original)), re.DOTALL)
        matches = list(record.finditer(data))
        assert len(matches) == 1, (column, len(matches))
        start = matches[0].end() - 8
        data = data[:start] + wanted + data[start + 8:]
    path.write_bytes(data)


def make_fixtures(directory):
    directory.mkdir(parents=True)
    xlsx = directory / "source.xlsx"
    xls = directory / "source.xls"
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Data"
    for row in [
        ["Synthetic", None, "Text", "Date", "Stale cache", "Blank cache", "Boolean", "Error"],
        [0, "0", "00123", datetime(2026, 11, 15), "=SUM(Refs!A1:A2)", "=SUM(Refs!A1:A2)+1", False, "#N/A"],
    ]:
        sheet.append(row)
    sheet.merge_cells("A1:B1")
    sheet["A1"].font = Font(bold=True)
    sheet["D2"].number_format = "yyyy-mm-dd"
    sheet["C2"].comment = Comment("synthetic note", "fixture")
    sheet.row_dimensions[1].height = 31
    sheet.column_dimensions["A"].width = 24
    sheet["J4"].font = Font(italic=True)  # Style-only cell extends the default used range.
    ref = book.create_sheet("Refs")
    ref.append([100])
    ref.append([7])
    book.save(xlsx)

    # Only create the synthetic BIFF fixture via COM. Production tests run the
    # Rust CLI separately with its configured COM worker unavailable.
    app = win32com.client.DispatchEx("Excel.Application")
    assert app.Workbooks.Count == 0
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    try:
        source = app.Workbooks.Open(str(xlsx), UpdateLinks=0, ReadOnly=True)
        try:
            app.Calculate()
            assert source.Worksheets("Data").Range("E2:F2").Value2 == ((107.0, 108.0),)
            source.SaveAs(str(xls), FileFormat=56)
        finally:
            source.Close(SaveChanges=False)
    finally:
        app.Quit()
    patch_xlsx_cache(xlsx)
    patch_biff_cache(xls)


def config(directory, task_specs):
    book = openpyxl.Workbook()
    global_sheet = book.active
    global_sheet.title = "全局配置"
    for row in [
        ["配置项", "值"], ["输入目录", str(directory)], ["输出目录", str(directory / "output")],
        ["日志目录", str(directory / "logs")], ["错误策略", "continue"],
    ]:
        global_sheet.append(row)
    tasks = book.create_sheet("打印配置")
    tasks.append(HEADERS)
    for source, target, mode, area, source_sheet, target_sheet, drop_comments in task_specs:
        tasks.append(["是", mode, str(source), source_sheet, area, str(target), target_sheet,
                      1, 1, "横向", "是", "否", "是" if drop_comments else "否", "是"])
    path = directory / "config.xlsx"
    assert not path.exists(), path
    book.save(path)
    return path


def specs(source, directory):
    return [(source, directory / f"mode{mode}.xlsx", mode, "A1:H2;A2:C2", "Data", "Print", False)
            for mode in (1, 2, 3)]


def run_cli(exe, directory, task_specs, menu=None, missing_python=False):
    directory.mkdir(parents=True, exist_ok=True)
    cfg = config(directory, task_specs)
    assert cfg.is_file()
    env = {**os.environ, "LOCALAPPDATA": str(directory / "localappdata"),
           "PYTHONIOENCODING": "utf-8", "PYTOOLS_CONFIG_PATH": str(cfg),
           "PYTOOLS_OPENPYXL_PYTHON": str(directory / "missing-python.exe") if missing_python else sys.executable,
           "PYTOOLS_OPENPYXL_WORKER": str(Path(__file__).resolve().parents[3] / "crates/excel-editor/python/openpyxl_worker.py"),
           "PYTOOLS_COM_WORKER": str(directory / "missing-com-worker.py"),
           "PYTHONDONTWRITEBYTECODE": "1"}
    command = [str(exe), "print-config-precheck" if menu is None else "menu", "--config", str(cfg)]
    if menu is not None:
        command.extend(["--log-view", "developer"])
    proc = subprocess.run(command, input=MENUS.get(menu, ""), cwd=directory, env=env,
                          capture_output=True, text=True, encoding="utf-8", timeout=90)
    (directory / "stdout.txt").write_text(proc.stdout, encoding="utf-8")
    (directory / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
    return proc


def snapshot(path):
    book = openpyxl.load_workbook(path)
    result = {}
    for s in book:
        result[s.title] = {
            "cells": {c.coordinate: [c.data_type, str(c.value)] for row in s for c in row if c.value is not None},
            "merges": sorted(map(str, s.merged_cells.ranges)),
            "print_area": s.print_area,
            "page": [s.page_setup.paperSize, s.page_setup.orientation, s.page_setup.fitToWidth,
                     s.page_setup.fitToHeight, s.print_options.horizontalCentered,
                     s.print_options.verticalCentered, s.oddFooter.center.text],
            "style": [s["A1"].font.bold, s["D2"].number_format,
                      {str(k): d.height for k, d in s.row_dimensions.items() if d.height},
                      {str(k): d.width for k, d in s.column_dimensions.items()}],
            "notes": {c.coordinate: c.comment.text for row in s for c in row if c.comment},
        }
    return result


def legacy_com(old_root, directory, source):
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(old_root / "PyCode"))
    from core import print_310_com

    directory.mkdir(parents=True)
    (directory / "logs").mkdir()
    # Pin the independent host, and make manual calculation active before
    # opening the cache-only synthetic source.
    original_start = print_310_com.start_excel_app

    def start_manual():
        app = original_start()
        assert app.Workbooks.Count == 0 and str(app.Name) == "Microsoft Excel", app.Name
        app.AutomationSecurity = 3
        app.Workbooks.Add()
        app.Calculation = -4135
        return app

    print_310_com.start_excel_app = start_manual
    try:
        tasks = [
            dict(enabled=True, row_no=mode + 1, mode=mode, source_wb=src,
                 source_ws=src_sheet, source_ranges=area, target_wb=dst, target_ws=dst_sheet,
                 drop_comments=drop, hide_zero=True, orientation="横向", fit_wide=1, fit_tall=1,
                 center_h=True, center_v=False)
            for src, dst, mode, area, src_sheet, dst_sheet, drop in specs(source, directory)
        ]
        summary = print_310_com.run_print_config_all_modes_com(tasks, directory / "logs")
    finally:
        print_310_com.start_excel_app = original_start
    assert summary["task_ok"] == 3 and summary["task_skip"] == 0, summary
    return {str(m): snapshot(directory / f"mode{m}.xlsx") for m in (1, 2, 3)}


def compare(a, b):
    result = {}
    for mode in a:
        result[mode] = {}
        for sheet, old in a[mode].items():
            new = b[mode].get(sheet, {})
            result[mode][sheet] = {k: {"reference": value, "actual": new.get(k)}
                                   for k, value in old.items() if value != new.get(k)}
    return result


def capture_modes(exe, root, source, menu, *, expect_outputs, missing_python=False, legacy_suffix=False):
    directory = root / menu
    proc = run_cli(exe, directory, specs(source, directory), menu, missing_python)
    suffix = "_Rust" if legacy_suffix and menu == "rust" else ""
    paths = {str(m): directory / f"mode{m}{suffix}.xlsx" for m in (1, 2, 3)}
    assert proc.returncode == 0, proc.stderr
    assert all(p.exists() == expect_outputs for p in paths.values()), (paths, proc.stdout, proc.stderr)
    if missing_python and expect_outputs:
        assert proc.stdout.count("无法完整读取 XLS 单元格格式") == 3, proc.stdout
        assert proc.stdout.count("无法读取 XLS 批注") == 2, proc.stdout
        log = next((directory / "logs").glob("*.run.json"))
        report = json.loads(log.read_text(encoding="utf-8"))
        assert report["task_stats"] == {"succeeded": 3, "failed": 0, "skipped": 0}, report
        assert len(report["business_issues"]) == 3, report
    return {mode: snapshot(path) for mode, path in paths.items()} if expect_outputs else {}


def write_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def verify_xls_modes(result, baseline):
    for mode, sheets in result.items():
        sheet = sheets["Print"]
        xlsx = baseline["xlsx_rust"][mode]["Print"]
        legacy = baseline["legacy_excel_com"][mode]["Print"]
        assert list(sheets) == ["Print"]
        for key in ("cells", "merges", "print_area", "page"):
            assert sheet[key] == xlsx[key], (mode, key, xlsx[key], sheet[key])
        for key in ("merges", "print_area", "page"):
            assert sheet[key] == legacy[key], (mode, key, legacy[key], sheet[key])
        cells = sheet["cells"]
        for address in set(cells) | set(legacy["cells"]):
            # Legacy COM hid text zero, coerced numeric text, and wrote an
            # HRESULT for #N/A. The current XLSX path preserves the cell types.
            if address not in {"B2", "B4", "C2", "C4", "H2"}:
                assert cells.get(address) == legacy["cells"].get(address), (mode, address)
        assert cells["E2"] == ["n", "901"] and "F2" not in cells
        assert all(kind != "f" for kind, _ in cells.values()), cells
        assert cells["B2"] == ["s", "0"] and cells["C2"] == ["s", "00123"]
        assert cells["H2"] == ["e", "#N/A"]
        assert not sheet["notes"] and not sheet["style"][0], sheet
        assert sheet["style"][1] == "yyyy-mm-dd", sheet
        verify_readable_column_widths(sheet)


def verify_readable_column_widths(sheet):
    widths = sheet["style"][3]
    assert widths.get("D", 0) >= 22, widths
    assert all(widths.get(col, 0) >= 14 for col in ("E", "F")), widths


def verify_xlsx_modes(result, baseline):
    for mode, sheets in result.items():
        expected = baseline[mode]
        assert list(sheets) == list(expected)
        for name, sheet in sheets.items():
            original = expected[name]
            if mode == "3":
                verify_readable_column_widths(sheet)
                sheet = {**sheet, "style": [*sheet["style"][:3], original["style"][3]]}
            assert sheet == original, (mode, name, original, sheet)


def verify_edges(exe, output, xls, xlsx):
    directory = output / "after-precheck"
    precheck = run_cli(exe, directory, specs(xls, directory))
    report = json.loads(precheck.stdout)
    assert precheck.returncode == 0 and report["ok"] == report["warnings"] == 3, report
    assert report["failed"] == 0 and len(list(directory.glob("mode*.xlsx"))) == 0, report
    assert sum("无法读取 XLS 批注" in issue for issue in report["issues"]) == 2, report

    directory = output / "after-precheck-menu"
    run = run_cli(exe, directory, specs(xls, directory), "precheck-menu")
    assert run.returncode == 0 and "无法读取 XLS 批注" in run.stdout, run.stdout
    assert not list(directory.glob("mode*.xlsx"))

    directory = output / "after-print-mixed"
    target = directory / "mixed.xlsx"
    default_area = directory / "auto.xlsx"
    run = run_cli(exe, directory, [
        (xlsx, target, 1, "A1:H2", "Data", "Reused", False),
        (xls, target, 3, "$A$1:$A$2", "Refs", "Reused", True),
        (xlsx, target, 1, "A1:H2", "Data", "XlsxKept", False),
        (xls, default_area, 1, "", "Data", "Print", True),
    ], "openpyxl", missing_python=True)
    assert run.returncode == 0 and target.exists() and default_area.exists(), run.stdout
    sheets = snapshot(target)
    assert list(sheets) == ["Reused", "XlsxKept"], sheets
    assert sheets["Reused"]["cells"] == {"A1": ["n", "100"], "A2": ["n", "7"]}
    assert not sheets["Reused"]["merges"] and not sheets["Reused"]["style"][2]
    assert sheets["Reused"]["style"][3] == {"A": 13.0}, sheets
    assert sheets["XlsxKept"]["notes"] == {"C2": "synthetic note"}, sheets
    assert snapshot(default_area)["Print"]["print_area"] == "'Print'!$A$1:$J$5"
    assert "无法读取 XLS 批注" not in run.stdout, run.stdout

    directory = output / "after-separate-targets"
    xlsx_target = directory / "from-xlsx.xlsx"
    xls_target = directory / "from-xls.xlsx"
    run = run_cli(exe, directory, [
        (xlsx, xlsx_target, 1, "A1:H2", "Data", "Print", False),
        (xls, xls_target, 1, "A1:H2", "Data", "Print", False),
    ], "openpyxl")
    assert run.returncode == 0, run.stdout
    assert snapshot(xlsx_target)["Print"]["notes"] == {"C2": "synthetic note"}
    assert not snapshot(xls_target)["Print"]["notes"]

    corrupt = output / "fixture" / "corrupt.xls"
    with corrupt.open("xb") as file:
        file.write(b"not a BIFF workbook")
    directory = output / "after-invalid-precheck"
    report = run_cli(exe, directory, [
        (corrupt, directory / "corrupt.xlsx", 1, "", "Data", "Print", False),
        (xls, directory / "missing.xlsx", 2, "", "Missing", "Print", False),
        (xls, directory / "invalid.xlsx", 3, "A0:IV12", "Data", "Print", False),
        (xls, directory / "bounds.xlsx", 3, "IW1:IW2", "Data", "Print", False),
        (xls, directory / "output.xls", 3, "", "Data", "Print", False),
    ])
    failed = json.loads(report.stdout)
    assert report.returncode != 0 and failed["failed"] == 5 and failed["ok"] == 0, failed
    assert list(directory.glob("*.xlsx")) == [directory / "config.xlsx"]

    directory = output / "after-invalid-print"
    target = directory / "existing.xlsx"
    directory.mkdir()
    shutil.copyfile(xlsx, target)
    preserved = sha(target)
    run = run_cli(exe, directory, [
        (xls, target, 1, "", "Data", "Print", False),
        (corrupt, directory / "corrupt.xlsx", 2, "", "Data", "Print", False),
        (xls, directory / "ok.xlsx", 3, "A1:H2", "Data", "Print", True),
    ], "openpyxl", missing_python=True)
    assert run.returncode == 0 and not (directory / "corrupt.xlsx").exists(), run.stdout
    assert sha(target) == preserved and (directory / "ok.xlsx").exists(), run.stdout
    report = json.loads(next((directory / "logs").glob("*.run.json")).read_text(encoding="utf-8"))
    assert report["task_stats"] == {"succeeded": 1, "failed": 2, "skipped": 0}, report

    directory = output / "after-xls-target"
    run = run_cli(exe, directory, [
        (xls, directory / "forbidden.xls", 1, "", "Data", "Print", False),
    ], "rust", missing_python=True)
    assert run.returncode == 0 and not list(directory.glob("*.xls")), run.stdout


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline", action="store_true")
    parser.add_argument("--exe", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--legacy", type=Path, required=True)
    args = parser.parse_args()
    output = args.output.resolve()
    fixture = output / "fixture"
    if args.baseline:
        make_fixtures(fixture)
    xls, xlsx = fixture / "source.xls", fixture / "source.xlsx"
    before = {path: sha(path) for path in (xls, xlsx)}
    if args.baseline:
        result = {"legacy_excel_com": legacy_com(args.legacy, output / "legacy", xls)}
        for menu in ("openpyxl", "rust"):
            result["xlsx_" + menu] = capture_modes(args.exe, output / "baseline-xlsx", xlsx, menu, expect_outputs=True, legacy_suffix=True)
            capture_modes(args.exe, output / "baseline-xls", xls, menu, expect_outputs=False)
        directory = output / "baseline-precheck"
        precheck = run_cli(args.exe, directory, specs(xls, directory))
        assert precheck.returncode != 0 and json.loads(precheck.stdout)["failed"] == 3, precheck.stdout
        write_json(output / "baseline.json", result)
    else:
        baseline = json.loads((output / "baseline.json").read_text(encoding="utf-8"))
        result = {}
        for menu in ("openpyxl", "rust"):
            xlsx_result = capture_modes(args.exe, output / "after-xlsx", xlsx, menu, expect_outputs=True)
            verify_xlsx_modes(xlsx_result, baseline["xlsx_" + menu])
            current = capture_modes(args.exe, output / "after-xls", xls, menu,
                                    expect_outputs=True, missing_python=True)
            verify_xls_modes(current, baseline)
            result[menu] = {
                "xls": current,
                "vs_xlsx_openpyxl": compare(baseline["xlsx_openpyxl"], current),
                "vs_xlsx_rust": compare(baseline["xlsx_rust"], current),
                "vs_legacy_excel_com": compare(baseline["legacy_excel_com"], current),
            }
        verify_edges(args.exe, output, xls, xlsx)
        write_json(output / "parity.json", result)
    assert {path: sha(path) for path in before} == before
    print(output)


if __name__ == "__main__":
    main()
