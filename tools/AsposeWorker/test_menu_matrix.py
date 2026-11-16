"""True CLI/menu differential checks on isolated synthetic working copies.

Every failed assertion/output/log is retained. No production config or dist is used.
"""
import argparse
import copy
import hashlib
import json
import math
import os
from pathlib import Path
import subprocess
import time
import traceback
import zipfile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def config(root):
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "全局配置"
    for row in [["配置项", "值"], ["输入目录", str(root)], ["输出目录", str(root / "output")],
                ["日志目录", str(root / "logs")], ["错误策略", "continue"]]:
        s.append(row)
    mapping = wb.create_sheet("机构映射表")
    for row in [["原始机构", "标准机构", "是否为外资行"], ["ABC Bank", "标准机构", 0], ["外资甲", "外资甲", 1]]:
        mapping.append(row)
    wb.create_sheet("重命名配置").append(["原表名", "新表名"])
    return wb


def save_config(wb, root):
    path = root / "config.xlsx"
    wb.save(path)
    return path


def standard_source(path, title="数据"):
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = title
    for row in [["机构名称", "ID", "余额"], ["ABC Bank", "00123", 12.125], ["ABC Bank", "00123", 12.125], ["普通行", "0", 0]]:
        s.append(row)
    s["AA12"] = "尾部保留"
    s["AA12"].font = Font(bold=True)
    s["AA12"].comment = Comment("合成批注", "测试")
    s.row_dimensions[12].height = 24
    s.column_dimensions["AA"].width = 22
    s.oddHeader.left.text = "单位：万元"
    s.oddFooter.center.text = "保留页脚"
    wb.save(path)


def color(v):
    if v is None:
        return None
    return (v.type, v.value, v.tint)


def style(c):
    return dict(font=(c.font.name, c.font.sz, bool(c.font.b), bool(c.font.i), color(c.font.color)),
                # Solid fills paint only fgColor; their bgColor is inactive.
                # Keep both colors for patterned fills, and the raw XLSX files.
                fill=(c.fill.patternType, color(c.fill.fgColor) if c.fill.patternType else None, color(c.fill.bgColor) if c.fill.patternType not in (None, "solid") else None),
                border=tuple((getattr(c.border, side).style, color(getattr(c.border, side).color)) for side in ["left", "right", "top", "bottom"]),
                alignment=(c.alignment.horizontal, c.alignment.vertical, bool(c.alignment.wrap_text), c.alignment.textRotation, c.alignment.indent),
                number_format=c.number_format, protection=(c.protection.locked, c.protection.hidden))


def snapshot(path):
    wb = openpyxl.load_workbook(path, data_only=False)
    cached = openpyxl.load_workbook(path, data_only=True)
    excluded = []
    sheets = {}
    for s in wb:
        if s.title == "Evaluation Warning":
            texts = [c.value for row in s.iter_rows(max_row=10, max_col=5) for c in row if isinstance(c.value, str)]
            assert any("Aspose" in t and "Evaluation" in t for t in texts), "Unattributed evaluation-like sheet"
            excluded.append(dict(sheet=s.title, marker_text=texts))
            continue
        cells = {}
        from openpyxl.cell.cell import Cell
        default_style = style(Cell(s, row=1, column=1))
        for row in s:
            for c in row:
                # An empty cell with only inactive/default style flags carries
                # no effective value or appearance. Raw ZIP output is retained;
                # this is semantic comparison, never byte-preservation evidence.
                if c.value is not None or (c.has_style and style(c) != default_style) or c.comment is not None:
                    value = c.value
                    if hasattr(value, "isoformat"):
                        value = value.isoformat()
                    cells[c.coordinate] = dict(value=value, data_type=c.data_type, cached=cached[s.title][c.coordinate].value,
                                               style=style(c), comment=None if c.comment is None else (c.comment.text, c.comment.author))
        sheets[s.title] = dict(cells=cells, merges=sorted(map(str, s.merged_cells.ranges)),
                               row_heights={str(k): v.height for k, v in s.row_dimensions.items() if v.height is not None and v.height != s.sheet_format.defaultRowHeight},
                               default_row_height=s.sheet_format.defaultRowHeight,
                               col_widths={k: v.width for k, v in s.column_dimensions.items() if v.width is not None},
                               freeze=s.freeze_panes, auto_filter=s.auto_filter.ref,
                               print_area=str(s.print_area), orientation=s.page_setup.orientation,
                               fit=(s.page_setup.fitToWidth, s.page_setup.fitToHeight, s.page_setup.scale),
                               headers=(str(s.oddHeader), str(s.evenHeader), str(s.firstHeader)),
                               footers=(str(s.oddFooter), str(s.evenFooter), str(s.firstFooter)))
    names = sorted((n.name, n.attr_text, n.localSheetId) for n in wb.defined_names.values())
    wb.close()
    cached.close()
    return dict(sheets=sheets, names=names), excluded


def differences(a, b, path=""):
    if isinstance(a, (int, float)) and not isinstance(a, bool) and isinstance(b, (int, float)) and not isinstance(b, bool):
        if math.isclose(a, b, rel_tol=1e-15, abs_tol=0.0): return []
    if isinstance(a, (list, tuple)) and isinstance(b, (list, tuple)) and len(a) == len(b):
        return [entry for i, (left, right) in enumerate(zip(a, b)) for entry in differences(left, right, path + "/" + str(i))]
    if isinstance(a, dict) and isinstance(b, dict):
        result = []
        for key in sorted(a.keys() | b.keys()):
            if key not in a or key not in b:
                result.append(dict(path=path + "/" + str(key), expected=a.get(key), actual=b.get(key)))
            else:
                result.extend(differences(a[key], b[key], path + "/" + str(key)))
        return result
    return [] if a == b else [dict(path=path, expected=a, actual=b)]


def prepare(root, case, aspose):
    root.mkdir(parents=True)
    wb = config(root)
    source, target = root / "source.xlsx", root / "target.xlsx"
    mode = "3" if aspose else "1"
    standard_source(source)
    if case == "1-1":
        b = openpyxl.Workbook(); s = b.active; s.title = "Report"
        for row in [["报表"], ["集团A"], ["地区", "存款", "存款", "贷款"], ["名称", "余额", "余额", "本息"], ["甲机构", 10, 11, 12]]: s.append(row)
        s.merge_cells("B3:C3"); b.save(source)
        b = openpyxl.Workbook(); s = b.active; s.title = "Append"; b.create_sheet("保留")["A1"] = "保留内容"; b.save(target)
        t = wb.create_sheet("时序提取规则")
        t.append(["是否启用", "规则名称", "工作簿关键字", "工作表关键字", "行头列", "列表头行", "必含列头", "必含行头", "数据起始行", "数据结束行", "数据起始列", "数据结束列", "跳过关键字", "目标工作簿路径", "目标工作表", "启用目标写入", "set区域", "目标去重列", "是否输出行序号"])
        t.append(["是", "WideRule", "source", "Report", "A", "3,4", "", "", 5, "", "B", "D", "", str(target), "Append", 1, "集团@A2", "", "是"])
        return ["wide-summary", "--config", str(save_config(wb, root)), "--modify-engine", "aspose" if aspose else "auto", str(source)], None, [target], [source]
    if case == "7-1":
        wb = openpyxl.Workbook(); s = wb.active; s.title = "全局配置"
        for row in [["旧键头", "旧值头"], ["输入目录", str(root)], ["输出目录", str(root / "output")], ["日志目录", str(root / "logs")]]: s.append(row)
        for name in ["校验规则生成配置", "校验规则映射表"]: wb.create_sheet(name)["A2"] = "existing user detail"
        c = save_config(wb, root)
        return ["menu", "--log-view", "developer", "--config", str(c)], f"7\n1\n1\n{mode}\n0\n0\n0\n", [root / "CONFIG_CANDIDATE"], [c]
    if case.startswith("2-"):
        if case == "2-2-plain":
            b = openpyxl.load_workbook(source); del b.active.row_dimensions[12]; b.save(source)
        if case == "2-3":
            b = openpyxl.Workbook(); s = b.active; s.title = "数据"
            for row in [["ID", "Value"], ["A", 12.125], ["A", 12.125], ["B", 0]]: s.append(row)
            b.save(source)
        wb.create_sheet("去重追加数据配置").append(["是否启用", "任务名", "源数据工作簿", "源数据工作表", "标识列序号", "目标工作簿", "目标工作表", "执行模式"])
        wb["去重追加数据配置"].append(["是", "合成去重", str(source), "数据", "2", str(target), "追加表", "追加"])
        b = openpyxl.Workbook(); s = b.active; s.title = "追加表"
        s.append(["ID", "Value"]); s.append(["OLD", 99]); b.save(target)
        return ["menu", "--config", str(save_config(wb, root))], f"2\n{case.split('-')[1]}\n{mode}\n0\n0\n", [source if case != "2-3" else target], [source] if case == "2-3" else []
    if case.startswith("3-"):
        m = int(case.split("-")[1])
        b = openpyxl.Workbook(); s = b.active; s.title = "合并源"; s["A1"] = "首条合并"; s.merge_cells("A1:B1"); s.row_dimensions[1].height = 32; s.column_dimensions["A"].width = 30
        s2 = b.create_sheet("独立源"); s2["A1"] = 0; s2["B1"] = "0"; s2["B1"].font = Font(bold=True); b.save(source)
        t = wb.create_sheet("打印配置")
        t.append(["是否启用", "打印模式", "源工作簿", "源工作表", "源工作表打印区域", "目标工作簿", "目标工作表", "FitToPagesWide", "FitToPagesTall", "打印方向", "水平居中", "垂直居中", "不输出批注", "零值不输出"])
        for sheet in ["合并源", "独立源"]:
            t.append(["是", m, str(source), sheet, "A1:B1", str(target), "打印结果", 1, 0, "横向", "是", "否", "是", "是"])
        if case.endswith("missing"):
            t.append(["是", m, str(source), "不存在的工作表", "A1:B1", str(target), "打印结果", 1, 0, "横向", "是", "否", "是", "是"])
        output = target.with_name("target_Aspose.xlsx") if aspose else target
        return ["menu", "--config", str(save_config(wb, root))], f"3\n1\n{'3' if aspose else '1'}\n0\n0\n", [output], [source]
    if case == "4-4":
        return ["convert-excel", "--config", str(save_config(wb, root)), "--engine", "aspose" if aspose else "auto", "--output-dir", str(root / "converted"), str(source)], None, [root / "converted/source.xlsx"], [source]
    if case == "4-5":
        s = wb["重命名配置"]; s.append(["数据", "新的名称"])
        return ["rename-sheets", "--config", str(save_config(wb, root)), "--modify-engine", "aspose" if aspose else "auto", str(source)], None, [source], []
    if case == "4-7":
        # Malformed pane state that the existing narrow repair is designed to fix.
        b = openpyxl.load_workbook(source); b.active.freeze_panes = "B2"; b.save(source)
        with zipfile.ZipFile(source) as z:
            parts = {n: z.read(n) for n in z.namelist()}
        xml = parts["xl/worksheets/sheet1.xml"].decode()
        parts["xl/worksheets/sheet1.xml"] = xml.replace('topLeftCell="B2"', 'topLeftCell="A1"').encode()
        with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in parts.items(): z.writestr(name, data)
        output = root / ("source-Excel修复_Aspose.xlsx" if aspose else "source-Excel修复.xlsx")
        return ["excel-open-repair", "--engine", "aspose" if aspose else "auto", str(source)], None, [output], [source]
    if case == "5-1":
        periods = []
        for name, factor in [("current", 1), ("prior", 0.5)]:
            p = root / (name + ".xlsx"); b = openpyxl.Workbook(); s = b.active; s.title = "村镇银行本外币"
            for row, base in [(5, 100), (6, 110)]:
                s.cell(row, 3, base * factor); s.cell(row, 4, "存款备注"); s.cell(row, 5, row)
                s.cell(row, 87, (base + 100) * factor); s.cell(row, 88, "贷款备注"); s.cell(row, 89, row + 1)
            b.save(p); periods.append(p)
        b = openpyxl.Workbook(); s = b.active; s.title = "广东分机构本外币"
        for r in [["机构"], ["其他"], ["村镇银行", 9], ["尾行"]]: s.append(r)
        b.save(target)
        return ["menu", "--config", str(save_config(wb, root))], f"5\n1\n{mode}\n{periods[0]}\n{periods[1]}\n1\n{target}\n0\n0\n", [target], periods
    if case == "5-2":
        standard_source(source, "甲分机构本外币")
        return ["institution-normalize", "--config", str(save_config(wb, root)), "--modify-engine", "aspose" if aspose else "auto", "--apply", "--in-place", "auto", str(source)], None, [source], []
    if case == "5-3":
        b = openpyxl.Workbook(); s = b.active; s.title = "甲分机构本外币"
        for row in [["机构", "余额"], ["普通甲", 10], ["外资甲", 20], ["普通乙", 30]]: s.append(row)
        b.save(source)
        return ["remove-foreign-bank", "--config", str(save_config(wb, root)), "--modify-engine", "aspose" if aspose else "auto", "--create-copy", str(source)], None, [root / "source(金融局).xlsx"], [source]
    if case in ("5-4", "5-4-menu"):
        standard_source(source, "外汇数据")
        if case == "5-4-menu":
            return ["menu", "--config", str(save_config(wb, root))], f"5\n4\n{mode}\n{source}\n0\n0\n", [source], []
        return ["fx-header", "--config", str(save_config(wb, root)), "--modify-engine", "aspose" if aspose else "auto", "--in-place", "auto", str(source)], None, [source], []
    if case == "5-6":
        t = wb.create_sheet("工作表提取")
        t.append(["是否启用", "工作簿关键字", "工作表关键字", "排除关键字", "提取区域", "多区域拼接方式", "输出文件夹", "文件名", "目标工作表", "备注"])
        t.append(["是", "source", "数据", "", "A1:C4", "纵向", str(root / "extracted"), "existing", "新增提取", ""])
        destination = root / "extracted/existing.xlsx"; destination.parent.mkdir()
        b = openpyxl.Workbook(); b.active.title = "原有表"; b.active["A1"] = "保留原有数据"; b.save(destination)
        return ["extract-sheets", "--config", str(save_config(wb, root)), "--modify-engine", "aspose" if aspose else "auto", "--append-existing", str(source)], None, [destination], [source]
    if case in ("5-7", "5-7-precision"):
        periods = []
        for name, amount in [("current", 14.125 if case == "5-7" else 14.123456789), ("prior", 9)]:
            p = root / (name + ".xlsx"); b = openpyxl.Workbook(); s = b.active; s.title = "惠州市涉农贷款分机构"
            s.append(["序号", "机构名称", "涉农贷款余额", "比上月"]); s.append(["序号", "机构名称", "涉农贷款余额", "比上月"]); s.append([1, "甲行", amount, 999])
            s = b.create_sheet("本外币涉农贷款分地区"); s.append(["地区", "涉农贷款比上月"]); s.append(["惠州市", 5.125 if case == "5-7" else 5.123456789]); b.save(p); periods.append(p)
        return ["menu", "--config", str(save_config(wb, root))], f"5\n7\n{mode}\n{periods[0]}\n{periods[1]}\n0\n0\n", [periods[0]], [periods[1]]
    raise ValueError(case)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--cases", nargs="+", default=["2-1", "2-2", "2-3", "3-1", "3-2", "3-3", "4-4", "4-5", "4-7", "5-1", "5-2", "5-3", "5-4", "5-6", "5-7"])
    args = parser.parse_args()
    root = Path(args.output).resolve(); root.mkdir(parents=True, exist_ok=False)
    results = []
    for case in args.cases:
        row = dict(case=case, status="FAIL", routes={})
        try:
            compared = []
            for aspose in [False, True]:
                route = "aspose" if aspose else "baseline"; directory = root / case / route
                command, stdin, outputs, readonly = prepare(directory, case, aspose)
                c = directory / "config.xlsx"; readonly = readonly + ([c] if c.exists() else [])
                hashes = {str(p): digest(p) for p in readonly}
                env = {**os.environ, "PYTHONIOENCODING": "utf-8", "LOCALAPPDATA": str(directory / "localappdata"), "PYTOOLS_CONFIG_PATH": str(c)}
                start = time.perf_counter()
                proc = subprocess.run([args.exe] + command, input=stdin, text=True, encoding="utf-8", stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=directory, env=env, timeout=180)
                elapsed = (time.perf_counter() - start) * 1000
                (directory / "stdout.txt").write_text(proc.stdout, encoding="utf-8"); (directory / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
                info = dict(command=command, returncode=proc.returncode, end_to_end_ms=elapsed, outputs=list(map(str, outputs)), source_hashes=hashes)
                row["routes"][route] = info
                for path, h in hashes.items(): assert digest(Path(path)) == h, ("Read-only source changed", path)
                assert proc.returncode == 0, proc.stderr[-3000:]
                if case.endswith("missing"):
                    history = list((directory / "logs").glob("*.run.json")); assert len(history) == 1
                    stats = json.loads(history[0].read_text("utf8"))["task_stats"]
                    assert stats == dict(failed=1, skipped=0, succeeded=2), stats
                snapshots = []
                for output in outputs:
                    if output.name == "CONFIG_CANDIDATE":
                        matches = list(directory.rglob("config-candidate.xlsx")); assert len(matches) == 1, matches
                        output = matches[0]; info["outputs"] = [str(output)]
                    assert output.exists(), f"Missing expected output {output}; found {list(directory.rglob('*.xlsx'))}"
                    value, excluded = snapshot(output)
                    if case == "7-1":
                        # Each isolated configuration intentionally points to its
                        # own directories. Normalize that fixture root only.
                        value = json.loads(json.dumps(value, ensure_ascii=False, default=str).replace(json.dumps(str(directory))[1:-1], "<CASE_ROOT>"))
                    snapshots.append(value)
                    if aspose: assert len(excluded) == 1, "Selected Aspose did not produce an evaluation output; route may not have executed"
                    info.setdefault("excluded_evaluation", []).extend(excluded)
                (directory / "snapshot.json").write_text(json.dumps(snapshots, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
                compared.append(snapshots)
            row["differences"] = differences(compared[0], compared[1])
            if case == "4-4" and row["differences"]:
                # User decision: Aspose preserves the input; do not copy the
                # local COM application's font/row-height normalization.
                source_view, _ = snapshot(root / case / "aspose/source.xlsx")
                row["com_appearance_differences"] = row["differences"]
                row["aspose_source_differences"] = differences(source_view, compared[1][0])
                assert not row["aspose_source_differences"], "Aspose did not preserve source semantics"
                assert all("/style/font/" in d["path"] or d["path"].endswith("/default_row_height") for d in row["differences"]), "COM difference is not an approved appearance difference"
                row["differences"] = []
            assert not row["differences"], f"{len(row['differences'])} semantic differences"
            row["status"] = "PASS"
        except Exception:
            row["error"] = traceback.format_exc()
        results.append(row)
        (root / "summary.json").write_text(json.dumps(dict(results=results), ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        print(json.dumps(dict(case=case, status=row["status"], error=row.get("error", "")[-700:]), ensure_ascii=False), flush=True)
    return int(any(r["status"] != "PASS" for r in results))


if __name__ == "__main__":
    raise SystemExit(main())
