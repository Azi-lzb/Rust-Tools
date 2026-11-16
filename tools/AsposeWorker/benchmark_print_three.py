"""Three actual print menu routes; correctness gates every timing cohort.

Only synthetic files under --output are written. Raw evaluation outputs and
all mismatches are retained. No output watermark is physically removed.
"""
import argparse
import copy
import hashlib
import importlib.util
import json
import os
import shutil
from pathlib import Path
import subprocess
import time
import traceback
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("original_print_oracle", HERE / "test_menu_matrix.py")
oracle = importlib.util.module_from_spec(spec); spec.loader.exec_module(oracle)
ENGINES = {"openpyxl": "1", "rust": "2", "aspose": "3"}


def print_style(c):
    def edge(side):
        value = getattr(c.border, side)
        return (None, None) if value is None else (value.style, oracle.color(value.color))
    return dict(font=(c.font.name, c.font.sz, bool(c.font.b), bool(c.font.i), oracle.color(c.font.color)),
                fill=(c.fill.patternType, oracle.color(c.fill.fgColor) if c.fill.patternType else None, oracle.color(c.fill.bgColor) if c.fill.patternType not in (None, "solid") else None),
                border=tuple(edge(side) for side in ["left", "right", "top", "bottom"]),
                alignment=(c.alignment.horizontal, c.alignment.vertical, bool(c.alignment.wrap_text), c.alignment.textRotation, c.alignment.indent),
                number_format=c.number_format, protection=(c.protection.locked, c.protection.hidden))


oracle.style = print_style


def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()


def rich_source(path):
    book = openpyxl.Workbook(); s = book.active; s.title = "数据"
    for r in range(1, 13):
        for c in range(1, 28): s.cell(r, c, r * 100 + c)
    s["A1"] = "合成标题"; s.merge_cells("A1:B1")
    s["A2"] = "00123"; s["B2"] = "12345678901234567890"
    s["C2"] = 0; s["D2"] = "0"; s["E2"] = False
    s["F2"] = datetime(2026, 11, 15); s["F2"].number_format = "yyyy-mm-dd"
    s["G2"] = "#N/A"; s["H2"] = None
    s["I2"] = "=SUM(J2:J3)"; s["J2"] = 12.125; s["J3"] = 3.5
    s["AA12"].comment = Comment("必须按配置保留的合成批注", "synthetic")
    for row in s:
        for c in row:
            if isinstance(c, openpyxl.cell.cell.MergedCell): continue
            c.font = Font(name="Calibri", size=11, bold=c.row == 1, color="FF123456")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(bottom=Side(style="thin", color="FF112233"))
            if c.row == 1: c.fill = PatternFill("solid", fgColor="FFDDEEFF")
    s.row_dimensions[1].height = 32; s.row_dimensions[12].height = 28
    s.column_dimensions["A"].width = 30; s.column_dimensions["AA"].width = 23.125
    book.save(path)
    # Correct cached result, independent of all engines under comparison.
    with zipfile.ZipFile(path) as package: parts = {n: package.read(n) for n in package.namelist()}
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"; ET.register_namespace("", ns[1:-1])
    xml = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    cell = next(c for c in xml.iter(ns + "c") if c.get("r") == "I2")
    cell.find(ns + "v").text = "15.625"
    parts["xl/worksheets/sheet1.xml"] = ET.tostring(xml)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for n, data in parts.items(): package.writestr(n, data)


def prepare(directory, fixture, mode, engine):
    command, _, _, sources = oracle.prepare(directory, f"3-{mode}", False)
    source = directory / "source.xlsx"; config = directory / "config.xlsx"
    if fixture != "reuse":
        if fixture == "rich": rich_source(source)
        elif fixture == "plain":
            book = openpyxl.Workbook(); s = book.active; s.title = "数据"
            for r in range(1, 101):
                for c in range(1, 11): s.cell(r, c, f"R{r}C{c}" if c == 1 else r * c)
            # Explicit common column widths avoid testing an implicit-width
            # serialization distinction as the only case in this cohort.
            for c in range(1, 11): s.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 13
            book.save(source)
        else: raise ValueError(fixture)
        book = openpyxl.load_workbook(config); tasks = book["打印配置"]; tasks.delete_rows(2, tasks.max_row)
        tasks.append(["是", mode, str(source), "数据", "A1:AA12" if fixture == "rich" else "A1:J100", str(directory / "target.xlsx"), "打印结果", 1, 0, "横向", "是", "否", "否", "是"])
        book.save(config)
    suffix = {"openpyxl": "", "rust": "_Rust", "aspose": "_Aspose"}[engine]
    output = directory / f"target{suffix}.xlsx"
    return command, f"3\n1\n{ENGINES[engine]}\n0\n0\n", output, [source, config]


def view(path):
    raw, excluded = oracle.snapshot(path)
    result = copy.deepcopy(raw)
    book = openpyxl.load_workbook(path)
    for name, sheet in result["sheets"].items():
        s = book[name]
        # These are effective defaults, not a tolerance for changed formatting.
        base = book._fonts[0]
        def effective(style):
            style = copy.deepcopy(style)
            name, size, bold, italic, color = style["font"]
            # The fixtures use the standard Office theme; an unspecified font
            # color and theme1/tint0 both resolve to its default black text.
            if color is None and oracle.color(base.color) == ("theme", 1, 0.0):
                color = ("theme", 1, 0.0)
            style["font"] = (name or base.name, size if size is not None else base.sz, bold, italic, color)
            h, v, wrap, rotation, indent = style["alignment"]
            style["alignment"] = (h or "general", v or "bottom", wrap, rotation or 0, indent or 0)
            style["border"] = tuple((line, color if line else None) for line, color in style["border"])
            locked, hidden = style["protection"]; style["protection"] = (True if locked is None else locked, bool(hidden))
            return style
        default = effective(oracle.style(openpyxl.cell.cell.Cell(s, row=1, column=1)))
        for cell in sheet["cells"].values(): cell["style"] = effective(cell["style"])
        sheet["cells"] = {a: c for a, c in sheet["cells"].items() if c["value"] is not None or c["comment"] is not None or c["style"] != default}
        sheet["margins"] = {n: getattr(s.page_margins, n) for n in ["left", "right", "top", "bottom", "header", "footer"]}
        sheet["paper"] = str(s.page_setup.paperSize) if s.page_setup.paperSize is not None else "1"
        sheet["center"] = (bool(s.print_options.horizontalCentered), bool(s.print_options.verticalCentered))
        sheet["fit_to_page"] = bool(s.sheet_properties.pageSetUpPr and s.sheet_properties.pageSetUpPr.fitToPage)
        width, height, scale = sheet["fit"]
        sheet["fit"] = (1 if width is None else width, 1 if height is None else height, scale)
        widths = {c: s.sheet_format.defaultColWidth or 8.43 for c in range(1, s.max_column + 1)}
        for dimension in s.column_dimensions.values():
            for c in range(dimension.min, dimension.max + 1): widths[c] = dimension.width
        sheet["col_widths"] = {openpyxl.utils.get_column_letter(c): width for c, width in widths.items()}
    book.close()
    return result, raw, excluded


def main():
    parser = argparse.ArgumentParser(); parser.add_argument("--output", required=True)
    parser.add_argument("--exe", default="F:/RustBuild/pytools/debug/RustTools.exe")
    parser.add_argument("--fixtures", nargs="+", default=["reuse", "rich", "plain"])
    parser.add_argument("--canonical-from", type=Path,
                        help="Reuse previously recorded source bytes for cross-build comparisons")
    args = parser.parse_args(); root = Path(args.output).resolve(); root.mkdir(parents=True, exist_ok=False)
    results = []
    for fixture in args.fixtures:
        canonical = root / "canonical" / fixture
        prepare(canonical, fixture, 1, "openpyxl")
        if args.canonical_from:
            shutil.copyfile(args.canonical_from / fixture / "source.xlsx", canonical / "source.xlsx")
        for mode in [1, 2, 3]:
            row = dict(fixture=fixture, mode=mode, engines={}); snapshots = {}
            for engine in ENGINES:
                directory = root / fixture / str(mode) / engine
                entry = dict(status="FAILED")
                try:
                    command, stdin, output, inputs = prepare(directory, fixture, mode, engine)
                    shutil.copyfile(canonical / "source.xlsx", directory / "source.xlsx")
                    before = {str(p): sha(p) for p in inputs}
                    env = {**os.environ, "PYTHONIOENCODING": "utf8", "LOCALAPPDATA": str(directory / "localappdata"), "PYTOOLS_CONFIG_PATH": str(directory / "config.xlsx")}
                    start = time.perf_counter_ns()
                    process = subprocess.run([args.exe] + command, input=stdin, capture_output=True, text=True, encoding="utf8", cwd=directory, env=env, timeout=180)
                    elapsed = (time.perf_counter_ns() - start) / 1e6
                    (directory / "stdout.txt").write_text(process.stdout, encoding="utf8"); (directory / "stderr.txt").write_text(process.stderr, encoding="utf8")
                    assert process.returncode == 0, process.stderr[-2000:]
                    assert output.is_file(), f"Missing output {output}"
                    assert all(sha(Path(p)) == digest for p, digest in before.items()), "Source/config changed"
                    value, raw, excluded = view(output)
                    assert len(excluded) == (1 if engine == "aspose" else 0)
                    (directory / "snapshot.json").write_text(json.dumps(dict(effective=value, raw=raw, evaluation=excluded), ensure_ascii=False, default=str, indent=2), encoding="utf8")
                    snapshots[engine] = value
                    entry = dict(status="GENERATED", command=[args.exe]+command, stdin=stdin, output=str(output), observed_e2e_ms=elapsed, timing_eligible=False, source_hashes=before)
                except Exception: entry["error"] = traceback.format_exc()
                row["engines"][engine] = entry
            if "openpyxl" in snapshots:
                for engine in ENGINES:
                    if engine not in snapshots: continue
                    diff = oracle.differences(snapshots["openpyxl"], snapshots[engine])
                    row["engines"][engine].update(status="ALIGNED" if not diff else "NOT_ALIGNED", differences=diff, timing_eligible=not diff)
            row["all_three_aligned"] = all(e["status"] == "ALIGNED" for e in row["engines"].values())
            results.append(row)
            (root / "PARITY.json").write_text(json.dumps(dict(results=results), ensure_ascii=False, indent=2, default=str), encoding="utf8")
            print(json.dumps(dict(fixture=fixture, mode=mode, statuses={k:v["status"] for k,v in row["engines"].items()}, differences={k:len(v.get("differences",[])) for k,v in row["engines"].items()})), flush=True)
    return int(not all(r["all_three_aligned"] for r in results))


if __name__ == "__main__": raise SystemExit(main())
