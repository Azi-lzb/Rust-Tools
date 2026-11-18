"""Targeted resident-Worker tests; every input is synthetic and stays under --output."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time
import traceback

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName


def fixture(path):
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "数据"
    for row in range(1, 13):
        for col in range(1, 28):
            s.cell(row, col, row * 100 + col)
    s["A1"] = "表头"
    s["B2"] = "00123"
    s["C2"] = "0"
    s["D2"] = False
    s["E2"] = "#N/A"
    s["F2"] = "=SUM(A3,A4)"  # Deliberately uncached: copying values must yield blank.
    s["AA12"].font = Font(bold=True, color="FF112233")
    s["AA12"].fill = PatternFill("solid", fgColor="FFABCDEF")
    s["AA12"].comment = Comment("合成批注", "test")
    s.row_dimensions[12].height = 28
    s.column_dimensions["AA"].width = 24
    wb.create_sheet("其他")["A1"] = "=SUM('数据'!A2:A3)"
    wb.defined_names.add(DefinedName("业务区域", attr_text="'数据'!$A$1:$B$2"))
    wb.save(path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dotnet", required=True)
    parser.add_argument("--dll", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=False)
    source, imported = root / "source.xlsx", root / "import-source.xlsx"
    fixture(source)
    fixture(imported)
    before = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in (source, imported)}
    err = (root / "worker.stderr.jsonl").open("w", encoding="utf-8")
    start = time.perf_counter()
    proc = subprocess.Popen([args.dotnet, args.dll, "--stdio"], stdin=subprocess.PIPE,
                            stdout=subprocess.PIPE, stderr=err, text=True, encoding="utf-8",
                            env={**os.environ, "DOTNET_ROOT": str(Path(args.dotnet).parent)})
    ready = json.loads(proc.stdout.readline())
    startup_ms = (time.perf_counter() - start) * 1000
    assert ready["ready"]
    results = []

    def run(label, ops, check, expected_error=False):
        candidate = root / (label + ".xlsx")
        request = dict(protocol_version=1, request_id=label, run_id="plan-primitives", method="Modify",
                       input=str(source), candidate=str(candidate),
                       operations=[dict(op="edit_plan_operation", operation=p) for p in ops])
        proc.stdin.write(json.dumps(request, ensure_ascii=False) + "\n")
        proc.stdin.flush()
        response = json.loads(proc.stdout.readline())
        (root / (label + ".response.json")).write_text(json.dumps(response, ensure_ascii=False, indent=2), encoding="utf-8")
        try:
            if expected_error:
                assert not response["ok"], response
                assert not candidate.exists(), "Unsupported operation published a candidate"
            else:
                assert response["ok"], response
                wb = openpyxl.load_workbook(candidate)
                assert "Evaluation Warning" not in wb.sheetnames, wb.sheetnames
                assert response["output_sanitation"]["validated"]
                check(wb)
                wb.close()
            for p, digest in before.items():
                assert hashlib.sha256(p.read_bytes()).hexdigest() == digest
            results.append(dict(case=label, status="PASS", response=response))
        except Exception:
            results.append(dict(case=label, status="FAIL", traceback=traceback.format_exc(), response=response))

    def assert_equal(actual, expected):
        assert actual == expected, (actual, expected)

    run("row-delete", [dict(operation="delete_rows", sheet="数据", first_row=10, count=2)],
        lambda w: (assert_equal(w["数据"]["A1"].value, "表头"), assert_equal(w["数据"]["AA10"].value, 1227),
                   assert_equal(w["其他"]["A1"].value, "=SUM('数据'!A2:A3)")))
    run("sheet-structure", [dict(operation="create_sheet_at", name="新表", index=0),
                            dict(operation="rename_sheet", old_name="数据", new_name="新名称"),
                            dict(operation="move_sheet", sheet="其他", new_index=1),
                            dict(operation="delete_sheet", name="新表")],
        lambda w: (assert_equal(w.sheetnames, ["其他", "新名称"]),
                   assert_equal(w["其他"]["A1"].value, "=SUM('新名称'!A2:A3)")))
    run("layout", [dict(operation="set_column_width", sheet="数据", column_index=27, width=22),
                   dict(operation="set_row_height", sheet="数据", row_index=12, height=31),
                   dict(operation="set_freeze_panes", sheet="数据", cell="AA10"),
                   dict(operation="set_auto_filter", sheet="数据", range="A1:AA12"),
                   dict(operation="set_margins", sheet="数据", margins=dict(left=0.3, right=0.4)),
                   dict(operation="set_page_settings", sheet="数据", settings=dict(orientation="landscape", fit_to_page=True, fit_to_width=1, fit_to_height=0, footer="&P"))],
        lambda w: (assert_equal(w["数据"].freeze_panes, "AA10"), assert_equal(w["数据"].row_dimensions[12].height, 31),
                   assert_equal(round(w["数据"].page_margins.left, 6), 0.3), assert_equal(w["数据"].oddFooter.center.text, "&P")))
    run("comments-names", [dict(operation="set_comment", sheet="数据", cell="AA12", text=None, author=None),
                           dict(operation="set_comment", sheet="数据", cell="Z10", text="新批注", author="测试"),
                           dict(operation="set_defined_name", name="新区域", formula="'数据'!$Z$10", local_sheet=None)],
        lambda w: (assert_equal(w["数据"]["AA12"].comment, None), assert_equal(w["数据"]["Z10"].comment.text, "新批注"),
                   assert_equal(w.defined_names["新区域"].attr_text.lstrip("="), "'数据'!$Z$10")))
    run("import", [dict(operation="import_range", source_workbook=str(imported), source_sheet="数据", source_range="Z10:AA12", target_sheet="其他", target_start="C2",
                        copy_values=True, copy_formulas=False, style_mode="full", copy_comments=True, copy_merges=True,
                        copy_row_heights=True, copy_column_widths=True, copy_row_hidden=True, copy_column_hidden=True,
                        copy_outline_levels=True, copy_all_row_dimensions=False, copy_all_column_dimensions=False,
                        materialize_blank_cells=True, hide_zero=False)],
        lambda w: (assert_equal(w["其他"]["D4"].value, 1227), assert_equal(w["其他"]["D4"].font.bold, True),
                   assert_equal(w["其他"]["D4"].comment.text, "合成批注"), assert_equal(w["其他"].row_dimensions[4].height, 28)))
    run("copy-overlap", [dict(operation="copy_range", source_sheet="数据", source_range="A2:B3", target_sheet="数据", target_start="B2", copy_values=True, copy_formulas=False, copy_styles=False, copy_comments=False)],
        lambda w: (assert_equal(w["数据"]["B2"].value, 201), assert_equal(w["数据"]["C2"].value, "00123")))
    run("copy-types-and-missing-cache", [dict(operation="copy_range", source_sheet="数据", source_range="B2:F2", target_sheet="其他", target_start="A10", copy_values=True, copy_formulas=False, copy_styles=True, copy_comments=True)],
        lambda w: (assert_equal(w["其他"]["A10"].value, "00123"), assert_equal(w["其他"]["B10"].value, "0"), assert_equal(w["其他"]["C10"].value, False),
                   assert_equal(w["其他"]["D10"].value, "#N/A"), assert_equal(w["其他"]["D10"].data_type, "e"), assert_equal(w["其他"]["E10"].value, None)))
    run("error-value", [dict(operation="set_cell_value", sheet="数据", cell="AA12", value=dict(type="error", value="#N/A"))],
        lambda w: (assert_equal(w["数据"]["AA12"].value, "#N/A"), assert_equal(w["数据"]["AA12"].data_type, "e")))
    run("decimal-date", [dict(operation="set_cell_value", sheet="数据", cell="AA12", value=dict(type="decimal", value="12.125")),
                         dict(operation="set_cell_value", sheet="数据", cell="Z10", value=dict(type="date_time", value=dict(serial="45000", number_format="yyyy-mm-dd")))],
        lambda w: (assert_equal(w["数据"]["AA12"].value, 12.125), assert_equal(w["数据"]["Z10"].number_format, "yyyy-mm-dd")))
    run("unsafe-precision", [dict(operation="set_cell_value", sheet="数据", cell="AA12", value=dict(type="decimal", value="1234567890123456"))], lambda _: None, True)
    run("unsupported-no-save", [dict(operation="unknown")], lambda _: None, True)
    proc.stdin.close()
    proc.wait(timeout=10)
    err.close()
    summary = dict(startup_ms=startup_ms, ready=ready, source_hashes={str(p): d for p, d in before.items()}, results=results)
    (root / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(dict(startup_ms=startup_ms, cases=[dict(case=r["case"], status=r["status"]) for r in results]), ensure_ascii=False))
    return int(any(r["status"] != "PASS" for r in results))


if __name__ == "__main__":
    raise SystemExit(main())
