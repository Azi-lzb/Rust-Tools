"""Stationary grouped-formula preservation and explicit unsafe-write refusals."""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import zipfile
from xml.etree import ElementTree as ET

import openpyxl


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--worker", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    root = Path(args.output).resolve(); root.mkdir(parents=True, exist_ok=False)
    repo = Path(__file__).resolve().parents[2]
    spec = importlib.util.spec_from_file_location("group_fixture", repo / "pytools-core/tests/python/test_reported_111_23.py")
    fixture = importlib.util.module_from_spec(spec); spec.loader.exec_module(fixture)
    sources = {}
    for kind in ("shared", "array"):
        path = root / (kind + ".xlsx")
        book = openpyxl.Workbook(); book.active.title = "Data"; book.active["A1"] = "keep"
        book.create_sheet("Keep")["A1"] = "keep"
        book.save(path); book.close(); fixture.inject(path, kind); sources[kind] = path
    hashes = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in sources.values()}
    imported = root / "import.xlsx"
    book = openpyxl.Workbook(); book.active.title = "Values"; book.active["A1"] = 10
    book.save(imported); book.close()
    hashes[imported] = hashlib.sha256(imported.read_bytes()).hexdigest()
    results = []
    with (root / "worker.stderr.jsonl").open("w", encoding="utf-8") as error:
        proc = subprocess.Popen([str(Path(args.worker).resolve()), "--stdio"], stdin=subprocess.PIPE,
                                stdout=subprocess.PIPE, stderr=error, text=True, encoding="utf-8")
        try:
            assert json.loads(proc.stdout.readline())["ready"]
            def run(label, kind, operations, allowed, preexisting=False):
                candidate = root / (label + ".xlsx")
                if preexisting: candidate.write_bytes(b"existing target must survive")
                request = dict(protocol_version=1, request_id=label, run_id="grouped-regression",
                               method="Modify", input=str(sources[kind]), candidate=str(candidate), operations=operations)
                proc.stdin.write(json.dumps(request) + "\n"); proc.stdin.flush()
                response = json.loads(proc.stdout.readline())
                (root / (label + ".response.json")).write_text(json.dumps(response, indent=2), encoding="utf-8")
                assert response["ok"] == allowed, (label, response)
                assert candidate.exists() == (allowed or preexisting), (label, "failed request left a candidate")
                if preexisting: assert candidate.read_bytes() == b"existing target must survive"
                assert all(hashlib.sha256(p.read_bytes()).hexdigest() == digest for p, digest in hashes.items())
                if allowed:
                    book = openpyxl.load_workbook(candidate, data_only=False)
                    assert book.sheetnames[-1] == "Evaluation Warning"
                    if kind == "shared":
                        assert book["Keep"]["D5"].value == "=B5+C5"
                        assert book["Keep"]["D6"].value == "=B6+C6"
                    book.close()
                    with zipfile.ZipFile(candidate) as z:
                        xml = ET.fromstring(z.read("xl/worksheets/sheet2.xml"))
                        f = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f"
                        groups = {}
                        for formula in xml.iter(f):
                            if formula.get("t") == "shared": groups.setdefault(formula.get("si"), []).append(formula)
                        for group in groups.values():
                            assert len([m for m in group if m.get("ref") and m.text]) == 1
                results.append(dict(case=label, status="PASS", accepted=allowed, originals_unchanged=True))
            for kind in ("shared", "array"):
                run(kind + "-unrelated-write", kind, [dict(op="set_value", sheet="Data", cell="AA100", value=dict(type="text", value="append"))], True)
                run(kind + "-inside-member", kind, [dict(op="set_value", sheet="Keep", cell="D6" if kind == "shared" else "E6", value=dict(type="number", value=99))], False)
                run(kind + "-row-movement", kind, [dict(op="edit_plan_operation", operation=dict(operation="insert_rows", sheet="Keep", first_row=1, count=1))], False)
            run("shared-dependent-write", "shared", [dict(op="set_value", sheet="Keep", cell="B5", value=dict(type="number", value=10))], True)
            candidate = root / "shared-dependent-write.xlsx"
            book = openpyxl.load_workbook(candidate, data_only=True)
            assert book["Keep"]["D5"].value is None, "stale shared-formula cache reused"
            assert book["Keep"]["D6"].value == 7
            book.close()
            run("array-dependent-write", "array", [dict(op="set_value", sheet="Keep", cell="B6", value=dict(type="number", value=10))], True)
            book = openpyxl.load_workbook(root / "array-dependent-write.xlsx", data_only=True)
            assert book["Keep"]["E5"].value is None and book["Keep"]["E6"].value is None
            assert book.calculation.fullCalcOnLoad and book.calculation.forceFullCalc
            book.close()
            def copy(sheet, source_range, target_sheet, at):
                return dict(op="edit_plan_operation", operation=dict(operation="copy_range", source_sheet=sheet,
                    source_range=source_range, target_sheet=target_sheet, target_start=at,
                    copy_values=True, copy_formulas=False, copy_styles=False, copy_comments=False))
            for kind in ("shared", "array"):
                run(kind + "-unrelated-range-copy", kind, [copy("Keep", "B5:C6", "Data", "AA100")], True)
                run(kind + "-dependent-range-copy", kind, [copy("Keep", "B5", "Keep", "B6")], True)
                cached = openpyxl.load_workbook(root / (kind + "-dependent-range-copy.xlsx"), data_only=True)
                if kind == "shared":
                    assert cached["Keep"]["D5"].value == 3 and cached["Keep"]["D6"].value is None
                else:
                    assert cached["Keep"]["E5"].value is None and cached["Keep"]["E6"].value is None
                assert cached.calculation.fullCalcOnLoad and cached.calculation.forceFullCalc
                cached.close()
                run(kind + "-range-copy-over-group", kind, [copy("Keep", "B5:B6", "Keep", "D5" if kind == "shared" else "E5")], False)
                operation = dict(operation="import_range", source_workbook=str(imported), source_sheet="Values", source_range="A1",
                    target_sheet="Keep", target_start="B6", copy_values=True, copy_formulas=False, style_mode="full",
                    copy_comments=False, copy_merges=False, copy_row_heights=False, copy_column_widths=False,
                    copy_row_hidden=False, copy_column_hidden=False, copy_outline_levels=False, copy_all_row_dimensions=False,
                    copy_all_column_dimensions=False, materialize_blank_cells=True, hide_zero=False)
                run(kind + "-dependent-import", kind, [dict(op="edit_plan_operation", operation=operation)], True)
                cached = openpyxl.load_workbook(root / (kind + "-dependent-import.xlsx"), data_only=True)
                assert cached["Keep"]["B6"].value == 10
                if kind == "shared": assert cached["Keep"]["D5"].value == 3 and cached["Keep"]["D6"].value is None
                else: assert cached["Keep"]["E5"].value is None and cached["Keep"]["E6"].value is None
                cached.close()
            run("array-copy-stored-results", "array", [copy("Keep", "E5:E6", "Data", "A10")], True)
            book = openpyxl.load_workbook(root / "array-copy-stored-results.xlsx", data_only=False)
            assert book["Data"]["A10"].value == 2 and book["Data"]["A11"].value == 6
            assert book["Data"]["A10"].data_type != "f" and book["Data"]["A11"].data_type != "f"
            book.close()
            formula_copy = copy("Keep", "E5:E6", "Data", "A10")
            formula_copy["operation"]["copy_formulas"] = True
            run("array-formula-copy-remains-unsupported", "array", [formula_copy], False)
            run("existing-candidate", "shared", [dict(op="set_value", sheet="Data", cell="A10", value=dict(type="text", value="append"))], False, preexisting=True)
        finally:
            proc.stdin.close()
            try: proc.wait(timeout=30)
            except subprocess.TimeoutExpired:
                proc.kill(); proc.wait(); raise
    (root / "results.json").write_text(json.dumps(results, indent=2), encoding="utf-8")
    print(json.dumps(dict(status="PASS", cases=results)))


if __name__ == "__main__": main()
