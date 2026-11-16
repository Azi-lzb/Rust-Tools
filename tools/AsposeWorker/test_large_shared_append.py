"""Large stationary shared groups, append safety and cache invalidation.

Synthetic files only. Timings describe this regression, not an engine benchmark.
"""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import time
import zipfile
from xml.etree import ElementTree as ET

import openpyxl

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--worker", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--exe", help="Also exercise the current Rust 2-3 menu/validation/commit chain")
    args = parser.parse_args()
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=False)
    source = root / "source.xlsx"
    book = openpyxl.Workbook()
    book.active.title = "Data"
    book.active["A1"] = "original header"
    sheet = book.create_sheet("Keep")
    for row in range(5, 10005):
        sheet.cell(row, 2, row)
        sheet.cell(row, 3, 2 * row)
        sheet.cell(row, 4, f"=B{row}+C{row}")
        sheet.cell(row, 27, f"=$B{row}+C{row}")
    book.save(source)
    book.close()
    with zipfile.ZipFile(source) as z:
        parts = [(info, z.read(info.filename)) for info in z.infolist()]
    with zipfile.ZipFile(source, "w") as z:
        for info, data in parts:
            if info.filename == "xl/worksheets/sheet2.xml":
                xml = ET.fromstring(data)
                for cell in xml.iter(S + "c"):
                    address = cell.get("r")
                    if address.startswith("D") or address.startswith("AA"):
                        row = int("".join(c for c in address if c.isdigit()))
                        formula = cell.find(S + "f")
                        formula.attrib.update(t="shared", si="0" if address.startswith("D") else "1")
                        if row == 5:
                            column = "D" if address.startswith("D") else "AA"
                            formula.set("ref", f"{column}5:{column}10004")
                        else:
                            formula.text = None
                        value = cell.find(S + "v")
                        if value is None:
                            value = ET.SubElement(cell, S + "v")
                        value.text = str(3 * row)
                data = ET.tostring(xml, encoding="utf-8")
            z.writestr(info, data)
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    writes = [dict(op="edit_plan_operation", operation=dict(operation="set_cell_value", sheet="Data",
                   cell=f"A{row}", value=dict(type="text", value=f"append-{row}"))) for row in range(100, 1100)]
    requests = []
    for name, ops in [
        ("append", writes),
        ("dependency", [dict(op="set_value", sheet="Keep", cell="B99", value=dict(type="number", value=100))]),
        ("inside-group", [dict(op="set_value", sheet="Keep", cell="AA100", value=dict(type="number", value=100))]),
    ]:
        requests.append(dict(protocol_version=1, request_id=name, run_id="large-shared-append", method="Modify",
                             input=str(source), candidate=str(root / f"{name}.xlsx"), operations=ops))
    started = time.perf_counter()
    with (root / "worker.stderr.jsonl").open("w", encoding="utf-8") as error:
        result = subprocess.run([str(Path(args.worker).resolve()), "--stdio"],
                                input="".join(json.dumps(r) + "\n" for r in requests),
                                stdout=subprocess.PIPE, stderr=error, text=True, encoding="utf-8", timeout=180)
    (root / "worker.stdout.jsonl").write_text(result.stdout, encoding="utf-8")
    responses = [json.loads(line) for line in result.stdout.splitlines()]
    assert result.returncode == 0 and responses[0]["ready"], result.stdout
    replies = {r["request_id"]: r for r in responses[1:]}
    assert replies["append"]["ok"] and replies["dependency"]["ok"] and not replies["inside-group"]["ok"], replies
    assert not (root / "inside-group.xlsx").exists()
    for label in ("append", "dependency"):
        path = root / f"{label}.xlsx"
        formulas = openpyxl.load_workbook(path, data_only=False)
        cached = openpyxl.load_workbook(path, data_only=True)
        assert formulas.sheetnames == ["Data", "Keep", "Evaluation Warning"]
        assert formulas["Data"]["A1"].value == "original header"
        for row in range(5, 10005):
            assert formulas["Keep"].cell(row, 4).value == f"=B{row}+C{row}"
            assert formulas["Keep"].cell(row, 27).value == f"=$B{row}+C{row}"
            for col in (4, 27):
                expected = None if label == "dependency" and row == 99 else 3 * row
                assert cached["Keep"].cell(row, col).value == expected, (label, row, col)
            assert formulas["Keep"].cell(row, 2).value == (100 if label == "dependency" and row == 99 else row)
            assert formulas["Keep"].cell(row, 3).value == 2 * row
        if label == "append":
            assert [formulas["Data"].cell(r, 1).value for r in range(100, 1100)] == [f"append-{r}" for r in range(100, 1100)]
        formulas.close()
        cached.close()
    assert hashlib.sha256(source.read_bytes()).hexdigest() == original_hash
    cli = None
    if args.exe:
        repo = Path(__file__).resolve().parents[2]
        spec = importlib.util.spec_from_file_location("append_routes", repo / "pytools-core/tests/python/test_append_candidate_routes.py")
        helper = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(helper)
        case = helper.AppendCandidateRoutes("runTest")
        case.root = root
        directory, data_source, target, config, expected = case.fixture("cli")
        book = openpyxl.load_workbook(target)
        book.create_sheet("Keep")
        book.save(target)
        book.close()
        with zipfile.ZipFile(source) as z:
            shared_xml = z.read("xl/worksheets/sheet2.xml")
        with zipfile.ZipFile(target) as z:
            parts = [(i, z.read(i.filename)) for i in z.infolist()]
        with zipfile.ZipFile(target, "w") as z:
            for info, data in parts:
                z.writestr(info, shared_xml if info.filename == "xl/worksheets/sheet2.xml" else data)
        before = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in (data_source, config)}
        env = {k: v for k, v in os.environ.items() if not k.startswith(("PYTOOLS_", "RUSTTOOLS_ASPOSE_"))}
        env.update(PYTOOLS_OPENPYXL_PYTHON=os.sys.executable,
                   PYTOOLS_OPENPYXL_WORKER=str(repo / "crates/excel-editor/python/openpyxl_worker.py"),
                   RUSTTOOLS_ASPOSE_WORKER_EXE=str(Path(args.worker).resolve()),
                   LOCALAPPDATA=str(root / "state"), XDG_STATE_HOME=str(root / "state"))
        cli_started = time.perf_counter()
        result = subprocess.run([str(Path(args.exe).resolve()), "menu", "--config", str(config), "--log-view", "developer"],
                                input="2\n3\n3\n0\n", cwd=directory, env=env, capture_output=True, text=True,
                                encoding="utf-8", errors="replace", timeout=180)
        (directory / "terminal.stdout.txt").write_text(result.stdout, encoding="utf-8")
        (directory / "terminal.stderr.txt").write_text(result.stderr, encoding="utf-8")
        reports = sorted((directory / "logs").glob("pytools-dedup-config_append-*.json"), key=lambda p: p.stat().st_mtime_ns)
        assert result.returncode == 0 and reports, result.stdout + result.stderr
        run = json.loads(reports[-1].read_text(encoding="utf-8-sig"))
        assert run["metrics"]["committed_targets"] == 1 and run["metrics"]["committed_appended_rows"] == 100, run["metrics"]
        book = openpyxl.load_workbook(target, data_only=False)
        assert list(book["Combined"].values) == expected
        assert book.sheetnames == ["Combined", "Keep", "Evaluation Warning"]
        for row in range(5, 10005):
            assert book["Keep"].cell(row, 4).value == f"=B{row}+C{row}"
            assert book["Keep"].cell(row, 27).value == f"=$B{row}+C{row}"
            assert book["Keep"].cell(row, 2).value == row and book["Keep"].cell(row, 3).value == 2 * row
        book.close()
        assert all(hashlib.sha256(p.read_bytes()).hexdigest() == h for p, h in before.items())
        cli = dict(status="PASS", committed_rows=100, preserved_formulas=20000, wall_seconds=time.perf_counter() - cli_started)
    report = dict(status="PASS", cases=3, shared_formulas=20000, appended_cells=1000,
                  unchanged_source_sha256=original_hash, total_seconds=time.perf_counter() - started,
                  responses=replies, cli=cli)
    (root / "results.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report))


if __name__ == "__main__":
    main()
