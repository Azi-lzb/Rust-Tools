"""Real CLI regressions for the reported 1-1/2-3 errors; synthetic inputs only.

Run with --exe, --output and optionally --worker. Retains failed outputs/logs.
Never opens the business paths from the user's terminal transcripts.
"""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import zipfile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.styles import Alignment, PatternFill, Protection, Border, Side

REPO = Path(__file__).resolve().parents[3]
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def inject(path, kind):
    with zipfile.ZipFile(path) as z:
        parts = {i.filename: (i, z.read(i.filename)) for i in z.infolist()}
    name = "xl/worksheets/sheet2.xml"
    xml = ET.fromstring(parts[name][1])
    data = xml.find("{%s}sheetData" % S)
    def cell(address, value=None, formula=None):
        row_number = int(''.join(c for c in address if c.isdigit()))
        row = next((r for r in data if int(r.get("r")) == row_number), None)
        if row is None:
            row = ET.SubElement(data, "{%s}row" % S, r=str(row_number))
        for existing in list(row):
            if existing.get("r") == address: row.remove(existing)
        c = ET.SubElement(row, "{%s}c" % S, r=address)
        if formula:
            attrs, text = formula
            ET.SubElement(c, "{%s}f" % S, attrs).text = text
        if value is not None:
            ET.SubElement(c, "{%s}v" % S).text = str(value)
        return c
    if kind == "empty":
        c = cell("AC126")
        c.set("t", "inlineStr")
        ET.SubElement(ET.SubElement(c, "{%s}is" % S), "{%s}t" % S).text = ""
    else:
        cell("B5", 1); cell("B6", 3); cell("C5", 2); cell("C6", 4)
        if kind == "shared":
            cell("D5", 3, ({"t": "shared", "si": "0", "ref": "D5:D6"}, "B5+C5"))
            cell("D6", 7, ({"t": "shared", "si": "0"}, None))
        elif kind == "array":
            cell("E5", 2, ({"t": "array", "ref": "E5:E6"}, "B5:B6*2"))
            cell("E6", 6)
        elif kind == "dependent":
            cell("D5", 999, ({}, "COUNTA(Combined!A2:A101)"))
    # Keep cell/row order valid. There are no extension-prefix attributes in this fixture.
    for row in data:
        row[:] = sorted(row, key=lambda c: openpyxl.utils.cell.coordinate_to_tuple(c.get("r"))[1])
    data[:] = sorted(data, key=lambda r: int(r.get("r")))
    info, _ = parts[name]
    parts[name] = (info, ET.tostring(xml, encoding="utf-8"))
    temporary = path.with_suffix(".fixture.tmp")
    with zipfile.ZipFile(temporary, "w") as z:
        for info, payload in parts.values(): z.writestr(info, payload)
    os.replace(temporary, path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--worker")
    args = parser.parse_args()
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=False)
    exe = str(Path(args.exe).resolve())
    matrix = module("reported_matrix", REPO / "tools/AsposeWorker/test_menu_matrix.py")
    append = module("reported_append", REPO / "pytools-core/tests/python/test_append_candidate_routes.py")
    env = {k: v for k, v in os.environ.items() if not k.startswith(("PYTOOLS_", "RUSTTOOLS_ASPOSE_"))}
    env.update(PYTOOLS_OPENPYXL_PYTHON=str(Path(os.sys.executable)),
               PYTOOLS_OPENPYXL_WORKER=str(REPO / "crates/excel-editor/python/openpyxl_worker.py"),
               LOCALAPPDATA=str(root / "state"), XDG_STATE_HOME=str(root / "state"))
    if args.worker: env["RUSTTOOLS_ASPOSE_WORKER_EXE"] = str(Path(args.worker).resolve())
    results = []
    def invoke(label, directory, arguments, source, target, config, keys=None):
        before = {p: digest(p) for p in (source, target, config)}
        original = openpyxl.load_workbook(target)
        keep_name = "保留" if "保留" in original.sheetnames else "Keep"
        keep_value = original[keep_name]["A1"].value
        original.close()
        run_env = dict(env, PYTOOLS_CONFIG_PATH=str(config))
        result = subprocess.run([exe, *map(str, arguments)], input=keys, cwd=directory,
                                env=run_env, capture_output=True, text=True, encoding="utf-8",
                                errors="replace", timeout=180)
        (directory / "terminal.stdout.txt").write_text(result.stdout, encoding="utf-8")
        (directory / "terminal.stderr.txt").write_text(result.stderr, encoding="utf-8")
        assert before[source] == digest(source) and before[config] == digest(config)
        if keys is None:
            payload = json.loads(result.stdout)
            reported_success = payload.get("target_write_status") == "succeeded"
        else:
            reports = sorted((directory / "logs").glob("pytools-dedup-config_append-*.json"), key=lambda p: p.stat().st_mtime_ns)
            assert reports, result.stdout + result.stderr
            payload = json.loads(reports[-1].read_text(encoding="utf-8-sig"))
            reported_success = payload["metrics"]["committed_targets"] == 1 and payload["metrics"]["failed_targets"] == 0
            if reported_success:
                assert payload["metrics"]["committed_appended_rows"] == 100
        changed = before[target] != digest(target)
        assert reported_success or not changed, "Failed transaction changed its original target"
        committed = reported_success and changed
        if committed:
            book = openpyxl.load_workbook(target)
            assert book[keep_name]["A1"].value == keep_value
            if keys is not None:
                assert book["Combined"].max_row == 101
                assert list(book["Combined"].values) == expected
            if "shared" in label:
                assert book[keep_name]["D5"].value == "=B5+C5"
                assert book[keep_name]["D6"].value == "=B6+C6"
            if "dependent" in label:
                assert book[keep_name]["D5"].value == "=COUNTA(Combined!A2:A101)"
            book.close()
            if "empty" in label:
                worker_module = module("empty_check", REPO / "crates/excel-editor/python/openpyxl_worker.py")
                assert ("Keep", "AC126") in worker_module._explicit_empty_strings(target)
        row = dict(case=label, status="PASS" if committed and result.returncode == 0 else "FAIL",
                   source_and_config_unchanged=True, committed=committed,
                   failed_target_unchanged=not committed, exit=result.returncode)
        results.append(row)
        (root / "results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(row, ensure_ascii=False), flush=True)
        return committed

    for style in ("alignment", "number_format", "fill", "protection", "border", "shared", "array"):
        label = "1-1-aspose-" + style
        directory = root / label
        arguments, _, targets, sources = matrix.prepare(directory, "1-1", True)
        target, source, config = targets[0], sources[0], directory / "config.xlsx"
        book = openpyxl.load_workbook(target); sheet = book["保留"]
        sheet["C5"] = 123.45
        if style == "alignment": sheet["C5"].alignment = Alignment(horizontal="general", vertical="bottom", indent=0)
        if style == "number_format": sheet["C5"].number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        if style == "fill": sheet["C5"].fill = PatternFill(patternType="solid", fgColor="FFAABBCC", bgColor="FF000000")
        if style == "protection": sheet["C5"].protection = Protection(locked=False, hidden=True)
        if style == "border": sheet["C5"].border = Border(left=Side(style="thin", color="FF112233"))
        book.save(target); book.close()
        if style in ("shared", "array"): inject(target, style)
        invoke(label, directory, arguments, source, target, config)

    for kind, modes in (("empty", ("auto", "openpyxl", "aspose")), ("dependent", ("auto", "openpyxl", "aspose")), ("shared", ("aspose",)), ("array", ("aspose",))):
        for mode in modes:
            label = "2-3-" + kind + "-" + mode
            case = append.AppendCandidateRoutes("runTest"); case.root = root
            directory, source, target, config, expected = case.fixture(label)
            book = openpyxl.load_workbook(target); book.create_sheet("Keep")["A1"] = "keep"
            book.save(target); book.close()
            inject(target, kind)
            keys = "2\n3\n" + {"auto": "1", "openpyxl": "2", "aspose": "3"}[mode] + "\n0\n"
            committed = invoke(label, directory, ["menu", "--config", config, "--log-view", "developer"], source, target, config, keys)
            if committed and mode in ("auto", "openpyxl"):
                before = {p: digest(p) for p in (source, target, config)}
                repeat = subprocess.run([exe, "menu", "--config", str(config), "--log-view", "developer"],
                                        input=keys, cwd=directory, env=dict(env, PYTOOLS_CONFIG_PATH=str(config)),
                                        capture_output=True, text=True, encoding="utf-8", timeout=180)
                (directory / "repeat.stdout.txt").write_text(repeat.stdout, encoding="utf-8")
                (directory / "repeat.stderr.txt").write_text(repeat.stderr, encoding="utf-8")
                assert repeat.returncode == 0
                reports = sorted((directory / "logs").glob("pytools-dedup-config_append-*.json"), key=lambda p: p.stat().st_mtime_ns)
                report = json.loads(reports[-1].read_text(encoding="utf-8-sig"))
                assert report["metrics"]["committed_appended_rows"] == 0
                assert all(digest(p) == value for p, value in before.items()), "repeat changed committed target/source/config"
                results[-1]["repeat_noop_verified"] = True
                (root / "results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    return int(any(r["status"] != "PASS" for r in results))


if __name__ == "__main__":
    raise SystemExit(main())
