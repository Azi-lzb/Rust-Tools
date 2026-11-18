"""Print the stored data-only view; do not require or refresh formula caches."""
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--modes", nargs="+", type=int, choices=[1, 2, 3], default=[1, 2, 3])
    parser.add_argument("--engines", nargs="+", choices=["openpyxl", "rust"], default=["openpyxl", "rust"])
    args = parser.parse_args()
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=False)
    repo = Path(__file__).resolve().parents[3]
    spec = importlib.util.spec_from_file_location("print_fixture", repo / "tools/AsposeWorker/test_menu_matrix.py")
    fixture = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(fixture)
    results = []
    samples = [("empty", "str", True, None, None), ("missing", "str", False, None, None),
               ("numeric-empty", "n", True, None, None), ("stale", "n", True, "17", 17),
               ("error", "e", True, "#DIV/0!", "#DIV/0!"),
               ("shared", "n", True, "17", 17), ("array", "n", True, "17", 17),
               ("array-missing", "n", True, "17", 17)]
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    for mode in args.modes:
        for engine in args.engines:
            # Exercise all cache representations in mode 1; modes 2/3 reuse
            # decoding, so cover their block placement with stale/missing data.
            selected = samples if mode == 1 else [s for s in samples if s[0] in ("stale", "missing")]
            for sample, cache_type, has_cache, cache_text, expected in selected:
                label = f"mode{mode}-{engine}-{sample}"
                directory = root / label
                command, _, outputs, sources = fixture.prepare(directory, f"3-{mode}", False)
                keys = f"3\n1\n{1 if engine == 'openpyxl' else 2}\n0\n0\n"
                source = sources[0]
                book = openpyxl.load_workbook(source)
                book["独立源"]["A1"] = "=1/0" if sample == "error" else "=2+2" if sample in ("stale", "shared", "array", "array-missing") else '=IF(1,"","x")'
                book.save(source)
                book.close()
                with zipfile.ZipFile(source) as z:
                    parts = [(i, z.read(i.filename)) for i in z.infolist()]
                with zipfile.ZipFile(source, "w") as z:
                    for info, payload in parts:
                        if info.filename == "xl/worksheets/sheet2.xml":
                            xml = ET.fromstring(payload)
                            cell = next(c for c in xml.iter(ns + "c") if c.get("r") == "A1")
                            cell.set("t", cache_type)
                            for value in list(cell.findall(ns + "v")):
                                cell.remove(value)
                            if has_cache:
                                ET.SubElement(cell, ns + "v").text = cache_text
                            if sample in ("shared", "array", "array-missing"):
                                formula = cell.find(ns + "f")
                                formula.set("t", "shared" if sample == "shared" else "array")
                                formula.set("ref", "A1:A2")
                                if sample == "shared":
                                    formula.set("si", "0")
                                data = xml.find(ns + "sheetData")
                                second_row = ET.SubElement(data, ns + "row", r="2")
                                follower = ET.SubElement(second_row, ns + "c", r="A2")
                                if sample == "shared":
                                    ET.SubElement(follower, ns + "f", t="shared", si="0")
                                if sample != "array-missing":
                                    ET.SubElement(follower, ns + "v").text = "23"
                            payload = ET.tostring(xml, encoding="utf-8")
                        z.writestr(info, payload)
                config = directory / "config.xlsx"
                book = openpyxl.load_workbook(config)
                for row in range(2, book["打印配置"].max_row + 1):
                    book["打印配置"].cell(row, 14, "否")
                if sample in ("shared", "array", "array-missing"):
                    book["打印配置"].cell(3, 5, "A1:B2")
                book.save(config)
                book.close()
                protected = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in (source, config)}
                env = {k: v for k, v in os.environ.items() if not k.startswith(("PYTOOLS_", "RUSTTOOLS_ASPOSE_"))}
                env.update(PYTHONIOENCODING="utf-8", PYTOOLS_OPENPYXL_PYTHON=os.sys.executable,
                           PYTOOLS_OPENPYXL_WORKER=str(repo / "crates/excel-editor/python/openpyxl_worker.py"),
                           LOCALAPPDATA=str(directory / "state"), XDG_STATE_HOME=str(directory / "state"))
                proc = subprocess.run([str(Path(args.exe).resolve()), *command], input=keys, cwd=directory,
                                      env=env, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
                (directory / "stdout.txt").write_text(proc.stdout, encoding="utf-8")
                (directory / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
                assert all(hashlib.sha256(p.read_bytes()).hexdigest() == h for p, h in protected.items())
                published = outputs[0].exists()
                error = ""
                if published:
                    book = openpyxl.load_workbook(outputs[0])
                    try:
                        sheet = book["打印结果"]
                        markers = [cell for row in sheet for cell in row if cell.value == "0"]
                        assert len(markers) == 1, "literal text zero marker lost"
                        actual = sheet.cell(markers[0].row, markers[0].column - 1).value
                        assert actual in (None, "") if expected is None else actual == expected, (expected, actual)
                        if sample in ("shared", "array", "array-missing"):
                            follower = sheet.cell(markers[0].row + 1, markers[0].column - 1).value
                            assert follower in (None, "") if sample == "array-missing" else follower == 23, follower
                        assert all(c.data_type != "f" for row in sheet for c in row), "print output retained formulas"
                    except AssertionError as exc:
                        error = str(exc)
                    finally:
                        book.close()
                else:
                    error = "stored-value print did not publish"
                row = dict(case=label, status="PASS" if published and not error and proc.returncode == 0 else "FAIL",
                           published=published, expected_cache_value=expected, source_and_config_unchanged=True, error=error)
                results.append(row)
                (root / "results.json").write_text(json.dumps(results, indent=2), encoding="utf-8")
                print(json.dumps(row), flush=True)
    assert all(r["status"] == "PASS" for r in results), [r for r in results if r["status"] != "PASS"]


if __name__ == "__main__":
    main()
