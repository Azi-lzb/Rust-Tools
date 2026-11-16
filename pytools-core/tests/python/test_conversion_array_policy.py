"""4-4's explicit Aspose route may recalculate arrays without changing data.

Synthetic, isolated files only. Retains evaluation output and terminal evidence.
"""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", required=True)
    parser.add_argument("--worker", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    root = Path(args.output).resolve()
    assert not root.exists(), "use a fresh isolated output directory"
    repo = Path(__file__).resolve().parents[3]
    spec = importlib.util.spec_from_file_location("menu_fixture", repo / "tools/AsposeWorker/test_menu_matrix.py")
    helper = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(helper)
    command, stdin, outputs, sources = helper.prepare(root, "4-4", True)
    source = sources[0]
    book = openpyxl.load_workbook(source)
    sheet = book.create_sheet("Array")
    sheet["A1"] = 2
    sheet["A2"] = 3
    sheet["B1"] = ArrayFormula(ref="B1:B2", text="=A1:A2*2")
    book.save(source)
    book.close()
    config = root / "config.xlsx"
    protected = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in (source, config)}
    env = {k: v for k, v in os.environ.items() if not k.startswith(("PYTOOLS_", "RUSTTOOLS_ASPOSE_"))}
    env.update(PYTHONIOENCODING="utf-8", RUSTTOOLS_ASPOSE_WORKER_EXE=str(Path(args.worker).resolve()),
               PYTOOLS_CONFIG_PATH=str(config), LOCALAPPDATA=str(root / "state"), XDG_STATE_HOME=str(root / "state"))
    result = subprocess.run([str(Path(args.exe).resolve()), *command], input=stdin, cwd=root, env=env,
                            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
    (root / "stdout.txt").write_text(result.stdout, encoding="utf-8")
    (root / "stderr.txt").write_text(result.stderr, encoding="utf-8")
    assert result.returncode == 0 and outputs[0].exists(), result.stdout + result.stderr
    original = openpyxl.load_workbook(source, data_only=False)
    candidate = openpyxl.load_workbook(outputs[0], data_only=False)
    assert candidate.sheetnames == original.sheetnames + ["Evaluation Warning"]
    for sheet in original:
        for row in sheet:
            for cell in row:
                if cell.value is None:
                    continue
                actual = candidate[sheet.title][cell.coordinate].value
                if isinstance(cell.value, ArrayFormula):
                    assert isinstance(actual, ArrayFormula)
                    assert (actual.ref, actual.text) == (cell.value.ref, cell.value.text)
                else:
                    assert actual == cell.value, (sheet.title, cell.coordinate, cell.value, actual)
    original.close()
    candidate.close()
    cached = openpyxl.load_workbook(outputs[0], data_only=True)
    assert cached["Array"]["B1"].value == 4 and cached["Array"]["B2"].value == 6
    cached.close()
    assert all(hashlib.sha256(p.read_bytes()).hexdigest() == h for p, h in protected.items())
    report = dict(status="PASS", output=str(outputs[0]), array_results=[4, 6], originals_unchanged=True, evaluation_retained=True)
    (root / "results.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report))


if __name__ == "__main__":
    main()
