"""Validate the packaged CLI with its own runtime and isolated synthetic data."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import zipfile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.workbook.defined_name import DefinedName


bundle = Path(sys.argv[1]).resolve()
evidence = Path(sys.argv[2]).resolve()
repo = Path(sys.argv[3]).resolve()
evidence.mkdir(parents=True, exist_ok=False)
exe = bundle / ("RustTools.exe" if os.name == "nt" else "RustTools")
env = os.environ.copy()
for key in list(env):
    if key.startswith("PYTOOLS_OPENPYXL") or key in {
        "PYTOOLS_COM_PYTHON", "PYTOOLS_COM_WORKER", "PYTOOLS_EXCEL_EDIT_ROUTE",
        "RUST_WORKSPACE_OPENPYXL_PYTHON", "RUSTTOOLS_TEST_COM",
    }:
        env.pop(key, None)
env["XDG_STATE_HOME"] = str(evidence / "state")
env["LOCALAPPDATA"] = str(evidence / "state")
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()


def run(arguments, name, stdin=None, cwd=None):
    result = subprocess.run([str(exe), *map(str, arguments)], cwd=cwd or evidence, env=env,
                            input=stdin, capture_output=True, text=True,
                            encoding="utf-8", errors="replace", timeout=180)
    (evidence / (name + ".stdout.txt")).write_text(result.stdout, encoding="utf-8")
    (evidence / (name + ".stderr.txt")).write_text(result.stderr, encoding="utf-8")
    assert result.returncode == 0, (name, result.returncode, result.stderr[-1500:])
    return result


run(["--help"], "help")
run([], "double-click-menu", "0\n", cwd=bundle)
config = evidence / "config.xlsx"
shutil.copy2(repo / "pytools-core/tests/fixtures/sheet_rename/config.xlsx", config)
config_hash = sha(config)
workbook = evidence / "rename.xlsx"
book = openpyxl.Workbook()
book.active.title = "A"
book.active["A1"] = 7
book.create_sheet("B")["A1"] = 9
summary = book.create_sheet("Summary")
summary["A1"] = "='B'!A1"
summary["A2"] = "=a!A1"
summary["A3"] = '="B!A1"'
summary["A4"] = "B!A1"
book.defined_names.add(DefinedName("Lookup", attr_text="'b'!$A$1"))
book.save(workbook)
book.close()

cached = evidence / "cached.xlsx"
ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
with zipfile.ZipFile(workbook) as source, zipfile.ZipFile(cached, "w", zipfile.ZIP_DEFLATED) as target:
    for item in source.infolist():
        data = source.read(item.filename)
        if item.filename == "xl/worksheets/sheet3.xml":
            sheet = ET.fromstring(data)
            for address, value, kind in [("A1", "9", "n"), ("A2", "7", "n"), ("A3", "B!A1", "str")]:
                cell = sheet.find(f".//{{{ns}}}c[@r='{address}']")
                cell.set("t", kind)
                cache = cell.find(f"{{{ns}}}v")
                if cache is None:
                    cache = ET.SubElement(cell, "{" + ns + "}v")
                cache.text = value
            ET.register_namespace("", ns)
            data = ET.tostring(sheet, encoding="utf-8", xml_declaration=True)
        target.writestr(item, data)
os.replace(cached, workbook)

report = json.loads(run(["rename-sheets", "--config", config, workbook], "rename").stdout)
assert report["rename_ok"] == 2, report
assert report["files"][0]["engine"] == "openpyxl", report
after = openpyxl.load_workbook(workbook, data_only=False)
assert after.sheetnames == ["RenamedA", "A", "Summary"], after.sheetnames
assert (after["RenamedA"]["A1"].value, after["A"]["A1"].value) == (7, 9)
assert [after["Summary"][f"A{i}"].value for i in range(1, 5)] == ["=A!A1", "=RenamedA!A1", '="B!A1"', "B!A1"]
assert after.defined_names["Lookup"].attr_text == "A!$A$1"
after.close()
after = openpyxl.load_workbook(workbook, data_only=True)
assert [after["Summary"][f"A{i}"].value for i in range(1, 5)] == [9, 7, "B!A1", "B!A1"]
after.close()
assert sha(config) == config_hash
renamed_hash = sha(workbook)
repeat = json.loads(run(["rename-sheets", "--config", config, workbook], "repeat").stdout)
assert repeat["rename_ok"] == 0 and sha(workbook) == renamed_hash, repeat
assert not list(evidence.glob(".pytools-sheet-rename-*")), "candidate leaked"

# These tests select three append engines explicitly; no source Worker override.
append_env = env.copy()
append_env["RUSTTOOLS_TEST_EXE"] = str(exe)
with (evidence / "append-routes.log").open("w", encoding="utf-8") as output:
    append = subprocess.run([sys.executable, "-X", "utf8", "-I", "-m", "unittest", "discover",
        "-s", str(repo / "pytools-core/tests/python"), "-p", "test_append_candidate_routes.py", "-v"],
        cwd=evidence, env=append_env, stdout=output, stderr=subprocess.STDOUT, timeout=300)
assert append.returncode == 0, (evidence / "append-routes.log").read_text(encoding="utf-8")

legacy_dir = evidence / "legacy-append-refusal"
legacy_dir.mkdir()
legacy_fixture = repo / "test-runs/platform-validation-20261113/excel-final-fixtures"
for name in ("source.xlsx", "target.xls"):
    shutil.copy2(legacy_fixture / name, legacy_dir / name)
legacy_config = legacy_dir / "config.xlsx"
book = openpyxl.Workbook()
sheet = book.active
sheet.title = "全局配置"
sheet.append(["配置项", "值"])
for key, value in [("输入目录", legacy_dir), ("输出目录", legacy_dir / "output"), ("日志目录", legacy_dir / "logs")]:
    sheet.append([key, str(value)])
sheet = book.create_sheet("去重追加数据配置")
sheet.append(["是否启用", "源数据工作簿", "源数据工作表", "标识列序号", "目标工作簿", "目标工作表", "执行模式"])
sheet.append(["Y", str(legacy_dir / "source.xlsx"), "Data", "1", str(legacy_dir / "target.xls"), "Combined", "追加"])
book.save(legacy_config)
book.close()
before = {p: sha(p) for p in legacy_dir.iterdir() if p.is_file()}
refused = run(["menu", "--config", legacy_config, "--log-view", "user"], "legacy-append-refusal", "2\n3\n1\n0\n")
assert "既有 XLS 去重追加已安全停用" in refused.stdout + refused.stderr
assert all(sha(path) == value for path, value in before.items())
assert not any("candidate" in path.name.lower() for path in legacy_dir.rglob("*"))

result = {"status": "PASS", "exe_sha256": sha(exe), "source_worker_overrides": False,
          "startup": "PASS", "rename_chain_case_names_cache_strings": "PASS",
          "repeat_byte_invariance": "PASS", "append_routes": "PASS",
          "existing_biff_append_safely_refused": "PASS",
          "runtime_resolution": "packaged-layout", "business_originals_used": False}
(evidence / "result.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
print(json.dumps(result), flush=True)
