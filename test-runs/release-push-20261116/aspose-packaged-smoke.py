"""Prove packaged workers execute with no SDK or source-path overrides."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import zipfile

import openpyxl
from docx import Document

bundle, root = map(lambda value: Path(value).resolve(), sys.argv[1:3])
root.mkdir(parents=True, exist_ok=False)
env = {key: value for key, value in os.environ.items() if not key.startswith(("RUSTTOOLS_ASPOSE_", "PYTOOLS_", "DOTNET_")) and key != "RUST_WORKSPACE_OPENPYXL_PYTHON"}
env.update(LOCALAPPDATA=str(root / "state"), PYTHONIOENCODING="utf-8")
config = root / "config.xlsx"
book = openpyxl.Workbook()
sheet = book.active
sheet.title = "全局配置"
sheet.append(["配置项", "值"])
for key, value in [("输出目录", str(root / "output")), ("日志目录", str(root / "logs"))]:
    sheet.append([key, value])
book.save(config)
book.close()
env["PYTOOLS_CONFIG_PATH"] = str(config)
exe = bundle / "RustTools.exe"


def invoke(arguments, name):
    result = subprocess.run([str(exe), *map(str, arguments)], cwd=root, env=env,
                            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
    (root / f"{name}.stdout.txt").write_text(result.stdout, encoding="utf-8")
    (root / f"{name}.stderr.txt").write_text(result.stderr, encoding="utf-8")
    assert result.returncode == 0, (name, result.stderr)
    return result


source = root / "外汇.xlsx"
book = openpyxl.Workbook()
sheet = book.active
sheet.title = "外汇合成表"
sheet["A1"] = 7
sheet["AA12"] = "00123"
sheet.oddHeader.right.text = "单位：万元"
book.save(source)
book.close()
invoke(["fx-header", "--modify-engine", "aspose", "--in-place", "auto", "--output-dir", root / "header-output", source], "aspose-header")
book = openpyxl.load_workbook(source)
assert book.sheetnames[-1] == "Evaluation Warning"
assert book["外汇合成表"].oddHeader.right.text == "单位：万美元"
assert book["外汇合成表"]["AA12"].value == "00123"
assert book["外汇合成表"]["A1"].value == 7
book.close()

document = root / "synthetic.docx"
doc = Document()
doc.add_paragraph("Synthetic packaged Aspose.Words fixture 00123")
doc.save(document)
before = hashlib.sha256(document.read_bytes()).hexdigest()
invoke(["convert-word", "--config", config, "--engine", "aspose", "--output-dir", root / "word-output", document], "aspose-word")
target = root / "word-output/synthetic.docx"
assert hashlib.sha256(document.read_bytes()).hexdigest() == before
actual = Document(target)
assert any("Synthetic packaged Aspose.Words fixture 00123" in p.text for p in actual.paragraphs)
with zipfile.ZipFile(target) as package:
    assert any(b"Evaluation" in package.read(name) for name in package.namelist() if name.endswith(".xml"))
result = dict(status="PASS", packaged_discovery=True, external_dotnet_or_worker_overrides=False,
              header_modified=True, word_source_unchanged=True, evaluation_output_retained=True)
(root / "result.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
print(json.dumps(result))
