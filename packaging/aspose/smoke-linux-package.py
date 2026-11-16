"""Extract-only Linux package smoke; synthetic files, no system installation."""
import hashlib
import json
import os
from pathlib import Path
import select
import subprocess
import sys
import time
import zipfile

import openpyxl

bundle, root = (Path(value).resolve() for value in sys.argv[1:3])
root.mkdir(parents=True, exist_ok=False)
env = {key: value for key, value in os.environ.items()
       if not key.startswith(("RUSTTOOLS_ASPOSE_", "DOTNET_", "PYTOOLS_"))
       and key != "RUST_WORKSPACE_OPENPYXL_PYTHON"}
env["XDG_STATE_HOME"] = str(root / "state")
config = root / "config.xlsx"
book = openpyxl.Workbook()
sheet = book.active
sheet.title = "全局配置"
sheet.append(["配置项", "值"])
sheet.append(["输出目录", str(root / "output")])
sheet.append(["日志目录", str(root / "logs")])
book.save(config)
book.close()
env["PYTOOLS_CONFIG_PATH"] = str(config)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def invoke(arguments, name, success=True):
    result = subprocess.run([str(bundle / "RustTools"), *map(str, arguments)],
                            cwd=root, env=env, capture_output=True, text=True,
                            encoding="utf-8", errors="replace", timeout=180)
    (root / (name + ".stdout.txt")).write_text(result.stdout, encoding="utf-8")
    (root / (name + ".stderr.txt")).write_text(result.stderr, encoding="utf-8")
    assert (result.returncode == 0) == success, (name, result.stdout, result.stderr)
    return result


source = root / "formula.xlsx"
book = openpyxl.Workbook()
sheet = book.active
sheet.title = "外汇合成表"
sheet["A1"] = 7
sheet["B1"] = "=SUM(A1,2)"
sheet["AA12"] = "00123"
sheet.oddHeader.right.text = "单位：万元"
book.save(source)
book.close()
before = sha(source)

with (root / "cells-worker.stderr.txt").open("w", encoding="utf-8") as stderr:
    started = time.perf_counter()
    worker = subprocess.Popen([str(bundle / "workers/aspose/AsposeWorker"), "--stdio"],
                              stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                              stderr=stderr, text=True, encoding="utf-8", env=env)
    def receive():
        assert select.select([worker.stdout], [], [], 180)[0], "Worker response timed out"
        return json.loads(worker.stdout.readline())

    try:
        ready = receive()
        startup_ms = (time.perf_counter() - started) * 1000
        assert ready["ready"] and ready["evaluation_mode"]
        candidate = root / "calculated.xlsx"
        request = dict(protocol_version=1, request_id="package-calc", run_id="linux-package-smoke",
                       method="Calculate", input=str(source), candidate=str(candidate))
        worker.stdin.write(json.dumps(request) + "\n")
        worker.stdin.flush()
        response = receive()
        (root / "calculation-response.json").write_text(json.dumps(response, indent=2), encoding="utf-8")
        assert response["ok"], response
        wb = openpyxl.load_workbook(candidate, data_only=True)
        assert wb["外汇合成表"]["B1"].value == 9
        assert wb.sheetnames[-1] == "Evaluation Warning"
        assert wb["外汇合成表"]["AA12"].value == "00123"
        wb.close()
    finally:
        worker.stdin.close()
        try:
            worker.wait(timeout=30)
        except subprocess.TimeoutExpired:
            worker.kill()
            worker.wait()
            raise
    assert worker.returncode == 0
assert sha(source) == before

invoke(["convert-excel", "--config", config, "--engine", "aspose",
        "--output-dir", root / "excel-output", source], "aspose-excel")
assert sha(source) == before
wb = openpyxl.load_workbook(root / "excel-output/formula.xlsx")
assert wb["外汇合成表"]["B1"].value == "=SUM(A1,2)"
assert wb["外汇合成表"]["AA12"].value == "00123"
assert wb.sheetnames[-1] == "Evaluation Warning"
wb.close()

invoke(["fx-header", "--modify-engine", "aspose", "--in-place", "auto",
        "--output-dir", root / "header-output", source], "aspose-header")
wb = openpyxl.load_workbook(source)
assert wb["外汇合成表"].oddHeader.right.text == "单位：万美元"
assert wb["外汇合成表"]["AA12"].value == "00123"
assert wb["外汇合成表"]["A1"].value == 7
assert wb["外汇合成表"]["B1"].value == "=SUM(A1,2)"
assert wb.sheetnames[-1] == "Evaluation Warning"
wb.close()

doc = root / "synthetic.docx"
with zipfile.ZipFile(doc, "w", zipfile.ZIP_DEFLATED) as archive:
    archive.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Synthetic Linux packaged Aspose.Words 00123</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')
doc_before = sha(doc)
invoke(["convert-word", "--config", config, "--engine", "aspose",
        "--output-dir", root / "word-output", doc], "aspose-word")
assert sha(doc) == doc_before
with zipfile.ZipFile(root / "word-output/synthetic.docx") as archive:
    document = archive.read("word/document.xml")
    assert b"Synthetic Linux packaged Aspose.Words 00123" in document
    assert any(b"Evaluation" in archive.read(name) for name in archive.namelist() if name.endswith(".xml"))

result = dict(status="PASS", platform=sys.platform, packaged_discovery=True,
              external_dotnet_or_worker_overrides=False, calculation_result=9,
              excel_source_unchanged=True,
              header_modified=True, word_source_unchanged=True,
              evaluation_output_retained=True, worker_startup_ms=startup_ms)
(root / "result.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
print(json.dumps(result))
