"""Synthetic DOCX probe. Evaluation output is retained byte-for-byte."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time
import zipfile
from xml.etree import ElementTree as ET

from docx import Document
from docx.shared import Pt


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dotnet", required=True)
    parser.add_argument("--dll", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=False)
    source = root / "source.docx"
    doc = Document()
    doc.add_heading("合成 Word 转换测试", 1)
    run = doc.add_paragraph().add_run("中文、数字 00123，普通正文。")
    run.bold = True
    run.font.size = Pt(14)
    doc.sections[0].header.paragraphs[0].text = "合成页眉"
    doc.sections[0].footer.paragraphs[0].text = "合成页脚"
    table = doc.add_table(rows=2, cols=2)
    for r in range(2):
        for c in range(2):
            table.cell(r, c).text = f"单元格{r},{c}"
    doc.save(source)
    digest = hashlib.sha256(source.read_bytes()).hexdigest()
    log = (root / "worker.stderr.jsonl").open("w", encoding="utf-8")
    clock = time.perf_counter()
    env = {**os.environ, "DOTNET_ROOT": str(Path(args.dotnet).parent)}
    proc = subprocess.Popen([args.dotnet, args.dll, "--stdio"], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=log, text=True, encoding="utf-8", env=env)
    ready = json.loads(proc.stdout.readline())
    startup = (time.perf_counter() - clock) * 1000
    assert ready["ready"]
    candidate = root / "watermarked.docx"
    request = dict(protocol_version=1, request_id="word-first", run_id="word-synthetic", method="Convert", input=str(source), candidate=str(candidate))
    clock = time.perf_counter()
    proc.stdin.write(json.dumps(request, ensure_ascii=False) + "\n")
    proc.stdin.flush()
    reply = json.loads(proc.stdout.readline())
    request_ms = (time.perf_counter() - clock) * 1000
    proc.stdin.close()
    proc.wait(timeout=10)
    log.close()
    assert reply["ok"], reply
    assert hashlib.sha256(source.read_bytes()).hexdigest() == digest
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs = {}
    with zipfile.ZipFile(candidate) as package:
        for name in package.namelist():
            if name == "word/document.xml" or (name.startswith(("word/header", "word/footer")) and name.endswith(".xml")):
                xml = ET.fromstring(package.read(name))
                paragraphs[name] = ["".join(p.itertext()) for p in xml.findall(".//w:p", ns)]
    result = dict(startup_ms=startup, request_ms=request_ms, ready=ready, result=reply,
                  source_sha256=digest, paragraphs=paragraphs)
    (root / "summary.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
