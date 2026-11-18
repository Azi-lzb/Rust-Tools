"""Real external Workers plus a post-response fault proxy; synthetic files only.

Checks final host validation, review destinations and unchanged inputs. Evaluation
content is retained byte-for-byte outside the injected business cell/paragraph.
This does not simulate Worker-internal errors or grant a failed candidate approval.
"""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

import openpyxl

ROOT = Path(__file__).resolve().parents[3]
SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def inject(path, document):
    with zipfile.ZipFile(path) as z:
        entries = [(info, z.read(info)) for info in z.infolist()]
    parts = {info.filename: data for info, data in entries}
    if document:
        member = "word/document.xml"
        tree = ET.fromstring(parts[member])
        body = tree.find(f"{{{W}}}body")
        p = ET.Element(f"{{{W}}}p")
        ET.SubElement(ET.SubElement(p, f"{{{W}}}r"), f"{{{W}}}t").text = "FAULT INJECTION: changed business paragraph"
        body.insert(0, p)
    else:
        book = ET.fromstring(parts["xl/workbook.xml"])
        sheets = book.find(f"{{{SS}}}sheets")
        sheet = next(s for s in sheets if s.get("name") != "Evaluation Warning")
        rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
        target = next(r.get("Target") for r in rels if r.get("Id") == rid)
        member = target.lstrip("/") if target.startswith("/") else "xl/" + target
        tree = ET.fromstring(parts[member])
        data = tree.find(f"{{{SS}}}sheetData")
        row = next((r for r in data if r.get("r") == "1"), None)
        if row is None:
            row = ET.Element(f"{{{SS}}}row", r="1")
            data.insert(0, row)
        for cell in list(row):
            if cell.get("r") == "A1":
                row.remove(cell)
        cell = ET.Element(f"{{{SS}}}c", r="A1", t="inlineStr")
        ET.SubElement(ET.SubElement(cell, f"{{{SS}}}is"), f"{{{SS}}}t").text = "FAULT INJECTION: changed business value"
        row.insert(0, cell)
    changed = ET.tostring(tree, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(path, "w") as z:
        for info, data in entries:
            z.writestr(info, changed if info.filename == member else data)


def proxy(worker, document, fault):
    child = subprocess.Popen([worker, "--stdio"], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                             stderr=sys.stderr, text=True, encoding="utf-8")
    def emit(message):
        if "pid" in message:
            message["pid"] = os.getpid()
        print(json.dumps(message, ensure_ascii=False), flush=True)
    try:
        emit(json.loads(child.stdout.readline()))
        for line in sys.stdin:
            request = json.loads(line)
            child.stdin.write(line)
            child.stdin.flush()
            reply = json.loads(child.stdout.readline())
            if fault and request.get("candidate") and reply.get("ok"):
                inject(Path(request["candidate"]), document)
            emit(reply)
    finally:
        child.stdin.close()
        child.wait(timeout=30)


def document_fixture(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>Original business paragraph</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')


def calculation_fixture(root, check_fails=False):
    root.mkdir(parents=True)
    source, template = root / "source.xlsx", root / "template.xlsx"
    b = openpyxl.Workbook()
    s = b.active
    s.title = "数据"
    s.append(["项目", "金额"])
    s.append(["甲", 10])
    s.append(["乙", 20])
    b.save(source)
    s["D2"] = f'=IF(SUM(B2:B3)={999 if check_fails else 30},"通过","未通过")'
    from openpyxl.workbook.defined_name import DefinedName
    b.defined_names.add(DefinedName("表结构区域", attr_text="'数据'!$A$1:$B$1"))
    b.defined_names.add(DefinedName("校验区域", attr_text="'数据'!$D$2"))
    b.save(template)
    return ["summary-named-validation", "--calc-engine", "aspose", "--template", str(template),
            "--output-dir", str(root / "output"), "--log-dir", str(root / "logs"), str(source)], None, root / "output" / "副本"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", required=True)
    parser.add_argument("--worker", required=True)
    parser.add_argument("--words-worker", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--cases", nargs="+", default=["1-1", "2-1", "2-2", "2-3", "3-1", "3-2", "3-3", "4-4", "4-5", "4-7", "5-1", "5-2", "5-3", "5-4", "5-6", "5-7", "1-4", "words"])
    parser.add_argument("--no-fault", action="store_true")
    args = parser.parse_args()
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=False)
    spec = importlib.util.spec_from_file_location("review_paths_fixture", ROOT / "tools/AsposeWorker/test_menu_matrix.py")
    helper = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(helper)
    results = []
    for case in args.cases:
        directory = root / case
        row = dict(case=case, status="FAIL", fault_injected=not args.no_fault)
        try:
            document = case == "words"
            if case.startswith("1-4"):
                command, stdin, expected_parent = calculation_fixture(directory, case.endswith("check-failed"))
                # Use the same independent configuration fixture as other menus.
                config = helper.config(directory)
                helper.save_config(config, directory)
            elif document:
                directory.mkdir(parents=True)
                document_fixture(directory / "source.docx")
                helper.save_config(helper.config(directory), directory)
                expected_parent = directory / "converted"
                command, stdin = ["convert-word", "--config", str(directory / "config.xlsx"), "--engine", "aspose", "--output-dir", str(expected_parent), str(directory / "source.docx")], None
            else:
                command, stdin, outputs, _ = helper.prepare(directory, case, True)
                expected_parent = outputs[0].parent
            protected = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in directory.rglob("*") if p.suffix in (".xlsx", ".docx")}
            if args.no_fault and not case.startswith("1-4") and case != "words":
                protected = {p: h for p, h in protected.items() if p not in outputs}
            wrapper = directory / "worker-proxy.py"
            worker = Path(args.words_worker if document else args.worker).resolve()
            wrapper.write_text(f"import runpy\nm=runpy.run_path({str(Path(__file__).resolve())!r})\nm['proxy']({str(worker)!r},{document!r},{not args.no_fault!r})\n", encoding="utf-8")
            env = {k: v for k, v in os.environ.items() if not k.startswith(("PYTOOLS_", "RUSTTOOLS_ASPOSE_"))}
            env.update(PYTHONIOENCODING="utf-8", LOCALAPPDATA=str(directory / "state"), XDG_STATE_HOME=str(directory / "state"),
                       PYTOOLS_CONFIG_PATH=str(directory / "config.xlsx"), RUSTTOOLS_ASPOSE_DOTNET=sys.executable)
            env["RUSTTOOLS_ASPOSE_WORDS_DLL" if document else "RUSTTOOLS_ASPOSE_WORKER_DLL"] = str(wrapper)
            proc = subprocess.run([str(Path(args.exe).resolve()), *command], input=stdin, cwd=directory,
                                  env=env, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=240)
            (directory / "stdout.txt").write_text(proc.stdout, encoding="utf-8")
            (directory / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
            for p, before in protected.items():
                assert hashlib.sha256(p.read_bytes()).hexdigest() == before, f"Protected input changed: {p}"
            reviews = list(directory.rglob("*_待核对_业务校验不符_*"))
            reviews = [p for p in reviews if p.suffix in (".xlsx", ".docx")]
            if args.no_fault:
                assert not reviews, proc.stdout + proc.stderr
                assert proc.returncode == 0, proc.stdout + proc.stderr
                if case.startswith("1-4"):
                    cached = []
                    for path in expected_parent.rglob("*.xlsx"):
                        b = openpyxl.load_workbook(path, data_only=True)
                        try:
                            if "数据" in b.sheetnames and "Evaluation Warning" in b.sheetnames:
                                cached.append(b["数据"]["D2"].value)
                        finally:
                            b.close()
                    assert ("未通过" if case.endswith("check-failed") else "通过") in cached, proc.stdout + proc.stderr
                elif document:
                    normal = expected_parent / "source.docx"
                    assert normal.is_file(), proc.stdout + proc.stderr
                    with zipfile.ZipFile(normal) as z:
                        assert b"Aspose.Words" in z.read("word/document.xml")
                else:
                    assert all(p.is_file() for p in outputs), proc.stdout + proc.stderr
                    for path in outputs:
                        with zipfile.ZipFile(path) as z:
                            assert b"Evaluation Warning" in z.read("xl/workbook.xml")
            else:
                assert len(reviews) == 1, proc.stdout + proc.stderr
                review = reviews[0]
                assert review.parent.samefile(expected_parent), (review, expected_parent)
                audit = json.loads(Path(str(review) + ".run.json").read_text(encoding="utf-8"))
                assert audit["status"] == "review_required"
                assert audit["business_validation_passed"] is False
                assert audit["original_target_replaced"] is False
                assert audit["validation_issue"]
                assert Path(audit["review_copy"]).samefile(review), (audit["review_copy"], str(review))
                assert Path(audit["source"]).is_file(), "Review record points at a deleted private seed"
                if case == "4-7":
                    assert review.name.startswith("source-Excel修复_Aspose_待核对"), review
                    assert Path(audit["source"]).samefile(directory / "source.xlsx")
                    assert audit["source_changed"] is False
                if case.startswith("3-"):
                    assert review.name.startswith("target_Aspose_待核对"), review
                with zipfile.ZipFile(review) as z:
                    if document:
                        assert b"Aspose.Words" in z.read("word/document.xml")
                    else:
                        assert b"Evaluation Warning" in z.read("xl/workbook.xml")
                row["review_copy"] = str(review)
                row["validation_issue"] = audit["validation_issue"]
            row.update(status="PASS", original_inputs_unchanged=True, review_copies=len(reviews), evaluation_retained=True)
        except Exception as error:
            row["error"] = str(error)[-4000:]
        results.append(row)
        print(json.dumps(row, ensure_ascii=False), flush=True)
    (root / "results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    return int(any(r["status"] != "PASS" for r in results))


if __name__ == "__main__":
    sys.exit(main())
