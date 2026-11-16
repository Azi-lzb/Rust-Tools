"""Synthetic DOC -> DOCX parity and real CLI timing, with raw evaluation output."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import statistics
import subprocess
import time
import openpyxl
from docx import Document
from docx.shared import Pt
import win32com.client


def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()


def snapshot(path):
    document = Document(path)
    def paragraphs(items):
        result = []
        for p in items:
            if p.text in ["Evaluation Only. Created with Aspose.Words. Copyright 2003-2026 Aspose Pty Ltd.", "Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License https://products.aspose.com/words/temporary-license/"]: continue
            if not p.text: continue
            runs = []
            for r in p.runs:
                if not r.text: continue
                value = dict(text=r.text, bold=r.bold, italic=r.italic, size=None if r.font.size is None else r.font.size.pt)
                # Word may split Chinese and Latin text into adjacent runs.
                # Coalesce only exactly equal formatting, never actual styles.
                if runs and all(runs[-1][key] == value[key] for key in ["bold", "italic", "size"]): runs[-1]["text"] += value["text"]
                else: runs.append(value)
            result.append(dict(text=p.text, runs=runs))
        return result
    return dict(body=paragraphs(document.paragraphs), tables=[[[c.text for c in r.cells] for r in t.rows] for t in document.tables],
                headers=[paragraphs(s.header.paragraphs) for s in document.sections], footers=[paragraphs(s.footer.paragraphs) for s in document.sections], sections=len(document.sections))


def main():
    p = argparse.ArgumentParser(); p.add_argument("--output", required=True); a = p.parse_args()
    root = Path(a.output).resolve(); root.mkdir(parents=True, exist_ok=False)
    original = root / "synthetic.docx"; doc = Document(); doc.add_heading("合成 Word 对拍", 1)
    r = doc.add_paragraph().add_run("中文、00123、正文"); r.bold = True; r.font.size = Pt(14)
    doc.sections[0].header.paragraphs[0].text = "保留页眉"; doc.sections[0].footer.paragraphs[0].text = "保留页脚"
    table = doc.add_table(rows=2, cols=2)
    for row in range(2):
        for col in range(2): table.cell(row, col).text = f"表格{row},{col}"
    doc.save(original); source = root / "synthetic.doc"
    # A separate Office instance only creates the synthetic legacy DOC
    # fixture. No active user document, global process termination or save occurs.
    app = win32com.client.DispatchEx("Word.Application"); app.Visible = False; app.DisplayAlerts = 0; app.AutomationSecurity = 3
    try:
        opened = app.Documents.Open(str(original), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
        try: opened.SaveAs2(str(source), FileFormat=0, AddToRecentFiles=False)
        finally: opened.Close(False)
    finally: app.Quit(False)
    before = {str(f): sha(f) for f in [original, source]}; records = []
    for i in range(13):
        views = []
        for engine in (["auto", "aspose"] if i % 2 == 0 else ["aspose", "auto"]):
            directory = root / f"{i:02}-{engine}"; directory.mkdir()
            config = directory / "config.xlsx"; book = openpyxl.Workbook(); sheet = book.active; sheet.title = "全局配置"
            sheet.append(["配置项", "值"]); sheet.append(["输出目录", str(directory)]); sheet.append(["日志目录", str(directory / "logs")]); book.save(config)
            command = ["F:/RustBuild/pytools/debug/RustTools.exe", "convert-word", "--config", str(config), "--engine", engine, "--output-dir", str(directory), str(source)]
            env = {**os.environ, "LOCALAPPDATA": str(directory / "localappdata"), "PYTOOLS_CONFIG_PATH": str(config), "PYTHONIOENCODING": "utf8"}
            start = time.perf_counter_ns(); result = subprocess.run(command, capture_output=True, text=True, encoding="utf8", env=env, cwd=directory, timeout=120)
            elapsed = (time.perf_counter_ns() - start) / 1e6
            (directory / "stdout.txt").write_text(result.stdout, encoding="utf8"); (directory / "stderr.txt").write_text(result.stderr, encoding="utf8")
            assert result.returncode == 0, result.stderr[-2000:]
            target = directory / "synthetic.docx"; assert target.exists()
            views.append(snapshot(target))
            assert all(sha(Path(path)) == digest for path, digest in before.items())
            row = dict(engine=engine, round=i, phase="warmup" if i < 3 else "measured", e2e_ms=elapsed, status="PASS", output=str(target), source_sha256=before)
            records.append(row)
            with (root / "rounds.jsonl").open("a", encoding="utf8") as f: f.write(json.dumps(row) + "\n")
            print(json.dumps(row), flush=True)
        (root / f"snapshot-{i:02}.json").write_text(json.dumps(views, ensure_ascii=False, indent=2), encoding="utf8")
        assert views[0] == views[1], "Word body/table/run/header/footer mismatch; inspect retained snapshots"
    summary = [dict(engine=e, n=10, median_e2e_ms=statistics.median(r["e2e_ms"] for r in records if r["engine"] == e and r["phase"] == "measured")) for e in ["auto", "aspose"]]
    (root / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf8")


if __name__ == "__main__": main()
