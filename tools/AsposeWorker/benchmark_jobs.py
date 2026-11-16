"""Real CLI batch and named-validation calculation benchmark on synthetic copies."""
import argparse
import importlib.util
import json
import os
from pathlib import Path
import statistics
import subprocess
import time
import zipfile
from xml.etree import ElementTree as ET
import openpyxl
from openpyxl.workbook.defined_name import DefinedName

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("matrix", HERE / "test_menu_matrix.py")
matrix = importlib.util.module_from_spec(spec); spec.loader.exec_module(matrix)


def main():
    p = argparse.ArgumentParser(); p.add_argument("--output", required=True); p.add_argument("--phase", choices=["batch", "calc"], required=True)
    a = p.parse_args(); root = Path(a.output).resolve(); root.mkdir(parents=True, exist_ok=False)
    exe = "F:/RustBuild/pytools/debug/RustTools.exe"; records = []
    repetitions = 1 if a.phase == "batch" else 13
    for i in range(repetitions):
        views = []
        for engine in (["auto", "aspose"] if i % 2 == 0 else ["aspose", "auto"]):
            directory = root / f"{i:02}-{engine}"; directory.mkdir()
            config = matrix.save_config(matrix.config(directory), directory)
            sources = []; before = {}
            if a.phase == "batch":
                for n in range(100):
                    source = directory / f"s{n:03}.xlsx"; matrix.standard_source(source, "甲分机构本外币"); sources.append(source)
                command = ["institution-normalize", "--config", str(config), "--modify-engine", engine, "--apply", "--in-place", "auto"] + list(map(str, sources))
                outputs = sources
            else:
                template = directory / "template.xlsx"; b = openpyxl.Workbook(); s = b.active; s.title = "数据"
                s.append(["项目", "金额"])
                s["D2"] = '=IF(SUM(\'数据\'!B2:B3)=\'参数\'!A1,"通过","校验失败：金额不平")'
                s["E2"] = '=SUMIF(\'数据\'!A2:A3,"甲",\'数据\'!B2:B3)+VLOOKUP("乙",\'数据\'!A2:B3,2,FALSE)'
                b.create_sheet("参数")["A1"] = 30
                b.defined_names.add(DefinedName("表结构区域", attr_text="'数据'!$A$1:$B$1")); b.defined_names.add(DefinedName("校验区域", attr_text="'数据'!$D$2:$E$2")); b.save(template)
                # Deliberately wrong cache, preserving formula text.
                with zipfile.ZipFile(template) as z: parts = {n: z.read(n) for n in z.namelist()}
                ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
                ET.register_namespace("", ns[1:-1])
                xml = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
                for c in xml.iter(ns + "c"):
                    if c.find(ns + "f") is not None:
                        v = c.find(ns + "v")
                        if v is None: v = ET.SubElement(c, ns + "v")
                        v.text = "-999"; c.attrib.pop("t", None)
                parts["xl/worksheets/sheet1.xml"] = ET.tostring(xml)
                with zipfile.ZipFile(template, "w", zipfile.ZIP_DEFLATED) as z:
                    for n, v in parts.items(): z.writestr(n, v)
                source = directory / "source.xlsx"; b = openpyxl.Workbook(); s = b.active; s.title = "数据"
                for row in [["项目", "金额"], ["甲", 10], ["乙", 20]]: s.append(row)
                b.save(source); sources = [source, template]; before = {str(f): matrix.digest(f) for f in sources}
                command = ["summary-named-validation", "--config", str(config), "--calc-engine", "aspose" if engine == "aspose" else "com", "--template", str(template), "--output-dir", str(directory / "output"), "--log-dir", str(directory / "logs"), "--log-view", "developer", str(source)]
                outputs = [directory / "output/副本/source.xlsx"]
            env = {**os.environ, "LOCALAPPDATA": str(directory / "localappdata"), "PYTOOLS_CONFIG_PATH": str(config), "PYTHONIOENCODING": "utf8"}
            if a.phase == "batch" and engine == "auto":
                env["PYTOOLS_OPENPYXL_WORKER"] = str(HERE / "openpyxl_launch_trace.py")
                env["ASPOSE_BENCH_OPENPYXL_TRACE"] = str(directory / "python-launches.jsonl")
            start = time.perf_counter_ns(); proc = subprocess.run([exe] + command, capture_output=True, text=True, encoding="utf8", cwd=directory, env=env, timeout=240)
            elapsed = (time.perf_counter_ns() - start) / 1e6
            (directory / "stdout.txt").write_text(proc.stdout, encoding="utf8"); (directory / "stderr.txt").write_text(proc.stderr, encoding="utf8")
            assert proc.returncode == 0, proc.stderr[-3000:]
            for f, digest in before.items(): assert matrix.digest(Path(f)) == digest
            current = []
            for f in outputs:
                view, excluded = matrix.snapshot(f); assert len(excluded) == (1 if engine == "aspose" else 0)
                if a.phase == "calc":
                    assert view["sheets"]["数据"]["cells"]["D2"]["cached"] == "通过"
                    assert view["sheets"]["数据"]["cells"]["E2"]["cached"] == 30
                else: assert view["sheets"]["甲分机构本外币"]["cells"]["A2"]["value"] == "标准机构"
                current.append(view)
            views.append(current)
            host = []
            for f in directory.rglob("host*.jsonl"):
                for line in f.read_text("utf8").splitlines():
                    try: host.append(json.loads(line))
                    except json.JSONDecodeError: pass
            row = dict(phase="warmup" if a.phase == "calc" and i < 3 else "measured", round=i, engine=engine, status="PASS", files=len(outputs), end_to_end_ms=elapsed, host=host, source_sha256=before)
            records.append(row)
            with (root / "rounds.jsonl").open("a", encoding="utf8") as f: f.write(json.dumps(row, ensure_ascii=False) + "\n")
            print(json.dumps({k: row[k] for k in ["round", "engine", "files", "end_to_end_ms", "status"]}), flush=True)
        if a.phase == "batch": assert not matrix.differences(*views)
        # COM's typography normalization remains separately reported, not copied.
    summary = []
    for engine in ["auto", "aspose"]:
        rows = [r for r in records if r["engine"] == engine and r["phase"] == "measured"]
        summary.append(dict(engine=engine, n=len(rows), files_per_round=rows[0]["files"], median_e2e_ms=statistics.median(r["end_to_end_ms"] for r in rows)))
    (root / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf8")


if __name__ == "__main__": main()
