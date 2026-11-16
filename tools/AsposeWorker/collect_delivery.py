"""Read existing captures and write a reproducible delivery index; runs no tests."""
import csv
import hashlib
import json
from pathlib import Path
import re
import statistics
import subprocess

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "test-runs/aspose-full-integration-20261115"


def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    before = json.loads((ROOT / "before.json").read_text("utf-8-sig"))
    changes = []
    # Start snapshots identify task-local deltas even in already-dirty files.
    for source in (ROOT / "prior").rglob("*"):
        if not source.is_file(): continue
        relative = source.relative_to(ROOT / "prior"); current = REPO / relative
        if current.is_file() and source.read_bytes() != current.read_bytes():
            changes.append(dict(path=relative.as_posix(), attribution="本轮叠加在启动快照之上", before_sha256=sha(source), after_sha256=sha(current)))
    for source in (ROOT / "prior-tests").glob("*"):
        relative = Path("crates/excel-editor/python/openpyxl_worker.py") if source.name == "openpyxl_worker.py" else Path("pytools-core/tests") / source.name
        current = REPO / relative
        if current.is_file() and source.read_bytes() != current.read_bytes():
            changes.append(dict(path=relative.as_posix(), attribution="本轮测试/Worker增量，保留修复前快照", before_sha256=sha(source), after_sha256=sha(current)))
    new_paths = ["pytools-core/src/aspose_conversion.rs", "pytools-core/src/aspose_documents.rs", "docs/ASPOSE_FULL_MENU_TRIAL_20261115.md"]
    new_paths += ["tools/AsposeWorker/" + name for name in ["PlanOperations.cs", "PlanSerialization.cs", "FormulaPreservation.cs", "FormulaCachePolicy.cs", "FontPreservation.cs", "StoredValues.cs", "test_plan_operations.py", "test_menu_matrix.py", "test_openpyxl_row_dimensions.py", "benchmark_pipeline.py", "benchmark_jobs.py", "openpyxl_launch_trace.py", "collect_delivery.py"]]
    new_paths += ["tools/AsposeWordsWorker/" + name for name in ["AsposeWordsWorker.csproj", "Program.cs", "test_worker.py", "benchmark_word.py"]]
    for name in new_paths:
        path = REPO / name
        if path.exists(): changes.append(dict(path=name, attribution="本轮新建（依据任务操作记录）", after_sha256=sha(path)))
    changes.append(dict(path="docs/ASPOSE_CALC_INTEGRATION.md", attribution="此前已存在的未跟踪报告；本轮仅追加新报告索引，未将全文归于本轮"))
    (ROOT / "CHANGES.json").write_text(json.dumps(changes, ensure_ascii=False, indent=2), encoding="utf8")
    # A harness-only missing optional config interrupted the menu run before
    # 4-7. Completed cohorts are retained; the remaining cohorts run serially.
    menu_roots = [Path("F:/RustBuild/pytools/aspose-benchmark-menus-final"), Path("F:/RustBuild/pytools/aspose-benchmark-menus-rest"), Path("F:/RustBuild/pytools/aspose-benchmark-fx-menu")]
    groups = {}
    for folder in menu_roots:
        log = folder / "rounds.jsonl"
        if log.exists():
            for line in log.read_text("utf8").splitlines():
                row = json.loads(line)
                if row["phase"] == "measured": groups.setdefault((row["case"], row["engine"]), []).append(row)
    fixed_print = Path("F:/RustBuild/pytools/aspose-benchmark-print-fixed/rounds.jsonl")
    if fixed_print.exists():
        replacement = {}
        for line in fixed_print.read_text("utf8").splitlines():
            row = json.loads(line)
            if row["phase"] == "measured": replacement.setdefault((row["case"], row["engine"]), []).append(row)
        groups.update(replacement)
    if groups:
        combined = []
        for (case, engine), rows in groups.items():
            values = [r["end_to_end_ms"] for r in rows]
            combined.append(dict(case=case, engine=engine, n=len(values), status="PASS" if len(values)==10 else "INCOMPLETE", end_to_end_ms=dict(median=statistics.median(values), min=min(values), max=max(values))))
        (menu_roots[0] / "summary.json").write_text(json.dumps(combined, indent=2), encoding="utf8")
    performance = []
    folders = dict(pure="aspose-benchmark-pure-final", menus="aspose-benchmark-menus-final", batch="aspose-benchmark-batch-final", calc="aspose-benchmark-calc-final-fixed", word="aspose-benchmark-word-verified")
    for kind, name in folders.items():
        folder = Path("F:/RustBuild/pytools") / name
        summary = folder / "summary.json"
        if not summary.exists(): continue
        data = json.loads(summary.read_text("utf8"))
        (ROOT / f"performance-{kind}.json").write_text(json.dumps(dict(root=str(folder), summary=data), ensure_ascii=False, indent=2), encoding="utf8")
        for row in data:
            performance.append(dict(group=kind, case=row.get("case", kind), engine=row["engine"], n=row["n"], processing_ms=row.get("processing_ms", {}).get("median"), end_to_end_ms=row.get("end_to_end_ms", {}).get("median", row.get("median_e2e_ms")), files=row.get("files_per_round", 1), evidence=str(summary)))
    with (ROOT / "PERFORMANCE.csv").open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["group", "case", "engine", "n", "processing_ms", "end_to_end_ms", "files", "evidence"]); writer.writeheader(); writer.writerows(performance)
    tests = []
    for name in ["modify-delivery.log", "calc-delivery.log", "print-delivery.log", "fx-delivery.log", "dedup-verified.log", "foreign-verified.log", "namespace-verified.log"]:
        p = ROOT / name
        if p.exists(): tests.append(dict(log=name, summaries=re.findall(r"test result: .*", p.read_text("utf8"))))
    for p in ROOT.glob("existing-*.log"):
        tests.append(dict(log=p.name, superseded=("dedup_menu_cli" in p.name or p.name.endswith("foreign.log") or "print_menu_cli" in p.name), summaries=re.findall(r"test result: .*", p.read_text("utf8"))))
    protected = {name: sha(REPO / name) for name in ["config.xlsx", "dist/RustTools.exe", "dist/config/config.xlsx"]}
    assert protected == before["protected"], "Protected file changed"
    delivery = dict(head=subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO, text=True).strip(), start_head=before["head"], branch=subprocess.check_output(["git", "branch", "--show-current"], cwd=REPO, text=True).strip(), staged_diff=subprocess.check_output(["git", "diff", "--cached", "--stat"], cwd=REPO, text=True), protected_sha256=protected, tests=tests, performance=performance,
                    not_tested=["full workspace tests", "Release/dist/packages", "Linux/WSL/Docker/Kylin", "WPS acceptance", "FAT/exFAT/shared drives", "real business files", "licensed mode", "watermark removal", "XLS append (still refused)", "XLSM/macro modification", "shared/array formula modification"])
    (ROOT / "DELIVERY.json").write_text(json.dumps(delivery, ensure_ascii=False, indent=2), encoding="utf8")
    print(json.dumps(dict(performance_rows=len(performance), task_delta_files=len(changes), head=delivery["head"], protected=protected), ensure_ascii=False))


if __name__ == "__main__": main()
