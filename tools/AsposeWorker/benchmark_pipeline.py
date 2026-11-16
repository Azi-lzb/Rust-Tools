"""Serial, isolated benchmarks of the protected production adapters.

Raw evaluation outputs are retained. No license workaround, cache flushing,
process-wide kill, business fixture, or distribution executable is used.
"""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import statistics
import sys
import time


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)
    return module


HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent


def streamed_digest(path, rows, columns, expected_edit=False):
    import openpyxl
    book = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        assert book.sheetnames in (["Data"], ["Data", "Evaluation Warning"])
        if len(book.sheetnames) == 2:
            assert any("Aspose" in str(cell.value) for row in book["Evaluation Warning"].iter_rows(max_row=10, max_col=5) for cell in row)
        sheet = book["Data"]; digest = hashlib.sha256()
        for r, row in enumerate(sheet.iter_rows(max_row=rows, max_col=columns), 1):
            values = []
            for c, cell in enumerate(row, 1):
                value, typ = ("BENCHMARK-修改", "s") if expected_edit and (r, c) == (10, 7) else (cell.value, cell.data_type)
                if value is not None:
                    values.append((c, value, typ, cell.number_format, cell.font.name, cell.font.sz, bool(cell.font.b), bool(cell.font.i)))
            digest.update(json.dumps((r, values), ensure_ascii=False, default=str, separators=(",", ":")).encode("utf8"))
        return digest.hexdigest(), len(book.sheetnames) == 2
    finally: book.close()


def python_worker():
    import openpyxl  # Load before READY, not in the API clock.
    worker = load("production_worker", REPO / "crates/excel-editor/python/openpyxl_worker.py")
    print(json.dumps(dict(ready=True, engine="protected-openpyxl", pid=os.getpid())), flush=True)
    for line in sys.stdin:
        q = json.loads(line); start = time.perf_counter_ns()
        try:
            response = worker._apply(q)
            response.update(ok=response.get("status") == "applied")
        except Exception as error:
            response = dict(ok=False, error=str(error))
        response["processing_ms"] = (time.perf_counter_ns() - start) / 1e6
        print(json.dumps(response), flush=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True)
    parser.add_argument("--exe", default="F:/RustBuild/pytools/debug/RustTools.exe")
    parser.add_argument("--phase", choices=["pure", "menus"], required=True)
    parser.add_argument("--cases", nargs="+")
    args = parser.parse_args()
    root = Path(args.output).resolve(); root.mkdir(parents=True, exist_ok=False)
    matrix = load("menu_matrix", HERE / "test_menu_matrix.py")
    rounds = []
    def record(row):
        rounds.append(row)
        with (root / "rounds.jsonl").open("a", encoding="utf8") as log: log.write(json.dumps(row, ensure_ascii=False, default=str) + "\n")
        print(json.dumps({k: row.get(k) for k in ["case", "engine", "round", "processing_ms", "end_to_end_ms", "status"]}), flush=True)
    if args.phase == "pure":
        transport = load("benchmark_transport", REPO / "benchmark/excel-engine-evaluation/baselines/four_dim/run.py")
        dotnet = "F:/RustBuild/pytools/aspose-worker-sdk-10.0.404/dotnet.exe"
        dll = "F:/RustBuild/pytools/aspose-worker/bin/AsposeWorker.dll"
        workers = {}
        try:
            for name, command in [("openpyxl", [sys.executable, "-u", str(Path(__file__).resolve()), "--python-worker"]), ("aspose", [dotnet, dll, "--stdio"])]:
                workers[name] = transport.Worker(command, root / (name + ".stderr.jsonl"), os.environ.copy())
            (root / "startup.json").write_text(json.dumps({k: dict(launch_to_ready_ms=w.launch_ms, ready=w.ready) for k, w in workers.items()}, indent=2), encoding="utf8")
            fixtures = json.loads((REPO / "benchmark/excel-engine-evaluation/baselines/fixtures/generated/manifest.json").read_text("utf8"))
            for item in fixtures:
                source = Path(item["input"]); before = matrix.digest(source)
                expected_digest, _ = streamed_digest(source, item["rows"], item["cols"], True)
                for i in range(13):
                    snapshots = []
                    for engine in (["openpyxl", "aspose"] if i % 2 == 0 else ["aspose", "openpyxl"]):
                        candidate = root / f'{item["id"]}-{engine}-{i:02}.xlsx'
                        operation = dict(operation="set_cell_value", sheet="Data", cell="G10", value=dict(type="text", value="BENCHMARK-修改"))
                        if engine == "aspose":
                            request = dict(protocol_version=1, request_id=f'{item["id"]}-{i}', run_id="protected-benchmark", method="Modify", input=str(source), candidate=str(candidate), operations=[dict(op="edit_plan_operation", operation=operation)])
                        else:
                            request = dict(schema_version=1, request_id=f'{item["id"]}-{i}', input=str(source), output=str(candidate), format="xlsx", plan=dict(request_id=f'{item["id"]}-{i}', workbook=dict(inventory_complete=True), operations=[operation]))
                        result = workers[engine].request(request)
                        assert result.get("ok"), result
                        digest, evaluation = streamed_digest(candidate, item["rows"], item["cols"])
                        assert evaluation == (engine == "aspose")
                        assert digest == expected_digest, (engine, item["id"], digest, expected_digest)
                        snapshots.append(digest)
                        assert matrix.digest(source) == before
                        record(dict(case=item["id"] + "-modify", engine=engine, round=i, phase="warmup" if i < 3 else "measured", status="PASS", response=result, processing_ms=result["processing_ms"], end_to_end_ms=result["request_e2e_ms"], output=str(candidate), source_sha256=before))
                    assert snapshots[0] == snapshots[1]
        finally:
            for worker in workers.values(): worker.close()
    else:
        cases = args.cases or ["2-1", "2-2", "2-3", "3-1", "3-2", "3-3", "4-4", "4-5", "4-7", "5-1", "5-2", "5-3", "5-4", "5-6", "5-7"]
        import subprocess
        for case in cases:
            for i in range(13):
                snapshots = []
                for aspose in ([False, True] if i % 2 == 0 else [True, False]):
                    engine = "aspose" if aspose else "baseline"; directory = root / case / f'{i:02}-{engine}'
                    command, stdin, outputs, readonly = matrix.prepare(directory, case, aspose)
                    config = directory / "config.xlsx"
                    if config.exists(): readonly += [config]
                    hashes = {str(p): matrix.digest(p) for p in readonly}
                    env = {**os.environ, "PYTHONIOENCODING": "utf8", "LOCALAPPDATA": str(directory / "localappdata"), "PYTOOLS_CONFIG_PATH": str(config)}
                    start = time.perf_counter_ns()
                    proc = subprocess.run([args.exe] + command, input=stdin, text=True, encoding="utf8", capture_output=True, cwd=directory, env=env, timeout=180)
                    elapsed = (time.perf_counter_ns() - start) / 1e6
                    (directory / "stdout.txt").write_text(proc.stdout, encoding="utf8"); (directory / "stderr.txt").write_text(proc.stderr, encoding="utf8")
                    assert proc.returncode == 0, proc.stderr[-2000:]
                    assert all(matrix.digest(Path(p)) == h for p, h in hashes.items())
                    views = []
                    for output in outputs:
                        if output.name == "CONFIG_CANDIDATE":
                            matches = list(directory.rglob("config-candidate.xlsx")); assert len(matches) == 1
                            output = matches[0]
                        view, excluded = matrix.snapshot(output); views.append(view)
                        assert len(excluded) == (1 if aspose else 0), str(output)
                    if case == "7-1":
                        views = json.loads(json.dumps(views, ensure_ascii=False, default=str).replace(json.dumps(str(directory))[1:-1], "<CASE_ROOT>"))
                    if case == "4-4" and aspose:
                        original, _ = matrix.snapshot(directory / "source.xlsx")
                        assert not matrix.differences(original, views[0])
                    snapshots.append(views)
                    record(dict(case=case, engine=engine, round=i, phase="warmup" if i < 3 else "measured", status="PASS", end_to_end_ms=elapsed, outputs=list(map(str, outputs)), source_hashes=hashes))
                diff = matrix.differences(*snapshots)
                if case == "4-4":
                    assert all("/style/font/" in d["path"] or d["path"].endswith("/default_row_height") for d in diff), diff[:5]
                else: assert not diff, diff[:10]
    groups = {}
    for row in rounds:
        if row["phase"] == "measured": groups.setdefault((row["case"], row["engine"]), []).append(row)
    summary = []
    for (case, engine), rows in groups.items():
        row = dict(case=case, engine=engine, n=len(rows), status="PASS" if len(rows) == 10 else "INCOMPLETE")
        for metric in ["processing_ms", "end_to_end_ms"]:
            values = [v[metric] for v in rows if metric in v]
            if values: row[metric] = dict(median=statistics.median(values), min=min(values), max=max(values))
        summary.append(row)
    (root / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf8")


if __name__ == "__main__":
    if "--python-worker" in sys.argv: python_worker()
    else: main()
