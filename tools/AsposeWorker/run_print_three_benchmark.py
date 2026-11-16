"""Correctness-gated real printing pipeline benchmark; never alters inputs/dist.

Timing excludes fixture preparation and assertions. All raw evaluation outputs
are retained. Worker startup is measured by the real production READY handshake,
not estimated by subtracting an unrelated empty process. Rust uses a test-only
resident API driver; actual menu wall-clock times are measured independently.
"""
import argparse
import csv
import json
import os
from pathlib import Path
import platform
import shutil
import statistics
import subprocess
import time

import benchmark_print_three as check


def dump(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2, default=str), encoding="utf8")


def validate(output, engine, expected, before):
    assert output.is_file(), f"Missing {output}"
    actual, raw, excluded = check.view(output)
    differences = check.oracle.differences(expected, actual)
    dump(output.with_suffix(".snapshot.json"), dict(effective=actual, raw=raw, evaluation=excluded, differences=differences))
    assert not differences, differences[:8]
    assert len(excluded) == (1 if engine == "aspose" else 0)
    assert all(check.sha(Path(p)) == digest for p, digest in before.items()), "Input modified"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True)
    parser.add_argument("--parity", required=True)
    parser.add_argument("--driver", default="F:/RustBuild/pytools/debug/examples/print_benchmark.exe")
    parser.add_argument("--exe", default="F:/RustBuild/pytools/debug/RustTools.exe")
    parser.add_argument("--warmup", type=int, default=3)
    parser.add_argument("--rounds", type=int, default=10)
    parser.add_argument("--cli-rounds", type=int, default=3)
    args = parser.parse_args()
    parity_root = Path(args.parity).resolve()
    parity = json.loads((parity_root / "PARITY.json").read_text(encoding="utf8"))
    assert len(parity["results"]) == 9 and all(r["all_three_aligned"] for r in parity["results"]), "Parity gate closed"
    root = Path(args.output).resolve(); root.mkdir(parents=True, exist_ok=False)
    env = {**os.environ, "PYTHONIOENCODING": "utf8", "PYTOOLS_PRINT_BENCH_TIMING": "1", "LOCALAPPDATA": str(root / "localappdata")}
    dump(root / "ENVIRONMENT.json", dict(platform=platform.platform(), python=platform.python_version(),
         exe_sha256=check.sha(Path(args.exe)), driver_sha256=check.sha(Path(args.driver)),
         worker_sha256=check.sha(Path(env["RUSTTOOLS_ASPOSE_WORKER_DLL"])),
         arguments=vars(args), timing_clock="perf_counter_ns (Python), Instant (Rust), Stopwatch (.NET)",
         timing_boundary="pipeline_ms includes config/read/prepare/copy/save/validate/commit/teardown; processing = pipeline - measured READY startup",
         evaluation_policy="Raw output preserved; only attributed evaluation worksheet excluded from correctness/data counts"))
    samples = []
    stderr = (root / "driver.stderr.txt").open("w", encoding="utf8")
    began = time.perf_counter_ns()
    driver = subprocess.Popen([args.driver], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=stderr, text=True, encoding="utf8", cwd=root, env=env, bufsize=1)
    ready_line = driver.stdout.readline()
    driver_start_ms = (time.perf_counter_ns() - began) / 1e6
    assert json.loads(ready_line)["ready"], ready_line
    dump(root / "DRIVER_READY.json", dict(startup_ms=driver_start_ms, ready=json.loads(ready_line)))
    try:
        for row in parity["results"]:
            fixture, mode = row["fixture"], row["mode"]
            # Read the original XLSX oracle, not its human-readable JSON dump:
            # JSON's default=str loses cached datetime types and would create
            # a spurious mismatch even for Openpyxl against itself.
            baseline, _, _ = check.view(Path(row["engines"]["openpyxl"]["output"]))
            for round_no in range(-args.warmup, args.rounds):
                engines = list(check.ENGINES)
                shift = (round_no + args.warmup) % len(engines)
                engines = engines[shift:] + engines[:shift]
                for engine in engines:
                    sample_id = f"{fixture}-m{mode}-{engine}-{'warm' if round_no < 0 else 'measure'}-{abs(round_no)}"
                    directory = root / "api" / sample_id
                    _, _, output, inputs = check.prepare(directory, fixture, mode, engine)
                    shutil.copyfile(parity_root / "canonical" / fixture / "source.xlsx", directory / "source.xlsx")
                    before = {str(p): check.sha(p) for p in inputs}
                    query = dict(run_id=sample_id, engine=engine, config=str(directory / "config.xlsx"))
                    request_start = time.perf_counter_ns()
                    driver.stdin.write(json.dumps(query) + "\n"); driver.stdin.flush()
                    response = driver.stdout.readline()
                    request_ms = (time.perf_counter_ns() - request_start) / 1e6
                    (directory / "response.json").write_text(response, encoding="utf8")
                    reply = json.loads(response)
                    assert reply["ok"] and not reply["task_failed"] and not reply["task_skipped"], reply
                    assert reply["run_id"] == sample_id
                    startup = 0.0; inner = None
                    if engine == "openpyxl":
                        timings = [e["details"] for e in reply["events"] if e.get("message") == "print benchmark timing"]
                        assert len(timings) == 1, reply
                        startup = timings[0]["worker_startup_ms"]; inner = timings[0]["worker_processing_ms"]
                    elif engine == "aspose":
                        host = root / "localappdata" / "rust-tools" / "logs" / "aspose-modify" / sample_id / "host.jsonl"
                        entries = [json.loads(s) for s in host.read_text(encoding="utf8").splitlines() if s.strip()]
                        assert entries and all(e.get("error") is None and e["result"] for e in entries), entries
                        startup = sum(e["result"]["timings"]["startup_ms"] for e in entries)
                        inner = sum(e["result"]["timings"]["worker_ms"] for e in entries)
                    assert 0 <= startup <= reply["pipeline_ms"], reply
                    sample = dict(fixture=fixture, mode=mode, engine=engine, phase="warmup" if round_no < 0 else "measured",
                                  round=round_no, boundary="resident_api", pipeline_ms=reply["pipeline_ms"],
                                  worker_startup_ms=startup, processing_ms=reply["pipeline_ms"] - startup,
                                  worker_internal_ms=inner, request_ms=request_ms, run_id=sample_id,
                                  inputs_sha256=before, output=str(output), correctness="ALIGNED")
                    validate(output, engine, baseline, before)
                    samples.append(sample)
                    with (root / "SAMPLES.jsonl").open("a", encoding="utf8") as f: f.write(json.dumps(sample, ensure_ascii=False) + "\n")
            # The actual product CLI is measured as a separate clock/boundary.
            for round_no in range(args.cli_rounds):
                engines = list(check.ENGINES); engines = engines[round_no % 3:] + engines[:round_no % 3]
                for engine in engines:
                    sample_id = f"{fixture}-m{mode}-{engine}-cli-{round_no}"
                    directory = root / "cli" / sample_id
                    command, stdin, output, inputs = check.prepare(directory, fixture, mode, engine)
                    shutil.copyfile(parity_root / "canonical" / fixture / "source.xlsx", directory / "source.xlsx")
                    before = {str(p): check.sha(p) for p in inputs}
                    child_env = {**env, "LOCALAPPDATA": str(directory / "localappdata"), "PYTOOLS_CONFIG_PATH": str(directory / "config.xlsx")}
                    start = time.perf_counter_ns()
                    process = subprocess.run([args.exe] + command, input=stdin, text=True, encoding="utf8", capture_output=True, cwd=directory, env=child_env, timeout=180)
                    elapsed = (time.perf_counter_ns() - start) / 1e6
                    (directory / "stdout.txt").write_text(process.stdout, encoding="utf8"); (directory / "stderr.txt").write_text(process.stderr, encoding="utf8")
                    assert process.returncode == 0, process.stderr
                    validate(output, engine, baseline, before)
                    sample = dict(fixture=fixture, mode=mode, engine=engine, phase="measured", round=round_no,
                                  boundary="real_menu_cli", e2e_ms=elapsed, inputs_sha256=before, output=str(output), correctness="ALIGNED")
                    samples.append(sample)
                    with (root / "SAMPLES.jsonl").open("a", encoding="utf8") as f: f.write(json.dumps(sample, ensure_ascii=False) + "\n")
            print(json.dumps(dict(fixture=fixture, mode=mode, status="ALIGNED_BENCHMARKED")), flush=True)
    finally:
        driver.stdin.close()
        try: driver.wait(timeout=30)
        except subprocess.TimeoutExpired: driver.terminate(); driver.wait(timeout=10)
        stderr.close()
    assert driver.returncode == 0, driver.returncode
    summary = []
    for fixture in ["reuse", "rich", "plain"]:
        for mode in [1, 2, 3]:
            for engine in check.ENGINES:
                group = [s for s in samples if s["fixture"] == fixture and s["mode"] == mode and s["engine"] == engine and s["phase"] == "measured"]
                api = [s for s in group if s["boundary"] == "resident_api"]; cli = [s for s in group if s["boundary"] == "real_menu_cli"]
                assert len(api) == args.rounds and len(cli) == args.cli_rounds
                result = dict(fixture=fixture, mode=mode, engine=engine, api_n=len(api), cli_n=len(cli))
                for metric in ["processing_ms", "worker_startup_ms", "pipeline_ms"]:
                    result[metric + "_median"] = statistics.median(s[metric] for s in api)
                    result[metric + "_min"] = min(s[metric] for s in api)
                    result[metric + "_max"] = max(s[metric] for s in api)
                result["cli_e2e_ms_median"] = statistics.median(s["e2e_ms"] for s in cli)
                summary.append(result)
    dump(root / "SUMMARY.json", dict(driver_startup_ms=driver_start_ms, results=summary))
    with (root / "SUMMARY.csv").open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(summary[0])); writer.writeheader(); writer.writerows(summary)
    return 0


if __name__ == "__main__": raise SystemExit(main())
