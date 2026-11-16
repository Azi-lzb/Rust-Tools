"""Cold one-file measurements plus explicitly projected serial 100-file costs."""
import hashlib
import json
import os
import statistics
import sys
import time
from pathlib import Path

from run import Worker, calc_validation, target_validation


def main():
    root = Path(sys.argv[1]).resolve()
    original = root / "run"
    evidence = root / "e2e-followup"
    evidence.mkdir(exist_ok=False)
    raw = evidence / "raw"
    raw.mkdir()
    env = json.loads((original / "environment.json").read_text("utf-8"))
    experiment = json.loads((root / "experiment.json").read_text("utf-8"))
    manifest = json.loads((Path(env["fixtures_root"]) / "manifest.json").read_text("utf-8"))
    rounds = [json.loads(s) for s in (original / "rounds.jsonl").read_text("utf-8").splitlines()]
    keys = list(dict.fromkeys((r["condition"], r["engine"]) for r in rounds))
    inputs = [Path(x["input"]) for x in manifest["read"]]
    inputs += [Path(env["fixtures_root"]) / ("calc-stale." + ext) for ext in ("xls", "xlsx")]
    inputs += [Path(experiment["modify_input"])]
    hashes = {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in inputs}
    samples = []
    with (evidence / "cold-rounds.jsonl").open("w", encoding="utf-8", buffering=1) as log:
        for condition, engine in keys:
            proto = next(r for r in rounds if (r["condition"], r["engine"]) == (condition, engine))
            for i in range(3):
                q = {"op": proto["op"]}
                if q["op"] == "read":
                    f = next(f for f in manifest["read"] if "read-" + f["id"] == condition)
                    q.update(input=f["input"], rows=f["rows"], cols=f["cols"])
                elif q["op"] == "calc":
                    q["input"] = str(Path(env["fixtures_root"]) / ("calc-stale." + condition.split("-")[1]))
                elif q["op"] in ("modify", "header"):
                    q["input"] = experiment["modify_input"]
                if q["op"] in ("write", "modify", "header"):
                    q["output"] = str(raw / f"{condition}-{engine}-{i}.xlsx")
                process_env = dict(os.environ)
                if engine == "aspose":
                    process_env.update(DOTNET_TieredCompilation="0", DOTNET_TieredPGO="0")
                start = time.perf_counter_ns()
                worker = Worker(env["commands"][engine], evidence / f"{condition}-{engine}-{i}.stderr.txt", process_env)
                try:
                    result = worker.request(q)
                    response_ms = (time.perf_counter_ns() - start) / 1e6
                finally:
                    worker.close()
                total_ms = (time.perf_counter_ns() - start) / 1e6
                if not result.get("ok"):
                    raise RuntimeError(result)
                if q["op"] == "read":
                    validation = {"ok": result["digest"] == f["digest"]}
                elif q["op"] == "calc":
                    validation = calc_validation(result["values"], manifest["calc_expected"])
                elif q["op"] in ("modify", "header"):
                    validation = target_validation(q["output"], q["op"])
                else:
                    # The same full readback verifier used by the original run,
                    # after stopping the timed writer; not charged to this task.
                    verifier = Worker(env["commands"]["rust"], evidence / f"verify-{engine}-{i}.stderr.txt")
                    try:
                        v = verifier.request(dict(op="verify", input=q["output"], rows=50000, cols=20))
                        validation = {"ok": v.get("ok") and v.get("digest") == manifest["write"]["digest"]}
                    finally:
                        verifier.close()
                record = dict(condition=condition, engine=engine, sample=i,
                              launch_to_ready_ms=worker.launch_ms, cold_to_response_ms=response_ms,
                              cold_through_exit_ms=total_ms, shutdown_ms=total_ms-response_ms,
                              request=result, validation=validation, query=q)
                log.write(json.dumps(record) + "\n")
                samples.append(record)
                print(f"{condition} {engine} {i+1}/3: {total_ms:.2f} ms validation={validation['ok']}", flush=True)
                assert validation["ok"]
    out = []
    for condition, engine in keys:
        s = [r for r in samples if (r["condition"], r["engine"]) == (condition, engine)]
        warm = [r["request_e2e_ms"] for r in rounds if (r["condition"], r["engine"]) == (condition, engine) and r["phase"] == "measured"]
        cold = statistics.median(r["cold_through_exit_ms"] for r in s)
        w = statistics.median(warm)
        out.append(dict(condition=condition, engine=engine, cold_n=3,
                        cold_single_ms=cold, cold_min_ms=min(r["cold_through_exit_ms"] for r in s),
                        cold_max_ms=max(r["cold_through_exit_ms"] for r in s),
                        response_ms=statistics.median(r["cold_to_response_ms"] for r in s),
                        launch_ms=statistics.median(r["launch_to_ready_ms"] for r in s),
                        shutdown_ms=statistics.median(r["shutdown_ms"] for r in s),
                        warm_request_median_ms=w, estimated_100_fresh_s=cold*100/1000,
                        estimated_100_resident_s=(cold+99*w)/1000,
                        batch_status="ESTIMATE_NOT_100_FILE_RUN"))
    (evidence / "summary.json").write_text(json.dumps(out, indent=2), "utf-8")
    integrity = {p: hashlib.sha256(Path(p).read_bytes()).hexdigest() == h for p, h in hashes.items()}
    (evidence / "input-integrity.json").write_text(json.dumps(integrity, indent=2), "utf-8")
    assert all(integrity.values())
    print("COMPLETE: 51 cold single-file runs; 100-file totals are projections, not observations", flush=True)


if __name__ == "__main__":
    main()
