"""Save-stage evidence: raw SDK archives retain their original evaluation content.

No production menus are invoked. Final candidates follow Worker sanitation.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import zipfile

import openpyxl


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def run(dotnet, dll, requests, root, archive):
    env = dict(os.environ, DOTNET_ROOT=str(Path(dotnet).parent))
    env.pop("RUSTTOOLS_ASPOSE_RAW_ARCHIVE_DIR", None)
    if archive is not None:
        env["RUSTTOOLS_ASPOSE_RAW_ARCHIVE_DIR"] = str(archive)
    proc = subprocess.run([dotnet, dll, "--stdio"], input="".join(json.dumps(r) + "\n" for r in requests),
                          env=env, capture_output=True, text=True, encoding="utf-8", timeout=180)
    (root / "stdout.jsonl").write_text(proc.stdout, encoding="utf-8")
    (root / "stderr.jsonl").write_text(proc.stderr, encoding="utf-8")
    assert proc.returncode == 0, proc.stderr
    rows = [json.loads(line) for line in proc.stdout.splitlines()]
    assert rows[0]["ready"] is True
    assert len(rows) == len(requests) + 1
    return rows[1:]


def request(source, output, index, method="Calculate", operations=None):
    return dict(protocol_version=1, request_id=f"task-{index}", run_id="save-archive-test",
                method=method, input=str(source), candidate=str(output), operations=operations or [])


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--dotnet", required=True)
    p.add_argument("--dll", required=True)
    p.add_argument("--output", required=True)
    args = p.parse_args()
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=False)
    source = root / "source.xlsx"
    b = openpyxl.Workbook()
    s = b.active
    s.title = "Data"
    s["A1"], s["A2"], s["B1"] = 7, 3, "=SUM(A1:A2)"
    s.row_dimensions[1].height = 27
    s.column_dimensions["A"].width = 24
    b.save(source)
    before = digest(source)
    archive = root / "evidence/raw-sdk-save"
    cases = [request(source, root / "calc.xlsx", 1),
             request(source, root / "modify.xlsx", 2, "Modify", [dict(op="set_value", sheet="Data", cell="C3", value=dict(type="text", value="updated"))]),
             request(source, root / "convert.xlsx", 3, "Convert"),
             request(source, root / "calc.xlsx", 4),
             request(source, root / "missing-parent/result.xlsx", 5),
             request(source, root / "invalid-operation.xlsx", 6, "Modify", [dict(op="not_supported")])]
    rows = run(args.dotnet, args.dll, cases, root, archive)
    results = []
    for req, reply in zip(cases[:3], rows[:3]):
        assert reply["ok"] is True, reply
        raw = reply["sdk_save_archive"]
        assert raw["stage"] == "sdk_save_before_worker_preservation"
        path = Path(raw["path"])
        assert path.parent.samefile(archive)
        assert digest(path) == raw["sha256"]
        record = json.loads(Path(raw["manifest_path"]).read_text(encoding="utf-8"))
        assert record["archive_sha256"] == raw["sha256"]
        assert record["request_id"] == req["request_id"]
        assert record["evaluation_content_removed"] is False
        for file in (path, Path(req["candidate"])):
            book = openpyxl.load_workbook(file)
            try:
                assert ("Evaluation Warning" in book.sheetnames) is (file == path)
                assert book["Data"]["B1"].value.lstrip("=") == "SUM(A1:A2)"
                if req["method"] == "Modify":
                    assert book["Data"]["C3"].value == "updated"
            finally:
                book.close()
            if req["method"] in ("Calculate", "Convert"):
                cache = openpyxl.load_workbook(file, data_only=True)
                try:
                    assert cache["Data"]["B1"].value == 10
                finally:
                    cache.close()
        with zipfile.ZipFile(path) as a, zipfile.ZipFile(req["candidate"]) as b:
            raw_parts = {n: a.read(n) for n in a.namelist()}
            preserved_parts = {n: b.read(n) for n in b.namelist()}
        assert raw_parts != preserved_parts, "Fixture must distinguish SDK Save from later preservation"
        assert reply["raw_sdk_archive_ms"] >= 0
        results.append(dict(case=req["method"], status="PASS", raw_archive=raw, raw_evaluation_retained=True, candidate_sanitized=True))
    assert all(not reply["ok"] for reply in rows[3:]), rows[3:]
    assert len(list(archive.glob("*.xlsx"))) == 3
    assert not list(archive.glob("*.partial"))
    assert not (root / "invalid-operation.xlsx").exists()
    results.append(dict(case="existing_target_save_failure_and_invalid_operation", status="PASS", valid_archives_added=0))
    old = {path: digest(path) for path in archive.iterdir()}
    second = root / "second"
    second.mkdir()
    another = run(args.dotnet, args.dll, [request(source, second / "calc.xlsx", 1)], second, archive)[0]
    assert another["ok"]
    assert Path(another["sdk_save_archive"]["path"]) not in old
    for path, hash_value in old.items():
        assert digest(path) == hash_value
    results.append(dict(case="repeated_names_keep_previous_archives", status="PASS"))
    legacy = root / "direct-without-archive-option"
    legacy.mkdir()
    reply = run(args.dotnet, args.dll, [request(source, legacy / "output.xlsx", 1)], legacy, None)[0]
    assert reply["ok"] and reply["sdk_save_archive"] is None
    results.append(dict(case="direct_worker_without_archive_option", status="PASS", archive_reported=False))
    failed = root / "archive-io-failure"
    failed.mkdir()
    blocker = failed / "evidence-is-a-file"
    blocker.write_bytes(b"keep existing file")
    reply = run(args.dotnet, args.dll, [request(source, failed / "output.xlsx", 1)], failed, blocker)[0]
    assert not reply["ok"]
    assert blocker.read_bytes() == b"keep existing file"
    assert not (failed / "output.xlsx").exists()
    results.append(dict(case="archive_io_failure_preserves_input_and_existing_file", status="PASS"))
    assert digest(source) == before
    result = dict(status="PASS", cases=results, source_sha256=before, source_unchanged=True,
                  raw_archive_unchanged=True, candidate_sanitized=True)
    (root / "RESULT.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
