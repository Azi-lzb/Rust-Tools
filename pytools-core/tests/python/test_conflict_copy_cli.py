"""Deterministic concurrent-save regression using a real isolated Openpyxl worker.

The test worker saves a synthetic 'user edit' after producing its candidate.
No production hooks, real business inputs, COM, or dist updates are involved.
"""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess

import openpyxl

ROOT = Path(__file__).resolve().parents[3]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=False)
    spec = importlib.util.spec_from_file_location("conflict_fixtures", ROOT / "tools/AsposeWorker/test_menu_matrix.py")
    fixtures = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(fixtures)
    results = []
    cases = [("5-2", False, True), ("2-3", False, True), ("1-1", False, True),
             ("5-2", True, True), ("2-3", True, True), ("1-1", True, True),
             ("5-2", True, False), ("5-2", "unreadable", True)]
    for case, corrupt, change_source in cases:
        label = case + ("-unreadable" if corrupt == "unreadable" else "-invalid" if corrupt else "-valid")
        if not change_source: label += "-source-unchanged"
        directory = root / label
        command, stdin, outputs, readonly = fixtures.prepare(directory, case, False)
        target = outputs[0]
        config = directory / "config.xlsx"
        protected = {p: digest(p) for p in [config, *readonly] if p != target}
        original_bytes = target.read_bytes()
        original = openpyxl.load_workbook(target)
        original.active["Z99"] = "NEW USER EDIT; DO NOT OVERWRITE"
        newer = directory / "new-user-edit.xlsx"
        original.save(newer)
        original.close()
        new_bytes = newer.read_bytes()
        wrapper = directory / "controlled-worker.py"
        wrapper.write_text(
            "import importlib.util,sys,shutil\n"
            "from pathlib import Path\n"
            f"spec=importlib.util.spec_from_file_location('real_worker', {str(ROOT / 'crates/excel-editor/python/openpyxl_worker.py')!r})\n"
            "w=importlib.util.module_from_spec(spec);sys.modules[spec.name]=w;spec.loader.exec_module(w)\n"
            "apply=w._apply\n"
            "def controlled(request):\n"
            " result=apply(request)\n"
            + (f" shutil.copyfile({str(newer)!r}, {str(target)!r})\n" if change_source else "")
            + (" Path(request['output']).write_bytes(b'BROKEN CONTAINER')\n" if corrupt == "unreadable" else
               " import openpyxl\n b=openpyxl.load_workbook(request['output'])\n b.active['A1']='CORRUPTED NON-TARGET HEADER'\n b.save(request['output']);b.close()\n" if corrupt else "")
            + " return result\n"
            "w._apply=controlled\n"
            "sys.exit(w.main())\n", encoding="utf-8")
        env = {k: v for k, v in os.environ.items() if not k.startswith(("PYTOOLS_", "RUSTTOOLS_ASPOSE_"))}
        env.update(PYTOOLS_OPENPYXL_PYTHON=os.sys.executable, PYTOOLS_OPENPYXL_WORKER=str(wrapper),
                   PYTHONIOENCODING="utf-8", LOCALAPPDATA=str(directory / "state"), XDG_STATE_HOME=str(directory / "state"))
        proc = subprocess.run([str(Path(args.exe).resolve()), *command], input=stdin, cwd=directory,
                              env=env, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
        (directory / "stdout.txt").write_text(proc.stdout, encoding="utf-8")
        (directory / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
        assert target.read_bytes() == (new_bytes if change_source else original_bytes), (label, "original file was overwritten", proc.stdout, proc.stderr)
        for path, before in protected.items(): assert digest(path) == before, str(path)
        copies = list(directory.glob("*_处理结果_源文件已变更_*.xlsx"))
        reviews = list(directory.glob("*_待核对_业务校验不符_*.xlsx"))
        if corrupt:
            assert not copies, "invalid candidate must never become a result copy"
            assert "failed" in proc.stdout.lower() or "失败" in proc.stdout or proc.returncode != 0, proc.stdout
            if corrupt == "unreadable":
                assert not reviews, "unreadable bytes are not a review workbook"
            else:
                assert len(reviews) == 1, (label, proc.stdout, proc.stderr)
                audit = json.loads(Path(str(reviews[0]) + ".run.json").read_text(encoding="utf-8"))
                assert audit["status"] == "review_required"
                assert audit["business_validation_passed"] is False
                assert audit["original_target_replaced"] is False
                assert audit["source_changed"] == change_source, audit
                assert Path(audit["source"]) == target
                assert audit["validation_issue"]
                pending = openpyxl.load_workbook(reviews[0])
                try: assert pending.active["A1"].value == "CORRUPTED NON-TARGET HEADER"
                finally: pending.close()
                assert "待核对副本" in proc.stderr
        else:
            assert not reviews, "validated results must not produce review artifacts"
            assert proc.returncode == 0, proc.stderr
            assert len(copies) == 1, (label, proc.stdout, proc.stderr)
            audit = json.loads(Path(str(copies[0]) + ".run.json").read_text(encoding="utf-8"))
            assert audit["original_target_replaced"] is False
            assert audit["new_edits_merged"] is False
            assert Path(audit["commit"]["target"]) == copies[0]
            assert "未合并" in proc.stderr
            book = openpyxl.load_workbook(copies[0])
            try:
                assert book.active["Z99"].value is None
                if case == "5-2":
                    assert book.active["A2"].value == "标准机构"
                    report = json.loads(proc.stdout)
                    assert Path(report["files"][0]["output_path"]) == copies[0]
                if case == "2-3":
                    assert book.active["A2"].value == "OLD"
                    assert book.active["A3"].value == "A"
                    history = list((directory / "logs").glob("2-3_*.run.json"))
                    assert len(history) == 1
                    run = json.loads(history[0].read_text(encoding="utf-8-sig"))
                    metrics = run["business_metrics"]
                    assert metrics["写入目标数"] == 0, metrics
                    assert metrics["追加行数"] == 0, metrics
                    assert metrics["冲突另存副本数"] == 1, metrics
                    assert metrics["另存副本追加行数"] == 2, metrics
                    # Unified summaries intentionally redact paths to basenames;
                    # the adjacent commit record above retains the full path.
                    assert [Path(p).name for p in run["output_paths"]] == [p.name for p in copies]
                if case == "1-1": assert book["Append"].max_row >= 2
            finally: book.close()
        results.append(dict(case=label, status="PASS", original_preserved=True,
                            conflict_copies=len(copies), review_copies=len(reviews),
                            invalid_candidate_rejected=bool(corrupt)))
        print(json.dumps(results[-1], ensure_ascii=False), flush=True)
    (root / "results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
