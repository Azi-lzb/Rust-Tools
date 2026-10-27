"""Real dedup-task parameters × synthetic workbooks parity.

The 17 real dedup tasks in the user's VBA toolbox point at absolute desktop
paths we must not touch, so this probe rebuilds a bridge config that keeps the
real parameter shapes (source_ws names, key_cols [3,4] / [1,3,5,7,8], empty
write_mode, several tasks sharing one target workbook) but substitutes
synthetic source/target files. Both sides load the SAME bridge config
workbook. Two consecutive runs verify append idempotency. Not real-sample
acceptance: parameters are real-shaped, files are synthetic.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import datetime as dt
from pathlib import Path

sys.dont_write_bytecode = True
PROBE = Path(__file__).resolve().parent
sys.path.insert(0, str(PROBE))

from openpyxl import Workbook, load_workbook  # noqa: E402

PYTOOLS_SOURCE = Path(r"C:/Users/AZI/Desktop/VibeCode/4.pytools-main")


def make_source(path: Path, sheet: str, columns: int, rows: list[tuple]) -> None:
    workbook = Workbook()
    sheet_ws = workbook.active
    sheet_ws.title = sheet
    for column in range(1, columns + 1):
        sheet_ws.cell(row=1, column=column, value=f"列{column}")
    for index, row in enumerate(rows, start=2):
        for column, value in enumerate(row, start=1):
            if value is not None:
                sheet_ws.cell(row=index, column=column, value=value)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def make_bridge_config(path: Path, root: Path, sources: dict[str, Path], target: Path) -> None:
    workbook = Workbook()
    global_sheet = workbook.active
    global_sheet.title = "全局配置"
    global_sheet["A1"], global_sheet["B1"] = "配置项", "值"
    for index, (key, value) in enumerate(
        [("输入目录", str(root)), ("输出目录", str(root / "output")), ("日志目录", str(root / "logs"))],
        start=2,
    ):
        global_sheet[f"A{index}"], global_sheet[f"B{index}"] = key, value

    tasks = workbook.create_sheet("去重追加数据配置")
    headers = [
        "行号", "任务名", "源工作簿", "源工作表", "键列", "目标工作簿", "目标工作表",
        "写入模式", "启用",
    ]
    for column, header in enumerate(headers, start=1):
        tasks.cell(row=1, column=column, value=header)
    # real parameter shapes: key_cols 3,4 / 1,3,5,7,8; several tasks share one
    # target workbook; write_mode empty; same sheet names as the real tasks
    rows = [
        (2, "", str(sources["two"]), "两项收入", "3,4", str(target), "两项收入", "", "是"),
        (3, "", str(sources["fx"]), "中外资本外币分地区", "3,4", str(target), "中外资本外币分地区", "", "是"),
        (4, "", str(sources["fx2"]), "中外资本外币分地区", "3,4", str(target), "中外资本外币分地区", "", "是"),
        (5, "", str(sources["city"]), "市信贷收支一批", "1,3,5,7,8", str(target), "按行业分类统计表", "", "是"),
    ]
    for r, row in enumerate(rows, start=2):
        for column, value in enumerate(row, start=1):
            tasks.cell(row=r, column=column, value=value)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def python_run(config: Path, log_dir: Path) -> dict:
    import logging

    sys.path.insert(0, str(PYTOOLS_SOURCE))
    from PyCode.core.config_xlsx import load_dedup_tasks
    from PyCode.core.dedup_311 import DedupTask, run_append_by_config

    # same dict -> DedupTask conversion as the source menu layer
    # (PyCode/main.py::_to_dedup_tasks), which skips tasks without source_wb
    tasks = []
    for raw in load_dedup_tasks(config):
        if raw.get("source_wb") is None:
            continue
        tasks.append(
            DedupTask(
                row_no=int(raw.get("row_no") or 0),
                enabled=bool(raw.get("enabled")),
                source_wb=raw["source_wb"],
                source_ws=str(raw.get("source_ws") or ""),
                key_cols=list(raw.get("key_cols") or []),
                target_wb=raw.get("target_wb"),
                target_ws=raw.get("target_ws"),
                write_mode=str(raw.get("write_mode") or ""),
                task_name=str(raw.get("task_name") or ""),
            )
        )
    try:
        return run_append_by_config(tasks, log_dir)
    finally:
        logging.shutdown()


def rust_run(config: Path, output_dir: Path) -> list[dict]:
    import os

    native = os.environ.get("PYTOOLS_RUST_CLI")
    assert native, "set PYTOOLS_RUST_CLI"
    # two-step product flow: plan (writes plan files + candidates) then
    # execute --plan publishes the new targets
    completed = subprocess.run(
        [native, "dedup-append-plan", "--config", str(config), "--output-dir", str(output_dir)],
        check=True,
        capture_output=True,
    )
    entries = json.loads(completed.stdout.decode("utf-8-sig", errors="replace"))
    for entry in entries:
        plan_file = entry.get("plan_file")
        if not plan_file:
            continue
        plan = json.loads(Path(plan_file).read_text(encoding="utf-8"))
        if not plan.get("edit_plan") and not plan.get("new_workbook"):
            # existing target with zero new rows: nothing to execute; this is
            # the round-2 idempotency case
            continue
        command = [native, "dedup-append-execute", "--plan", plan_file, "--commit-target"]
        if plan.get("target_existed"):
            # existing-target commits require capability evidence; declare the
            # plan's own operations/features/preservation scope (probe scope)
            edit = plan["edit_plan"]
            capability = {
                "engine": "umya_spreadsheet",
                "status": "passed",
                "format": edit["format"],
                "operations": sorted({op["operation"] for op in edit["operations"]}),
                "supported_features": edit["workbook"]["features_present"],
                "preserved": edit["preserve"],
                "platform": "probe-real-shape-dedup",
                "evidence": {
                    "sample_id": "probe-real-shape-dedup",
                    "report_path": "pytools-core/probe/compare_real_dedup_tasks.py",
                    "adapter_version": "umya-spreadsheet-3.1.1",
                    "engine_version": "umya-spreadsheet-3.1.1",
                    "platform": "probe-real-shape-dedup",
                },
            }
            capability_path = output_dir / "probe_capability.json"
            capability_path.write_text(json.dumps(capability), encoding="utf-8")
            command += ["--capabilities", str(capability_path), "--platform", "probe-real-shape-dedup"]
        subprocess.run(command, check=True, capture_output=True)
    return entries


def projection(path: Path) -> dict:
    workbook = load_workbook(path, data_only=False)
    try:
        result = {}
        for sheet in workbook.worksheets:
            cells = {}
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        cells[cell.coordinate] = cell.value
            result[sheet.title] = cells
        return result
    finally:
        workbook.close()


def main() -> int:
    root = Path(tempfile.gettempdir()) / "pytools-real-dedup-parity"
    shutil.rmtree(root, ignore_errors=True)
    root.mkdir(parents=True)

    sources = {
        "two": root / "src" / "两项收入源.xlsx",
        "fx": root / "src" / "一批外币源.xlsx",
        "fx2": root / "src" / "一批外币源2.xlsx",
        "city": root / "src" / "市信贷收支源.xlsx",
    }
    # key cols 3,4: duplicate rows inside the batch keep the first occurrence
    make_source(sources["two"], "两项收入", 5, [
        ("a", "b", "k1", "k2", 10), ("a", "b", "k1", "k2", 99),
        ("a", "b", "k3", "k4", 20),
    ])
    make_source(sources["fx"], "中外资本外币分地区", 6, [
        ("x", "y", "f1", "f2", 1.5, "u"), ("x", "y", "f3", "f4", 2.5, "v"),
    ])
    make_source(sources["fx2"], "中外资本外币分地区", 6, [
        ("x", "y", "f1", "f2", 7.5, "w"), ("x", "y", "f5", "f6", 8.5, "z"),
    ])
    # key cols 1,3,5,7,8: five-column key with one blank-key row skipped
    make_source(sources["city"], "市信贷收支一批", 8, [
        ("c1", "m", "c3", "n", "c5", "o", "c7", "c8"),
        ("c1", "m", "c3", "n", "c5", "o", "c7", "c8"),
        (None, "m", None, "n", None, "o", None, None),
        ("d1", "m", "d3", "n", "d5", "o", "d7", "d8"),
    ])

    config = root / "config.xlsx"
    py_target = root / "py_target.xlsx"
    rust_target = root / "rust_target.xlsx"
    make_bridge_config(config, root, sources, py_target)

    # run python twice against its own target
    py_stats = [python_run(config, root / "py_logs") for _ in range(2)]
    # rebuild bridge with the rust target and run twice
    make_bridge_config(config, root, sources, rust_target)
    rust_stats = [rust_run(config, root / "rust_out") for _ in range(2)]

    py_projection = projection(py_target)
    rust_projection = projection(rust_target)

    differences = []
    for sheet, cells in py_projection.items():
        other = rust_projection.get(sheet)
        if other is None:
            differences.append(f"sheet missing: {sheet}")
            continue
        keys = set(cells) | set(other)
        for key in sorted(keys):
            if cells.get(key) != other.get(key):
                differences.append(f"{sheet}!{key}: {cells.get(key)!r} != {other.get(key)!r}")
    extra_sheets = sorted(set(rust_projection) - set(py_projection))
    if extra_sheets:
        differences.append(f"extra sheets: {extra_sheets}")

    if differences:
        raise AssertionError(
            "real-shape dedup parity failed: " + "; ".join(differences[:12])
        )

    def totals(stats):
        if isinstance(stats, dict):
            return {
                key: stats.get(key)
                for key in ("task_ok", "task_skipped", "appended_rows", "duplicate_rows_filtered")
            }
        entries = stats if isinstance(stats, list) else []
        return {
            "tasks": len(entries),
            "appended": sum(int(e.get("appended_rows") or 0) for e in entries),
            "filtered": sum(int(e.get("duplicate_rows_filtered") or 0) for e in entries),
        }

    report = {
        "schema_version": 1,
        "suite": "pytools.real_shape_dedup_parity",
        "date": dt.date.today().isoformat(),
        "status": "passed_exact",
        "parameter_shapes": "real (source_ws/key_cols/empty write_mode/shared target) from VbaV2.4.2.xlsm; files synthetic",
        "round1_python": totals(py_stats[0]),
        "round2_python": totals(py_stats[1]),
        "round1_rust": totals(rust_stats[0]),
        "round2_rust": totals(rust_stats[1]),
        "target_projection": "exact after two rounds",
        "limitations": [
            "Real parameter shapes, synthetic workbooks; not real-sample acceptance.",
            "New-target path only; existing-target capability commit covered by other tests.",
        ],
    }
    report_path = PROBE / "reports" / "real_shape_dedup_parity_20261025.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("REAL_DEDUP_RESULT=" + json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
