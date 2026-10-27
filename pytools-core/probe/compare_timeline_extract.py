"""Compare PT-1-2 Rust extraction with the read-only PyCode implementation."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

WORKSPACE = Path(__file__).resolve().parents[2]
PYTOOLS_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\4.pytools-main")
sys.dont_write_bytecode = True
sys.path.insert(0, str(PYTOOLS_SOURCE))

from PyCode.core.config_xlsx import (  # noqa: E402
    PATH_MAP_COLS,
    TIMELINE_COLS,
    load_global,
    load_path_maps,
    load_timeline_rules,
)
from PyCode.core.timeline_396 import run_timeline_slim  # noqa: E402


def make_source(path: Path, month: int) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "本外币"
    sheet.append([f"2026年{month}月本外币存款", None, None])
    sheet.append(["项目", "本月余额", "比上月"])
    sheet.append(["活期存款", 100 + month, 5])
    sheet.append(["活期存款", 110 + month, 6])
    sheet.append(
        ["定期存款", 200 + month, datetime(2026, month, 15, 12, 30, 0, 123000)]
    )
    sheet.append(["活期存款", "文本余额", True])
    sheet.append(["合计", 410 + month * 3, 18])

    appendix = workbook.create_sheet("本外币附表")
    appendix.append([f"2026年{month}月附表", None, None])
    appendix.append(["项目", "本月余额", "比上月"])
    appendix.append(["活期存款", 20 + month, 2])
    workbook.save(path)


def make_config(
    path: Path,
    input_dir: Path,
    output_dir: Path,
    log_dir: Path,
    target_path: Path | None = None,
    error_policy: str = "continue",
) -> None:
    import pandas as pd

    global_rows = [
        ("输入目录", str(input_dir), ""),
        ("输出目录", str(output_dir), ""),
        ("日志目录", str(log_dir), ""),
        ("错误策略", error_policy, ""),
        ("默认编码", "utf-8", ""),
        ("源文件扩展名", ".xlsx", ""),
        ("时序规则命中策略", "all_match", ""),
    ]
    rule = {column: "" for column in TIMELINE_COLS}
    rule.update(
        {
            "是否启用": "是",
            "规则名称": "活期提取",
            "工作簿关键字": "存款表",
            "工作表关键字": "本外币",
            "行头列": 1,
            "列表头行": "2",
            "必含列头": "本月余额",
            "必含行头": "活期",
            "数据起始行": 3,
            "数据起始列": 2,
            "跳过关键字": "合计",
            "目标工作簿路径": str(target_path) if target_path else "",
            "目标工作表": "追加目标" if target_path else "",
            "启用目标写入": "是" if target_path else "否",
        }
    )
    duplicate_rule = dict(rule)
    duplicate_rule["规则名称"] = "重复规则"
    path_map = [
        "是",
        "活期标准化",
        "",
        "存款表",
        "本外币",
        "行头",
        "精确",
        "活期存款",
        "活期",
        "",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(global_rows, columns=["键", "值", "备注"]).to_excel(
            writer, sheet_name="全局配置", index=False
        )
        pd.DataFrame([rule, duplicate_rule], columns=TIMELINE_COLS).to_excel(
            writer, sheet_name="时序提取规则", index=False
        )
        pd.DataFrame([path_map], columns=PATH_MAP_COLS).to_excel(
            writer, sheet_name="路径标准化映射", index=False
        )


def write_test_capability(plan_path: Path, capability_path: Path) -> str:
    append_plan = json.loads(plan_path.read_text(encoding="utf-8"))
    edit_plan = append_plan.get("edit_plan")
    if not isinstance(edit_plan, dict):
        raise AssertionError("PT-1-2 target plan did not contain an ExcelEditPlan")
    platform = "synthetic-pt-1-2-umya-test"
    capability = {
        "engine": "umya_spreadsheet",
        "status": "passed",
        "format": edit_plan["format"],
        "operations": sorted(
            {operation["operation"] for operation in edit_plan["operations"]}
        ),
        "supported_features": sorted(edit_plan["workbook"]["features_present"]),
        "preserved": sorted(edit_plan["preserve"]),
        "platform": platform,
        "evidence": {
            "sample_id": "synthetic-pt-1-2-existing-target",
            "report_path": "pytools-core/probe/compare_timeline_extract.py",
            "adapter_version": "workspace-test-capability",
            "engine_version": "umya-spreadsheet-3.1.1",
            "platform": platform,
        },
    }
    capability_path.write_text(
        json.dumps(capability, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return platform


def normalize_value(value):
    if isinstance(value, (date, datetime)):
        return [type(value).__name__, value.isoformat()]
    return value


def workbook_projection(path: Path) -> dict:
    workbook = load_workbook(path, data_only=False)
    sheets = []
    for sheet in workbook.worksheets:
        cells = []
        for row in sheet.iter_rows():
            row_cells = []
            for cell in row:
                if cell.value is None and not cell.has_style:
                    row_cells.append(None)
                else:
                    comment = cell.comment
                    row_cells.append(
                        {
                            "value": normalize_value(cell.value),
                            "type": cell.data_type,
                            "bold": cell.font.bold,
                            "number_format": cell.number_format,
                            "fill": [
                                cell.fill.fill_type,
                                cell.fill.fgColor.type,
                                cell.fill.fgColor.rgb,
                                cell.fill.fgColor.indexed,
                            ],
                            "comment": [comment.text, comment.author] if comment else None,
                        }
                    )
            cells.append(row_cells)
        sheets.append(
            {
                "name": sheet.title,
                "state": sheet.sheet_state,
                "dimension": sheet.calculate_dimension(),
                "freeze_panes": str(sheet.freeze_panes) if sheet.freeze_panes else None,
                "merged_ranges": sorted(str(item) for item in sheet.merged_cells.ranges),
                "row_dimensions": {
                    str(index): [dimension.height, dimension.hidden]
                    for index, dimension in sheet.row_dimensions.items()
                    if dimension.height is not None or dimension.hidden
                },
                "column_dimensions": {
                    key: [dimension.width, dimension.hidden]
                    for key, dimension in sheet.column_dimensions.items()
                    if dimension.width != 13.0 or dimension.hidden
                },
                "page_setup": [
                    sheet.page_setup.orientation,
                    sheet.page_setup.fitToWidth,
                    sheet.page_setup.fitToHeight,
                    sheet.page_setup.scale,
                    sheet.sheet_properties.pageSetUpPr.fitToPage
                    if sheet.sheet_properties.pageSetUpPr
                    else None,
                    str(sheet.print_area) if sheet.print_area else None,
                    sheet.oddHeader.center.text,
                    sheet.oddFooter.center.text,
                ],
                "cells": cells,
            }
        )
    return {"sheetnames": workbook.sheetnames, "sheets": sheets}


def run() -> dict:
    rust_cli = Path(
        os.environ.get("PYTOOLS_RUST_CLI", WORKSPACE / "target" / "debug" / "pytools.exe")
    )
    if not rust_cli.is_file():
        raise FileNotFoundError(f"Rust CLI is missing: {rust_cli}")

    with tempfile.TemporaryDirectory(
        prefix="pytools-pt12-parity-", ignore_cleanup_errors=True
    ) as raw_temp:
        root = Path(raw_temp)
        source_dir = root / "sources"
        source_dir.mkdir()
        source_a = source_dir / "本外币存款表_202604.xlsx"
        source_b = source_dir / "本外币存款表_202605.xlsx"
        broken_source = source_dir / "本外币存款表_202606.xlsx"
        nested_source_dir = source_dir / "nested"
        nested_source_dir.mkdir()
        source_c = nested_source_dir / "本外币存款表_202607.xlsx"
        lock_source = source_dir / "~$本外币存款表_202608.xlsx"
        make_source(source_a, 4)
        make_source(source_b, 5)
        broken_source.write_bytes(b"synthetic invalid xlsx")
        make_source(source_c, 7)
        make_source(lock_source, 8)
        config = root / "config" / "config.xlsx"
        python_output = root / "python-output"
        rust_output = root / "rust-output"
        menu_output = root / "menu-output"
        logs = root / "logs"
        make_config(config, source_dir, menu_output, logs)

        global_config = load_global(config)
        global_config.output_dir = python_output
        rules = load_timeline_rules(config)
        path_maps = load_path_maps(config)
        python_path = run_timeline_slim(
            rules,
            path_maps,
            global_config,
            source_paths=[source_a, source_b, broken_source],
        )
        if python_path is None or not python_path.is_file():
            raise AssertionError("PyCode did not produce the expected timeline output")

        command = [
            str(rust_cli),
            "timeline-extract",
            "--config",
            str(config),
            "--output-dir",
            str(rust_output),
            "--log-dir",
            str(logs),
            "--log-view",
            "developer",
            str(source_a),
            str(source_b),
            str(broken_source),
        ]
        rust_run = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if rust_run.returncode != 0:
            raise AssertionError(
                f"Rust CLI exited {rust_run.returncode}: {rust_run.stderr}\n{rust_run.stdout}"
            )
        rust_result = json.loads(rust_run.stdout)
        rust_path = Path(rust_result["path"])
        python_projection = workbook_projection(python_path)
        rust_projection = workbook_projection(rust_path)
        if python_projection != rust_projection:
            raise AssertionError(
                json.dumps(
                    {"python": python_projection, "rust": rust_projection},
                    ensure_ascii=False,
                    default=str,
                )
            )

        python_discovery_global = load_global(config)
        python_discovery_global.output_dir = root / "python-discovery-output"
        python_discovery_path = run_timeline_slim(
            load_timeline_rules(config),
            load_path_maps(config),
            python_discovery_global,
        )
        if python_discovery_path is None or not python_discovery_path.is_file():
            raise AssertionError("PyCode default directory scan produced no output")
        discovery_run = subprocess.run(
            [
                str(rust_cli),
                "timeline-extract",
                "--config",
                str(config),
                "--output-dir",
                str(root / "rust-discovery-output"),
                "--log-dir",
                str(logs),
                "--log-view",
                "developer",
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if discovery_run.returncode != 0:
            raise AssertionError(
                "Rust default directory scan failed: "
                f"{discovery_run.stderr}\n{discovery_run.stdout}"
            )
        discovery_result = json.loads(discovery_run.stdout)
        if discovery_result["total_files"] != 4 or discovery_result["failed_files"] != 1:
            raise AssertionError(
                "Rust default directory scan did not include recursive configured files "
                "and the invalid workbook while excluding the lock file"
            )
        discovery_history = json.loads(
            Path(discovery_result["run_log_path"]).read_text(encoding="utf-8")
        )
        if discovery_history["progress"]["total_inputs"] != 4:
            raise AssertionError("source discovery progress did not report the scanned file count")
        if workbook_projection(python_discovery_path) != workbook_projection(
            Path(discovery_result["path"])
        ):
            raise AssertionError(
                "default input-directory scan differs from the PyCode output projection"
            )

        python_user_global = load_global(config)
        python_user_global.output_dir = root / "python-user-output"
        python_user_path = run_timeline_slim(
            load_timeline_rules(config),
            load_path_maps(config),
            python_user_global,
            source_paths=[source_a],
        )
        if python_user_path is None:
            raise AssertionError("PyCode did not produce output for user-view case")
        user_run = subprocess.run(
            [
                str(rust_cli),
                "timeline-extract",
                "--config",
                str(config),
                "--output-dir",
                str(root / "rust-user-output"),
                "--log-dir",
                str(logs),
                "--log-view",
                "user",
                str(source_a),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if user_run.returncode != 0:
            raise AssertionError(
                f"Rust user-view extraction failed: {user_run.stderr}\n{user_run.stdout}"
            )
        user_result = json.loads(user_run.stdout)
        if workbook_projection(Path(user_result["path"])) != workbook_projection(
            python_user_path
        ):
            raise AssertionError("user-view output differs from the PyCode workbook")
        if "时序提取完成" not in user_run.stderr:
            raise AssertionError("user view did not report the extraction completion step")
        if any(
            engine in (user_run.stdout + user_run.stderr)
            for engine in ("calamine", "rust_xlsxwriter", "umya_spreadsheet")
        ):
            raise AssertionError("user view leaked an Excel engine name")
        user_history = json.loads(
            Path(user_result["run_log_path"]).read_text(encoding="utf-8")
        )
        if user_history["status"] != "succeeded" or not {
            "calamine",
            "rust_xlsxwriter",
        }.issubset(set(user_history["actual_engines"])):
            raise AssertionError("user-view history omitted actual engines or success status")

        event_lines = [
            line for line in rust_run.stderr.splitlines() if line.startswith("{")
        ]
        developer_events = [json.loads(line) for line in event_lines]
        run_history_path = Path(rust_result["run_log_path"])
        history = json.loads(run_history_path.read_text(encoding="utf-8"))
        if history["events"] != developer_events:
            raise AssertionError("CLI developer event stream differs from persisted history")
        if history["feature_id"] != "PT-1-2" or history["status"] != "partial":
            raise AssertionError("unexpected PT-1-2 run-history status")
        rows = len(python_projection["sheets"][0]["cells"]) - 1
        if rows != 20:
            raise AssertionError(f"expected 20 extracted rows, got {rows}")

        menu_input = f"1\n2\n{source_a};{source_b};{broken_source}\n0\n"
        menu_run = subprocess.run(
            [str(rust_cli), "menu", "--config", str(config), "--log-view", "off"],
            input=menu_input,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if menu_run.returncode != 0:
            raise AssertionError(
                f"menu exited {menu_run.returncode}: {menu_run.stderr}\n{menu_run.stdout}"
            )
        if "选择时序提取源文件" not in menu_run.stdout:
            raise AssertionError("the real menu did not route through PT-1-2 source selection")
        result_field = menu_run.stdout.find('"configured_target_rules"')
        marker = menu_run.stdout.rfind("{", 0, result_field) if result_field >= 0 else -1
        if marker < 0:
            raise AssertionError(
                "menu output did not include the timeline extraction result:\n"
                + menu_run.stdout
                + "\n"
                + menu_run.stderr
            )
        menu_result, _ = json.JSONDecoder().raw_decode(menu_run.stdout[marker:])
        menu_projection = workbook_projection(Path(menu_result["path"]))
        if python_projection != menu_projection:
            raise AssertionError("menu 1→2 result differs from the PyCode output")
        if menu_result["failed_files"] != 1:
            raise AssertionError("menu path did not preserve continue-on-error behavior")

        python_new_target = root / "python-new-target.xlsx"
        python_new_config = root / "python-new-target-config" / "config.xlsx"
        make_config(
            python_new_config,
            source_dir,
            root / "python-new-target-output",
            logs,
            python_new_target,
        )
        python_new_global = load_global(python_new_config)
        python_new_summary = run_timeline_slim(
            load_timeline_rules(python_new_config),
            load_path_maps(python_new_config),
            python_new_global,
            source_paths=[source_a, source_b],
        )
        if python_new_summary is None or not python_new_target.exists():
            raise AssertionError("PyCode did not create the absent target workbook")

        rust_new_target = root / "rust-new-target.xlsx"
        rust_new_config = root / "rust-new-target-config" / "config.xlsx"
        make_config(
            rust_new_config,
            source_dir,
            root / "rust-new-target-output",
            logs,
            rust_new_target,
        )
        new_target_run = subprocess.run(
            [
                str(rust_cli),
                "timeline-extract",
                "--config",
                str(rust_new_config),
                "--output-dir",
                str(root / "rust-new-target-output"),
                "--log-view",
                "off",
                str(source_a),
                str(source_b),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if new_target_run.returncode != 0:
            raise AssertionError(
                f"Rust did not create the absent target workbook: {new_target_run.stderr}"
            )
        new_target_result = json.loads(new_target_run.stdout)
        new_target_append = new_target_result["target_append_results"]
        if (
            len(new_target_append) != 1
            or new_target_append[0]["status"] != "succeeded"
            or new_target_append[0]["target_existed"]
            or new_target_append[0]["engine"] != "rust_xlsxwriter"
        ):
            raise AssertionError(f"unexpected absent-target execution: {new_target_append}")
        if workbook_projection(rust_new_target) != workbook_projection(python_new_target):
            raise AssertionError("Rust absent-target output differs from PyCode")

        target_path = root / "existing-target.xlsx"
        target_workbook = Workbook()
        target_workbook.active.title = "保留工作表"
        preserved_sheet = target_workbook.active
        preserved_sheet["A1"] = "must remain unchanged"
        preserved_sheet["A1"].comment = Comment("retain this note", "PyTools test")
        preserved_sheet["B1"].fill = PatternFill("solid", fgColor="FF00FF00")
        preserved_sheet.merge_cells("C1:D1")
        preserved_sheet["C1"] = "retained merge"
        preserved_sheet["A3"] = "=1+1"
        preserved_sheet.freeze_panes = "B2"
        preserved_sheet.row_dimensions[2].height = 24
        preserved_sheet.column_dimensions["B"].width = 30
        preserved_sheet.page_setup.orientation = "landscape"
        preserved_sheet.page_setup.fitToWidth = 1
        preserved_sheet.sheet_properties.pageSetUpPr.fitToPage = True
        preserved_sheet.oddFooter.center.text = "Page &P"
        target_sheet = target_workbook.create_sheet("追加目标")
        target_sheet.append(
            ["数据日期", "源文件", "工作表名", "规则名称", "行头路径", "列头路径", "数值", "旧备注"]
        )
        target_sheet.append(
            [
                datetime(2026, 4, 30),
                source_a.name,
                "本外币",
                "活期提取",
                "活期存款_1",
                "本月余额",
                104,
                "duplicate from previous run",
            ]
        )
        target_sheet.append(
            [
                datetime(2025, 12, 31),
                "history.xlsx",
                "历史工作表",
                "旧规则",
                "历史路径",
                "历史列",
                7,
                "retained history",
            ]
        )
        target_workbook.save(target_path)
        target_seed = target_path.read_bytes()
        target_before = workbook_projection(target_path)
        target_config = root / "target-config" / "config.xlsx"
        make_config(target_config, source_dir, root / "target-output", logs, target_path)
        deferred_run = subprocess.run(
            [
                str(rust_cli),
                "timeline-extract",
                "--config",
                str(target_config),
                "--log-view",
                "off",
                str(source_a),
                str(source_b),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if deferred_run.returncode != 0:
            raise AssertionError(f"target-contract check failed: {deferred_run.stderr}")
        deferred_result = json.loads(deferred_run.stdout)
        if not deferred_result["target_writes_deferred"]:
            raise AssertionError("configured target write was not surfaced as deferred")
        if workbook_projection(target_path) != target_before:
            raise AssertionError("deferred target write modified its target workbook")

        plan_file = Path(deferred_result["target_append_results"][0]["plan_file"])
        test_capability = root / "pt12-test-capability.json"
        test_platform = write_test_capability(plan_file, test_capability)
        python_target = root / "python-existing-target.xlsx"
        python_target.write_bytes(target_seed)
        python_target_config = root / "python-target-config" / "config.xlsx"
        python_target_output = root / "python-target-output"
        make_config(
            python_target_config,
            source_dir,
            python_target_output,
            logs,
            python_target,
        )
        python_global = load_global(python_target_config)
        python_rules = load_timeline_rules(python_target_config)
        python_maps = load_path_maps(python_target_config)
        python_target_summary = run_timeline_slim(
            python_rules,
            python_maps,
            python_global,
            source_paths=[source_a, source_b],
        )
        if python_target_summary is None or not python_target.exists():
            raise AssertionError("PyCode did not append to the existing target workbook")

        committed_run = subprocess.run(
            [
                str(rust_cli),
                "menu",
                "--config",
                str(target_config),
                "--capabilities",
                str(test_capability),
                "--platform",
                test_platform,
                "--log-view",
                "developer",
            ],
            input=f"1\n2\n{source_a};{source_b}\ny\n0\n",
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if committed_run.returncode != 0:
            raise AssertionError(
                f"Rust target append failed: {committed_run.stderr}\n{committed_run.stdout}"
            )
        result_field = committed_run.stdout.find('"target_append_results"')
        marker = committed_run.stdout.rfind("{", 0, result_field)
        if result_field < 0 or marker < 0:
            raise AssertionError(
                "PyTools menu did not report the PT-1-2 target result:\n"
                + committed_run.stdout
            )
        committed_result, _ = json.JSONDecoder().raw_decode(
            committed_run.stdout[marker:]
        )
        append_results = committed_result["target_append_results"]
        if len(append_results) != 1 or append_results[0]["status"] != "succeeded":
            raise AssertionError(f"unexpected PT-1-2 target execution: {append_results}")
        if append_results[0]["engine"] != "umya_spreadsheet":
            raise AssertionError(f"unexpected existing-target engine: {append_results[0]}")
        if append_results[0]["appended_rows"] != 19:
            raise AssertionError(f"unexpected existing-row de-duplication: {append_results[0]}")
        if append_results[0]["duplicate_rows_filtered"] != 1:
            raise AssertionError(f"existing duplicate was not filtered: {append_results[0]}")
        rust_target_projection = workbook_projection(target_path)
        python_target_projection = workbook_projection(python_target)
        if rust_target_projection != python_target_projection:
            raise AssertionError(
                "Rust existing-target output differs from PyCode: "
                + json.dumps(
                    {"python": python_target_projection, "rust": rust_target_projection},
                    ensure_ascii=False,
                    default=str,
                )
            )
        if target_path.read_bytes() == target_seed:
            raise AssertionError("confirmed Rust target append did not update the target")
        event_lines = [
            line for line in committed_run.stderr.splitlines() if line.startswith("{")
        ]
        developer_events = [json.loads(line) for line in event_lines]
        history_path = Path(committed_result["run_log_path"])
        target_history = json.loads(history_path.read_text(encoding="utf-8"))
        if target_history["events"] != developer_events:
            raise AssertionError("menu target event stream differs from persisted history")
        target_events = [
            event
            for event in developer_events
            if event.get("phase") == "target_append_execute"
        ]
        if (
            len(target_events) != 1
            or target_events[0].get("actual_engine") != "umya_spreadsheet"
            or not isinstance(target_events[0].get("stage_elapsed_ms"), int)
        ):
            raise AssertionError(f"unexpected target execution event: {target_events}")

        mismatch_target = root / "mismatched-target.xlsx"
        mismatch_workbook = Workbook()
        mismatch_sheet = mismatch_workbook.active
        mismatch_sheet.title = "追加目标"
        mismatch_sheet.append(
            ["源文件", "工作表名", "规则名称", "数据日期", "行头路径", "错误列头"]
        )
        mismatch_sheet.append(
            ["must remain", "历史工作表", "旧规则", datetime(2025, 12, 31), "旧路径", "old"]
        )
        mismatch_workbook.save(mismatch_target)
        mismatch_target_seed = mismatch_target.read_bytes()
        mismatch_config = root / "mismatch-config" / "config.xlsx"
        make_config(
            mismatch_config,
            source_dir,
            root / "mismatch-output",
            logs,
            mismatch_target,
        )
        mismatch_python_global = load_global(mismatch_config)
        mismatch_python_global.output_dir = root / "mismatch-python-output"
        mismatch_python_path = run_timeline_slim(
            load_timeline_rules(mismatch_config),
            load_path_maps(mismatch_config),
            mismatch_python_global,
            source_paths=[source_a],
        )
        if mismatch_python_path is None:
            raise AssertionError("PyCode did not produce output for header-mismatch case")
        if mismatch_target.read_bytes() != mismatch_target_seed:
            raise AssertionError("PyCode changed a target with a missing required header")

        mismatch_run = subprocess.run(
            [
                str(rust_cli),
                "timeline-extract",
                "--config",
                str(mismatch_config),
                "--output-dir",
                str(root / "mismatch-output"),
                "--log-view",
                "off",
                str(source_a),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if mismatch_run.returncode != 0:
            raise AssertionError(
                f"Rust header-mismatch target run failed: {mismatch_run.stderr}"
            )
        mismatch_result = json.loads(mismatch_run.stdout)
        mismatch_targets = mismatch_result["target_append_results"]
        if (
            len(mismatch_targets) != 1
            or mismatch_targets[0]["status"] != "skipped"
            or not any("header_mismatch" in issue for issue in mismatch_targets[0]["issues"])
        ):
            raise AssertionError(f"unexpected header-mismatch result: {mismatch_targets}")
        if workbook_projection(Path(mismatch_result["path"])) != workbook_projection(
            mismatch_python_path
        ):
            raise AssertionError("Rust summary differs from PyCode on header-mismatch case")
        if mismatch_target.read_bytes() != mismatch_target_seed:
            raise AssertionError("Rust changed a target with a missing required header")

        fail_fast_output = root / "fail-fast-output"
        fail_fast_logs = root / "fail-fast-logs"
        fail_fast_config = root / "fail-fast-config" / "config.xlsx"
        make_config(
            fail_fast_config,
            source_dir,
            fail_fast_output,
            fail_fast_logs,
            error_policy="fail_fast",
        )
        fail_fast_global = load_global(fail_fast_config)
        fail_fast_rules = load_timeline_rules(fail_fast_config)
        fail_fast_maps = load_path_maps(fail_fast_config)
        try:
            run_timeline_slim(
                fail_fast_rules,
                fail_fast_maps,
                fail_fast_global,
                source_paths=[broken_source, source_a],
            )
        except Exception:
            pass
        else:
            raise AssertionError("PyCode fail_fast baseline did not stop on the first error")
        python_fail_fast_outputs = list(fail_fast_output.glob("*.xlsx"))
        if python_fail_fast_outputs:
            raise AssertionError("PyCode fail_fast unexpectedly wrote an output workbook")

        fail_fast_run = subprocess.run(
            [
                str(rust_cli),
                "timeline-extract",
                "--config",
                str(fail_fast_config),
                "--log-view",
                "developer",
                str(broken_source),
                str(source_a),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if fail_fast_run.returncode == 0:
            raise AssertionError("Rust fail_fast did not return a failure status")
        fail_fast_history_paths = list(fail_fast_logs.glob("PT-1-2_*.run.json"))
        if len(fail_fast_history_paths) != 1:
            raise AssertionError("Rust fail_fast did not persist exactly one run history")
        fail_fast_history = json.loads(
            fail_fast_history_paths[0].read_text(encoding="utf-8")
        )
        if fail_fast_history["status"] != "failed":
            raise AssertionError("Rust fail_fast history did not record failure")
        if fail_fast_history["progress"]["processed_inputs"] != 1:
            raise AssertionError("Rust fail_fast processed inputs beyond the first failure")
        if list(fail_fast_output.glob("*.xlsx")):
            raise AssertionError("Rust fail_fast unexpectedly wrote an output workbook")

        return {
            "schema_version": 1,
            "suite": "pytools.timeline.extract_python_rust_parity",
            "status": "passed",
            "feature_id": "PT-1-2",
            "platform": "windows-x64-native",
            "inputs": 3,
            "valid_inputs": 2,
            "matched_rule_sheets": rust_result["matched_rule_sheets"],
            "output_rows": rust_result["rows"],
            "full_workbook_projection_equal": True,
            "checked": [
                "sheet name, cell values and types",
                "numeric, text, boolean, and datetime extraction values",
                "header bold and cell number formats",
                "worksheet dimensions and freeze panes",
                "duplicate-rule global deduplication",
                "row suffix and path mapping",
                "developer event stream and persisted history",
                "omitted source paths recursively scan configured input directory and skip Excel lock files; full output matches PyCode",
                "user view reports progress, hides engine names, and matches PyCode output",
                "menu 1→2 manual multi-file selection and continue-on-error behavior",
                "absent target is created through rust_xlsxwriter and matches the PyCode workbook",
                "existing target append is plan-only by default and leaves the target untouched",
                "confirmed menu target commit preserves unrelated data and matches PyCode",
                "header-mismatched existing target is skipped without changing the target; summary still matches PyCode",
                "unmodified worksheet formula, comment, fill, merge, dimensions and page settings are retained",
                "menu target execution event matches persisted developer history",
                "fail_fast stops at the first failed workbook and does not create output",
            ],
            "rust_engines": history["actual_engines"],
            "user_view": {
                "full_workbook_projection_equal": True,
                "progress_reported": "时序提取完成" in user_run.stderr,
                "engine_names_hidden": True,
                "actual_engines_in_history": user_history["actual_engines"],
            },
            "default_source_discovery": {
                "full_workbook_projection_equal": True,
                "total_files": discovery_result["total_files"],
                "failed_files": discovery_result["failed_files"],
                "recursive_source_included": True,
                "excel_lock_file_skipped": True,
                "actual_engines_in_history": discovery_history["actual_engines"],
            },
            "elapsed_ms": history["total_elapsed_ms"],
            "deferred_target_write_rules": rust_result["configured_target_rules"],
            "differences": [],
            "notes": [
                "Synthetic workbooks only; no source-project workbooks were modified.",
                "Existing-target edit uses a temporary synthetic test capability; no production capability was registered.",
            ],
        }


if __name__ == "__main__":
    report = run()
    output = WORKSPACE / "pytools-core" / "probe" / "reports" / "timeline_extract_parity_20261025.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.is_file():
        previous_report = json.loads(output.read_text(encoding="utf-8"))
        manual_dialog_case = previous_report.get("manual_native_dialog_case")
        if manual_dialog_case is not None:
            report["manual_native_dialog_case"] = manual_dialog_case
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
