"""Realistic synthetic parity suite for PT-1-3/1-4 wide summary and PT-6-2 comment summary.

Scenario 1 exercises a config-driven multi-rule wide summary (two rules with different
sheet keywords, bounds, skip keywords, required headers and a set region) over two
source workbooks, with the Python baseline loading the SAME config workbook the Rust
CLI reads. Scenario 2 exercises a multi-sheet comment-summary template with two row
regions and two set columns per sheet, over two source workbooks with unmatched sheets.
All inputs are synthetic; this does not constitute real-sample acceptance.
"""

from __future__ import annotations

import datetime as dt
import json
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True
PROBE = Path(__file__).resolve().parent
WORKSPACE = PROBE.parents[1]
sys.path.insert(0, str(PROBE))
sys.path.insert(0, str(WORKSPACE))

import compare_summary_tools as cst  # noqa: E402
import compare_wide_summary as cws  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.comments import Comment  # noqa: E402
from openpyxl.styles import Font  # noqa: E402

PYTOOLS_SOURCE = cst.PYTOOLS_SOURCE
INSTITUTIONS = ["甲分行", "乙分行", "丙分行", "丁分行", "戊分行"]


def make_report_workbook(path: Path, month: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = f"{month}Report"
    sheet["A1"] = f"{month} 报表"
    sheet.merge_cells("A2:B2")
    sheet["A2"] = "集团A"
    sheet["A3"] = "地区"
    sheet["B3"] = "存款"
    sheet.merge_cells("B3:C3")
    sheet["D3"] = "贷款"
    sheet["E3"] = "票据"
    sheet["F3"] = "理财"
    sheet["A4"] = "名称"
    sheet["B4"] = "余额"
    sheet["C4"] = "余额"
    sheet["D4"] = "本息"
    sheet["E4"] = "余额"
    sheet["F4"] = "余额"
    for cell in ("A3", "B3", "D3", "E3", "F3", "A4", "B4", "C4", "D4", "E4", "F4"):
        sheet[cell].font = Font(bold=True)
    row = 5
    for index in range(24):
        name = INSTITUTIONS[index % len(INSTITUTIONS)]
        if index % 7 == 3:
            name = f"{name}小计"
        if index % 11 == 5:
            name = "合计"
        sheet.cell(row=row, column=1, value=name)
        sheet.cell(row=row, column=2, value=100.5 + index if index % 5 else None)
        sheet.cell(row=row, column=3, value=index + 1)
        sheet.cell(row=row, column=4, value=200.25 + index * 2 if index % 3 else "待定")
        sheet.cell(row=row, column=5, value=None if index % 4 else 50 + index)
        sheet.cell(row=row, column=6, value=dt.date(2026, 6, 30) if index == 0 else f"note{index}")
        row += 1

    balance = workbook.create_sheet("余额表")
    for column, header in enumerate(("机构", "币种", "科目", "余额", "备注"), start=1):
        balance.cell(row=1, column=column, value=header)
        balance.cell(row=1, column=column).font = Font(bold=True)
    for index in range(18):
        balance.cell(row=2 + index, column=1, value=INSTITUTIONS[index % len(INSTITUTIONS)])
        balance.cell(row=2 + index, column=2, value="CNY" if index % 2 else "USD")
        balance.cell(row=2 + index, column=3, value=f"科目{index % 6 + 1}")
        balance.cell(row=2 + index, column=4, value=1000.75 + index if index % 6 else None)
        balance.cell(row=2 + index, column=5, value=f"备注{index}" if index % 5 == 0 else None)

    workbook.create_sheet("说明")
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def make_wide_config(path: Path, output: Path) -> None:
    workbook = Workbook()
    global_sheet = workbook.active
    global_sheet.title = "全局配置"
    global_sheet.append(["配置项", "值"])
    for key, value in (
        ("输入目录", cws.rust_path_argument(output / "unused_input")),
        ("输出目录", cws.rust_path_argument(output)),
        ("日志目录", cws.rust_path_argument(output / "logs")),
        ("错误策略", "fail_fast"),
        ("源文件扩展名", ".xlsx"),
        ("时序规则匹配模式", "all_match"),
    ):
        global_sheet.append([key, value])

    timeline = workbook.create_sheet("时序提取规则")
    timeline.append([
        "是否启用", "规则名称", "工作簿关键字", "工作表关键字", "行头列", "列表头行",
        "必含列头", "必含行头", "数据起始行", "数据结束行", "数据起始列", "数据结束列",
        "跳过关键字", "目标工作簿路径", "目标工作表", "启用目标写入", "set区域", "目标去重列",
        "是否输出行序号",
    ])
    timeline.append([
        "是", "WideReport", "sample", "Report", "A", "3,4", "", "", 5, "", "B", "F",
        "小计;合计", "", "", "否", "集团@B2", "", "是",
    ])
    timeline.append([
        "是", "WideBalance", "sample", "余额", "A", "1", "科目", "", 2, "", "B", "D",
        "", "", "", "否", "", "", "否",
    ])
    maps = workbook.create_sheet("路径标准化映射")
    maps.append([
        "是否启用", "映射名称", "适用规则名", "工作簿关键字", "工作表关键字", "作用对象",
        "匹配方式", "原始路径", "标准路径", "备注",
    ])
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def python_wide_baseline(config: Path, sources: list[Path], row_suffix: bool) -> Path | None:
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.config_xlsx import load_global, load_path_maps, load_timeline_rules
    from core.wide_397_398 import run_wide_summary

    global_config = load_global(config)
    rules = load_timeline_rules(config)
    maps = load_path_maps(config)
    try:
        return run_wide_summary(
            rules, maps, global_config, row_suffix, source_paths=sources
        )
    finally:
        import logging
        logging.shutdown()


def _sheet_dimensions(path: Path) -> list[dict[str, object]]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, data_only=False)
    try:
        return [
            {"name": sheet.title, "rows": sheet.max_row, "columns": sheet.max_column}
            for sheet in workbook.worksheets
        ]
    finally:
        workbook.close()


def run_wide_case(root: Path, row_suffix: bool, workbooks: list[Path]) -> dict[str, object]:
    case = root / ("wide_suffix" if row_suffix else "wide_no_suffix")
    output = case / "output"
    config = case / "config.xlsx"
    make_wide_config(config, output)
    python_output = python_wide_baseline(config, workbooks, row_suffix)
    rust_output, plans = cws.run_rust_sources(config, workbooks, row_suffix)
    if python_output is None or rust_output is None:
        raise AssertionError(
            f"wide summary did not save: python={python_output}, rust={rust_output}"
        )
    if plans:
        raise AssertionError(f"target writes must be disabled in this suite, got plans: {plans}")
    cws.assert_parity(python_output, rust_output)
    projection = cws.normalized(rust_output)
    return {
        "row_suffix": row_suffix,
        "python_rust_summary": "exact",
        "sheet_names": projection["sheet_names"],
        "dimensions": projection["dimensions"],
        "python_summary_dimensions": _sheet_dimensions(python_output),
    }


def make_comment_template(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "余额表"
    sheet["A1"] = "某银行 余额表模板"
    sheet["B1"] = "机构"
    sheet["B1"].comment = Comment("set(机构)", "probe")
    sheet["C1"] = "币种"
    sheet["C1"].comment = Comment("set(币种)", "probe")
    sheet["D1"] = "存款"
    sheet["D1"].comment = Comment("列区域1", "probe")
    sheet["E1"] = "中间业务"
    sheet["F1"] = "贷款"
    sheet["F1"].comment = Comment("列区域#1", "probe")
    sheet["A3"] = "机构名称"
    sheet["A3"].comment = Comment("行区域1", "probe")
    sheet["A26"].comment = Comment("行区域#1", "probe")
    sheet["A28"] = "机构名称"
    sheet["A28"].comment = Comment("行区域2", "probe")
    sheet["A45"].comment = Comment("行区域#2", "probe")

    flow = workbook.create_sheet("流水表")
    flow["A1"] = "机构名称"
    flow["A1"].comment = Comment("行区域1", "probe")
    flow["A61"].comment = Comment("行区域#1", "probe")
    flow["B1"] = "机构"
    flow["B1"].comment = Comment("set(机构)", "probe")
    flow["C1"] = "摘要"
    flow["C1"].comment = Comment("列区域1", "probe")
    flow["F1"] = "金额"
    flow["F1"].comment = Comment("列区域#1", "probe")

    workbook.create_sheet("说明")
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def make_comment_source(path: Path, tag: str) -> None:
    workbook = Workbook()
    balance = workbook.active
    balance.title = "余额表"
    for index in range(24):
        row = 4 + index if index < 22 else 29 + (index - 22)
        balance.cell(row=row, column=1, value=INSTITUTIONS[index % len(INSTITUTIONS)] + tag)
        balance.cell(row=row, column=2, value=f"机构{index}")
        balance.cell(row=row, column=3, value="CNY" if index % 2 else "USD")
        balance.cell(row=row, column=4, value=100.5 + index if index % 5 else None)
        balance.cell(row=row, column=5, value=f"middle{index}" if index % 3 == 0 else None)
        balance.cell(row=row, column=6, value=300 + index * 1.5)

    flow = workbook.create_sheet("流水表")
    for index in range(58):
        flow.cell(row=3 + index, column=1, value=f"流水{tag}{index}")
        flow.cell(row=3 + index, column=2, value=f"机构{index % 9}")
        flow.cell(row=3 + index, column=3, value=f"摘要{index}")
        flow.cell(row=3 + index, column=4, value=dt.date(2026, 7, 1 + index % 28))
        flow.cell(row=3 + index, column=5, value=None if index % 7 == 0 else f"remark{index}")
        flow.cell(row=3 + index, column=6, value=round(index * 3.33, 2) if index % 4 else "冲正")

    workbook.create_sheet("说明")
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def run_rust_comments_multi(template: Path, sources: list[Path], output: Path):
    import os
    import subprocess

    native_cli = os.environ.get("PYTOOLS_RUST_CLI")
    if not native_cli:
        raise RuntimeError("this suite requires PYTOOLS_RUST_CLI to point at a native pytools.exe")
    invocation = [
        native_cli, "summary-comments",
        "--template", str(template),
        "--output-dir", str(output),
    ] + [str(source) for source in sources]
    completed = subprocess.run(invocation, check=True, capture_output=True)
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for index, character in enumerate(stdout):
        if character != "{":
            continue
        try:
            result = json.loads(stdout[index:])
        except json.JSONDecodeError:
            continue
        path_value = result.get("path")
        return (Path(path_value) if path_value else None), result
    raise RuntimeError(f"Rust comment CLI emitted no JSON: {stdout}\n{stderr}")


def run_comment_case(root: Path, sources: list[Path]) -> dict[str, object]:
    case = root / "comments"
    template = case / "template.xlsx"
    make_comment_template(template)
    python_output = cst.python_comments_only(template, sources, case / "python_output")
    rust_output, rust_report = run_rust_comments_multi(
        template, sources, case / "rust_output"
    )
    if python_output is None or rust_output is None:
        raise AssertionError(
            f"comment summary did not save: python={python_output}, rust={rust_output}"
        )
    cst.assert_equivalent(python_output, rust_output, "PT-6-2 realistic multi-sheet template")
    projection = cws.normalized(rust_output)
    return {
        "python_rust_output": "exact",
        "sheet_names": projection["sheet_names"],
        "dimensions": projection["dimensions"],
        "rust_rows": rust_report.get("rows"),
        "unmatched_source_sheet": "说明 skipped by both sides",
    }


def make_mode_config(path: Path, output: Path, mode: str) -> None:
    workbook = Workbook()
    global_sheet = workbook.active
    global_sheet.title = "全局配置"
    global_sheet.append(["配置项", "值"])
    for key, value in (
        ("输入目录", cws.rust_path_argument(output / "unused_input")),
        ("输出目录", cws.rust_path_argument(output)),
        ("日志目录", cws.rust_path_argument(output / "logs")),
        ("错误策略", "fail_fast"),
        ("源文件扩展名", ".xlsx"),
        ("时序规则命中策略", mode),
    ):
        global_sheet.append([key, value])

    timeline = workbook.create_sheet("时序提取规则")
    timeline.append([
        "是否启用", "规则名称", "工作簿关键字", "工作表关键字", "行头列", "列表头行",
        "必含列头", "必含行头", "数据起始行", "数据结束行", "数据起始列", "数据结束列",
        "跳过关键字", "目标工作簿路径", "目标工作表", "启用目标写入", "set区域", "目标去重列",
        "是否输出行序号",
    ])
    timeline.append([
        "是", "WideAll", "sample", "Report", "A", "3,4", "", "", 5, "", "B", "F",
        "小计;合计", "", "", "否", "集团@B2", "", "是",
    ])
    timeline.append([
        "是", "WideRep", "sample", "Rep", "A", "3", "", "甲分行;乙分行", 4, 20, "B", "D",
        "", "", "", "否", "", "", "是",
    ])
    maps = workbook.create_sheet("路径标准化映射")
    maps.append([
        "是否启用", "映射名称", "适用规则名", "工作簿关键字", "工作表关键字", "作用对象",
        "匹配方式", "原始路径", "标准路径", "备注",
    ])
    maps.append(["是", "标准化甲", "WideAll", "sample", "Report", "行头", "包含", "甲分行", "甲标准化分行", ""])
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def run_rule_mode_case(root: Path, mode: str, workbooks: list[Path]) -> dict[str, object]:
    case = root / f"rule_mode_{mode}"
    output = case / "output"
    config = case / "config.xlsx"
    make_mode_config(config, output, mode)
    python_output = python_wide_baseline(config, workbooks, row_suffix=True)
    rust_output, plans = cws.run_rust_sources(config, workbooks, row_suffix=True)
    if python_output is None or rust_output is None:
        raise AssertionError(
            f"wide summary ({mode}) did not save: python={python_output}, rust={rust_output}"
        )
    if plans:
        raise AssertionError(f"target writes must be disabled in this suite, got plans: {plans}")
    cws.assert_parity(python_output, rust_output)
    projection = cws.normalized(rust_output)
    return {
        "mode": mode,
        "python_rust_summary": "exact",
        "sheet_names": projection["sheet_names"],
        "dimensions": projection["dimensions"],
        "python_summary_dimensions": _sheet_dimensions(python_output),
    }


def make_set_conflict_workbook(path: Path, month: str, empty_set: bool) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = f"{month}Report"
    sheet["B1"] = "" if empty_set else f"{month}期"
    sheet["B2"] = "" if empty_set else "集团A"
    sheet["A3"] = "地区"
    sheet["B3"] = "存款"
    sheet["C3"] = "存款"
    sheet["D3"] = "贷款"
    sheet["A4"] = "名称"
    sheet["B4"] = "余额"
    sheet["C4"] = "余额"
    sheet["D4"] = "本息"
    # Duplicate col paths 存款_余额 (B/C) with different values: the first
    # occurrence must win on both sides; D keeps a unique path.
    sheet["A5"] = "甲分行"
    sheet["B5"] = 10.25
    sheet["C5"] = 999
    sheet["D5"] = 23
    sheet["A6"] = "乙分行"
    sheet["B6"] = 14
    sheet["C6"] = 888
    sheet["D6"] = 29
    sheet["A7"] = "甲分行"
    sheet["B7"] = 31
    sheet["C7"] = 777
    sheet["D7"] = 35
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def make_set_conflict_config(path: Path, output: Path) -> None:
    workbook = Workbook()
    global_sheet = workbook.active
    global_sheet.title = "全局配置"
    global_sheet.append(["配置项", "值"])
    for key, value in (
        ("输入目录", cws.rust_path_argument(output / "unused_input")),
        ("输出目录", cws.rust_path_argument(output)),
        ("日志目录", cws.rust_path_argument(output / "logs")),
        ("错误策略", "fail_fast"),
        ("源文件扩展名", ".xlsx"),
        ("时序规则命中策略", "all_match"),
    ):
        global_sheet.append([key, value])
    timeline = workbook.create_sheet("时序提取规则")
    timeline.append([
        "是否启用", "规则名称", "工作簿关键字", "工作表关键字", "行头列", "列表头行",
        "必含列头", "必含行头", "数据起始行", "数据结束行", "数据起始列", "数据结束列",
        "跳过关键字", "目标工作簿路径", "目标工作表", "启用目标写入", "set区域", "目标去重列",
        "是否输出行序号",
    ])
    timeline.append([
        "是", "WideSet", "sample", "Report", "A", "3,4", "", "", 5, "", "B", "D",
        "", "", "", "否", "集团@B2;期间@B1", "", "是",
    ])
    maps = workbook.create_sheet("路径标准化映射")
    maps.append([
        "是否启用", "映射名称", "适用规则名", "工作簿关键字", "工作表关键字", "作用对象",
        "匹配方式", "原始路径", "标准路径", "备注",
    ])
    # Both numbered 存款_余额 columns map to the same standard path, which
    # happens after header numbering in PyCode and must trigger the
    # duplicate-column-path conflict rule (keep first value, single header).
    maps.append(["是", "合并存款1", "WideSet", "sample", "Report", "列头", "精确", "存款_余额_1", "公共存款", ""])
    maps.append(["是", "合并存款2", "WideSet", "sample", "Report", "列头", "精确", "存款_余额_2", "公共存款", ""])
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def run_set_conflict_case(root: Path, workbooks: list[Path]) -> dict[str, object]:
    case = root / "wide_sets_conflict"
    output = case / "output"
    config = case / "config.xlsx"
    make_set_conflict_config(config, output)
    python_output = python_wide_baseline(config, workbooks, row_suffix=False)
    rust_output, plans = cws.run_rust_sources(config, workbooks, row_suffix=False)
    if python_output is None or rust_output is None:
        raise AssertionError(
            f"set/conflict wide summary did not save: python={python_output}, rust={rust_output}"
        )
    if plans:
        raise AssertionError(f"target writes must be disabled in this suite, got plans: {plans}")
    cws.assert_parity(python_output, rust_output)
    from openpyxl import load_workbook

    workbook = load_workbook(rust_output, data_only=False)
    try:
        sheet = workbook.worksheets[0]
        headers = [sheet.cell(1, column).value for column in range(1, sheet.max_column + 1)]
        sample_rows = [
            [cst.normalize(sheet.cell(row, column).value) for column in range(1, sheet.max_column + 1)]
            for row in range(2, min(sheet.max_row, 8) + 1)
        ]
    finally:
        workbook.close()
    return {
        "python_rust_summary": "exact",
        "sheet_names": [item["name"] for item in _sheet_dimensions(rust_output)],
        "dimensions": [
            [item["rows"], item["columns"]] for item in _sheet_dimensions(rust_output)
        ],
        "headers": headers,
        "sample_rows": sample_rows,
        "conflict_policy": "first value kept for duplicate column paths (PyCode 保留首值)",
    }


def main() -> int:
    with tempfile.TemporaryDirectory(
        prefix="pytools-realistic-suite-", ignore_cleanup_errors=True
    ) as temp:
        root = Path(temp)
        workbooks = []
        for month in ("2026-06", "2026-07"):
            workbook = root / "sources" / f"sample_{month}-30.xlsx"
            if month == "2026-07":
                workbook = root / "sources" / "sample_2026-07-31.xlsx"
            make_report_workbook(workbook, month)
            workbooks.append(workbook)

        wide_results = [
            run_wide_case(root, row_suffix, workbooks) for row_suffix in (False, True)
        ]

        rule_mode_results = [
            run_rule_mode_case(root, mode, workbooks) for mode in ("all_match", "first_match")
        ]

        set_conflict_workbooks = []
        for month, empty in (("2026-06", False), ("2026-07", True)):
            workbook = root / "set_sources" / f"sample_{month}-30.xlsx"
            make_set_conflict_workbook(workbook, month, empty)
            set_conflict_workbooks.append(workbook)
        set_conflict_result = run_set_conflict_case(root, set_conflict_workbooks)

        comment_sources = []
        for tag in ("一", "二"):
            source = root / "comment_sources" / f"report_{tag}.xlsx"
            make_comment_source(source, tag)
            comment_sources.append(source)
        comment_result = run_comment_case(root, comment_sources)

        report = {
            "schema_version": 1,
            "suite": "pytools.realistic_synthetic_parity_suite",
            "date": dt.date.today().isoformat(),
            "status": "passed_synthetic_realistic_parity",
            "scope_note": (
                "Synthetic but business-shaped inputs created per user instruction; "
                "config-driven both sides for wide summary. Not real-sample acceptance."
            ),
            "wide_summary_multi_rule": wide_results,
            "wide_summary_rule_match_modes": rule_mode_results,
            "wide_summary_sets_and_conflicts": set_conflict_result,
            "comment_summary_multi_sheet": comment_result,
            "rust_runtime": {
                "host": "Windows 11 x64, native MSVC target",
                "entrypoint": "pytools wide-summary[-row-suffix] / pytools summary-comments",
            },
            "limitations": [
                "Synthetic inputs only; real business samples for wide summary and comment templates remain unavailable.",
                "No target append in this suite; existing-target semantics are covered by compare_wide_summary.py cases.",
                "Report records dimensions and counts only; outputs were temporary.",
            ],
        }
        report_path = PROBE / "reports" / "realistic_synthetic_suite_parity_20261025.json"
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print("SUITE_RESULT=" + json.dumps({
            "status": report["status"],
            "wide_cases": [result["dimensions"] for result in wide_results],
            "rule_modes": [result["dimensions"] for result in rule_mode_results],
            "comment_sheets": comment_result["dimensions"],
            "report": str(report_path),
        }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
