"""Capture real legacy PyTools and Rust results using private synthetic BIFF fixtures."""

import argparse
from contextlib import redirect_stdout
from copy import deepcopy
from datetime import datetime
import importlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import traceback

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils.datetime import CALENDAR_MAC_1904, from_excel
import win32com.client

from run_special_reports_xls_baseline import ROOT, compare, create_biff, sha, write_json


CASES = ("1-2", "1-2-styled", "1-2-epoch", "1-4", "4-3")


def make_source(path, case):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Data"
    if case.startswith("1-2"):
        sheet["B3"] = "00123"
        sheet["C3"] = True
        sheet["B5"] = 0
        sheet["D5"] = "=SUM(2,3)"
        sheet.merge_cells("B6:C7")
        sheet["B6"] = "合并"
        if case == "1-2-styled":
            sheet["F8"].font = Font(bold=True)
            book.create_sheet("空表")
        if case == "1-2-epoch":
            book.epoch = CALENDAR_MAC_1904
            sheet["E5"] = datetime(2025, 4, 8, 14, 30)
    elif case == "1-4":
        sheet.title = "Branch_A"
        sheet["B1"] = "余额"
        sheet["C1"] = "余额"
        sheet["D1"] = "华东"
        sheet.merge_cells("B2:C3")
        sheet["B2"] = "合并锚点"
        sheet["B4"] = 0
        sheet["C4"] = "00123"
        other = book.create_sheet("Branch_B")
        other["D1"] = "华南"
        other["B2"] = 7
        other["C2"] = True
    else:
        sheet["A1"] = "甲机构"
    book.save(path)
    book.close()


def make_config(path):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "重命名配置"
    for column, (header, value) in enumerate(
        (("内容重命名启用", "Y"), ("内容重命名取值规则", "Data@A1"), ("备注", "合成用例")), 15
    ):
        sheet.cell(1, column, header)
        sheet.cell(2, column, value)
    book.save(path)
    book.close()


def write_cached_fixture(app, path, to_xls):
    if to_xls:
        return create_biff(app, path)
    # Round-trip paired fixture through the same host, including formula caches.
    book = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
    try:
        book.Save()
    finally:
        book.Close(SaveChanges=False)
    return path


def snapshot(path):
    book = openpyxl.load_workbook(path)
    try:
        return {s.title: [
            [{"value": c.value.isoformat() if isinstance(c.value, datetime) else c.value,
              "type": c.data_type} if c.value is not None else None for c in row]
            for row in s.rows
        ] for s in book}
    finally:
        book.close()


def run_case(root, route, case, exe, com, legacy_comment_shadow, replay):
    directory = root / case / route
    directory.mkdir(parents=True)
    (directory / "logs").mkdir()
    (directory / "output").mkdir()
    if replay:
        for path, expected in replay["input_sha256"].items():
            saved = root_path_for_baseline(path) / "before" / Path(path).name
            assert sha(saved) == expected, saved
        path = next(iter(replay["input_sha256"]))
        src = root_path_for_baseline(path) / "before" / Path(path).name
    else:
        authored = directory / "sample.xlsx"
        make_source(authored, case)
        app = com.start_excel_app()
        try:
            src = write_cached_fixture(app, authored, route.endswith("-xls"))
        finally:
            com.safe_quit(app)
    input_dir = directory / "input"
    input_dir.mkdir()
    src = Path(shutil.copy2(src, input_dir / src.name))
    template = None
    config = None
    if case == "1-4":
        template = directory / "template.xlsx"
        if replay:
            original_template = next(p for p in replay["input_sha256"] if Path(p).name == "template.xlsx")
            template_source = root_path_for_baseline(original_template) / "before" / template.name
        else:
            template_source = ROOT / "pytools-core/tests/fixtures/summary_comments/comment_template.xlsx"
        shutil.copy2(template_source, template)
    if case == "4-3":
        config = directory / "config.xlsx"
        if replay:
            original_config = next(p for p in replay["input_sha256"] if Path(p).name == "config.xlsx")
            shutil.copy2(root_path_for_baseline(original_config) / "before" / config.name, config)
        else:
            make_config(config)
        if replay is None:
            src = src.rename(src.with_name("甲机构_" + src.name))
    inputs = [path for path in (src, template, config) if path is not None]
    before = directory / "before"
    before.mkdir()
    input_hashes = {str(p): sha(p) for p in inputs}
    for p in inputs:
        shutil.copy2(p, before / p.name)
    commands = {"1-2": "summary-used-range", "1-4": "summary-comments", "4-3": "unrename-by-content"}
    row = dict(case=case, route=route, input_sha256=input_hashes)
    if route == "legacy-xls":
        legacy_summary = importlib.import_module("PyCode.core.summary_321_322_com")
        convert = legacy_summary._ensure_openpyxl_compatible
        if case.startswith("1-2"):
            args = ([src], directory / "output", directory / "logs")
            action = legacy_summary.run_summary_by_usedrange_com
        elif case == "1-4":
            args = (template, [src], directory / "output", directory / "logs")
            action = legacy_summary.run_summary_by_comment_com
            if legacy_comment_shadow:
                # Excel returns the already-open workbook for a second Open of the
                # same path. The legacy merge extractor then closes that object.
                # Isolate its conversion input without changing the legacy extraction.
                shadow = directory / "merge-read.xls"
                shutil.copy2(src, shadow)
                legacy_summary._ensure_openpyxl_compatible = (
                    lambda path, *rest: convert(shadow if path == src else path, *rest)
                )
                row["legacy_workaround"] = "separate BIFF copy for merge conversion"
        else:
            args = (config, [src], directory / "logs")
            action = importlib.import_module(
                "PyCode.core.convert_33_34_35_37"
            ).run_batch_unrename_by_content_47
        row["call"] = action.__module__ + "." + action.__name__
        try:
            with (directory / "stdout.txt").open("w", encoding="utf-8") as out, redirect_stdout(out):
                row["result"] = action(*args)
        except Exception:
            row["error"] = traceback.format_exc()
            (directory / "stderr.txt").write_text(row["error"], encoding="utf-8")
        finally:
            legacy_summary._ensure_openpyxl_compatible = convert
    else:
        command = [str(exe), commands["1-2" if case.startswith("1-2") else case],
                   "--log-view", "developer", "--log-dir", str(directory / "logs")]
        if case == "4-3":
            command.extend(["--config", str(config)])
        else:
            command.extend(["--output-dir", str(directory / "output")])
        if template:
            command.extend(["--template", str(template)])
        command.append(str(src))
        env = {k: v for k, v in os.environ.items() if not k.startswith(("PYTOOLS_", "RUSTTOOLS_ASPOSE_"))}
        env.update(PYTOOLS_CONVERT_ENGINES_DISABLED="1", PYTOOLS_COM_WORKER=str(directory / "missing.py"),
                   LOCALAPPDATA=str(directory / "state"))
        row["command"] = command
        proc = subprocess.run(command, cwd=directory, env=env, capture_output=True, text=True,
                              encoding="utf-8", errors="replace", timeout=120)
        row["exit"] = proc.returncode
        (directory / "stdout.txt").write_text(proc.stdout, encoding="utf-8")
        (directory / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
    if case == "4-3":
        names = sorted(input_dir.iterdir())
        assert len(names) == 1
        row["name"] = names[0].stem
        row["renamed"] = names[0] != src
        row["bytes_preserved"] = sha(names[0]) == input_hashes[str(src)]
    else:
        outputs = sorted((directory / "output").glob("*.xlsx"))
        row["outputs"] = [{"path": str(p), "sha256": sha(p)} for p in outputs]
        if len(outputs) == 1:
            row["snapshot"] = snapshot(outputs[0])
    row["readonly_inputs_unchanged"] = all(
        p.is_file() and sha(p) == input_hashes[str(p)]
        for p in inputs if not (case == "4-3" and p == src)
    )
    return row


def root_path_for_baseline(path):
    path = Path(path)
    return path.parent.parent if path.parent.name == "input" else path.parent


def compare_case(case, current, prior, key):
    expected = prior.get(key)
    actual = current.get(key)
    raw_delta = compare(expected, actual)
    expected = deepcopy(expected)
    if prior["route"] == "legacy-xls" and case == "1-2-styled":
        # Rust's worksheet contract reports empty sheets and trims unused
        # trailing columns, even when a serialized blank sets the row extent.
        rows = expected["汇总"]
        assert rows[0][-2:] == [{"value": "列5", "type": "s"}, {"value": "列6", "type": "s"}]
        assert all(row[-2:] == [None, None] for row in rows[1:])
        expected["汇总"] = [row[:-2] for row in rows] + [
            [{"value": "sample", "type": "s"}, {"value": "空表", "type": "s"},
             {"value": 1, "type": "n"}, None, None, None, None]
        ]
    if prior["route"] == "legacy-xls" and case == "1-2-epoch":
        # Legacy COM reports raw Value2 from the source's epoch. Current
        # summaries write the same calendar value as an Excel date.
        old_cell = expected["汇总"][3][7]
        assert old_cell["type"] == "n"
        value = from_excel(old_cell["value"], CALENDAR_MAC_1904)
        expected["汇总"][3][7] = {
            "value": value.isoformat(),
            "type": "d",
        }
    return raw_delta, compare(expected, actual)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--exe", type=Path, required=True)
    parser.add_argument("--legacy", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--routes", nargs="+", default=["legacy-xls", "rust-xlsx", "rust-xls"],
                        choices=["legacy-xls", "rust-xlsx", "rust-xls"])
    parser.add_argument("--host", default="Excel.Application", choices=["Excel.Application", "ket.Application"])
    parser.add_argument("--cases", nargs="+", choices=CASES, default=CASES)
    parser.add_argument("--legacy-comment-shadow", action="store_true")
    parser.add_argument("--baseline", type=Path, nargs="+")
    args = parser.parse_args()
    args.exe = args.exe.resolve(strict=True)
    args.output = args.output.resolve()
    args.output.mkdir()
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(args.legacy.resolve(strict=True)))
    com = importlib.import_module("PyCode.core.com_engine")
    dispatch = win32com.client.DispatchEx
    host_info = {}

    def pinned(prog_id, *values, **options):
        if prog_id.casefold() != args.host.casefold():
            raise RuntimeError("Alternate host excluded from this comparison")
        app = dispatch(prog_id, *values, **options)
        if app.Workbooks.Count:
            raise RuntimeError("COM session already contains workbooks")
        app.AutomationSecurity = 3
        host_info.update(prog_id=prog_id, name=app.Name, version=app.Version)
        return app

    win32com.client.DispatchEx = pinned
    rows = []
    baseline = [row for path in args.baseline or ()
                for row in json.loads(path.read_text(encoding="utf-8"))]
    replay = {(row["case"], row["route"]): row for row in baseline}
    try:
        for case in args.cases:
            for route in args.routes:
                rows.append(run_case(args.output, route, case, args.exe, com,
                                     args.legacy_comment_shadow, replay.get((case, route))))
                write_json(args.output / "runs.json", rows)
    finally:
        win32com.client.DispatchEx = dispatch
        write_json(args.output / "host.json", host_info)
    if args.baseline:
        comparisons = []
        for row in rows:
            for prior in baseline:
                if prior["case"] != row["case"] or prior["route"] == "rust-xls":
                    continue
                key = "name" if row["case"] == "4-3" else "snapshot"
                delta, business_delta = compare_case(row["case"], row, prior, key)
                comparisons.append(dict(case=row["case"], current=row["route"], reference=prior["route"],
                                        difference_paths=delta, business_difference_paths=business_delta))
        write_json(args.output / "comparisons.json", comparisons)
        assert all(not entry["business_difference_paths"] for entry in comparisons), comparisons
        for row in rows:
            if row["route"].startswith("rust-"):
                assert row["exit"] == 0 and row["readonly_inputs_unchanged"], row
                if row["case"] == "4-3":
                    assert row["renamed"] and row["bytes_preserved"], row
                else:
                    assert "snapshot" in row, row
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
