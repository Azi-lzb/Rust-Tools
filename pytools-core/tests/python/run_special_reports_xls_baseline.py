"""Record XLS COM and XLSX OUO results using separate synthetic files.

Run only on Windows with an approved RustTools executable and Excel/WPS COM.
The legacy functions are imported read-only and receive no business paths.
"""

import argparse
import hashlib
import importlib
import json
import os
from pathlib import Path
import runpy
import shutil
import subprocess
import sys
import traceback
from zipfile import ZipFile

import openpyxl
import win32com.client


ROOT = Path(__file__).resolve().parents[3]
LEGACY = {
    "5-1": ("feature_1_1_split_village_bank_com", "run_split_village_bank_com"),
    "5-2": ("feature_1_2_normalize_institution_com", "run_normalize_institution_com"),
    "5-3": ("feature_1_3_remove_foreign_bank_com", "run_remove_foreign_bank_com"),
    "5-4": ("feature_1_4_fx_header_com", "run_fx_header_fix_com"),
    "5-5": ("feature_1_5_region_sum_check_com", "run_region_sum_check_com"),
    "5-6": ("feature_1_6_extract_sheets_com", "run_extract_sheets_com"),
    "5-7": ("feature_1_8_adjust_rural_loan_com", "run_adjust_rural_loan_com"),
}
VARIANTS = {
    "5-1-validation-failure": "5-1",
    "5-2-validation-failure": "5-2",
    "5-3-multiple-runs": "5-3",
    "5-3-multiple-runs-values": "5-3",
    "5-3-zero": "5-3",
    "5-3-com-disabled": "5-3",
    "5-3-target-exists": "5-3",
    "5-3-save-failure": "5-3",
    "5-3-validation-failure": "5-3",
    "5-3-source-changed": "5-3",
    "5-4-validation-failure": "5-4",
    "5-4-source-changed": "5-4",
    "5-4-zero": "5-4",
    "5-4-pages": "5-4",
    "5-5-clean": "5-5",
    "5-6-full": "5-6",
    "5-6-mixed": "5-6",
    "5-6-new": "5-6",
    "5-6-collision": "5-6",
    "5-6-formula-cache": "5-6",
    "5-7-mixed": "5-7",
    "5-7-source-changed": "5-7",
    "5-7-validation-failure": "5-7",
    "5-7-unmatched": "5-7",
}


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_json(path, data):
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False, default=str), encoding="utf-8")


def create_biff(app, path):
    workbook = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
    try:
        target = path.with_suffix(".xls")
        workbook.SaveAs(str(target), FileFormat=56)
    finally:
        workbook.Close(SaveChanges=False)
    assert target.read_bytes().startswith(bytes.fromhex("d0cf11e0a1b11ae1")), target
    return target


def prepare(directory, scenario, helper):
    case = VARIANTS.get(scenario, scenario)
    if case != "5-5":
        _, _, products, _ = helper["prepare"](directory, "5-4" if case == "5-4" else case, False)
    else:
        directory.mkdir()
        helper["save_config"](helper["config"](directory), directory)
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "本外币分地区"
        sheet.append(["合成数据"])
        sheet.append(["序号", "地区", "一致列", "差异列"])
        sheet.append([1, "惠州市", 21, 21 if scenario == "5-5-clean" else 99])
        for i, region in enumerate(("惠阳区", "博罗县", "惠东县", "龙门县", "惠州市本部", "大亚湾"), 1):
            sheet.append([i + 1, region, i, i])
        book.save(directory / "source.xlsx")
        products = []
    if scenario.startswith("5-3-multiple-runs") or scenario == "5-3-zero":
        path = directory / "source.xlsx"
        book = openpyxl.load_workbook(path)
        sheet = book.active
        if scenario == "5-3-zero":
            sheet["A3"] = "普通丙"
        else:
            for row in range(5, 14):
                sheet.append(["外资甲" if row in (10, 11) else f"普通{row}", row])
            sheet["B13"] = "=SUM(7,6)" if scenario == "5-3-multiple-runs" else 13
            book.create_sheet("其他统计")["A1"] = "不处理"
        book.save(path)
        book.close()
    if case == "5-4":
        path = directory / "source.xlsx"
        with path.open("rb") as stream:
            book = openpyxl.load_workbook(stream)
        sheet = book.active
        sheet.HeaderFooter.differentOddEven = True
        sheet.HeaderFooter.differentFirst = True
        unit = "单位：万美元" if scenario == "5-4-zero" else "单位：万元"
        sheet.oddHeader.left.text = unit
        sheet.evenHeader.center.text = unit
        sheet.firstHeader.right.text = unit
        sheet.evenFooter.right.text = "保留偶数页页脚"
        if scenario == "5-4-pages":
            sheet.print_area = "A1:C4"
            sheet.page_setup.orientation = "landscape"
        book.save(path)
        book.close()
    names = {"5-1": ("current", "prior", "target"), "5-7": ("current", "prior")}
    inputs = [directory / (name + ".xlsx") for name in names.get(case, ("source",))]
    if scenario == "5-7-unmatched":
        book = openpyxl.load_workbook(inputs[0])
        book["惠州市涉农贷款分机构"].append([2, "新增机构", 3, 99])
        book.save(inputs[0])
        book.close()
    if case == "5-6":
        config = directory / "config.xlsx"
        book = openpyxl.load_workbook(config)
        # The legacy config contract has no destination sheet override.
        book["工作表提取"]["I2"] = None
        if scenario == "5-6-full":
            book["工作表提取"]["E2"] = None
        book.save(config)
        book.close()
    if scenario == "5-6-mixed":
        second = directory / "source-second.xlsx"
        helper["standard_source"](second)
        inputs.append(second)
    elif scenario == "5-6-new":
        products[0].unlink()
    elif scenario == "5-6-collision":
        book = openpyxl.load_workbook(products[0])
        book.create_sheet("数据")["A1"] = "不要覆盖"
        book.save(products[0])
        book.close()
    elif scenario == "5-6-formula-cache":
        source = inputs[0]
        book = openpyxl.load_workbook(source)
        book.active["C4"] = "=SUM(1,2)"
        book.save(source)
        book.close()
        seed = source.with_suffix(".uncached.xlsx")
        source.rename(seed)
        formula = b"<f>SUM(1,2)</f><v></v>"
        with ZipFile(seed) as before, ZipFile(source, "w") as after:
            for info in before.infolist():
                content = before.read(info)
                if info.filename == "xl/worksheets/sheet1.xml":
                    assert content.count(formula) == 1
                    content = content.replace(formula, b"<f>SUM(1,2)</f><v>3</v>")
                after.writestr(info, content)
    return inputs, products


def legacy_run(case, paths, directory, append):
    module, name = LEGACY[case]
    action = getattr(importlib.import_module(f"PyCode.core.{module}"), name)
    config = importlib.import_module("PyCode.core.config_xlsx")
    cfg = directory / "config.xlsx"
    log = directory / "legacy-logs"
    log.mkdir()
    output = directory / "output"
    output.mkdir()
    if case == "5-1":
        return action(paths[0], paths[1], [paths[2]], log)
    if case == "5-2":
        return action(paths, config.load_institution_mapping(cfg), log)
    if case == "5-3":
        return action(paths, config.load_foreign_banks(cfg), log)
    if case == "5-4":
        return action(paths, log)
    if case == "5-5":
        return action(paths, output, log)
    if case == "5-6":
        tasks = config.load_extract_tasks(cfg)
        native_inputs = [path for path in paths if path.suffix == ".xlsx"]
        if native_inputs:
            native = importlib.import_module("PyCode.core.feature_1_6_extract_sheets")
            native_result = native.run_extract_sheets(
                native_inputs, tasks, output, log, append_existing=append
            )
        else:
            native_result = None
        com_inputs = [path for path in paths if path.suffix == ".xls"]
        return {"native": native_result, "com": action(
            com_inputs, tasks, output, log, append_existing=append or bool(native_inputs)
        )}
    return action(paths[0], paths[1], log)


def menu_input(case, paths, append):
    file_list = ";".join(map(str, paths))
    number = case[-1]
    prompts = {
        "5-1": [*map(str, paths[:2]), "1", *map(str, paths[2:])],
        "5-2": ["1", file_list],
        "5-3": ["1", file_list],
        "5-4": [file_list],
        "5-5": ["1", file_list],
        "5-6": ["y" if append else "n", file_list],
        "5-7": [*map(str, paths)],
    }
    return "\n".join(["5", number, *prompts[case], "0", "0", ""])


def rust_worker_trace(directory, selected_host, fault=None, source=None, replacement=None):
    worker = ROOT / "pytools-core/python/pywin32_excel_conversion_worker.py"
    wrapper = directory / "test_com_worker.py"
    trace = directory / "com-hosts.jsonl"
    wrapper.write_text(
        "import json, runpy, sys\nfrom pathlib import Path\nimport win32com.client\n"
        "dispatch = win32com.client.DispatchEx\n"
        "def pinned_dispatch(prog_id, *args, **kwargs):\n"
        f"    if prog_id.casefold() != {selected_host.casefold()!r}:\n"
        "        raise RuntimeError('Host excluded by the parity test: ' + prog_id)\n"
        "    app = dispatch(prog_id, *args, **kwargs)\n"
        f"    with open({str(trace)!r}, 'a', encoding='utf-8') as log:\n"
        "        log.write(json.dumps({'prog_id': prog_id, 'application': app.Name, 'version': app.Version}) + '\\n')\n"
        "    return app\n"
        "win32com.client.DispatchEx = pinned_dispatch\n"
        f"sys.argv[0] = {str(worker)!r}\n"
        f"ns = runpy.run_path({str(worker)!r})\n"
        "session = ns['ConversionSession']\n"
        f"fault = {fault!r}\n"
        f"source = {str(source) if source else None!r}\n"
        f"replacement = {str(replacement) if replacement else None!r}\n"
        "if fault == 'save-failure':\n"
        "    open_book = session._open_workbook\n"
        "    class FailSave:\n"
        "        def __init__(self, book): self.book = book\n"
        "        def __getattr__(self, name): return getattr(self.book, name)\n"
        "        def Save(self): raise PermissionError('synthetic COM Save failure')\n"
        "    def open_with_fault(self, *args, **kwargs):\n"
        "        return FailSave(open_book(self, *args, **kwargs))\n"
        "    session._open_workbook = open_with_fault\n"
        "elif fault in ('validation-failure', 'source-changed'):\n"
        "    def post_response(action):\n"
        "        def run(*args, **kwargs):\n"
        "            result = action(*args, **kwargs)\n"
        "            if result.get('ok') and result.get('verified'):\n"
        "                if fault == 'validation-failure':\n"
        "                    result.update(verified=False, saved=True, validation_issue='synthetic validation mismatch')\n"
        "                else:\n"
        "                    Path(source).write_bytes(Path(replacement).read_bytes())\n"
        "            return result\n"
        "        return run\n"
        "    session.delete_rows_com = post_response(session.delete_rows_com)\n"
        "    session.fx_header_com = post_response(session.fx_header_com)\n"
        "    session.village_split_com = post_response(session.village_split_com)\n"
        "    session.institution_normalize_com = post_response(session.institution_normalize_com)\n"
        "    session.rural_loan_com = post_response(session.rural_loan_com)\n"
        "raise SystemExit(ns['main']())\n",
        encoding="utf-8",
    )
    return wrapper, trace


def observe(app, path, directory):
    # WPS may touch even a read-only file, so observe a disposable copy.
    copy = directory / path.name
    shutil.copy2(path, copy)
    workbook = app.Workbooks.Open(str(copy), UpdateLinks=0, ReadOnly=True)
    try:
        result = {}
        for sheet in workbook.Worksheets:
            used = sheet.UsedRange
            end_row = used.Row + used.Rows.Count - 1
            end_col = used.Column + used.Columns.Count - 1
            cells = sheet.Range(sheet.Cells(1, 1), sheet.Cells(end_row, end_col))
            result[sheet.Name] = {"values": cells.Value2, "formulas": cells.Formula}
            if "外汇" in sheet.Name:
                page = sheet.PageSetup
                props = ("LeftHeader", "CenterHeader", "RightHeader", "CenterFooter",
                         "DifferentFirstPageHeaderFooter", "OddAndEvenPagesHeaderFooter",
                         "PrintArea", "Orientation")
                result[sheet.Name]["page"] = {name: getattr(page, name, None) for name in props}
                for name in ("EvenPage", "FirstPage"):
                    variant = getattr(page, name, None)
                    if variant is not None:
                        result[sheet.Name]["page"][name] = {
                            side: getattr(variant, side).Text
                            for side in ("LeftHeader", "CenterHeader", "RightHeader", "RightFooter")
                        }
        return result
    finally:
        workbook.Close(SaveChanges=False)


def observe_xlsx_report(path):
    workbook = openpyxl.load_workbook(path)
    try:
        return {sheet.title: {"values": list(sheet.values),
                              "merges": sorted(map(str, sheet.merged_cells.ranges)),
                              "comments": {cell.coordinate: [cell.comment.text, cell.comment.author]
                                           for row in sheet for cell in row if cell.comment}}
                for sheet in workbook}
    finally:
        workbook.close()


def compare(left, right, path="$"):
    """Compare values with the repository's cross-engine float tolerance."""
    if isinstance(left, bool) or isinstance(right, bool):
        return [] if type(left) is type(right) and left == right else [path]
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        if abs(left - right) <= 1e-15 * max(abs(left), abs(right)):
            return []
    if isinstance(left, dict) and isinstance(right, dict):
        missing = [f"{path}.{key}" for key in sorted(left.keys() ^ right.keys())]
        return missing + [diff for key in sorted(left.keys() & right.keys())
                          for diff in compare(left[key], right[key], f"{path}.{key}")]
    if isinstance(left, list) and isinstance(right, list):
        size = [f"{path}.length"] if len(left) != len(right) else []
        return size + [diff for i, (a, b) in enumerate(zip(left, right))
                       for diff in compare(a, b, f"{path}[{i}]")]
    return [] if left == right else [path]


def compare_captures(legacy, ouo, current):
    def lookup(directory, route):
        return {r["case"]: r for r in json.loads((directory / "record.json").read_text(encoding="utf-8"))
                if r["route"] == route}

    def snapshot(directory, case, route):
        path = directory / case / route / "readback/0.json"
        return json.loads(path.read_text(encoding="utf-8")) if path.is_file() else None

    results = []
    references = [(directory, route, lookup(directory, route))
                  for directory, route in ((legacy, "legacy-xls"), (ouo, "rust-xlsx"))]
    for case, run in lookup(current, "rust-xls").items():
        current_snapshot = snapshot(current, case, "rust-xls")
        for directory, route, index in references:
            prior = index.get(case)
            left = snapshot(directory, case, route)
            if current_snapshot is None or prior is None or left is None:
                results.append({"case": case, "reference": route, "status": "NOT_TESTED"})
                continue
            differences = compare(left, current_snapshot)
            results.append({"case": case, "reference": route,
                            "snapshot_equal": not differences, "difference_paths": differences,
                            "sheet_order_equal": list(left) == list(current_snapshot),
                            "rust_status": run["status"], "reference_status": prior["status"],
                            "reference_output_com_open_error": prior.get("output_com_open_error"),
                            "rust_output_com_open_error": run.get("output_com_open_error")})
    return results


def verify_rust_xls(row, scenario, paths, original_files, directory, protected_target, replacement):
    if scenario == "5-7-mixed":
        assert row["exit_code"] != 0
        assert "必须均为 .xls 或均为 .xlsx" in (directory / "stderr.txt").read_text(encoding="utf-8")
    elif scenario == "5-7-validation-failure":
        assert row["exit_code"] != 0
    else:
        assert row["exit_code"] == 0, "menu subprocess failed"
    assert not row.get("output_com_open_error"), "Office could not reopen the result"
    assert not row["readonly_input_changes"], "read-only source/configuration changed"
    case = VARIANTS.get(scenario, scenario)
    source = paths[-1] if case == "5-1" else paths[0]
    normal_target = source.with_stem(f"{source.stem}(金融局)") if case == "5-3" else source
    outputs = directory / "readback"
    review = list(source.parent.glob("*_待核对_业务校验不符_*.xls"))
    conflicts = list(source.parent.glob("*_处理结果_源文件已变更_*.xls"))
    if scenario.endswith("validation-failure"):
        assert sha(source) == original_files[source]
        assert len(review) == 1 and not conflicts
        report = json.loads(review[0].with_suffix(".xls.run.json").read_text(encoding="utf-8"))
        assert not report["original_target_replaced"] and not report["business_validation_passed"]
        assert Path(report["planned_target"]) == normal_target
        assert report["review_copy_sha256"] == sha(review[0])
        assert report["expected_source_sha256"] == original_files[source]
        assert (outputs / "retained-0.json").is_file()
        if case == "5-3":
            assert not normal_target.exists()
    elif scenario.endswith("source-changed"):
        assert sha(source) == sha(replacement)
        assert len(conflicts) == 1 and not review
        report = json.loads(conflicts[0].with_suffix(".xls.run.json").read_text(encoding="utf-8"))
        assert not report["original_target_replaced"] and not report["new_edits_merged"]
        assert Path(report["commit"]["target"]) == conflicts[0]
        assert report["commit"]["sha256_after"] == sha(conflicts[0])
        assert (outputs / "retained-0.json").is_file()
        if case == "5-3":
            assert not normal_target.exists()
    elif scenario in ("5-3-com-disabled", "5-3-target-exists", "5-3-save-failure", "5-7-mixed"):
        assert sha(source) == original_files[source]
        assert not review and not conflicts
        if protected_target:
            assert sha(protected_target) == original_files[protected_target]
        elif case == "5-3":
            assert not normal_target.exists()
        if scenario == "5-3-save-failure":
            assert "synthetic COM Save failure" in (directory / "stdout.txt").read_text(encoding="utf-8")
        else:
            assert not row["actual_com_hosts"]
    elif scenario == "5-4-zero":
        assert sha(source) == original_files[source]
        assert row["actual_com_hosts"]
        assert (outputs / "0.json").is_file()
    else:
        assert row["outputs"] and all(output["changed_or_created"] for output in row["outputs"]), "result missing"
        assert (outputs / "0.json").is_file(), "result unreadable"
        if case in ("5-5", "5-6") or scenario == "5-3-zero":
            assert not row["actual_com_hosts"]
        else:
            assert row["actual_com_hosts"], "expected COM host was not called"
        if case == "5-3":
            body = json.loads((outputs / "0.json").read_text(encoding="utf-8"))["甲分机构本外币"]
            assert all(row_values[0] != "外资甲" for row_values in body["values"])
            if scenario.startswith("5-3-multiple-runs"):
                assert len(body["values"]) == 10
            if scenario == "5-3-zero":
                assert sha(normal_target) == original_files[source]
        if scenario == "5-6-formula-cache":
            result = json.loads((outputs / "0.json").read_text(encoding="utf-8"))
            assert result["数据"]["values"][3][2] == 3
    row["assertions_passed"] = True


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--exe", type=Path)
    parser.add_argument("--legacy", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--host", choices=("Excel.Application", "ket.Application"), required=True)
    parser.add_argument("--cases", nargs="+", choices=[*LEGACY, *VARIANTS], default=list(LEGACY))
    routes = ("legacy-xls", "rust-xlsx", "rust-xls")
    parser.add_argument("--routes", nargs="+", choices=routes, default=routes)
    parser.add_argument("--compare-legacy", type=Path)
    parser.add_argument("--compare-ouo", type=Path)
    parser.add_argument("--verify-rust-xls", action="store_true")
    args = parser.parse_args()
    if any(route.startswith("rust-") for route in args.routes):
        if args.exe is None:
            parser.error("--exe is required for a Rust route")
        args.exe = args.exe.resolve(strict=True)
    if bool(args.compare_legacy) != bool(args.compare_ouo):
        parser.error("provide both comparison baselines")
    args.legacy = args.legacy.resolve(strict=True)
    args.output = args.output.resolve()
    args.output.mkdir(parents=True)
    sys.path.insert(0, str(args.legacy))
    sys.dont_write_bytecode = True
    helper = runpy.run_path(str(ROOT / "tools/AsposeWorker/test_menu_matrix.py"))
    com = importlib.import_module("PyCode.core.com_engine")
    original_dispatch = win32com.client.DispatchEx
    def test_dispatch(prog_id, *a, **kw):
        # Make the legacy branch actually use the selected host, with no fallback.
        if prog_id.casefold() != args.host.casefold():
            raise RuntimeError(f"Excluded host for this run: {prog_id}")
        app = original_dispatch(prog_id, *a, **kw)
        if int(app.Workbooks.Count) != 0:
            raise RuntimeError("COM session already contains workbooks")
        app.AutomationSecurity = 3
        return app

    win32com.client.DispatchEx = test_dispatch
    records = []
    try:
        for scenario in args.cases:
            case = VARIANTS.get(scenario, scenario)
            append = scenario != "5-6-new"
            for route in args.routes:
                directory = args.output / scenario / route
                directory.parent.mkdir(parents=True, exist_ok=True)
                row = {"case": scenario, "route": route, "host": args.host}
                records.append(row)
                try:
                    paths, products = prepare(directory, scenario, helper)
                    if route.endswith("-xls"):
                        app = com.start_excel_app()
                        try:
                            if scenario == "5-6-mixed":
                                paths = [paths[1], create_biff(app, paths[0])]
                            elif scenario == "5-7-mixed":
                                paths = [create_biff(app, paths[0]), paths[1]]
                            else:
                                paths = [create_biff(app, path) for path in paths]
                        finally:
                            com.safe_quit(app)
                    cfg = directory / "config.xlsx"
                    modified_source = paths[-1] if case == "5-1" else paths[0]
                    replacement = None
                    protected_target = None
                    fault = next((name for name in ("save-failure", "validation-failure", "source-changed")
                                  if scenario.endswith(name)), None)
                    if fault == "source-changed":
                        replacement = directory / "concurrent.xls"
                        app = com.start_excel_app()
                        try:
                            book_path = directory / "concurrent.xlsx"
                            book = openpyxl.Workbook()
                            book.active["A1"] = "模拟运行期间的用户修改"
                            book.save(book_path)
                            book.close()
                            create_biff(app, book_path)
                        finally:
                            com.safe_quit(app)
                    if scenario == "5-3-target-exists":
                        protected_target = paths[0].with_stem(f"{paths[0].stem}(金融局)")
                        shutil.copy2(paths[0], protected_target)
                    row["input_sha256"] = {str(path): sha(path) for path in [*paths, cfg]}
                    original_files = {path: sha(path) for path in directory.rglob("*.xls*")}
                    original = directory / "before"
                    original.mkdir()
                    for index, path in enumerate([*paths, cfg]):
                        shutil.copy2(path, original / f"{index}-{path.name}")
                    if route == "legacy-xls":
                        row["legacy_callable"] = LEGACY[case]
                        row["legacy_result"] = legacy_run(case, paths, directory, append)
                    else:
                        wrapper, com_trace = rust_worker_trace(directory, args.host, fault, modified_source, replacement)
                        env = {k: v for k, v in os.environ.items()
                               if not k.startswith(("PYTOOLS_", "RUSTTOOLS_ASPOSE_"))}
                        env.update(PYTHONIOENCODING="utf-8",
                                   LOCALAPPDATA=str(directory / "state"),
                                   PYTOOLS_COM_PYTHON=sys.executable,
                                   PYTOOLS_COM_WORKER=str(wrapper),
                                   PYTOOLS_OPENPYXL_PYTHON=sys.executable,
                                   PYTOOLS_OPENPYXL_WORKER=str(ROOT / "crates/excel-editor/python/openpyxl_worker.py"),
                                   PYTOOLS_COM_PREFER_WPS="1" if args.host == "ket.Application" else "0")
                        row["command"] = [str(args.exe), "menu", "--config", str(cfg), "--log-view", "developer"]
                        if case in ("5-5", "5-6") or scenario == "5-3-com-disabled":
                            env["PYTOOLS_CONVERT_ENGINES_DISABLED"] = "1"
                            row["external_conversion_engines_disabled"] = True
                        row["stdin"] = menu_input(case, paths, append)
                        proc = subprocess.run(row["command"], input=row["stdin"], cwd=directory, env=env,
                                              capture_output=True, text=True, encoding="utf-8", errors="replace",
                                              timeout=180)
                        row["exit_code"] = proc.returncode
                        (directory / "stdout.txt").write_text(proc.stdout, encoding="utf-8")
                        (directory / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
                        row["actual_com_hosts"] = [json.loads(line) for line in
                                                   com_trace.read_text(encoding="utf-8").splitlines()] if com_trace.exists() else []
                        if case in ("5-5", "5-6"):
                            assert not row["actual_com_hosts"], "Rust read-only source path called COM"
                    if case in ("5-1", "5-2", "5-4", "5-7"):
                        outputs = [paths[-1] if case == "5-1" else paths[0]]
                    elif case == "5-6":
                        outputs = products
                    else:
                        outputs = sorted(path for path in directory.rglob("*.xls*")
                                         if path not in original_files and
                                         ("(金融局)" in path.stem or "地区数值核对结果" in path.stem))
                    row["outputs"] = [{"path": str(path), "sha256": sha(path),
                                       "changed_or_created": original_files.get(path) != sha(path)}
                                      for path in outputs if path.is_file()]
                    row["readonly_input_changes"] = [str(path) for path in (*paths, cfg)
                                                     if path not in outputs and sha(path) != original_files[path]
                                                     and not (path == paths[0] and fault == "source-changed")]
                    evidence = directory / "readback"
                    evidence.mkdir()
                    app = com.start_excel_app()
                    try:
                        for index, path in enumerate(outputs):
                            if path.is_file():
                                if case in ("5-5", "5-6"):
                                    write_json(evidence / f"{index}.json", observe_xlsx_report(path))
                                    try:
                                        observe(app, path, evidence)
                                    except Exception:
                                        row["output_com_open_error"] = traceback.format_exc()
                                else:
                                    write_json(evidence / f"{index}.json", observe(app, path, evidence))
                        if fault in ("validation-failure", "source-changed"):
                            retained = sorted([*modified_source.parent.glob("*_待核对_业务校验不符_*.xls"),
                                               *modified_source.parent.glob("*_处理结果_源文件已变更_*.xls")])
                            for index, path in enumerate(retained):
                                write_json(evidence / f"retained-{index}.json", observe(app, path, evidence))
                    finally:
                        com.safe_quit(app)
                    row["status"] = "captured"
                    if args.verify_rust_xls and route == "rust-xls":
                        verify_rust_xls(row, scenario, paths, original_files, directory, protected_target, replacement)
                    if route != "legacy-xls" and row["readonly_input_changes"]:
                        row["status"] = "error"
                        row["error"] = "Rust route changed read-only inputs"
                except Exception:
                    row["status"] = "error"
                    row["error"] = traceback.format_exc()
                write_json(args.output / "record.json", records)
    finally:
        win32com.client.DispatchEx = original_dispatch
    if args.compare_legacy:
        write_json(args.output / "comparison.json",
                   compare_captures(args.compare_legacy, args.compare_ouo, args.output))
    return int(any(r["status"] == "error" for r in records))


if __name__ == "__main__":
    raise SystemExit(main())
