"""50-file initial references through existing Windows Release entry points.

One measured batch per route, after paired one-file correctness gates. Single-
task-only menus execute 50 real invocations, explicitly labelled as such. Never
multiply a single-file time, change a production route, or remove watermarks.
"""
import argparse
import copy
import importlib.util
import json
import os
import re
from pathlib import Path
import shutil
import subprocess
import sys
import time
import traceback
import zipfile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.workbook.defined_name import DefinedName
import benchmark_print_three as oracle

matrix = oracle.oracle
HERE = Path(__file__).resolve().parent
CASES = {
    "1-1": ["openpyxl", "aspose"],
    "1-5": ["com", "betteroffice", "aspose"],
    "2-1": ["openpyxl", "aspose"], "2-2": ["openpyxl", "aspose"],
    "2-3": ["openpyxl", "aspose"], "3-1": ["openpyxl", "rust", "aspose"],
    "4-4-xlsx": ["auto", "aspose"], "4-4-xls": ["auto", "aspose"],
    "4-4-word": ["auto", "aspose"], "4-5": ["openpyxl", "aspose"],
    "4-7": ["auto", "aspose"], "5-1": ["openpyxl", "aspose"],
    "5-2": ["openpyxl", "aspose"], "5-3": ["openpyxl", "aspose"],
    "5-4": ["auto", "openpyxl", "aspose"],
    "5-6": ["openpyxl", "aspose"], "5-7": ["openpyxl", "aspose"],
    "7-1": ["openpyxl", "aspose"],
}
CHOICE = {"auto": "1", "openpyxl": "2", "aspose": "3"}
GLOBAL_POLICIES = {"umya": "UMYAOnly", "betteroffice_modify": "BetterOfficeOnly",
                   "rustnative": "RustNativeOnly", "route_auto": "Auto"}
CHOICE.update({key: "1" for key in GLOBAL_POLICIES})


def dump(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, default=str, indent=2), encoding="utf8")


def rewrite_config(path, old, new):
    book = openpyxl.load_workbook(path)
    for sheet in book:
        for row in sheet:
            for cell in row:
                if isinstance(cell.value, str): cell.value = cell.value.replace(str(old), str(new))
    book.save(path); book.close()


def make_legacy(root, word=False):
    import win32com.client
    root.mkdir(parents=True, exist_ok=True)
    if word:
        from docx import Document
        from docx.shared import Pt
        original = root / "original.docx"; doc = Document(); doc.add_heading("合成 Word 对拍", 1)
        run = doc.add_paragraph().add_run("中文、00123、正文"); run.bold = True; run.font.size = Pt(14)
        doc.sections[0].header.paragraphs[0].text = "保留页眉"
        doc.sections[0].footer.paragraphs[0].text = "保留页脚"
        table = doc.add_table(rows=2, cols=2)
        for r in range(2):
            for c in range(2): table.cell(r, c).text = f"表格{r},{c}"
        doc.save(original)
        target = root / "source.doc"
        app = win32com.client.DispatchEx("Word.Application")
        try:
            app.Visible = False; app.DisplayAlerts = 0; app.AutomationSecurity = 3
            book = app.Documents.Open(str(original), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
            try: book.SaveAs2(str(target), FileFormat=0, AddToRecentFiles=False)
            finally: book.Close(False)
        finally: app.Quit(False)
    else:
        original = root / "original.xlsx"; matrix.standard_source(original)
        target = root / "source.xls"
        app = win32com.client.DispatchEx("Excel.Application")
        try:
            app.Visible = False; app.DisplayAlerts = False; app.AutomationSecurity = 3
            book = app.Workbooks.Open(str(original), UpdateLinks=0, ReadOnly=True, AddToMru=False)
            try: book.SaveAs(str(target), FileFormat=56)
            finally: book.Close(False)
        finally: app.Quit()
    return target


def calc_fixture(root, count):
    root.mkdir(parents=True)
    config = matrix.save_config(matrix.config(root), root)
    template = root / "template.xlsx"; book = openpyxl.Workbook(); s = book.active; s.title = "数据"
    s.append(["项目", "金额"])
    s["D2"] = '=IF(SUM(\'数据\'!B2:B3)=\'参数\'!A1,"通过","校验失败：金额不平")'
    s["E2"] = '=SUMIF(\'数据\'!A2:A3,"甲",\'数据\'!B2:B3)+VLOOKUP("乙",\'数据\'!A2:B3,2,FALSE)'
    book.create_sheet("参数")["A1"] = 30
    for name, address in [("表结构区域", "$A$1:$B$1"), ("校验区域", "$D$2:$E$2")]:
        book.defined_names.add(DefinedName(name, attr_text=f"'数据'!{address}"))
    book.save(template)
    with zipfile.ZipFile(template) as z: parts = {n: z.read(n) for n in z.namelist()}
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"; ET.register_namespace("", ns[1:-1])
    xml = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    for cell in xml.iter(ns + "c"):
        if cell.find(ns + "f") is not None:
            value = cell.find(ns + "v")
            if value is None: value = ET.SubElement(cell, ns + "v")
            value.text = "-999"; cell.attrib.pop("t", None)
    parts["xl/worksheets/sheet1.xml"] = ET.tostring(xml)
    with zipfile.ZipFile(template, "w", zipfile.ZIP_DEFLATED) as z:
        for name, value in parts.items(): z.writestr(name, value)
    for i in range(count):
        book = openpyxl.Workbook(); s = book.active; s.title = "数据"
        for row in [["项目", "金额"], ["甲", 10], ["乙", 20]]: s.append(row)
        book.save(root / f"source_{i:03}.xlsx")


def canonical(root, case, count, legacy):
    if case == "1-5": calc_fixture(root, count); return
    if case in ("4-4-xls", "4-4-word"):
        root.mkdir(parents=True); matrix.save_config(matrix.config(root), root)
        for i in range(count): shutil.copyfile(legacy, root / f"source_{i:03}{legacy.suffix}")
        return
    alias = {"4-4-xlsx": "4-4", "5-4": "5-4-menu"}.get(case, case)
    for i in range(count):
        job = root / f"j{i:03}"; matrix.prepare(job, alias, False)
        if not (job / "config.xlsx").exists(): matrix.save_config(matrix.config(job), job)
        source = job / "source.xlsx"
        if source.exists():
            renamed = job / f"source_{i:03}.xlsx"; source.rename(renamed)
            rewrite_config(job / "config.xlsx", source, renamed)
            if case in ("1-1", "5-6"):
                book = openpyxl.load_workbook(job / "config.xlsx")
                name = "时序提取规则" if case == "1-1" else "工作表提取"
                book[name].cell(2, 3 if case == "1-1" else 2, renamed.stem)
                if case == "1-1": book[name].cell(2, 2, f"WideRule{i:03}")
                book.save(job / "config.xlsx")


def merged_config(root, jobs, sheets):
    book = openpyxl.load_workbook(jobs[0] / "config.xlsx")
    for name in sheets:
        sheet = book[name]; sheet.delete_rows(2, sheet.max_row)
        for job in jobs:
            source = openpyxl.load_workbook(job / "config.xlsx")
            for row in list(source[name].values)[1:]: sheet.append(row)
            source.close()
    book.save(root / "config.xlsx"); book.close()
    return root / "config.xlsx"


def prepare_run(canonical_root, root, case, engine, count):
    shutil.copytree(canonical_root, root)
    for config in root.rglob("config.xlsx"): rewrite_config(config, canonical_root, root)
    jobs = [root / f"j{i:03}" for i in range(count)]
    sources = [job / f"source_{i:03}.xlsx" for i, job in enumerate(jobs)]
    choice = CHOICE.get(engine)
    commands = []; outputs = []; readonly = []; originals = []
    config = root / "config.xlsx"
    if case == "1-5":
        sources = sorted(root.glob("source_*.xlsx")); template = root / "template.xlsx"
        commands.append((["summary-named-validation", "--config", str(config), "--calc-engine", engine, "--template", str(template), "--output-dir", str(root / "output"), "--log-dir", str(root / "logs")] + list(map(str, sources)), None))
        outputs = [root / "output/副本" / p.name for p in sources]; readonly = sources + [template, config]
    elif case in ("4-4-xls", "4-4-word"):
        suffix = ".xls" if case.endswith("xls") else ".doc"
        sources = sorted(root.glob("source_*" + suffix))
        commands.append((["convert-excel" if suffix == ".xls" else "convert-word", "--config", str(config), "--engine", engine, "--output-dir", str(root / "converted")] + list(map(str, sources)), None))
        outputs = [root / "converted" / (p.stem + (".xlsx" if suffix == ".xls" else ".docx")) for p in sources]
        readonly = sources + [config]
    elif case in ("2-1", "2-2", "2-3", "3-1", "1-1", "5-6"):
        task_sheet = {"1-1": "时序提取规则", "3-1": "打印配置", "5-6": "工作表提取"}.get(case, "去重追加数据配置")
        config = merged_config(root, jobs, [task_sheet])
        readonly.append(config)
        if case.startswith("2-"):
            commands.append((["menu", "--config", str(config)], f"2\n{case[-1]}\n{choice}\n0\n0\n"))
            outputs = [job / "target.xlsx" for job in jobs] if case == "2-3" else sources
            if case == "2-3": readonly += sources
        elif case == "3-1":
            pick = {"openpyxl": "1", "rust": "2", "aspose": "3"}[engine]
            commands.append((["menu", "--config", str(config)], f"3\n1\n{pick}\n0\n0\n"))
            outputs = [job / {"openpyxl": "target.xlsx", "rust": "target_Rust.xlsx", "aspose": "target_Aspose.xlsx"}[engine] for job in jobs]
            readonly += sources
        elif case == "1-1":
            commands.append((["wide-summary", "--config", str(config), "--modify-engine", engine] + list(map(str, sources)), None))
            outputs = [job / "target.xlsx" for job in jobs]; readonly += sources
        else:
            commands.append((["extract-sheets", "--config", str(config), "--modify-engine", engine, "--append-existing"] + list(map(str, sources)), None))
            outputs = [job / "extracted/existing.xlsx" for job in jobs]; readonly += sources
    elif case in ("5-7", "7-1"):
        for job in jobs:
            cfg = job / "config.xlsx"
            if case != "7-1": readonly.append(cfg)
            if case == "5-7":
                commands.append((["menu", "--config", str(cfg)], f"5\n7\n{choice}\n{job / 'current.xlsx'}\n{job / 'prior.xlsx'}\n0\n0\n"))
                outputs.append(job / "current.xlsx"); readonly.append(job / "prior.xlsx")
            else:
                # This is an isolated synthetic config, explicitly complete
                # the safe commit rather than timing a cancelled preview.
                commands.append((["menu", "--config", str(cfg)], f"7\n1\n1\n{choice}\ny\n0\n0\n"))
                outputs.append(cfg)
    else:
        config = jobs[0] / "config.xlsx"; readonly.append(config)
        if case == "5-1":
            commands.append((["menu", "--config", str(config)], f"5\n1\n{choice}\n{jobs[0] / 'current.xlsx'}\n{jobs[0] / 'prior.xlsx'}\n1\n" + ";".join(str(job / "target.xlsx") for job in jobs) + "\n0\n0\n"))
            outputs = [job / "target.xlsx" for job in jobs]
            readonly += [jobs[0] / "current.xlsx", jobs[0] / "prior.xlsx"]
        elif case == "5-4":
            commands.append((["menu", "--config", str(config)], f"5\n4\n{choice}\n" + ";".join(map(str, sources)) + "\n0\n0\n"))
            outputs = sources
        else:
            mapping = {"4-4-xlsx": "convert-excel", "4-5": "rename-sheets", "4-7": "excel-open-repair", "5-2": "institution-normalize", "5-3": "remove-foreign-bank"}
            command = [mapping[case], "--config", str(config)]
            command += ["--engine", engine] if case in ("4-4-xlsx", "4-7") else ["--modify-engine", engine]
            if case == "4-4-xlsx":
                command += ["--output-dir", str(root / "converted")]; outputs = [root / "converted" / p.name for p in sources]; readonly += sources
            elif case == "4-7":
                outputs = [p.with_name(p.stem + ("-Excel修复_Aspose.xlsx" if engine == "aspose" else "-Excel修复.xlsx")) for p in sources]; readonly += sources
            elif case == "5-2": command += ["--apply", "--in-place", "auto"]; outputs = sources
            elif case == "5-3":
                command += ["--create-copy"]; outputs = [p.with_name(p.stem + "(金融局).xlsx") for p in sources]; readonly += sources
            else: outputs = sources
            commands.append((command + list(map(str, sources)), None))
    if os.environ.get("MENU50_REAL_MENU") == "1" and case in (
        "1-1", "1-5", "4-4-xlsx", "4-4-xls", "4-4-word", "4-5", "4-7", "5-2", "5-3", "5-6"
    ):
        selected = ";".join(map(str, sources))
        if case == "1-1": stdin = f"1\n1\n{choice}\n1\n{selected}\n0\n0\n"
        elif case == "1-5":
            pick = {"com":"2", "betteroffice":"3", "aspose":"4"}[engine]
            stdin = f"1\n5\n{pick}\n{root / 'template.xlsx'}\n1\n{selected}\n0\n0\n"
        elif case.startswith("4-4"):
            word = case == "4-4-word"
            stdin = f"4\n4\n{'2' if word else '1'}\n{'2' if engine == 'aspose' else '1'}\n1\n{selected}\n0\n0\n"
            extension = "docx" if word else "xlsx"
            outputs = [p.parent / extension / (p.stem + "." + extension) for p in sources]
        elif case in ("4-5", "4-7"):
            pick = choice if case == "4-5" else ("2" if engine == "aspose" else "1")
            stdin = f"4\n{case[-1]}\n{pick}\n1\n{selected}\ny\n0\n0\n"
        elif case == "5-6": stdin = f"5\n6\n{choice}\ny\n{selected}\n0\n0\n"
        else: stdin = f"5\n{case[-1]}\n{choice}\n1\n{selected}\n0\n0\n"
        commands = [(["menu", "--config", str(config)], stdin)]
    for p in outputs:
        if p.is_file() and p.suffix == ".xlsx": originals.append(oracle.view(p)[0])
        else: originals.append(None)
    return commands, outputs, list(set(readonly)), originals


def normalized(value, root):
    return json.loads(json.dumps(value, ensure_ascii=False, default=str).replace(json.dumps(str(root))[1:-1], "<CASE_ROOT>"))


def execute(exe, canonical_root, root, case, engine, count):
    commands, outputs, readonly, originals = prepare_run(canonical_root, root, case, engine, count)
    before = {str(p): oracle.sha(p) for p in readonly}
    env = {**os.environ, "LOCALAPPDATA": str(root / "localappdata"), "PYTHONIOENCODING": "utf8", "PYTHONDONTWRITEBYTECODE": "1"}
    # Absence is significant for the original narrow FX-header patch path.
    # Explicit Openpyxl selection is scoped by the real menu/CLI instead.
    env.pop("PYTOOLS_EXCEL_EDIT_ROUTE", None)
    if engine in GLOBAL_POLICIES:
        env["PYTOOLS_EXCEL_EDIT_ROUTE"] = GLOBAL_POLICIES[engine]
        for args, _ in commands:
            if "--modify-engine" in args: args[args.index("--modify-engine") + 1] = "auto"
    elapsed = 0.0; executions = []
    for i, (args, stdin) in enumerate(commands):
        clock = time.perf_counter_ns()
        p = subprocess.run([exe] + args, input=stdin, capture_output=True, text=True, encoding="utf8", cwd=root, env=env, timeout=1200)
        duration = (time.perf_counter_ns() - clock) / 1e6; elapsed += duration
        (root / f"stdout-{i:03}.txt").write_text(p.stdout, encoding="utf8"); (root / f"stderr-{i:03}.txt").write_text(p.stderr, encoding="utf8")
        executions.append(dict(arguments=args, stdin=stdin, e2e_ms=duration, returncode=p.returncode))
        dump(root / "EXECUTIONS.json", executions)
        assert p.returncode == 0, p.stderr[-3000:]
        if len(commands) > 1 and i % 10 == 9: print(json.dumps(dict(case=case, engine=engine, completed_tasks=i+1, tasks=len(commands))), flush=True)
    assert all(oracle.sha(Path(p)) == h for p,h in before.items()), "read-only input/config changed"
    views = []; raw_eval = []
    for index, path in enumerate(outputs):
        if path.name == "CONFIG_CANDIDATE":
            matches = list(path.parent.rglob("config-candidate.xlsx")); assert len(matches) == 1, matches
            path = matches[0]
        assert path.is_file(), f"Missing expected output: {path}"
        if path.suffix == ".docx":
            spec = importlib.util.spec_from_file_location("word_oracle", HERE.parent / "AsposeWordsWorker/benchmark_word.py")
            module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)
            value = module.snapshot(path); excluded = []
            with zipfile.ZipFile(path) as z:
                assert any(b"Aspose.Words" in z.read(n) for n in z.namelist() if n.endswith(".xml")) == (engine == "aspose")
        else:
            value, raw, excluded = oracle.view(path)
            assert len(excluded) == (1 if engine == "aspose" else 0), (path,excluded)
            if case == "1-5":
                cells = value["sheets"]["数据"]["cells"]
                assert cells["D2"]["cached"] == "通过" and cells["E2"]["cached"] == 30, cells
            if case == "4-5": assert "新的名称" in value["sheets"]
            if case == "4-4-xlsx" and engine == "aspose":
                source_view = oracle.view(root / f"j{index:03}" / f"source_{index:03}.xlsx")[0]
                assert not matrix.differences(source_view,value), "Aspose must preserve source, not inherit COM typography changes"
            if case == "5-2": assert value["sheets"]["甲分机构本外币"]["cells"]["A2"]["value"] == "标准机构"
            if case == "3-1":
                sheet=value["sheets"]["打印结果"]
                assert not sheet["merges"] and sheet["cells"]["B1"]["value"] == "0"
            if originals[index] is not None:
                assert matrix.differences(originals[index], value), f"No business change: {path}"
        raw_eval.append(dict(path=str(path), evaluation=excluded)); views.append(normalized(value, root))
    assert len(views) == count
    dump(root / "OUTPUT_SNAPSHOTS.json", views); dump(root / "EVALUATION.json", raw_eval)
    return dict(case=case, engine=engine, files=count, invocation_count=len(commands),
                entry="menu" if all(args[0] == "menu" for args,_ in commands) else "cli",
                batch_kind="one_batch" if len(commands)==1 else "50_independent_tasks",
                total_e2e_ms=elapsed, verified_outputs=len(views), readonly_sha256=before, outputs=[str(p) for p in outputs]), views


def compare(case, baseline, actual):
    raw_diff = matrix.differences(baseline, actual)
    if case == "1-5":
        # Only normalize RANGE tokens naming these two known synthetic sheets.
        # Quoted string literals, other names, operators and caches stay strict.
        from openpyxl.formula import Tokenizer
        baseline, actual = copy.deepcopy(baseline), copy.deepcopy(actual)
        for view in baseline + actual:
            for sheet in view["sheets"].values():
                for cell in sheet["cells"].values():
                    if cell["data_type"] != "f": continue
                    tokens = Tokenizer(cell["value"]).items
                    for token in tokens:
                        if token.type == "OPERAND" and token.subtype == "RANGE":
                            token.value = re.sub(r"^'(数据|参数)'!", r"\1!", token.value)
                    cell["value"] = "=" + "".join(t.value for t in tokens)
    diff = matrix.differences(baseline, actual)
    allowed = []
    if case == "1-5":
        remaining_paths = {d["path"] for d in diff}
        allowed += [dict(d, classification="equivalent_known_sheet_RANGE_quoting") for d in raw_diff if d["path"] not in remaining_paths]
    if case in ("4-4-xlsx", "4-4-xls", "1-5"):
        appearance = [d for d in diff if (
            d["path"].endswith("/style/font/0")
            and {d.get("expected"), d.get("actual")} in ({"Calibri", "宋体"}, {"Calibri", "SimSun"})
        ) or (
            d["path"].endswith("/default_row_height")
            and {d.get("expected"), d.get("actual")} == {13.5, 15.0}
        )]
        allowed += appearance
        diff = [d for d in diff if d not in appearance]
    return diff, allowed


def main():
    p = argparse.ArgumentParser(); p.add_argument("--output", required=True)
    p.add_argument("--exe", default="F:/RustBuild/pytools/release/RustTools.exe")
    p.add_argument("--cases", nargs="+", default=list(CASES)); args = p.parse_args()
    root = Path(args.output).resolve(); root.mkdir(parents=True, exist_ok=False)
    dump(root / "ENVIRONMENT.json", dict(exe=args.exe, exe_sha256=oracle.sha(Path(args.exe)), arguments=vars(args),
         worker_sha256=oracle.sha(Path(os.environ["RUSTTOOLS_ASPOSE_WORKER_DLL"])),
         timing="one preliminary measured 50-file batch; startup/processing/safety/logging all included; no user waiting; never single-file x50"))
    results = []
    for case in args.cases:
        entry = dict(case=case, baseline=CASES[case][0], routes=[])
        try:
            legacy = None
            if case in ("4-4-xls", "4-4-word"): legacy=make_legacy(root / "legacy" / case, case.endswith("word"))
            for n in (1,50): canonical(root / "canonical" / case / str(n),case,n,legacy)
            preflight = {}; baseline = None
            for engine in CASES[case]:
                try:
                    info, views = execute(args.exe, root / "canonical" / case / "1", root / case / "preflight" / engine, case, engine,1)
                    if baseline is None: baseline = views
                    diff, allowed = compare(case, baseline, views)
                    dump(root / case / "preflight" / engine / "DIFFERENCES.json",dict(differences=diff,approved_appearance_differences=allowed))
                    assert not diff, diff[:8]
                    preflight[engine] = True
                    print(json.dumps(dict(case=case, engine=engine, stage="preflight",status="PASS")), flush=True)
                except Exception:
                    preflight[engine] = False
                    entry["routes"].append(dict(engine=engine,status="PREFLIGHT_FAILED",error=traceback.format_exc()))
                    print(json.dumps(dict(case=case,engine=engine,stage="preflight",status="FAILED")),flush=True)
            if not preflight.get(entry["baseline"]): raise AssertionError("Baseline preflight failed; no ratios")
            batch_baseline=None
            for engine in CASES[case]:
                if not preflight[engine]:continue
                try:
                    info, views = execute(args.exe,root / "canonical" / case / "50",root / case / "batch50" / engine,case,engine,50)
                    if engine==entry["baseline"]:batch_baseline=views
                    assert batch_baseline is not None
                    diff, allowed=compare(case,batch_baseline,views)
                    dump(root / case / "batch50" / engine / "DIFFERENCES.json",dict(differences=diff,approved_appearance_differences=allowed))
                    assert not diff,diff[:8]
                    info.update(status="PASS",approved_appearance_differences=len(allowed),evidence=str(root / case / "batch50" / engine))
                    entry["routes"].append(info)
                    print(json.dumps({k:info[k] for k in ["case","engine","files","invocation_count","total_e2e_ms","status"]}),flush=True)
                except Exception: entry["routes"].append(dict(engine=engine,status="BATCH_FAILED",error=traceback.format_exc()))
        except Exception:entry["error"]=traceback.format_exc()
        results.append(entry); dump(root / "RESULTS.json",dict(results=results))
        print(json.dumps(dict(case=case,stage="complete",statuses={r['engine']:r['status'] for r in entry['routes']})),flush=True)
    return 0 if all(all(r['status']=="PASS" for r in c['routes']) and not c.get('error') for c in results) else 1


if __name__ == "__main__":raise SystemExit(main())
