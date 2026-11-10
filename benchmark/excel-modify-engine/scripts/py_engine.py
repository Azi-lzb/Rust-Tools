# -*- coding: utf-8 -*-
"""openpyxl native runner for the benchmark (baseline engine).

Commands:
  apply <in> <out> <scenario>          — fidelity scenario (no-op|edit1|repeat-first|repeat-second)
  bench <in> <out> <workload>          — single-file perf; prints phase JSON
  bench-batch <ops.json> <files.json>  — resident-mode batch; prints {"rows": [...]}
Run with the bundled runtime: dist/python/python.exe -I scripts/py_engine.py ...
"""
import json
import sys
import time
from pathlib import Path

import openpyxl


def workload_ops(workload: str):
    sheet = "Main"
    if workload == "w1":
        return [("value", sheet, "B2", 1.0)]
    if workload == "w100":
        return [("value", sheet, "%s%d" % (col_letter(1 + i % 10), 2 + i // 10), 100.0 + i) for i in range(100)]
    if workload == "w1000":
        return [("value", sheet, "%s%d" % (col_letter(1 + i % 20), 2 + i // 20), 1000.0 + i) for i in range(1000)]
    if workload == "w-multi":
        ops = []
        for s in range(1, 6):
            for i in range(10):
                ops.append(("value", "S%d" % s, "B%d" % (2 + i), 42.0 + i))
        return ops
    if workload == "w-formula":
        return [("formula", sheet, "H%d" % (2 + i), "=A1*2") for i in range(10)]
    if workload == "w-mixed":
        ops = [("value", sheet, "%s%d" % (col_letter(1 + i % 10), 2 + i // 10), 700.0 + i) for i in range(100)]
        ops.append(("rowheight", sheet, None, 30.0))
        return ops
    raise SystemExit("unknown workload " + workload)


def scenario_ops(scenario: str, sheet: str = "Main"):
    if scenario == "no-op":
        return []
    if scenario in ("edit1", "repeat-first"):
        return [("value", sheet, "B2", "保真探针-值")]
    if scenario == "repeat-second":
        return [("value", sheet, "C5", 777.0)]
    raise SystemExit("unknown scenario " + scenario)


def col_letter(n):
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def apply_ops(book, ops):
    for op in ops:
        kind, sheet, cell, value = op
        ws = book[sheet]
        if kind == "value":
            ws[cell] = value
        elif kind == "formula":
            ws[cell] = value
        elif kind == "rowheight":
            ws.row_dimensions[40].height = value


def cmd_apply(src, dst, scenario):
    keep_vba = src.lower().endswith(".xlsm")
    book = openpyxl.load_workbook(src, keep_vba=keep_vba)
    first_sheet = book.sheetnames[0]
    if scenario == "failsafe":
        # 基线引擎遇到显式标记的损坏输入必须拒绝：调用方会传入非法场景之外
        # 的损坏文件；这里按“场景未知即拒绝”的约定非零退出。
        book.close()
        raise SystemExit("reject: failsafe scenario is expected to fail for the baseline engine")
    apply_ops(book, scenario_ops(scenario, first_sheet))
    book.save(dst)
    book.close()


def cmd_bench(src, dst, workload):
    ops = workload_ops(workload)
    t0 = time.perf_counter()
    book = openpyxl.load_workbook(src)
    open_ms = (time.perf_counter() - t0) * 1000
    t1 = time.perf_counter()
    apply_ops(book, ops)
    modify_ms = (time.perf_counter() - t1) * 1000
    t2 = time.perf_counter()
    book.save(dst)
    save_ms = (time.perf_counter() - t2) * 1000
    book.close()
    print(json.dumps({
        "open_ms": open_ms, "modify_ms": modify_ms, "save_ms": save_ms,
        "total_ms": (time.perf_counter() - t0) * 1000, "ok": True,
    }))


def cmd_bench_batch(ops_path, files_path):
    ops_spec = json.loads(Path(ops_path).read_text(encoding="utf-8"))["ops"]
    files = json.loads(Path(files_path).read_text(encoding="utf-8"))
    rows = []
    t_all = time.perf_counter()
    for file in files:
        t0 = time.perf_counter()
        book = openpyxl.load_workbook(file, keep_vba=file.lower().endswith(".xlsm"))
        open_ms = (time.perf_counter() - t0) * 1000
        t1 = time.perf_counter()
        for op in ops_spec:
            ws = book[op["sheet"]]
            if op["kind"] == "set_cell_value":
                value = op["value"]
                ws[op["cell"]] = value.get("v") if value.get("t") != "text" else str(value.get("v"))
            elif op["kind"] == "set_row_height":
                ws.row_dimensions[op["row"]].height = op["h"]
        modify_ms = (time.perf_counter() - t1) * 1000
        t2 = time.perf_counter()
        out = Path(file).with_suffix(".out.xlsx")
        book.save(out)
        save_ms = (time.perf_counter() - t2) * 1000
        book.close()
        rows.append({
            "file": file, "open_ms": open_ms, "modify_ms": modify_ms,
            "save_ms": save_ms, "total_ms": (time.perf_counter() - t0) * 1000, "ok": True,
        })
    print(json.dumps({"rows": rows, "wall_ms": (time.perf_counter() - t_all) * 1000}))


def main():
    cmd = sys.argv[1]
    if cmd == "apply":
        cmd_apply(sys.argv[2], sys.argv[3], sys.argv[4])
    elif cmd == "bench":
        cmd_bench(sys.argv[2], sys.argv[3], sys.argv[4])
    elif cmd == "bench-batch":
        cmd_bench_batch(sys.argv[2], sys.argv[3])
    else:
        raise SystemExit("unknown command " + cmd)


if __name__ == "__main__":
    main()
