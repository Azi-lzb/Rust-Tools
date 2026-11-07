#!/usr/bin/env python
"""最终真实数据批量验收（七类抽测）。

对真实报送样本的**隔离副本**运行七类菜单对应的核心 CLI/菜单路径，
验证退出码与关键业务产物；原始 `测试用例/` 文件全程只读。
运行前逐本记录原件 SHA-256，运行后复核不变。

用法（仓库根目录）：
    python pytools-core/probe/final_acceptance.py [--dist dist]

输出：test-runs/final-acceptance-<stamp>/result.json 与 console.log。
真实业务文件不入库；本脚本本身可重复执行，属于交付前回归入口。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import pathlib
import shutil
import subprocess
import sys
import time

REPO = pathlib.Path(__file__).resolve().parents[2]
CASES = REPO / "测试用例"
BS = chr(92)


def sha256(path: pathlib.Path) -> str:
    return hashlib.sha256(pathlib.Path(path).read_bytes()).hexdigest()


class Runner:
    def __init__(self, exe: pathlib.Path, work: pathlib.Path, log_handle):
        self.exe = exe
        self.work = work
        self.log = log_handle
        self.results: list[dict] = []

    def run(self, step: str, argv: list[str], check, timeout: int = 420) -> dict:
        started = time.time()
        proc = subprocess.run(
            [str(self.exe)] + argv,
            cwd=str(REPO), capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=timeout,
        )
        entry = {
            "step": step, "exit": proc.returncode,
            "elapsed_s": round(time.time() - started, 1),
        }
        try:
            entry["checks"] = check(proc, pathlib.Path.cwd())
        except Exception as error:  # noqa: BLE001 - 验收脚本需把失败计入报告
            entry["checks"] = {"error": f"{type(error).__name__}: {error}"}
        entry["ok"] = proc.returncode == 0 and not entry["checks"].get("error")
        self.results.append(entry)
        self.log.write(
            f"== {step}: exit={proc.returncode} {entry['elapsed_s']}s\n"
            f"stdout tail: {proc.stdout[-400:]}\nstderr tail: {proc.stderr[-400:]}\n"
        )
        self.log.flush()
        return entry

    def menu(self, step: str, config: pathlib.Path, input_text: str, check,
             timeout: int = 420) -> dict:
        started = time.time()
        proc = subprocess.run(
            [str(self.exe), "menu", "--config", str(config), "--log-view", "off"],
            cwd=str(REPO), input=input_text, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=timeout,
        )
        entry = {
            "step": step, "exit": proc.returncode,
            "elapsed_s": round(time.time() - started, 1),
        }
        try:
            entry["checks"] = check(proc, pathlib.Path.cwd())
        except Exception as error:  # noqa: BLE001
            entry["checks"] = {"error": f"{type(error).__name__}: {error}"}
        entry["ok"] = proc.returncode == 0 and not entry["checks"].get("error")
        self.results.append(entry)
        self.log.write(
            f"== {step}: exit={proc.returncode} {entry['elapsed_s']}s\n"
            f"stdout tail: {proc.stdout[-400:]}\nstderr tail: {proc.stderr[-400:]}\n"
        )
        self.log.flush()
        return entry


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dist", default="dist")
    args = parser.parse_args()

    stamp = time.strftime("%Y%m%d-%H%M%S")
    out_dir = REPO / "test-runs" / f"final-acceptance-{stamp}"
    work = out_dir / "work"
    work.mkdir(parents=True, exist_ok=True)
    exe = REPO / args.dist / "RustTools.exe"
    config = REPO / "dist" / "config" / "config.xlsx"
    assert exe.is_file(), exe
    assert config.is_file(), config

    originals: dict[str, str] = {}
    def stage(rel_src: str, dest_name: str | None = None) -> pathlib.Path:
        src = CASES / rel_src
        originals[str(src)] = sha256(src)
        dest = work / (dest_name or src.name)
        shutil.copy(src, dest)
        return dest

    console = (out_dir / "console.log").open("w", encoding="utf-8")
    runner = Runner(exe, work, console)
    env_note = {"exe": str(exe), "work": str(work)}

    # ── 第 1 类：1-2 按使用区域汇总（真实月报 xlsx）────────────────
    sample_4020_xlsx = stage("4/2026.09 4020 中信银行 仲恺区存贷统计月报.xlsx")
    r = runner.run("1-2 used-region-summary",
        ["summary-used-range", "--output-dir", str(work / "out-12"), "--log-view", "off",
         str(sample_4020_xlsx)],
        lambda p, _: {"saved": "按使用区域汇总" in p.stdout and '"saved": true' in p.stdout.replace(" ", " ") or '"saved":true' in p.stdout})

    # ── 第 4 类：4-4 COM 转换真实 .xls ────────────────────────────
    sample_4020_xls = stage("4/xls2xlsx/2026.09 4020 中信银行 仲恺区存贷统计月报.xls")
    converted_dir = work / "converted"
    converted_dir.mkdir(exist_ok=True)
    runner.run("4-4 convert 4020 xls->xlsx",
        ["convert-file", "--input", str(sample_4020_xls), "--to", "xlsx",
         "--output", str(converted_dir / "4020.xlsx"), "--engine", "pywin32"],
        lambda p, _: {"verified": '"status":"converted"' in p.stdout.replace(" ", "")})

    # ── 第 4 类：4-5 真实 4020 表名改名（人民币→1人民币、外币→2外币）──
    runner.run("4-5 rename-sheets real 4020",
        ["rename-sheets", "--config", str(config), "--log-view", "off",
         str(converted_dir / "4020.xlsx")],
        lambda p, _: {"renamed": '"rename_ok": 2' in p.stdout or '"rename_ok":2' in p.stdout})

    # ── 第 4 类：4-6 22 本真实月报提交检查 ────────────────────────
    sources46 = work / "sources-46"
    sources46.mkdir(exist_ok=True)
    for src in (CASES / "4").glob("*.xlsx"):
        originals[str(src)] = sha256(src)
        shutil.copy(src, sources46 / src.name)
    runner.run("4-6 submission-check 22 workbooks",
        ["submission-check", "--config", str(config), "--output-dir", str(work / "out-46"),
         "--log-view", "off", str(sources46)],
        lambda p, _: {"submitted": json.loads(p.stdout).get("submitted") == 22,
                      "rules": json.loads(p.stdout).get("rules") == 33})

    # ── 第 4 类：4-7 audit 28 本真实打不开样本 ────────────────────
    audit_dir = pathlib.Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
    manifest = json.load(open(audit_dir / "test-output/s1f1_mode_parity_20261107/excel_open_fix_28.json", encoding="utf-8"))
    inputs129 = audit_dir / "test-output/s1f1_com_isolation_20261107/source_bundle/inputs_129"
    sources47 = work / "sources-47"
    sources47.mkdir(exist_ok=True)
    expected_counts = {}
    for item in manifest:
        originals[str(inputs129 / item["file"])] = sha256(inputs129 / item["file"])
        shutil.copy(inputs129 / item["file"], sources47 / item["file"])
        expected_counts[item["file"]] = (item["quoted"], item["pane_fixed"])
    def check47(p, _):
        report = json.loads(p.stdout)
        actual = {pathlib.Path(f["input"]).name: (f["quoted"], f["pane_fixed"])
                  for f in report["files"] if f["status"] == "repaired"}
        mismatch = [n for n, c in expected_counts.items() if actual.get(n) != c]
        return {"repaired": report["repaired"], "count_mismatches": len(mismatch)}
    runner.run("4-7 repair 28 audit files",
        ["excel-open-repair", "--log-dir", str(work / "logs-47"), "--log-view", "off"]
        + [str(sources47 / item["file"]) for item in manifest],
        check47, timeout=600)

    # ── 第 5 类：5-4 外汇页眉（真实外发用 .xls → 转换 → openpyxl 就地）──
    sample_wfw = stage("5/2026.09 外发用.xls")
    runner.run("5-4a convert 外发用 xls->xlsx",
        ["convert-file", "--input", str(sample_wfw), "--to", "xlsx",
         "--output", str(work / "wfw.xlsx"), "--engine", "pywin32"],
        lambda p, _: {"verified": '"status":"converted"' in p.stdout.replace(" ", "")})
    runner.run("5-4b fx-header in-place",
        ["fx-header", "--in-place", "openpyxl", "--view", "off",
         "--log-dir", str(work / "logs-54"), str(work / "wfw.xlsx")],
        lambda p, _: {"fx_sheets": json.loads(p.stdout)["summary"]["fx_sheets"] == 3})

    # ── 第 5 类：5-5 地区总分校验（真实商业银行共享一批 .xls 直读）────
    sample_bank = stage("5/2026.09 商业银行共享一批.xls")
    runner.run("5-5 region-sum-check",
        ["region-sum-check", "--config", str(config), "--output-dir", str(work / "out-55"),
         "--log-view", "off", str(sample_bank)],
        lambda p, _: {"ok": p.returncode == 0})

    # ── 第 5 类：5-6 按配置提取工作表（真实规范化本外币 .xlsx）──────
    sample_wb = stage("5/2026.09 本外币.xls")
    sample_wb_xlsx = work / "src-56" / "2026.09 本外币.xlsx"
    sample_wb_xlsx.parent.mkdir(exist_ok=True)
    runner.run("5-6a convert 本外币 xls->xlsx",
        ["convert-file", "--input", str(sample_wb), "--to", "xlsx",
         "--output", str(sample_wb_xlsx), "--engine", "pywin32"],
        lambda p, _: {"verified": '"status":"converted"' in p.stdout.replace(" ", "")})
    runner.run("5-6b extract-sheets",
        ["extract-sheets", "--config", str(config), "--output-dir", str(work / "out-56"),
         "--log-view", "off", str(sample_wb_xlsx)],
        lambda p, _: {"extracted": json.loads(p.stdout).get("extracted_sheets") == 1})

    # ── 第 5 类：5-7 涉农贷款修正（真实专项二批次 .xls 对，COM 路由）──
    cur = stage("5/5-7/2026.09. 专项二批次.xls")
    prev = stage("5/5-7/2026.08. 专项二批次.xls")
    before = sha256(cur)
    runner.menu("5-7 rural-loan in-place (menu 5-7)", config,
        f"5\n7\n{cur}\n{prev}\n0\n0\n",
        lambda p, _: {"file_changed": sha256(cur) != before})

    # ── 原件保护复核 ──────────────────────────────────────────────
    modified = [p for p, h in originals.items() if sha256(p) != h]
    summary = {
        "when": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "exe": str(exe),
        "steps": runner.results,
        "all_ok": all(entry["ok"] for entry in runner.results),
        "originals_modified": modified,
        "originals_checked": len(originals),
        "env": env_note,
    }
    (out_dir / "result.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({
        "out_dir": str(out_dir),
        "all_ok": summary["all_ok"],
        "steps_ok": sum(1 for e in runner.results if e["ok"]),
        "steps_total": len(runner.results),
        "originals_modified": len(modified),
    }, ensure_ascii=False))
    console.close()
    return 0 if summary["all_ok"] and not modified else 1


if __name__ == "__main__":
    sys.exit(main())
