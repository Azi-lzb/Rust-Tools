# -*- coding: utf-8 -*-
"""Full dual-engine e2e sweep: dates + wide tables + original 8 scenarios.

Runs each config through both engines (fresh targets each time), compares
outputs cell-by-cell (values/styles/widths/heights/merges/setup) and writes
per-scenario JSON + a summary line. Expects RustTools.exe next to scripts/..

Usage: python -I scripts/run_e2e_sweep.py [fixture_name ...]
"""
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
FIX = HERE.parent / "fixtures"
EXE = HERE.parent / "RustTools.exe"

SCENARIOS = [
    "cfg_dates",
    "cfg_wide186_m1", "cfg_wide186_m2",
    "cfg_wide210_m1", "cfg_wide210_m2",
    "cfg_wide222_m1", "cfg_wide222_m2",
    "cfg_mode1", "cfg_mode2", "cfg_mode3",
    "cfg_multirange", "cfg_emptyrange",
    "cfg_multitarget", "cfg_repeatsheet",
]


def main():
    only = set(sys.argv[1:])
    results_dir = HERE.parent / "results"
    results_dir.mkdir(exist_ok=True)
    summary = []
    for scenario in SCENARIOS:
        if only and scenario not in only:
            continue
        config = FIX / f"{scenario}.xlsx"
        # 清理目标（openpyxl 目标 = out_*.xlsx；Rust 目标 = *_Rust.xlsx）
        for stale in FIX.glob("out_*.xlsx"):
            stale.unlink()
        record = {}
        for engine in ("rust", "openpyxl"):
            proc = subprocess.run(
                [sys.executable, "-I", str(HERE / "run_print_menu.py"),
                 str(EXE), str(config), engine,
                 str(results_dir / "e2e_sweep.jsonl")],
                capture_output=True, text=True, timeout=1800)
            record[engine] = proc.stdout.strip().splitlines()[-1] if proc.stdout.strip() else "?"
        # 对比：目标 = 配置里首个任务目标名。宽表/日期按前缀约定。
        targets = {
            "cfg_dates": ["out_dates"],
            "cfg_mode1": ["out_m1"], "cfg_mode2": ["out_m2"], "cfg_mode3": ["out_m3"],
            "cfg_multirange": ["out_mr"], "cfg_emptyrange": ["out_er"],
            "cfg_multitarget": ["out_mt1", "out_mt2"],
            "cfg_repeatsheet": ["out_rs"],
        }.get(scenario)
        if targets is None:
            if scenario.startswith("cfg_wide"):
                # cfg_wide186_m1 -> out_wide186_m1
                targets = [scenario.replace("cfg_wide", "out_wide")]
            else:
                continue
        for target in targets:
            sink = results_dir / f"compare_{scenario}_{target}.json"
            proc = subprocess.run(
                [sys.executable, "-I", str(HERE / "compare_xlsx.py"),
                 str(FIX / f"{target}.xlsx"), str(FIX / f"{target}_Rust.xlsx"),
                 str(sink)],
                capture_output=True, text=True)
            line = proc.stdout.strip() or proc.stderr.strip()[-160:]
            summary.append(f"{scenario}/{target}: {line}")
            print(f"{scenario}/{target}: {line}")
    (results_dir / "e2e_sweep_summary.txt").write_text(
        "\n".join(summary) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
