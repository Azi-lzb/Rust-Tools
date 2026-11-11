# -*- coding: utf-8 -*-
"""Drive RustTools.exe through menu 3 (报表打印) for one config × engine.

Records wall time, exit code, and optional stderr stage-timing lines
(PYTOOLS_PRINT_STAGE_TIMING=1). One JSON record per run, appended to sink.

Usage:
  python -I scripts/run_print_menu.py <exe> <config> <openpyxl|rust> <sink.jsonl>
"""
import json
import os
import subprocess
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parents[3]
DIST_PYTHON = REPO / "dist" / "python" / "python.exe"

INPUT_BY_ENGINE = {
    "openpyxl": "3\n1\n0\n0\n",
    "rust": "3\n2\n0\n0\n",
}


def main():
    exe, config, engine, sink = sys.argv[1], Path(sys.argv[2]), sys.argv[3], Path(sys.argv[4])
    # 继承完整父环境（SystemRoot/TEMP/PATH 必须保留）：
    # 子进程缺 TEMP 时 unique_temp_json 写临时计划失败，openpyxl 引擎会快速失败。
    env = dict(os.environ)
    env.update({
        "PYTOOLS_CONFIG_PATH": str(config.resolve()),
        # 3-1 Openpyxl 引擎需要随包 Python（见 static_print_engines::resolve_print_python）
        "PYTOOLS_OPENPYXL_PYTHON": str(DIST_PYTHON),
        "PYTOOLS_OPENPYXL_WORKER": str(REPO / "dist" / "python" / "openpyxl_worker.py"),
        "PYTOOLS_PRINT_STAGE_TIMING": "1",
    })
    payload = INPUT_BY_ENGINE[engine].encode("gbk", errors="replace")
    started = time.perf_counter()
    proc = subprocess.run(
        [str(exe)],
        input=payload,
        capture_output=True,
        env=env,
        cwd=str(exe and Path(exe).parent),
        timeout=1800,
    )
    wall_ms = int((time.perf_counter() - started) * 1000)
    stdout = proc.stdout.decode("gbk", errors="replace")
    stderr = proc.stderr.decode("gbk", errors="replace")
    stages = []
    for line in stderr.splitlines():
        if "[3-1][stage-timing]" in line:
            try:
                stages.append(json.loads(line.split("[3-1][stage-timing] ", 1)[1]))
            except json.JSONDecodeError:
                pass
    tail = "\n".join(stdout.strip().splitlines()[-6:])
    record = {
        "config": config.name,
        "engine": engine,
        "exit": proc.returncode,
        "wall_ms": wall_ms,
        "stages": stages,
        "stdout_tail": tail,
        "stderr_tail": stderr.strip().splitlines()[-3:] if stderr.strip() else [],
    }
    with open(sink, "a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False) + "\n")
    print(engine, config.name, "exit=", proc.returncode, f"wall={wall_ms}ms",
          "stages=" + str(len(stages)))


if __name__ == "__main__":
    main()
