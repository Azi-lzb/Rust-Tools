"""Compare the 3-4 print-config precheck with the read-only Python baseline."""

from __future__ import annotations

import json
import shlex
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook

sys.dont_write_bytecode = True
WORKSPACE = Path(__file__).resolve().parents[2]
PYTOOLS_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\4.pytools-main")
HEADERS = [
    "是否启用", "打印模式", "源工作簿", "源工作表", "源工作表打印区域", "目标工作簿",
    "目标工作表", "FitToPagesWide", "FitToPagesTall", "打印方向", "水平居中",
    "垂直居中", "不输出批注", "零值不输出", "备注",
]


def wsl_path(path: Path) -> str:
    return f"/mnt/{path.drive.rstrip(':').lower()}/{path.as_posix()[2:].lstrip('/')}"


def write_source(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Name", "Value"])
    ws.append(["sample", 7])
    wb.save(path)


def write_config(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "打印配置"
    ws.append(HEADERS)
    # Zero height falls back to one in the Python precheck; ambiguous direction warns.
    ws.append(["是", 1, "source.xlsx", "Data", "A1:B2", "target.xlsx", "Print", 1, 0, "横向纵向", "否", "否", "否", "否", "warnings"])
    # Negative fit height is truthy in Python and produces a fit warning.
    ws.append(["是", 1, "source.xlsx", "Data", "A1:B2", "target.xlsx", "Print2", 1, -1, "纵向", "否", "否", "否", "否", "negative-fit"])
    # One failure for each validation branch exercised by the fixture.
    ws.append(["是", 9, "source.xlsx", "Data", "A1:B2", "target.xlsx", "Print", 1, 1, "", "否", "否", "否", "否", "mode"])
    ws.append(["是", 1, "source.xlsx", "Missing", "A1:B2", "target.xlsx", "Print", 1, 1, "", "否", "否", "否", "否", "sheet"])
    ws.append(["是", 1, "source.xlsx", "Data", "not-a-range", "target.xlsx", "Print", 1, 1, "", "否", "否", "否", "否", "range"])
    ws.append(["是", 1, "source.xlsx", "Data", "A1:B2", None, "Print", 1, 1, "", "否", "否", "否", "否", "target"])
    # Disabled invalid rows must not affect counts.
    ws.append(["否", 9, None, "", "bad", None, "", 0, 0, "bad", "否", "否", "否", "否", "disabled"])
    wb.save(path)


def rust_counts(config: Path) -> dict[str, object]:
    workspace = wsl_path(WORKSPACE)
    config_wsl = wsl_path(config)
    command = (
        f"cd {shlex.quote(workspace)} && cargo run --quiet -p pytools-core --bin pytools -- "
        f"print-config-precheck --config {shlex.quote(config_wsl)}"
    )
    result = subprocess.run(["wsl", "-e", "sh", "-lc", command], capture_output=True, check=False)
    stdout = result.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for index, char in enumerate(stdout):
        if char == "{":
            try:
                payload = json.loads(stdout[index:])
                return payload
            except json.JSONDecodeError:
                continue
    raise RuntimeError(f"Rust print precheck emitted no JSON (exit={result.returncode}): {stdout}")


def main() -> int:
    sys.dont_write_bytecode = True
    if not PYTOOLS_SOURCE.is_dir():
        raise FileNotFoundError(PYTOOLS_SOURCE)
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.config_xlsx import load_print_tasks
    from core.print_310 import run_print_config_precheck

    with tempfile.TemporaryDirectory(prefix="pytools-print-precheck-", ignore_cleanup_errors=True) as temp:
        root = Path(temp)
        (root / "PyCode").mkdir()
        config = root / "PyCode" / "config.xlsx"
        write_source(root / "source.xlsx")
        write_config(config)
        python = run_print_config_precheck(load_print_tasks(config), root / "python_logs")
        rust = rust_counts(config)
        expected = (python["ok"], python["warn"], python["fail"])
        actual = (int(rust["ok"]), int(rust["warnings"]), int(rust["failed"]))
        if expected != actual:
            raise AssertionError(f"print precheck counts differ: Python={expected}, Rust={actual}; report={rust}")
        report = {
            "status": "passed_synthetic_parity",
            "scope": "3-4 print config precheck only",
            "python_counts": {"ok": expected[0], "warn": expected[1], "fail": expected[2]},
            "rust_counts": {"ok": actual[0], "warn": actual[1], "fail": actual[2]},
            "scenarios": ["valid row with fit height zero", "ambiguous orientation", "invalid mode", "missing source sheet", "invalid A1 range", "empty target path", "disabled invalid row ignored"],
            "limitations": ["Synthetic XLSX config and source only; no real print task exists in the current source config.", "Print output and workbook modification were not executed; ExcelModifier capability evidence remains required.", "XLSM, comment-based print, full configuration print, Windows release build, Kylin/UOS and deb were not validated."],
        }
        reports = WORKSPACE / "pytools-core" / "probe" / "reports"
        reports.mkdir(parents=True, exist_ok=True)
        output = reports / "print_precheck_parity_20261023.json"
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
