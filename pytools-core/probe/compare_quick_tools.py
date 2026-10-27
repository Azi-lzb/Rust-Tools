"""Exercise quick tools with a fake xdg-open, without opening a real app."""

from __future__ import annotations

import json
import os
import shlex
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook

sys.dont_write_bytecode = True
WORKSPACE = Path(__file__).resolve().parents[2]


def wsl_path(path: Path) -> str:
    return f"/mnt/{path.drive.rstrip(':').lower()}/{path.as_posix()[2:].lstrip('/')}"


def make_config(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "全局配置"
    ws.append(["键", "值"])
    ws.append(["输入目录", "input"])
    ws.append(["输出目录", "output"])
    ws.append(["日志目录", "logs"])
    wb.save(path)
    wb.close()


def run_cli(action: str, config: Path, fake_bin: Path, log_path: Path) -> bytes:
    workspace = wsl_path(WORKSPACE)
    command = (
        f"cd {shlex.quote(workspace)} && "
        f"export PATH={shlex.quote(wsl_path(fake_bin))}:$PATH "
        f"PYTOOLS_QUICK_LOG={shlex.quote(wsl_path(log_path))} && "
        f"cargo run --quiet -p pytools-core --bin pytools -- {action} "
        f"--config {shlex.quote(wsl_path(config))}"
    )
    result = subprocess.run(["wsl", "-e", "sh", "-lc", command], capture_output=True, check=False)
    stdout = result.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    if result.returncode != 0:
        raise RuntimeError(f"Rust {action} failed (exit={result.returncode}): {stdout}\n{result.stderr.decode('utf-8', errors='replace')}")
    return result.stdout


def parse_latest(stdout: bytes) -> str | None:
    text = stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for index, char in enumerate(text):
        if char == "{":
            try:
                return json.loads(text[index:])["latest"]
            except json.JSONDecodeError:
                continue
    raise RuntimeError(f"open-latest-output emitted no JSON: {text}")


def assert_opened(log: Path, expected: str) -> None:
    if not log.is_file():
        raise AssertionError("fake xdg-open was not invoked")
    actual = log.read_text(encoding="utf-8").strip()
    if actual != expected:
        raise AssertionError(f"quick tool opened {actual!r}, expected {expected!r}")


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="pytools-quick-", ignore_cleanup_errors=True) as temp:
        root = Path(temp)
        config_dir = root / "PyCode" / "config"
        config_dir.mkdir(parents=True)
        config = config_dir / "config.xlsx"
        make_config(config)
        fake_bin = root / "fake-bin"
        fake_bin.mkdir()
        opener = fake_bin / "xdg-open"
        opener.write_text("#!/bin/sh\nprintf '%s\\n' \"$1\" > \"$PYTOOLS_QUICK_LOG\"\n", encoding="utf-8", newline="\n")
        log = root / "opened-path.txt"
        chmod = subprocess.run(["wsl", "-e", "sh", "-lc", f"chmod +x {shlex.quote(wsl_path(opener))}"], capture_output=True, check=False)
        if chmod.returncode != 0:
            raise RuntimeError(chmod.stderr.decode("utf-8", errors="replace"))

        run_cli("open-config", config, fake_bin, log)
        assert_opened(log, wsl_path(config))

        # Global relative directories resolve from the parent of PyCode.
        output = root / "PyCode" / "output"
        nested = output / "nested"
        nested.mkdir(parents=True)
        older = output / "older.xlsx"
        newest = nested / "newest.xlsx"
        ignored = output / "~$ignored.xlsx"
        older.write_bytes(b"older")
        newest.write_bytes(b"newest")
        ignored.write_bytes(b"ignored")
        os.utime(older, (1_700_000_000, 1_700_000_000))
        os.utime(newest, (1_700_000_100, 1_700_000_100))
        os.utime(ignored, (1_700_000_200, 1_700_000_200))
        log.unlink(missing_ok=True)
        latest_stdout = run_cli("open-latest-output", config, fake_bin, log)
        latest_json = parse_latest(latest_stdout)
        expected_latest = wsl_path(newest)
        if latest_json != expected_latest:
            raise AssertionError(f"latest path JSON differs: {latest_json!r} != {expected_latest!r}")
        assert_opened(log, expected_latest)

        empty = root / "empty-output"
        empty.mkdir()
        log.unlink(missing_ok=True)
        # Point the global setting at an empty output directory for the no-file case.
        wb = Workbook()
        ws = wb.active
        ws.title = "全局配置"
        ws.append(["键", "值"])
        ws.append(["输入目录", "input"])
        ws.append(["输出目录", wsl_path(empty)])
        ws.append(["日志目录", "logs"])
        empty_config = config_dir / "empty-config.xlsx"
        wb.save(empty_config)
        wb.close()
        empty_stdout = run_cli("open-latest-output", empty_config, fake_bin, log)
        if parse_latest(empty_stdout) is not None or log.exists():
            raise AssertionError("empty output directory should return null without opening a path")

        # Open-output-dir creates the configured directory if it does not exist.
        new_output = root / "new-output"
        out_config = config_dir / "output-config.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "全局配置"
        ws.append(["键", "值"])
        ws.append(["输入目录", "input"])
        ws.append(["输出目录", wsl_path(new_output)])
        ws.append(["日志目录", "logs"])
        wb.save(out_config)
        wb.close()
        log.unlink(missing_ok=True)
        run_cli("open-output-dir", out_config, fake_bin, log)
        if not new_output.is_dir():
            raise AssertionError("open-output-dir did not create the configured directory")
        assert_opened(log, wsl_path(new_output))

        report = {
            "status": "passed_synthetic_os_integration",
            "scope": ["9-1 open config", "9-2 open latest output", "9-3 open output directory"],
            "latest_selection": "recursive newest mtime; ignores names beginning with ~$",
            "empty_output": "returns null and does not invoke opener",
            "new_output_directory": "created before opener invocation",
            "system_opener": "fake xdg-open captured arguments; no desktop application was launched",
            "limitations": ["Linux/WSL command path only; Windows Explorer and macOS open were not tested.", "The old Python implementation uses os.startfile and was not launched on the Windows desktop; the probe verifies the requested path selection and Rust opener arguments.", "Real user output directories and target Linux desktop integration were not tested."],
        }
        output = WORKSPACE / "pytools-core" / "probe" / "reports" / "quick_tools_parity_20261023.json"
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
