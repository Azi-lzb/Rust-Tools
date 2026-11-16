"""Confirm safe rejection of no-overwrite rename on the Windows-mounted WSL volume."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys

import openpyxl

base = Path("F:/RustBuild/pytools/release-push-20261116")
root = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else base / "drvfs-safety-probe"
assert root.is_relative_to(base)
root.mkdir(exist_ok=False)
original_exe = Path("F:/RustBuild/pytools/x86_64-unknown-linux-gnu/debug/RustTools")
exe = root / "RustTools"
shutil.copyfile(original_exe, exe)
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
assert sha(exe) == sha(original_exe)
source, target, config = root / "original.xlsx", root / "A_original.xlsx", root / "config.xlsx"
book = openpyxl.Workbook()
book.active.title = "Data"
book.active["A1"] = "A"
book.save(source)
book.close()
target.write_bytes(b"preexisting destination must remain unchanged")
linux = lambda path: "/mnt/f/" + path.as_posix().split(":/", 1)[1]
book = openpyxl.Workbook()
sheet = book.active
sheet.title = "全局配置"
sheet.append(["配置项", "值"])
for key, value in [("输入目录", linux(root)), ("输出目录", linux(root / "output")), ("日志目录", linux(root / "logs"))]:
    sheet.append([key, value])
sheet = book.create_sheet("重命名配置")
sheet.cell(2, 15, "Y")
sheet.cell(2, 16, "Data@A1")
book.save(config)
book.close()
before = {str(path): sha(path) for path in [source, target, config]}
result = subprocess.run(["wsl", "-d", "Ubuntu", "-u", "root", "--", linux(exe), "menu", "--config", linux(config), "--log-view", "developer"],
                        input=f"4\n3\n1\n1\n{linux(source)}\n0\n0\n", capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
(root / "stdout.txt").write_text(result.stdout, encoding="utf-8")
(root / "stderr.txt").write_text(result.stderr, encoding="utf-8")
assert result.returncode == 0, result.stderr
assert all(sha(Path(path)) == digest for path, digest in before.items())
assert not list(root.glob("A_original_*.xlsx"))
diagnostics = result.stdout + result.stderr
for path in root.rglob("*"):
    if path.suffix in {".json", ".jsonl"}:
        diagnostics += path.read_text(encoding="utf-8", errors="replace")
assert "FILE_RENAME_FAILED" in diagnostics, diagnostics[-3000:]
script = Path(__file__).with_name("probe-drvfs-syscall.py").resolve()
linux_script = "/mnt/c/" + script.as_posix().split(":/", 1)[1]
syscall = subprocess.run(["wsl", "-d", "Ubuntu", "-u", "root", "--", "python3", linux_script, linux(root)], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30)
(root / "syscall.stdout.txt").write_text(syscall.stdout, encoding="utf-8")
(root / "syscall.stderr.txt").write_text(syscall.stderr, encoding="utf-8")
assert syscall.returncode == 0, syscall.stderr
raw_syscall = json.loads(syscall.stdout)
record = dict(status="KNOWN_LIMITATION_SAFE_REJECTION", filesystem="WSL DrvFS on Windows F drive", source_unchanged=True,
              existing_target_unchanged=True, configuration_unchanged=True, no_new_renamed_file=True, error="EINVAL / no-overwrite rename unavailable", syscall=raw_syscall, hashes=before)
(root / "result.json").write_text(json.dumps(record, indent=2), encoding="utf-8")
print(json.dumps(record))
