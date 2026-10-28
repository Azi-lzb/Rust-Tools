# -*- coding: utf-8 -*-
"""Windows ZIP 发布包验证脚本。

对一个已生成的 pytools-windows-x64-*.zip 执行解压后验证：

1. 解压到隔离临时目录；
2. 解压布局中的 bundled Python + openpyxl_worker.py 完成 --version 协议握手；
3. bundled 运行时实际 import openpyxl / et_xmlfile；
4. 用 bundled openpyxl 生成合成工作簿（仅作为测试夹具）；
5. 通过解压出的 pytools.exe `apply-edit-plan` 以合成 ExcelEditPlan 执行一次
   受限修改（openpyxl 引擎路由 + PlanOutputValidator 回读校验都在产品内完成）；
6. 校验输出与输入快照。

用法:
    python packaging/openpyxl/verify-windows-zip.py <path-to-zip>
    # 省略路径时自动选择 dist 下最新的 pytools-windows-x64-*.zip
退出码 0 表示全部通过；任何一步失败都以非零退出并打印原因。
"""

from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]


def fail(message: str) -> "NoReturn":
    print(f"verify-windows-zip: FAIL {message}", file=sys.stderr)
    raise SystemExit(1)


def find_zip(explicit: str | None) -> Path:
    if explicit:
        archive = Path(explicit)
        if not archive.is_file():
            fail(f"指定的 ZIP 不存在：{archive}")
        return archive
    candidates = sorted((REPO_ROOT / "dist").glob("pytools-windows-x64-*.zip"))
    if not candidates:
        fail("dist 下没有 pytools-windows-x64-*.zip；请先运行 package.bat")
    return candidates[-1]


def run(command: list[str], *, cwd: Path | None = None) -> subprocess.CompletedProcess:
    result = subprocess.run(
        command, cwd=cwd, capture_output=True, text=True, timeout=300,
        encoding="utf-8", errors="replace",
    )
    return result


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> None:
    archive = find_zip(sys.argv[1] if len(sys.argv) > 1 else None)
    print(f"verify archive: {archive}")
    work = Path(tempfile.mkdtemp(prefix="pytools-zip-verify-"))
    try:
        extract = work / "extract"
        with zipfile.ZipFile(archive) as bundle:
            names = bundle.namelist()
            bundle.extractall(extract)
        required = ["dist/pytools.exe", "dist/python/python.exe",
                    "dist/python/openpyxl_worker.py"]
        for name in required:
            if name not in names:
                fail(f"ZIP 缺少 {name}")
        dist = extract / "dist"
        python_exe = dist / "python/python.exe"
        worker = dist / "python/openpyxl_worker.py"
        pytools = dist / "pytools.exe"

        # 1) 握手
        probe = run([str(python_exe), "-I", str(worker), "--version"])
        if probe.returncode != 0 or "protocol=1" not in probe.stdout:
            fail(f"解压后 worker --version 握手失败：{probe.stdout!r} {probe.stderr!r}")
        print("PASS worker --version handshake (protocol=1)")
        if "openpyxl=missing" in probe.stdout:
            fail("解压后运行时缺少 openpyxl")

        # 2) import openpyxl
        check = run([str(python_exe), "-I", "-c",
                     "import openpyxl, et_xmlfile; print(openpyxl.__version__)"])
        if check.returncode != 0:
            fail(f"解压后 import openpyxl 失败：{check.stderr.strip()[:300]}")
        print(f"PASS import openpyxl ({check.stdout.strip()})")

        # 3) 合成工作簿（夹具）+ 计划 + 能力记录
        fixture = work / "fixture.xlsx"
        make = run([str(python_exe), "-I", "-c",
                    "import openpyxl, sys\n"
                    "wb = openpyxl.Workbook()\n"
                    "ws = wb.active\n"
                    "ws.title = 'Data'\n"
                    "ws['A1'] = 'original'\n"
                    "ws['B1'] = 7\n"
                    "wb.save(sys.argv[1])\n",
                    str(fixture)])
        if make.returncode != 0:
            fail(f"生成合成工作簿失败：{make.stderr.strip()[:300]}")
        before_hash = sha256(fixture)

        output = work / "result.xlsx"
        plan = {
            "request_id": "zip-verify-1",
            "input": str(fixture),
            "output": str(output),
            "format": "xlsx",
            "workbook": {
                "format": "xlsx",
                "inventory_complete": True,
                "features_present": ["relationships", "defined_names"],
            },
            "operations": [
                {"operation": "set_cell_value", "sheet": "Data", "cell": "A1",
                 "value": {"type": "text", "value": "packaged"}},
                {"operation": "set_column_width", "sheet": "Data",
                 "column_index": 1, "width": 22.0},
            ],
            "preserve": ["cell_values_and_types"],
            "drop_allowed": {},
        }
        capabilities = [{
            "engine": "openpyxl_worker",
            "status": "passed",
            "format": "xlsx",
            "operations": ["set_cell_value", "set_column_width"],
            "supported_features": ["relationships", "defined_names"],
            "preserved": ["cell_values_and_types"],
            "platform": "windows-x86_64",
            "evidence": {
                "sample_id": "zip-verify",
                "report_path": "packaging/openpyxl/verify-windows-zip.py",
                "adapter_version": "1",
                "engine_version": "openpyxl",
                "platform": "windows-x86_64",
            },
        }]
        plan_path = work / "plan.json"
        capability_path = work / "capabilities.json"
        plan_path.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
        capability_path.write_text(json.dumps(capabilities), encoding="utf-8")

        # 4) 通过解压出的 pytools.exe 执行（openpyxl 引擎路由 + 回读校验）
        applied = run([
            str(pytools), "apply-edit-plan",
            "--plan", str(plan_path),
            "--capabilities", str(capability_path),
            "--platform", "windows-x86_64",
        ])
        if applied.returncode != 0:
            fail(f"apply-edit-plan 失败：{(applied.stdout + applied.stderr).strip()[:600]}")
        report = json.loads(applied.stdout)
        if report.get("engine") != "openpyxl_worker":
            fail(f"执行引擎不是 openpyxl_worker：{report.get('engine')}")
        print("PASS apply-edit-plan via packaged pytools.exe (openpyxl_worker)")

        # 5) 输出回读 + 输入快照
        if not output.is_file():
            fail("输出工作簿不存在")
        readback = run([str(python_exe), "-I", "-c",
                        "import openpyxl, sys\n"
                        "wb = openpyxl.load_workbook(sys.argv[1])\n"
                        "ws = wb['Data']\n"
                        "print(ws['A1'].value, '|', ws.column_dimensions['A'].width)\n",
                        str(output)])
        if readback.returncode != 0 or "packaged" not in readback.stdout:
            fail(f"输出回读失败：{readback.stdout!r} {readback.stderr.strip()[:300]}")
        print(f"PASS output readback ({readback.stdout.strip()})")
        if sha256(fixture) != before_hash:
            fail("输入工作簿被修改（必须保持只读）")
        print("PASS input snapshot unchanged")

        print("verify-windows-zip: ALL PASS")
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
