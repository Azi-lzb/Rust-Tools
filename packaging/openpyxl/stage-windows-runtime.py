# -*- coding: utf-8 -*-
"""Windows openpyxl 运行时打包脚本。

把 CPython embeddable 发行版 + openpyxl + et-xmlfile + worker 脚本组装到
`python/` 目录，随 pytools.exe 一起分发。运行时布局：

    <dist>/pytools.exe
    <dist>/python/python.exe
    <dist>/python/python312._pth          (启用 site-packages)
    <dist>/python/Lib/site-packages/openpyxl/...
    <dist>/python/Lib/site-packages/et_xmlfile/...
    <dist>/python/openpyxl_worker.py

`inplace_edit::openpyxl_runtime_paths` 会优先识别 `<exe_dir>/python/` 布局。

用法:
    python packaging/openpyxl/stage-windows-runtime.py --dist dist
    # 可选 --runtime <已解压的 embeddable 目录>，默认自动下载
"""

from __future__ import annotations

import argparse
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_VERSION = (Path(__file__).resolve().parent / "python-version-windows.txt")
DEFAULT_VERSION = "3.12.10"
OPENPYXL_VERSION = "3.1.5"
ET_XMLFILE_VERSION = "2.0.0"
PYWIN32_VERSION = "312"
WINDOWS_REQUIREMENTS = Path(__file__).resolve().parent / "requirements-windows.txt"
OPENPYXL_WORKER_SOURCE = REPO_ROOT / "crates" / "excel-editor" / "python" / "openpyxl_worker.py"
PYWIN32_WORKER_SOURCE = REPO_ROOT / "pytools-core" / "python" / "pywin32_excel_conversion_worker.py"


def fail(message: str) -> "NoReturn":
    print(f"openpyxl windows runtime staging: {message}", file=sys.stderr)
    raise SystemExit(1)


def pinned_version() -> str:
    if PYTHON_VERSION.is_file():
        value = PYTHON_VERSION.read_text(encoding="utf-8").strip()
        if value:
            return value
    return DEFAULT_VERSION


def download_embeddable(version: str, workdir: Path) -> Path:
    url = (
        f"https://www.python.org/ftp/python/{version}/"
        f"python-{version}-embed-amd64.zip"
    )
    archive = workdir / f"python-{version}-embed-amd64.zip"
    print(f"download {url}")
    with urllib.request.urlopen(url, timeout=300) as response:
        archive.write_bytes(response.read())
    target = workdir / "runtime"
    with zipfile.ZipFile(archive) as bundle:
        bundle.extractall(target)
    return target


def install_packages(runtime: Path, workdir: Path) -> None:
    """下载 Windows x64 wheels，并由 pip 正确安装到 embeddable runtime。

    pywin32 wheel 带有 platlib 数据目录，需要 pip 安装而不能简单解压。
    安装目标固定为产品使用的 CPython 3.12 x64，不依赖暂存主机位数。
    """
    site_packages = runtime / "Lib" / "site-packages"
    site_packages.mkdir(parents=True, exist_ok=True)
    wheels = workdir / "wheels"
    wheels.mkdir(exist_ok=True)
    if not WINDOWS_REQUIREMENTS.is_file():
        fail(f"Windows 依赖清单缺失：{WINDOWS_REQUIREMENTS}")
    target_tags = [
        "--platform", "win_amd64",
        "--python-version", "3.12",
        "--implementation", "cp",
        "--abi", "cp312",
    ]
    subprocess.run(
        [
            sys.executable,
            "-m",
            "pip",
            "download",
            "--only-binary=:all:",
            "--no-deps",
            "--dest",
            str(wheels),
            *target_tags,
            "-r",
            str(WINDOWS_REQUIREMENTS),
        ],
        check=True,
    )
    subprocess.run(
        [
            sys.executable,
            "-m",
            "pip",
            "install",
            "--only-binary=:all:",
            "--no-deps",
            "--upgrade",
            "--no-index",
            "--find-links",
            str(wheels),
            "--target",
            str(site_packages),
            *target_tags,
            "-r",
            str(WINDOWS_REQUIREMENTS),
        ],
        check=True,
    )
    for package in ("openpyxl", "et_xmlfile", "win32com"):
        if not (site_packages / package).is_dir():
            fail(f"包未就位：{package}")


def enable_site_packages(runtime: Path) -> None:
    """embeddable 发行版默认隔离 site；_pth 中加入 Lib\\site-packages。"""
    pth_files = sorted(runtime.glob("python*._pth"))
    if not pth_files:
        fail("找不到 python*._pth（embeddable 发行版结构异常）")
    pth = pth_files[0]
    lines = pth.read_text(encoding="utf-8").splitlines()
    entry = "Lib\\site-packages"
    if entry not in lines:
        # 必须放在 `import site` 之前，否则 site 已按旧路径初始化。
        for index, line in enumerate(lines):
            if line.strip() == "import site":
                lines.insert(index, entry)
                break
        else:
            lines.extend([entry, "import site"])
    pth.write_text("\n".join(lines) + "\n", encoding="utf-8")


def verify(runtime: Path) -> None:
    python = runtime / "python.exe"
    if not python.is_file():
        fail(f"缺少解释器：{python}")
    worker = runtime / "openpyxl_worker.py"
    if not worker.is_file():
        fail(f"缺少 worker：{worker}")
    result = subprocess.run(
        [
            str(python),
            "-I",
            "-c",
            "import importlib.metadata, openpyxl, et_xmlfile, pythoncom, win32com.client; "
            "print(openpyxl.__version__ + '|' + et_xmlfile.__version__ + '|' + importlib.metadata.version('pywin32'))",
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0:
        fail(f"runtime 自检失败：{result.stderr.strip()[:400]}")
    actual = result.stdout.strip()
    expected = f"{OPENPYXL_VERSION}|{ET_XMLFILE_VERSION}|{PYWIN32_VERSION}"
    if actual != expected:
        fail(f"Python 依赖版本不符：{actual} != {expected}")
    for worker in (
        runtime / "openpyxl_worker.py",
        runtime / "pywin32_excel_conversion_worker.py",
    ):
        if not worker.is_file():
            fail(f"worker 未暂存：{worker}")
    worker_version = subprocess.run(
        [str(python), "-I", str(runtime / "pywin32_excel_conversion_worker.py"), "--version"],
        capture_output=True,
        text=True,
        timeout=120,
    )
    worker_source = PYWIN32_WORKER_SOURCE.read_text(encoding="utf-8")
    protocol_match = re.search(r"^PROTOCOL_VERSION\s*=\s*(\d+)\s*$", worker_source, re.MULTILINE)
    if protocol_match is None:
        fail("pywin32 worker 源码未声明 PROTOCOL_VERSION")
    expected_worker_version = (
        f"protocol={protocol_match.group(1)} pywin32={PYWIN32_VERSION}"
    )
    if worker_version.returncode != 0 or worker_version.stdout.strip() != expected_worker_version:
        fail(f"pywin32 worker 握手失败：{worker_version.stdout.strip()} {worker_version.stderr.strip()[:400]}")
    openpyxl_help = subprocess.run(
        [str(python), "-I", str(runtime / "openpyxl_worker.py"), "--help"],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if openpyxl_help.returncode != 0:
        fail(f"openpyxl worker 启动失败：{openpyxl_help.stderr.strip()[:400]}")


def stage(dist: Path, runtime_source: Path | None) -> None:
    for worker in (OPENPYXL_WORKER_SOURCE, PYWIN32_WORKER_SOURCE):
        if not worker.is_file():
            fail(f"worker 源码缺失：{worker}")
    target = dist / "python"
    stage_root = Path(tempfile.mkdtemp(prefix=".python-stage-", dir=dist))
    staged = stage_root / "python"
    backup = dist / f".python-backup-{os.getpid()}"
    try:
        with tempfile.TemporaryDirectory() as temporary:
            workdir = Path(temporary)
            if runtime_source is not None:
                source = runtime_source.resolve()
                if not (source / "python.exe").is_file():
                    fail(f"--runtime 目录缺少 python.exe：{source}")
                shutil.copytree(source, staged)
                install_packages(staged, workdir)
            else:
                runtime = download_embeddable(pinned_version(), workdir)
                shutil.copytree(runtime, staged)
                install_packages(staged, workdir)
            enable_site_packages(staged)
            shutil.copy2(OPENPYXL_WORKER_SOURCE, staged / "openpyxl_worker.py")
            shutil.copy2(PYWIN32_WORKER_SOURCE, staged / "pywin32_excel_conversion_worker.py")
            verify(staged)

        # Keep the current runnable runtime intact until the new tree passes
        # all checks. If publishing fails, restore the previous tree.
        if backup.exists():
            fail(f"发现未清理的运行时备份，拒绝覆盖：{backup}")
        if target.exists():
            os.replace(target, backup)
        try:
            os.replace(staged, target)
        except Exception:
            if backup.exists() and not target.exists():
                os.replace(backup, target)
            raise
        if backup.exists():
            shutil.rmtree(backup)
    finally:
        shutil.rmtree(stage_root, ignore_errors=True)

    digest = hashlib.sha256((target / "python.exe").read_bytes()).hexdigest()[:16]
    size = sum(item.stat().st_size for item in target.rglob("*") if item.is_file())
    print(f"staged openpyxl runtime -> {target}")
    print(f"  python.exe sha256[:16]={digest}  total={size / 1024 / 1024:.1f} MiB")


def main() -> None:
    parser = argparse.ArgumentParser(description="Stage the Windows openpyxl runtime")
    parser.add_argument("--dist", default="dist", help="目标发布目录（默认 dist）")
    parser.add_argument("--runtime", default=None, help="已解压的 embeddable 目录（离线）")
    args = parser.parse_args()
    dist = (REPO_ROOT / args.dist).resolve()
    dist.mkdir(parents=True, exist_ok=True)
    stage(dist, Path(args.runtime) if args.runtime else None)


if __name__ == "__main__":
    main()
