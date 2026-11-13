"""Finalize this task's local Windows test archive; no external applications."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import zipfile

root = Path(__file__).resolve().parents[2]
evidence = Path(__file__).resolve().parent
archive = Path(sys.argv[1]).resolve()
assert archive.parent == root / "dist"


def digest(data):
    return hashlib.sha256(data).hexdigest()


sources = json.loads((evidence / "source-start.json").read_text(encoding="utf-8-sig"))
for item in sources:
    assert digest((root / item["path"]).read_bytes()) == item["sha256"].lower(), item["path"]
files = {
    "RustTools.exe": root / "dist/RustTools.exe",
    "python/openpyxl_worker.py": root / "crates/excel-editor/python/openpyxl_worker.py",
    "python/pywin32_excel_conversion_worker.py": root / "pytools-core/python/pywin32_excel_conversion_worker.py",
    "config/config.xlsx": root / "config.xlsx",
}
manifest = {
    "package_kind": "manual-acceptance-test",
    "platform": "windows-x86_64",
    "source_base": "fa4ba947a64300bbcac1134529848e34565b5925",
    "build_command": "CARGO_TARGET_DIR=F:/RustBuild/pytools cargo build --release --offline -p pytools-core --bin RustTools",
    "source_sha256": {item["path"]: item["sha256"].lower() for item in sources},
    "artifact_sha256": {name: digest(path.read_bytes()) for name, path in files.items()},
    "existing_xls_append": "SAFELY_REFUSED_NOT_RESTORED",
    "linux_route_changed": False,
    "known_limit": "WPS BIFF strict byte-invariance previously FAILED; platform/user acceptance incomplete",
}
# Windows PowerShell Compress-Archive uses backslashes in both ZIP headers.
# Python normalizes the central names when appending, leaving the old local
# headers inconsistent. Rewrite a NEW archive with matching slash names instead.
normalized = archive.with_suffix(".normalized.zip")
assert not normalized.exists()
with zipfile.ZipFile(archive) as original:
    assert original.testzip() is None, "original archive is invalid"
    assert "BUILD_INFO.json" not in original.namelist(), "archive already finalized"
    with zipfile.ZipFile(normalized, "x", compression=zipfile.ZIP_DEFLATED) as package:
        for entry in original.infolist():
            name = entry.filename.replace("\\", "/")
            package.writestr(name, original.read(entry))
        package.writestr("BUILD_INFO.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        package.write(root / "dist/doc/TESTING_NOTES.md", "doc/TESTING_NOTES.md")
with zipfile.ZipFile(normalized) as package:
    assert package.testzip() is None
os.replace(normalized, archive)
with zipfile.ZipFile(archive) as package:
    assert package.testzip() is None
    for name, path in files.items():
        assert package.read(name) == path.read_bytes(), name

# Test the real extracted root layout from an unrelated working directory.
with tempfile.TemporaryDirectory(prefix="rusttools-package-start-") as folder:
    extracted = Path(folder) / "extract"
    with zipfile.ZipFile(archive) as package:
        for name in package.namelist():
            assert (extracted / name).resolve().is_relative_to(extracted.resolve())
        package.extractall(extracted)
    result = subprocess.run(
        [str(extracted / "RustTools.exe"), "menu", "--log-view", "off"],
        input="0\n", cwd=folder, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=90,
    )
    (evidence / "extracted-start.stdout.txt").write_text(result.stdout, encoding="utf-8")
    (evidence / "extracted-start.stderr.txt").write_text(result.stderr, encoding="utf-8")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "数据提取与汇总" in result.stdout, result.stdout

(evidence / "BUILD_INFO.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print("PASS ZIP integrity and EXE/Worker/config byte comparisons")
print("PASS extracted menu startup and config discovery from unrelated cwd")
print("ZIP SHA256", digest(archive.read_bytes()))
