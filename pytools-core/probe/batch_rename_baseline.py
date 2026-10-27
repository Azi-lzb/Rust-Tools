"""Run the read-only PyCode 3-3 baseline against the shared synthetic case."""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import platform
import sys
import tempfile
from pathlib import Path


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--source-root",
        type=Path,
        default=Path(r"C:\Users\AZI\Desktop\VibeCode\4.pytools-main"),
    )
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    sys.path.insert(0, str(args.source_root))
    from PyCode.core.convert_33_34_35_37 import run_batch_rename_33

    workspace = Path(__file__).parents[1]
    cases = json.loads(
        (workspace / "tests/fixtures/batch_rename_cases.json").read_text(encoding="utf-8")
    )
    config_path = workspace / "tests/fixtures/batch_rename_config.xlsx"
    config_hash = digest(config_path)

    with tempfile.TemporaryDirectory(prefix="pytools-batch-rename-") as temporary:
        root = Path(temporary)
        input_paths: list[Path] = []
        for entry in cases["inputs"]:
            path = root / entry["name"]
            input_paths.append(path)
            if not entry["exists"]:
                continue
            if entry["kind"] == "directory":
                path.mkdir(parents=True)
                (path / "marker.txt").write_text(entry["content"], encoding="utf-8")
            else:
                path.write_text(entry["content"], encoding="utf-8")
        for entry in cases["reserved"]:
            (root / entry["name"]).write_text(entry["content"], encoding="utf-8")

        result = run_batch_rename_33(config_path, input_paths, root / "logs")
        logger = logging.getLogger("PyCode.3_3_batch_rename")
        for handler in list(logger.handlers):
            handler.flush()
            handler.close()
            logger.removeHandler(handler)
        entries = []
        for path in sorted(root.iterdir(), key=lambda item: item.name):
            if path.name == "logs":
                continue
            entries.append({"name": path.name, "kind": "directory" if path.is_dir() else "file"})
        marker = root / "2026-01-31 0003 丙银行 存款表" / "marker.txt"
        output = {
            "schema_version": 1,
            "suite": "pytools.batch_rename_33",
            "source": "PyCode/core/convert_33_34_35_37.py::run_batch_rename_33",
            "config_fixture": "tests/fixtures/batch_rename_config.xlsx",
            "config_sha256_before": config_hash,
            "config_sha256_after": digest(config_path),
            "python_baseline": {
                "status": "保留",
                "python": platform.python_version(),
                "passed": 1,
                "total": 1,
            },
            "result": {
                "ok": result["ok"],
                "skip": result["skip"],
                "entries": entries,
                "renamed_directory_marker": marker.read_text(encoding="utf-8"),
                "reserved_collision_content": (
                    root / cases["reserved"][0]["name"]
                ).read_text(encoding="utf-8"),
            },
        }
        text = json.dumps(output, ensure_ascii=False, indent=2) + "\n"
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(text, encoding="utf-8")
        print(text, end="")


if __name__ == "__main__":
    main()
