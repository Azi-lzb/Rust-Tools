"""Parity probe for Unicode case mapping in PT-6-2 set comments."""

from __future__ import annotations

import contextlib
import json
import logging
import os
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

sys.dont_write_bytecode = True
WORKSPACE = Path(__file__).resolve().parents[2]
PYTOOLS_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\4.pytools-main")
sys.path.insert(0, str(Path(__file__).resolve().parent))

from compare_summary_tools import (  # noqa: E402
    assert_equivalent,
    make_comment_template,
    make_source,
    run_rust_comments,
    workbook_projection,
)


def run_python_comments(template: Path, source: Path, output: Path) -> tuple[Path | None, dict[str, object]]:
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.summary_321_322 import run_summary_by_comment

    output.mkdir(parents=True, exist_ok=True)
    with contextlib.redirect_stdout(None):
        result = run_summary_by_comment(
            template, [source], output / "python_comments", output / "python_logs"
        )
    logging.shutdown()
    return (Path(result["path"]) if result.get("saved") else None), result


def main() -> None:
    markers = ("İİset(区域)", "ſet(区域)")
    rust_cli = os.environ.get("PYTOOLS_RUST_CLI")
    if not rust_cli or not Path(rust_cli).is_file():
        raise RuntimeError("set PYTOOLS_RUST_CLI to the built Windows pytools.exe")

    with tempfile.TemporaryDirectory(prefix="pytools_pt6_2_unicode_set_") as temporary:
        root = Path(temporary)
        cases = []
        for index, marker in enumerate(markers):
            case_root = root / str(index)
            case_root.mkdir()
            template = case_root / "template.xlsx"
            source = case_root / "source.xlsx"
            make_comment_template(template)
            workbook = load_workbook(template)
            workbook["模板"]["D1"].comment.text = marker
            workbook.save(template)
            workbook.close()
            make_source(source)
            template_before = template.read_bytes()
            source_before = source.read_bytes()

            python_output, python_report = run_python_comments(
                template, source, case_root / "python"
            )
            rust_output, rust_report = run_rust_comments(
                template, source, case_root / "rust"
            )
            assert_equivalent(
                python_output, rust_output, f"PT-6-2 Unicode set marker {marker}"
            )
            if python_output is None or rust_output is None:
                raise AssertionError("Python and Rust must both save a summary")

            python_sheets, dimensions = workbook_projection(python_output)
            rust_sheets, _ = workbook_projection(rust_output)
            if template.read_bytes() != template_before:
                raise AssertionError("synthetic template changed during read-only summary")
            if source.read_bytes() != source_before:
                raise AssertionError("synthetic source changed during summary")
            cases.append(
                {
                    "marker": marker,
                    "summary_rows": rust_report.get("rows"),
                    "python_saved": bool(python_report.get("saved")),
                    "rust_saved": bool(rust_report.get("saved")),
                    "workbook_projection": "exact",
                    "python_sheet_count": len(python_sheets),
                    "rust_sheet_count": len(rust_sheets),
                    "python_dimensions": dimensions,
                    "template_and_source_unchanged": True,
                }
            )

        report = {
            "feature": "PT-6-2",
            "status": "passed_synthetic_python_rust_cli_parity",
            "scenarios": [
                "Unicode İ lowercasing before ASCII set marker",
                "Python re.IGNORECASE long-s matching",
            ],
            "python_runtime": sys.version.split()[0],
            "rust_entry": "pytools summary-comments",
            "cases": cases,
            "limitations": [
                "Synthetic workbook only; no authorized real annotated template/source sample was available.",
                "Does not complete PT-6-2 real-sample acceptance or target Linux/deb validation.",
            ],
        }
        report_path = WORKSPACE / "pytools-core/probe/reports/summary_comments_unicode_set_marker_parity_20261026.json"
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print(json.dumps({"report": str(report_path), **report}, ensure_ascii=True))


if __name__ == "__main__":
    main()
