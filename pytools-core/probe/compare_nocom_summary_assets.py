"""Aggregate PyCode/NoComPyTools PT-6-1 differences across original XLSX assets."""

from __future__ import annotations

import datetime as dt
import importlib.metadata
import json
import sys
import tempfile
from collections import Counter
from pathlib import Path

import compare_summary_tools as parity


def compare_projections(pycode_output: Path, nocom_output: Path):
    pycode_sheets, pycode_dimensions = parity.workbook_projection(pycode_output)
    nocom_sheets, nocom_dimensions = parity.workbook_projection(nocom_output)
    differences = []
    cell_counts: Counter[str] = Counter()
    if len(pycode_sheets) != len(nocom_sheets):
        differences.append("worksheet_count")
    for pycode_sheet, nocom_sheet in zip(pycode_sheets, nocom_sheets):
        for field in pycode_sheet:
            if pycode_sheet.get(field) != nocom_sheet.get(field):
                differences.append(f"{pycode_sheet['name']}.{field}")
        for kind in ("values", "types", "styles"):
            pycode_cells = pycode_sheet[kind]
            nocom_cells = nocom_sheet[kind]
            missing = object()
            cell_counts[kind] += sum(
                pycode_value != nocom_cells.get(coordinate, missing)
                for coordinate, pycode_value in pycode_cells.items()
            )
            cell_counts[kind] += sum(
                coordinate not in pycode_cells for coordinate in nocom_cells
            )
    if pycode_dimensions != nocom_dimensions:
        differences.append("worksheet_dimensions")
    return sorted(set(differences)), pycode_dimensions, nocom_dimensions, cell_counts


def main() -> int:
    source_root = parity.PYTOOLS_SOURCE / "xlsx"
    if not source_root.is_dir():
        raise FileNotFoundError(source_root)
    sources = sorted(
        path
        for path in source_root.rglob("*")
        if path.is_file()
        and path.suffix.lower() == ".xlsx"
        and not path.name.startswith("~$")
    )
    if not sources:
        raise RuntimeError("original PyTools xlsx tree has no XLSX inputs")

    source_metadata_before = [
        (path.stat().st_size, path.stat().st_mtime_ns) for path in sources
    ]
    batch_size = 3
    batches = [sources[start : start + batch_size] for start in range(0, len(sources), batch_size)]
    exact_batches = 0
    different_batches = 0
    inputs_in_exact_batches = 0
    inputs_in_different_batches = 0
    pycode_failures = 0
    nocom_failures = 0
    missing_outputs = 0
    difference_fields: Counter[str] = Counter()
    changed_cells: Counter[str] = Counter()
    pycode_summary_rows = 0
    nocom_summary_rows = 0

    with tempfile.TemporaryDirectory(
        prefix="pytools-nocom-summary-assets-", ignore_cleanup_errors=True
    ) as temp:
        root = Path(temp)
        for batch_index, batch in enumerate(batches, 1):
            batch_root = root / f"batch_{batch_index:02}"
            try:
                pycode_output = parity.python_used_range_only(batch, batch_root / "pycode")
            except Exception as error:  # noqa: BLE001 - aggregate failures without source data
                pycode_failures += len(batch)
                print(
                    f"batch {batch_index}/{len(batches)}: PyCode failed ({type(error).__name__})",
                    flush=True,
                )
                continue
            try:
                nocom_output = parity.python_nocom_used_range(batch, batch_root / "nocom")
            except Exception as error:  # noqa: BLE001 - aggregate failures without source data
                nocom_failures += len(batch)
                print(
                    f"batch {batch_index}/{len(batches)}: NoComPyTools failed ({type(error).__name__})",
                    flush=True,
                )
                continue
            if pycode_output is None or nocom_output is None:
                missing_outputs += len(batch)
                print(
                    f"batch {batch_index}/{len(batches)}: output missing",
                    flush=True,
                )
                continue

            fields, pycode_dimensions, nocom_dimensions, batch_cell_counts = compare_projections(
                pycode_output, nocom_output
            )
            pycode_summary_rows += sum(item["rows"] for item in pycode_dimensions)
            nocom_summary_rows += sum(item["rows"] for item in nocom_dimensions)
            if fields:
                different_batches += 1
                inputs_in_different_batches += len(batch)
                difference_fields.update(set(fields))
                changed_cells.update(batch_cell_counts)
            else:
                exact_batches += 1
                inputs_in_exact_batches += len(batch)

            print(
                f"batch {batch_index}/{len(batches)}: compared {len(batch)}; "
                f"{'different' if fields else 'exact'}; "
                f"PyCode rows={sum(item['rows'] for item in pycode_dimensions)}; "
                f"NoCom rows={sum(item['rows'] for item in nocom_dimensions)}",
                flush=True,
            )

    source_metadata_after = [
        (path.stat().st_size, path.stat().st_mtime_ns) for path in sources
    ]
    source_metadata_unchanged = source_metadata_before == source_metadata_after
    if not source_metadata_unchanged:
        raise RuntimeError("source XLSX size or modification time changed during comparison")

    try:
        pycalamine_version = importlib.metadata.version("python-calamine")
    except importlib.metadata.PackageNotFoundError:
        pycalamine_version = "absent"

    compared_batches = exact_batches + different_batches
    report = {
        "schema_version": 1,
        "suite": "pytools.summary_used_range_pycode_nocom_asset_comparison",
        "date": dt.date.today().isoformat(),
        "status": "completed_with_differences"
        if different_batches or pycode_failures or nocom_failures or missing_outputs
        else "completed_exact",
        "source_scope": "all XLSX assets recursively under the original PyTools xlsx directory",
        "input_workbooks": len(sources),
        "comparison_unit": "batch",
        "batch_size": batch_size,
        "batches": len(batches),
        "compared_batches": compared_batches,
        "exact_batches": exact_batches,
        "different_batches": different_batches,
        "inputs_in_exact_batches": inputs_in_exact_batches,
        "inputs_in_different_batches": inputs_in_different_batches,
        "pycode_failed_inputs": pycode_failures,
        "nocom_failed_inputs": nocom_failures,
        "missing_output_inputs": missing_outputs,
        "summary_rows_including_headers": {
            "pycode": pycode_summary_rows,
            "nocom_pytools": nocom_summary_rows,
        },
        "summary_row_delta_pycode_minus_nocom": pycode_summary_rows - nocom_summary_rows,
        "projection_difference_fields_by_batch": dict(sorted(difference_fields.items())),
        "changed_output_cells_by_projection": dict(sorted(changed_cells.items())),
        "source_size_and_mtime_unchanged": source_metadata_unchanged,
        "python_runtime": sys.version.split()[0],
        "pycode_environment": {
            "openpyxl": importlib.metadata.version("openpyxl"),
        },
        "nocom_environment": {
            "openpyxl": importlib.metadata.version("openpyxl"),
            "python_calamine": pycalamine_version,
        },
        "privacy": "No source workbook names, cell values, paths, or hashes are recorded.",
        "scope_note": "This compares the two original Python variants only; it does not change the confirmed PyCode production baseline or infer NoCom-only product scope.",
    }
    report_path = (
        Path(__file__).resolve().parent
        / "reports"
        / f"summary_used_range_nocom_assets_{dt.date.today():%Y%m%d}.json"
    )
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "status": report["status"],
                "input_workbooks": len(sources),
                "exact_batches": exact_batches,
                "different_batches": different_batches,
                "inputs_in_exact_batches": inputs_in_exact_batches,
                "inputs_in_different_batches": inputs_in_different_batches,
                "pycode_failed_inputs": pycode_failures,
                "nocom_failed_inputs": nocom_failures,
                "difference_fields": dict(sorted(difference_fields.items())),
                "report": str(report_path),
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
