from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True
WORKSPACE = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(WORKSPACE / "pytools-core" / "probe"))
import compare_summary_tools as compare


def run_batch(
    sources: list[Path],
    output_id: int,
    temporary_root: Path,
) -> dict[str, object]:
    python_output = compare.python_used_range_only(
        sources,
        temporary_root / f"python_{output_id}",
    )
    if python_output is None:
        raise RuntimeError("python_not_saved")
    rust_output, rust_result = compare.run_rust(
        "summary-used-range",
        sources,
        temporary_root / f"rust_{output_id}",
    )
    if rust_output is None:
        raise RuntimeError("rust_not_saved")
    python_dimensions, rust_dimensions = compare.assert_equivalent(
        python_output,
        rust_output,
        f"PT-6-1 backup batch {output_id}",
    )
    rows = python_dimensions[0]["rows"] - 1
    if (
        rust_result.get("rows") != rows
        or rust_result.get("files") != len(sources)
        or python_dimensions != rust_dimensions
    ):
        raise AssertionError("summary_dimensions_or_counts_differ")
    progress = rust_result.get("progress_log", {})
    if progress.get("status") != "succeeded" or not progress.get(
        "event_stream_matches_saved_history"
    ):
        raise AssertionError("developer_log_mismatch")
    return {
        "status": "passed",
        "input_workbooks": len(sources),
        "summary_rows": rows,
        "summary_columns": python_dimensions[0]["columns"],
        "developer_events": progress.get("event_count"),
        "event_stream_matches_history": True,
        "actual_engines": progress.get("actual_engines", []),
    }


def compare_with_bisection(
    sources: list[Path],
    temporary_root: Path,
    next_output_id: list[int],
) -> list[dict[str, object]]:
    output_id = next_output_id[0]
    next_output_id[0] += 1
    try:
        result = run_batch(sources, output_id, temporary_root)
        return [result]
    except Exception as error:
        if len(sources) == 1:
            return [
                {
                    "status": "failed",
                    "input_workbooks": 1,
                    "failed_at": type(error).__name__,
                }
            ]
        midpoint = len(sources) // 2
        left = compare_with_bisection(
            sources[:midpoint], temporary_root, next_output_id
        )
        right = compare_with_bisection(
            sources[midpoint:], temporary_root, next_output_id
        )
        return left + right


def main() -> int:
    backup_root = compare.PYTOOLS_SOURCE / "xlsx" / "备份"
    all_sources = sorted(
        path
        for path in backup_root.rglob("*.xlsx")
        if not path.name.startswith("~$")
    )
    if len(all_sources) < 3:
        raise RuntimeError("expected_backup_asset_set_missing")
    smallest = min(all_sources, key=lambda path: path.stat().st_size)
    largest = max(all_sources, key=lambda path: path.stat().st_size)
    sources = [path for path in all_sources if path not in (smallest, largest)]
    before = {
        path: (path.stat().st_size, path.stat().st_mtime_ns)
        for path in sources
    }
    status = []
    next_output_id = [1]
    with tempfile.TemporaryDirectory(
        prefix="pytools-summary-backup-parity-",
        ignore_cleanup_errors=True,
    ) as temp:
        temporary_root = Path(temp)
        initial_batches = [
            sources[index : index + 3]
            for index in range(0, len(sources), 3)
        ]
        for batch_number, batch in enumerate(initial_batches, 1):
            results = compare_with_bisection(
                batch,
                temporary_root,
                next_output_id,
            )
            status.extend(results)
            print(
                json.dumps(
                    {
                        "batch": batch_number,
                        "input_workbooks": len(batch),
                        "resolved_groups": len(results),
                        "passed_groups": sum(
                            item["status"] == "passed" for item in results
                        ),
                        "failed_groups": sum(
                            item["status"] != "passed" for item in results
                        ),
                    }
                ),
                flush=True,
            )
    after = {
        path: (path.stat().st_size, path.stat().st_mtime_ns)
        for path in sources
    }
    if before != after:
        raise RuntimeError("source_metadata_changed")
    report = {
        "schema_version": 1,
        "suite": "pytools.summary_used_range_backup_asset_parity",
        "date": "2026-10-24",
        "status": (
            "passed_real_backup_full_projection"
            if all(item["status"] == "passed" for item in status)
            else "partial_real_backup_parity"
        ),
        "input_format": "xlsx",
        "source_scope": "xlsx/backup workbooks excluding the smallest and largest previously covered",
        "previously_validated_endpoints_excluded": 2,
        "source_workbooks": len(sources),
        "source_bytes": sum(path.stat().st_size for path in sources),
        "source_metadata_unchanged": True,
        "initial_batches_of_at_most_three": len(initial_batches),
        "resolved_groups": len(status),
        "passed_groups": sum(item["status"] == "passed" for item in status),
        "failed_groups": sum(item["status"] != "passed" for item in status),
        "summary_rows": sum(
            int(item.get("summary_rows", 0))
            for item in status
            if item["status"] == "passed"
        ),
        "full_projection": (
            "values_rows_types_styles_dimensions_freeze_header_merges_formulas_defined_names_print_area_page_setup_margins_centering"
        ),
        "developer_event_history_matches": all(
            item.get("event_stream_matches_history") is True
            for item in status
            if item["status"] == "passed"
        ),
        "groups": status,
    }
    report_path = (
        WORKSPACE
        / "pytools-core"
        / "probe"
        / "reports"
        / "summary_used_range_backup_assets_20261024.json"
    )
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps({key: value for key, value in report.items() if key != "groups"}))
    return 0 if report["status"] == "passed_real_backup_full_projection" else 1


if __name__ == "__main__":
    raise SystemExit(main())