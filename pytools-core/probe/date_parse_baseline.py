#!/usr/bin/env python3
"""Test-only baseline runner for the original py_statistics_tools date parser."""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True)
    parser.add_argument(
        "--cases",
        default=str(
            Path(__file__).resolve().parents[1]
            / "tests"
            / "fixtures"
            / "date_parse_cases.json"
        ),
    )
    parser.add_argument(
        "--output",
        default=str(Path(__file__).resolve().parent / "reports" / "date_parse_baseline.json"),
    )
    args = parser.parse_args()

    sys.path.insert(0, str(Path(args.source_root).resolve()))
    from PyCode.core.date_parse import normalize_data_date_value, parse_data_date

    vectors = json.loads(Path(args.cases).read_text(encoding="utf-8"))
    results = []

    for case in vectors["normalize"]:
        value = case["value"]
        if isinstance(value, dict):
            parsed = datetime.fromisoformat(value["value"])
            value = parsed.date() if value["kind"] == "date" else parsed
        actual = normalize_data_date_value(value)
        actual_text = actual.isoformat() if actual is not None else None
        expected = case["expected"]
        results.append(
            {
                "id": case["id"],
                "operation": "normalize",
                "actual": actual_text,
                "expected": expected,
                "status": "保留" if actual_text == expected else "损坏",
            }
        )

    for case in vectors["parse"]:
        actual_date, actual_source = parse_data_date(
            case["file_name"], case["sheet_a1"]
        )
        actual = {
            "date": actual_date or None,
            "source": actual_source or None,
        }
        expected = {
            "date": case["expected_date"],
            "source": case["expected_source"],
        }
        results.append(
            {
                "id": case["id"],
                "operation": "parse",
                "actual": actual,
                "expected": expected,
                "status": "保留" if actual == expected else "损坏",
            }
        )

    failures = [item for item in results if item["status"] != "保留"]
    report = {
        "schema_version": 1,
        "suite": "pytools.date_parse",
        "baseline": "original PyCode.core.date_parse",
        "status": "损坏" if failures else "保留",
        "passed": len(results) - len(failures),
        "total": len(results),
        "cases": results,
    }
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps({key: report[key] for key in ("suite", "status", "passed", "total")}, ensure_ascii=False))
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
