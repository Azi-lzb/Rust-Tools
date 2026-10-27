#!/usr/bin/env python3
"""Test-only baseline runner for original py_statistics_tools path/config helpers."""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path


def python_value(value):
    if isinstance(value, dict):
        kind = value.get("kind")
        if kind == "nan":
            return float("nan")
        if kind == "integer":
            return int(value["value"])
        if kind == "number":
            return float(value["value"])
    return value


def path_map_rule(PathMapRule, values):
    defaults = {
        "enabled": True,
        "name": "map",
        "applicable_rule": "",
        "wb_keyword": "",
        "sheet_keyword": "",
        "target": "两者",
        "match_mode": "精确",
        "original": "旧路径",
        "standard": "新路径",
    }
    defaults.update(values)
    return PathMapRule(**defaults)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True)
    parser.add_argument(
        "--cases",
        default=str(
            Path(__file__).resolve().parents[1]
            / "tests"
            / "fixtures"
            / "path_map_cases.json"
        ),
    )
    parser.add_argument(
        "--output",
        default=str(Path(__file__).resolve().parent / "reports" / "path_map_baseline.json"),
    )
    args = parser.parse_args()

    sys.path.insert(0, str(Path(args.source_root).resolve()))
    from PyCode.core.config_xlsx import PathMapRule
    from PyCode.core.path_map import _match_kw, standardize
    from PyCode.core.rule_engine import match_all_keywords, parse_col_spec, parse_col_specs

    vectors = json.loads(Path(args.cases).read_text(encoding="utf-8"))
    results = []

    for case in vectors["keywords"]:
        source = case["source"]
        keywords = case["keywords"]
        actual = {
            "path_map": _match_kw(keywords, source),
            "rule_engine": match_all_keywords(source, keywords),
        }
        expected = {"path_map": case["expected"], "rule_engine": case["expected"]}
        results.append(
            {
                "id": case["id"],
                "operation": "keywords",
                "actual": actual,
                "expected": expected,
                "status": "保留" if actual == expected else "损坏",
            }
        )

    for case in vectors["columns_single"]:
        actual = parse_col_spec(python_value(case["value"]))
        expected = case["expected"]
        results.append(
            {
                "id": case["id"],
                "operation": "parse_col_spec",
                "actual": actual,
                "expected": expected,
                "status": "保留" if actual == expected else "损坏",
            }
        )

    for case in vectors["columns_multiple"]:
        actual = parse_col_specs(python_value(case["value"]))
        expected = case["expected"]
        results.append(
            {
                "id": case["id"],
                "operation": "parse_col_specs",
                "actual": actual,
                "expected": expected,
                "status": "保留" if actual == expected else "损坏",
            }
        )

    for case in vectors["standardize"]:
        maps = [path_map_rule(PathMapRule, item) for item in case["maps"]]
        actual = standardize(
            case["path"],
            case["target"],
            case["rule_name"],
            case["workbook_name"],
            case["sheet_name"],
            maps,
        )
        expected = case["expected"]
        results.append(
            {
                "id": case["id"],
                "operation": "standardize_path",
                "actual": actual,
                "expected": expected,
                "status": "保留" if actual == expected else "损坏",
            }
        )

    failures = [item for item in results if item["status"] != "保留"]
    report = {
        "schema_version": 1,
        "suite": "pytools.path_map",
        "baseline": "original PyCode.core.path_map and rule_engine",
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
