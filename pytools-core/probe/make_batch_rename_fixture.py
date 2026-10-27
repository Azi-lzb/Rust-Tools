"""Generate the synthetic xlsx config fixture for the Rust/Python parity test."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from openpyxl import Workbook


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).parents[1] / "tests/fixtures/batch_rename_config.xlsx",
    )
    args = parser.parse_args()

    case_path = Path(__file__).parents[1] / "tests/fixtures/batch_rename_cases.json"
    cases = json.loads(case_path.read_text(encoding="utf-8"))
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = cases["config_sheet"]
    sheet.append(cases["headers"])
    for row in cases["config_rows"]:
        sheet.append(row)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(args.output)
    workbook.close()
    print(f"wrote {args.output}")


if __name__ == "__main__":
    main()
