# pytools-core

This package is the Rust-side implementation target for py_statistics_tools. It depends on generic Excel primitives under workspace crates/ and never on audit-core.

- excel-reader: pure Rust Calamine values, formulas, sheet names and supported merges; namespace-aware OOXML supplements read comments, conditional formatting and data-validation structure.
- excel-writer: rust_xlsxwriter for new xlsx workbooks.
- excel-editor: ExcelEditPlan/ExcelModifier boundary for existing-workbook changes; OOXML Patch, umya-spreadsheet and a bundled Excelize worker remain operation-specific candidates selected by preservation evidence.
- ooxml-patch: narrow namespace-aware OOXML changes for verified gaps.

Keep pytools fixtures, probe configuration and reports under this directory. The Rust probe is shared, but its inputs, baseline results and reports stay separate from audit-core. Python comparison scripts remain baselines/test helpers; the only planned production Python exceptions are an isolated, capability-gated openpyxl worker for existing-workbook edits and a future Windows-only pywin32 worker for explicitly selected, validated Excel/WPS COM modification or calculation. Neither worker may contain PyTools business logic; Linux does not depend on COM.

The initial Rust modules include shared-vector date/path helpers, configuration loaders for global/timeline/path-map sheets, a single-worksheet timeline extractor, and a candidate workbook adapter for institution-name normalization. The normalization adapter matches the source helper on one synthetic xlsx fixture and preserves that fixture's values, formula text, merge, comment, styles, dimensions, and page metadata; it remains outside the default edit path pending real-template verification and does not yet support xlsm. Timeline batch scheduling, CSV input, result writing, target append, full menu integration, and real business templates remain unverified.

## Worksheet area extraction candidate

The `pytools` binary has an explicit area-only extraction command:

```text
pytools extract-sheets-area --config <config.xlsx> [--output-dir <directory>] <source.xlsx>...
```

It loads the existing worksheet-extraction config, reads values through the shared Calamine adapter, and creates new XLSX output with rust_xlsxwriter. The command accepts explicit source file paths; it is not wired to the original product menu or directory scanner. It rejects full-sheet style-preserving extraction, append-to-existing, dates, durations, error cells, and non-finite numbers. Its synthetic Python parity evidence is in `probe/reports/sheet_extract_area_parity.json`; see `MIGRATION_INVENTORY.md` for the tested scope and remaining gaps.

## Regional total check

The `pytools` binary also exposes the Rust non-COM version of the original regional total check:

```text
pytools region-sum-check [--config <workbook.xlsx>] [--output-dir <directory>] [--log-dir <directory>] <source.xlsx>...
```

When `--config` is supplied, output and log directories default to the global configuration; explicit directory arguments override them. Without a configuration workbook, defaults are `./output` and `./logs`. The command uses the shared Calamine reader and rust_xlsxwriter result writer. Its synthetic Python parity suite checks result statistics and workbook values/formulas, and verifies that source files remain unchanged. Real templates, non-XLSX inputs, the original GUI menu, and target Linux packaging remain unverified; see `MIGRATION_INVENTORY.md`.

## Local filename mapping

Operation 4-8 is callable from the Rust binary:

```text
pytools rename-local-map --config <workbook.xlsx> <file>...
```

The command reads ordered M:N mappings from the existing rename-configuration sheet and renames the explicitly listed files. It accepts regular files of any extension, reserves destinations without overwriting them, and returns JSON counts and per-file results. Its 10-path synthetic Python parity suite passes through the CLI; the original GUI menu, real configuration, target Linux filesystems, and non-UTF-8 filenames remain unverified.

## Configuration-driven batch rename

Operation 3-3 is available through the Rust CLI:

```text
pytools rename-batch --config <workbook.xlsx> [--log-dir <directory>] [<file-or-directory>...]
```

It reads the `重命名配置` sheet with the shared Calamine reader, applies ordered short-name matching and full-name/code mapping, and renames each explicitly supplied file or directory. Destination collisions receive `_1`, `_2`, ... suffixes. The command writes JSON counts/per-path results and a timestamped log. The command accepts an empty path list to validate configuration without renaming files. Synthetic Python parity is recorded in `probe/reports/batch_rename_parity.json`; a smoke test also used one valid alias from the original read-only configuration with a synthetic file in a temporary directory. No real business file was renamed. Target Linux behavior and the original Python GUI menu remain unverified.

## Content-prefix renaming

Operations 4-6 and 4-7 are available through the Rust CLI:

```text
pytools rename-by-content --config <workbook.xlsx> <file>...
pytools unrename-by-content --config <workbook.xlsx> <file>...
```

Both commands read the configured cell values from `.xlsx`/`.xlsm` files with the shared Calamine reader, rename files without overwriting occupied destinations, and emit JSON statistics and per-file results. They do not save workbook contents. Synthetic parity checks compare the Rust API and CLI with the Python baseline for the tested text-cell cases. The source `.xls`/`.xlt` COM path, real macro-enabled workbooks, date/formula-cache values, GUI dispatcher, real directories, Windows-native Rust, and target Linux filesystems remain unverified.

## File-sort prefix

Operation 4-9 is also callable from the Rust binary:

```text
pytools file-sort-prefix --config <workbook.xlsx> [--output-dir <directory>] [--log-dir <directory>] <file>...
```

It reads the existing file-sort rules, applies the first matching rule to each explicit file path, preserves occupied destinations, and writes the result workbook. `--output-dir` defaults to `./output`; `--log-dir` defaults to `./logs`. The CLI writes success/failure log events to a timestamped file and stderr. The synthetic CLI parity test verifies the generated report is readable, matches the Python baseline's counts and final filename set, and emits the expected success-log entries. Exact logger timestamp/concurrency behavior, the original GUI dispatcher, real directories, and target Linux deployment remain unverified.

## Remove file-sort prefix

Operation 4-10 is callable from the Rust binary:

```text
pytools remove-file-sort-prefix --config <workbook.xlsx> [--output-dir <directory>] [--log-dir <directory>] [<file>...]
```

It removes a configured sorting prefix, preserving a preceding content prefix when the configured workbook content matches. It writes the result workbook and timestamped success/failure logs; an empty file list produces the source-compatible blank report. The synthetic CLI parity tests cover ten file/path cases and the empty-input report. `.xls`/`.xlt` content-cell parity, the original GUI dispatcher, exact logger timing/concurrency, real directories, and target Linux deployment remain unverified.
