# BetterOffice XLSX modification audit

This standalone adapter is pinned to BetterOffice `0.5.1`; `Cargo.lock` freezes
the complete dependency graph. It uses the repository's existing patched
BetterOffice model and calculation crates and does not link into production.

## Reproduce

Run from the Rust-Tools-test root in PowerShell:

```powershell
$env:CARGO_TARGET_DIR = 'F:\RustBuild\pytools-test'
cargo test --release --offline --manifest-path benchmark/rust-xlsx-modifier/adapters/betteroffice/Cargo.toml
cargo run --release --offline --manifest-path benchmark/rust-xlsx-modifier/adapters/betteroffice/Cargo.toml -- audit benchmark/rust-xlsx-modifier/adapters/betteroffice/results
cargo run --release --offline --manifest-path benchmark/rust-xlsx-modifier/adapters/betteroffice/Cargo.toml -- gen100k benchmark/rust-xlsx-modifier/adapters/betteroffice/results
benchmark/rust-xlsx-modifier/adapters/betteroffice/scripts/run_perf.ps1
benchmark/rust-xlsx-modifier/adapters/betteroffice/scripts/run_excel_reference.ps1
benchmark/rust-xlsx-modifier/adapters/betteroffice/scripts/run_com_validation.ps1
benchmark/rust-xlsx-modifier/adapters/betteroffice/scripts/run_comment_roundtrip.ps1
benchmark/rust-xlsx-modifier/adapters/betteroffice/scripts/inventory_public_api.ps1
```

The COM validation scripts require Microsoft Excel installed on the test
machine. They open only generated files read-only except for the script that
creates its generated comment fixture, then save only below `results/`.
`run_excel_reference.ps1` executes the same row operations in Excel COM and
records its resulting conditional-format range as the correctness reference.
`inventory_public_api.ps1` scans the pinned crate sources and lists all public
`Op` variants plus hits for page/print-related API names.

`audit` creates a synthetic XLSX with a chart, conditional formatting, merge,
defined name, print settings, formulas and an opaque `customXml` part. The Rust
test asserts the structural/formula outcomes. The COM scripts independently
check Excel's interpretation. `perf` measures a generated 100,000-row x
3-column workbook; it performs one warm-up followed by three measured rounds
per workload. Each round starts from the same source bytes.

Generated workbooks, JSON and logs are retained in `results/` as test evidence.
