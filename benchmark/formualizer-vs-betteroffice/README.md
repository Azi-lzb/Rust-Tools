# Formualizer vs BetterOffice

Isolated formula-calculation experiment. It uses the same XLSX fixture for both
Rust engines, takes expected values from a local Microsoft Excel COM
`CalculateFullRebuild`, compares resident workbook calculation and Formualizer's
source-aware cache update as separate paths, and keeps outputs outside the
production runtime. Nothing here changes rustTools menus or the 1-4 route.

## Reproduction

1. On Windows with Microsoft Excel installed, run the fixture builders under
   `fixtures/synthetic/` to create or refresh the Excel COM-generated inputs,
   zero-based formula manifests, and `CalculateFullRebuild` oracle JSON.
   The tracked files are the deterministic experiment inputs; regeneration
   requires Excel COM.
2. From the repository root, build and run the basic fixture:

   ```powershell
   $env:CARGO_TARGET_DIR='F:\RustBuild\pytools-test'
   cargo run --release --locked --offline --manifest-path benchmark/formualizer-vs-betteroffice/Cargo.toml -- run benchmark/formualizer-vs-betteroffice/fixtures/synthetic/formula-suite.xlsx benchmark/formualizer-vs-betteroffice/results/synthetic benchmark/formualizer-vs-betteroffice/fixtures/synthetic/cases.json benchmark/formualizer-vs-betteroffice/fixtures/synthetic/excel-com-oracle.json benchmark/formualizer-vs-betteroffice/fixtures/synthetic/excel-com-incremental-oracle.json
   ```

For a scale fixture, use `generate_scale_fixture.ps1` (100–100,000 formula
targets; add `-IncludeIndirect` for a repeated INDIRECT group),
`prepare_scale_stale_cache.ps1`, and then
`run_performance_repetitions.ps1` with the matching `cases-scale-*-stale.json`
and Excel COM oracle files. It runs one warm-up and three measured rounds from
the same input and saves timing CSV/summary JSON. `verify_packages.ps1` checks
reopened formula caches and package-member preservation; `verify_fidelity.ps1`
uses Excel COM to inspect the small feature fixture.

To compare one and two calculation passes, use
`fixtures/synthetic/prepare_two_pass_fixture.ps1` for the dependency-chain
fixture, or generate a scale fixture with `-IncludeIndirect`. Then invoke the
binary's `run-two-pass` command with the input, spec, and separate input-25 and
input-35 Excel COM oracles. `run_two_pass_repetitions.ps1` runs the scale case
with a warm-up plus three measured rounds.

The isolated manifest pins `zip=7.2.0` for Formualizer 0.14.0's cache rewrite.
It also carries two narrowly scoped BetterOffice transitive-dependency
compatibility shims required to compile with Rocky's system Rust 1.92.0; see
`reports/01_ENGINE_CAPABILITIES.md`. These are experiment-only and do not
modify production code. See `reports/` for measured results and limitations.
