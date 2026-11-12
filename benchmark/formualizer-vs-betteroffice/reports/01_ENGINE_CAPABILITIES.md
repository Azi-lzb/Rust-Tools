# Engine capabilities and test boundary

## Tested implementations

| Implementation | Version | Path exercised | Evidence |
|---|---:|---|---|
| BetterOffice | 0.5.1 (the checked-in rustTools crates) | `Workbook::open` → `recalculate_all` twice → `save`; `edit_cell` for dirty dependency recalculation | Adapter output, Excel COM oracle, reopened XLSX cache inspection |
| Formualizer | 0.14.0 | `UmyaAdapter` resident workbook (`evaluate_all`, delta evaluation, `to_xlsx_bytes`) and source-aware `recalculate_xlsx_file` cache rewrite | Adapter output, strict/lenient cache-rewrite summaries, Excel COM oracle, OOXML part and cache inspection |

Formualizer uses `formualizer-workbook` with `umya`, `xlsx-recalc`, and
`system-clock`. Its cache-rewrite implementation needs `zip = 7.2.0`: Formualizer
0.14.0 calls the `ZipFile::data_start()` API that returns `u64` in 7.2; zip
7.4 changes that API and fails to compile this path. This exact pin is recorded
in the isolated `Cargo.lock`.

The crate manifests declare Formualizer `MIT OR Apache-2.0` and BetterOffice
`Apache-2.0`. This is a metadata check of the two direct libraries, not a full
third-party dependency-license audit.

For the Linux Rocky 8 system toolchain build (rustc 1.92.0), the experiment locally patches two
unrelated BetterOffice transitive crates: the render crate replaces three
unstable `usize::isolate_lowest_one` calls with the equivalent stable
`value & value.wrapping_neg()` operation; `yrs` rewrites one unstable `if let`
match guard into an equivalent nested match. These shims are isolated under
this benchmark, are not production changes, and do not alter formula
calculation. An unpatched Linux Release build failed on these Rust language/API
features; the evidence below is for the patched isolated harness.

## Observed API capabilities

| Capability | BetterOffice | Formualizer resident workbook | Formualizer source-aware cache path |
|---|---|---|---|
| Open and evaluate XLSX formulas | PASS for the tested fixtures | PASS for the tested fixtures | PASS for supported functions in tested fixtures |
| Recalculate after input mutation | PASS; `edit_cell` recalculates dirty dependencies | PASS; `evaluate_all_with_delta` returns a delta | Not a resident mutation API; recalculates source file caches |
| Write a new workbook package | PASS; exercised and reopened | API executes, but output fidelity/type checks FAIL | PASS for the tested cache-only edit; original parts preserved except the formula worksheet |
| Refuse incomplete formula coverage | No equivalent strict source-cache gate exercised | N/A | PASS with `require_complete=true`; unknown function prevents publication |
| XLS input | NOT_TESTED in this formula-only experiment | NOT_TESTED | NOT_TESTED |
| Real rustTools 1-4 workbook | NOT_TESTED | NOT_TESTED | NOT_TESTED |

The suite exercises SUM, SUMIF/SUMIFS, COUNTIF/COUNTIFS, IF/IFS, AND/OR,
VLOOKUP/XLOOKUP, INDEX/MATCH, ROUND variants, DATE/YEAR/MONTH/DAY, text
functions, cross-sheet references, a defined name, mixed references, and
error/unknown-function cases. Scale fixtures repeat a smaller seven-formula
group (SUM, SUMIF, SUMIFS, COUNTIF, VLOOKUP, INDEX/MATCH, ROUND) over 10,000
and 70,000 formula cells. Those synthetic scales establish capacity for this
formula mix; they do not establish broad Excel function compatibility.

This project tests formula calculation and formula-cache writing. It is not a
general workbook editor certification and does not test arbitrary XLSX
features, macros, external links, charts, or `.xls`.
