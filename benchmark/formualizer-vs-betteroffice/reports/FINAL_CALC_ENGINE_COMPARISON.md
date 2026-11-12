# Formualizer vs BetterOffice: final comparison

## Decision

Do not replace BetterOffice in rustTools with Formualizer's resident
load-calculate-save route. It calculates the tested formulas correctly and is
faster on synthetic formula-heavy files, but its current writer fails XLSX
cache typing and package fidelity checks. Keep BetterOffice as the production
default. The Formualizer source-aware formula-cache updater is promising as an
optional, strict-coverage, cache-only route, pending the real rustTools 1-4
formula corpus and workbook fidelity validation.

## Evidence summary

| Measure | BetterOffice | Formualizer |
|---|---|---|
| Mixed formula suite correctness | 33/33 full and 33/33 after input change | 33/33 full and 33/33 after input change |
| Unknown formula cache handling | N/A in tested save route | Strict mode refuses incomplete known-function coverage; lenient cache output 32/33 |
| Supported 10k scale correctness | 10,000/10,000 full and after change | 10,000/10,000 full and after change; source cache 10,000/10,000 |
| Supported 70k scale correctness | 70,000/70,000 full and after change | 70,000/70,000 full and after change; source cache 70,000/70,000 |
| 100k scale correctness | 100,000/100,000 full and after change | 100,000/100,000 full and after change; source cache 100,000/100,000 |
| 70k / 100k calculation medians | 1,783.6 / 2,524.6 ms | 423.0 / 569.1 ms resident calculation |
| 70k / 100k XLSX cache fidelity | PASS on the tested fixtures | Resident writer FAIL; source-aware cache path PASS for tested supported formulas |
| Linux glibc 2.28 | PASS for isolated harness with documented stable-toolchain compatibility shims | PASS as dependency in that harness; source-cache path executed offline |
| Kylin V10 / UOS V20 | NOT_TESTED | NOT_TESTED |
| Actual rustTools 1-4 business files | NOT_TESTED | NOT_TESTED |

The 70k and 100k timings are three-round medians after warm-up. Formualizer's
423 ms and 569 ms calculation results are not enough to justify replacing the
writer. Its source-cache updater finished the supported-function
recalc/cache-write paths at 1,000.7 ms (70k) and 1,325.7 ms (100k) median and
changed only one formula worksheet package part.
Linux evidence is Rocky Linux 8.9 x86_64/glibc 2.28 with networking disabled;
it is not a full deb or target Kylin/UOS validation.

## Scope and limitations

This is a formula-engine experiment on Excel COM-generated synthetic XLSX
fixtures. It is not a broad Excel compatibility certification. No `.xls`,
macros, dynamic array/spill formulas, shared formulas, circular/iterative
calculation, real 1-4 files, or business output were tested. The 70k and 100k files
repeat seven formula groups over a small data table; they measure graph/cell
scale, not realistic dependency diversity. Per-engine peak memory was not
isolated. See the six detailed reports for methods and evidence.

## Status

Implementation and experiments are isolated under
`benchmark/formualizer-vs-betteroffice/`. No rustTools production route was
changed. Linux glibc 2.28 experiment runtime passed; Kylin/UOS and actual
business samples remain open validation items.
