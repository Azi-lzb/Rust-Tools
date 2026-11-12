# XLSX fidelity

## Verification performed

The original and output packages were compared at the ZIP-member level. Excel
COM reopened the small fixture outputs and checked formulas/results, sheet
names, named ranges, print area, orientation, paper size, margins, merge state,
hyperlinks, representative widths/heights, and styles. Formula caches were
independently read from worksheet XML and compared with Excel COM oracle values.

## Findings

| Output path | Formula cache result | Package/feature result | Status |
|---|---|---|---|
| BetterOffice save | 33/33 cache values match; 10k/10k and 70k/70k scale caches match | Excel COM reopens output; test fixture sheet/print/style/merge/link/name checks preserved. Package writer updates expected workbook/calc metadata and worksheet parts | PASS for tested features |
| Formualizer resident `to_xlsx_bytes()` | 7/33 typed caches match in the mixed suite; all numeric scale formula caches reopen as text values (`t="str"`) | Rewrites many package members; loses the defined name used by `Calc!B28` and representative original worksheet features | FAIL as a production save path for these inputs |
| Formualizer source-aware cache rewrite | 32/33 in lenient unknown-function suite; strict supported-function fixture 33/33; 10k/10k, 70k/70k, and 100k/100k scale cache values match | Cache-only supported fixture changes only `xl/worksheets/sheet1.xml`; stale scale runs change one formula worksheet part | PASS for supported formulas and tested OOXML shape; strict gate required |

The resident Formualizer API's in-memory calculation correctness must not be
confused with its current XLSX serialization correctness. Its output converted
calculated numbers, booleans, and errors into string-typed cached values in the
observed files. Excel opening a workbook does not repair this fidelity failure
for downstream readers that use cached values.

The source-aware cache updater is materially different: it patches formula
cache values in the source worksheet XML instead of rebuilding the workbook.
For the tested supported formulas it preserves all other ZIP members and
updates numeric caches correctly. It refuses strict output if it encounters
an unknown function. This does not certify unsupported OOXML extensions,
signatures, macros, external links, or every Excel feature.
