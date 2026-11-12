# 1-4 validation formula cache check

## Scope

This check asks whether Formualizer's formula-cache type marker causes the
existing 1-4 validation formula result to be read incorrectly.

The repository's checked-in 1-4 timeline workbooks
`pytools-core/tests/fixtures/timeline/wide_summary_elapsed_duration.xls` and
`wide_summary_legacy.xls` contain no formula cells, so they cannot test a
formula cache. Instead, this check reuses the exact formula scenario from the
existing integration test
`libreoffice_recalculates_validation_formula_cache_on_a_working_copy` in
`pytools-core/src/conversion_engine.rs`:

```excel
=IF(A2>0,"错|校验项|差异","")
```

The test workbook has `检查表!A2` as input and `检查表!B2` as the copied
validation formula. Excel COM generated the expected results. All workbooks
and outputs are under the ignored experiment `results/` directory; no source
fixture or production file was changed.

## Results

| Input/cache case | Engine | Full calculation | After changing A2 from 0 to 1 | Saved cache vs Excel |
|---|---|---:|---:|---:|
| Normal cache | BetterOffice | MATCH | MATCH | MATCH 1/1 |
| Normal cache | Formualizer | MATCH | MATCH | MATCH 1/1 |
| Stale cache (`STALE`) | BetterOffice | MATCH | MATCH | MATCH 1/1 |
| Stale cache (`STALE`) | Formualizer | MATCH | MATCH | MATCH 1/1 |

The expected values were `TEXT:` for the empty-string result and
`TEXT:错|校验项|差异` for the validation message. Formualizer's serialized
formula cell uses `t="str"` for these text results, which is the correct
cache type for this 1-4 formula. The source-aware cache updater also passed
the same 1/1 checks and replaced the deliberately stale cached value.

## Conclusion and limits

The cache-type mismatch observed on numeric formulas in the general
Formualizer formula suite did **not** invalidate this existing 1-4 validation
formula scenario: its result is text, and the saved text cache type matched
Excel. Both Formualizer's resident output and source-aware cache output
reopened through the independent OOXML verifier with the correct result.

This is a focused test based on the formula used by the checked-in 1-4
LibreOffice integration test, not a real user 1-4 workbook. The repository's
tracked timeline XLS fixtures have zero formulas, and the real validation
templates/workbooks used by business runs are not present as test fixtures.
Therefore this does not establish parity for the full 1-4 formula set,
shared formulas, named ranges, external references, or any numeric formula
cache consumed by the feature. Those require the actual sanitized 1-4
workbook/template pair. The experiment outputs are in
`results/one-four-existing-test/`.
