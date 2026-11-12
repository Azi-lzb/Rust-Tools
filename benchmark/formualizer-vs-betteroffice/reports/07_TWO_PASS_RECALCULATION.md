# Formualizer two-pass recalculation

## Question

Does calling Formualizer a second time, like the current BetterOffice full
calculation policy, correct formulas such as `INDIRECT` that might settle
after the first pass? Does a second source-aware cache pass make another
change?

## Fixtures and result correctness

Excel COM generated two synthetic cases. The 37-target case adds a four-cell
chain with two `INDIRECT("Data!B2")` formulas and downstream references. The
10,000-target case contains eight formula groups, including 1,250
`INDIRECT("Data!B2")+ROW()` formulas. Both start with stale caches: the input
cell is changed from 1 to 25 without updating formula caches; then the adapters
mutate it to 35 and compare against separate `CalculateFullRebuild` oracles.

| Case | BetterOffice (two full passes) | Formualizer pass 1 | Formualizer pass 2 | After input change, pass 1 / pass 2 | Source cache pass 1 / pass 2 |
|---|---|---|---|---|---|
| 37 targets with INDIRECT dependency chain | 37/37 | 37/37 | 37/37 | 37/37 / 37/37 | Executed / executed |
| 10,000 targets with 1,250 INDIRECT formulas | 10,000/10,000 | 10,000/10,000 | 10,000/10,000 | 10,000/10,000 / 10,000/10,000 | Executed / executed |

All values matched Excel COM. On the first 10k cache pass, 5,000 caches
changed; the second pass changed zero. Independent OOXML checks found 10,000
of 10,000 cached values correct after each pass. Pass one changed only
`xl/worksheets/sheet1.xml`; pass two changed no ZIP members. The Formualizer
resident writer still failed typed-cache fidelity, so two passes do not fix
its separate serialization problem.

The 37-cell chain also showed identical results after its second full pass.
After mutation, Formualizer's first incremental evaluation changed 13 cells in
the chain and matched Excel; its second incremental evaluation changed zero
cells and returned the same values.

## Timings

10k formula fixture, Release process, one warm-up and three measured rounds.
Values are three-round medians, in milliseconds:

| Operation | BetterOffice | Formualizer |
|---|---:|---:|
| Full calculation passes | 356.94 (two passes) | 72.71 (pass 1) + 4.76 (pass 2) = 77.47 |
| Input mutation/recalculation | 94.83 (`edit_cell`, including dependency recalc) | 28.90 (first delta pass) + 6.05 (second delta pass) = 34.95 |
| Load + full calculation + save | 712.87 | 304.06 for resident writer; output fidelity FAIL |
| Source-aware cache update | N/A | 146.45 (pass 1) + 155.13 (pass 2) = 301.58 |

The source-cache second pass takes time to reopen and check the package even
though it changes no formulas. Its total is not a same-operation comparison to
BetterOffice's full open/calculate/extract/save path. Timing CSV and per-round
JSON are in `results/perf-two-pass-10k-indirect/`.

The two-pass CLI also built and ran offline in Rocky Linux 8.9 x86_64 with
glibc 2.28 and networking disabled. Both calculation engines matched the
10,000-cell Excel oracle after each full/incremental pass. Independent package
verification found the source-aware cache output correct for all 10,000
formula caches after both passes; its first pass changed only the formula
worksheet XML and its second pass changed no ZIP members. The resident
Formualizer writer still had typed-cache mismatches in the same run.

## Conclusion

Two Formualizer passes remain faster than BetterOffice's two full passes on
this 10k synthetic formula set. However, the first pass already returned all
Excel-correct `INDIRECT` and downstream values. The second pass added no
correctness improvement in either the full or incremental case; it only added
about 4.8 ms to the resident full calculation and about 6.0 ms to the
incremental calculation at the measured median. The cache updater likewise
reached the correct result in pass one and changed zero values in pass two.

So a second Formualizer pass is safe as a defensive consistency check for this
tested workload, but the evidence does not show that it is needed to make these
`INDIRECT` results correct. It still leaves Formualizer faster here. Do not
extrapolate this to real rustTools formulas until their actual workbooks,
including indirect-reference patterns and output readers, are tested.
