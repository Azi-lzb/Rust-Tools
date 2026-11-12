# Incremental recalculation

## Procedure

For small cases, Excel COM calculated the workbook after the initial input and
again after mutating `Data!B2` from 10 to 25. For scale fixtures, the original
workbook was calculated with `Data!B2=1`, then a stale-cache input was created
by changing only that input cell's OOXML value to 25 while retaining the
formula caches from the original value. The engines recalculated against the
Excel COM result for input 25, then the adapter changed the input to 35 and
calculated again against a separately generated Excel COM oracle.

BetterOffice follows the rustTools calculation policy (`recalculate_all`
twice) and its `edit_cell` path recalculates dirty dependencies. Formualizer's
resident route uses `evaluate_all` and then `evaluate_all_with_delta`. Its
source-aware cache route starts from the file's existing package and recalculates
formula caches.

## Results

| Scenario | BetterOffice | Formualizer resident | Formualizer source-aware cache |
|---|---|---|---|
| Full calculation from stale caches | MATCH vs Excel COM | MATCH vs Excel COM | MATCH for supported-function fixtures; strict mode refuses incomplete coverage |
| Input mutation and dependent results | MATCH vs Excel COM | MATCH vs Excel COM | The cache-file API does not expose the same in-memory mutation path; not ranked as resident incremental calculation |
| 10,000 formulas | 10,000/10,000 at input 25 and 35 | 10,000/10,000 at input 25 and 35 | 10,000 formula cells evaluated; cache output matched 10,000/10,000 |
| 70,000 formulas | 70,000/70,000 at input 25 and 35 | 70,000/70,000 at input 25 and 35 | 70,000 formula cells evaluated; cache output matched 70,000/70,000 |
| 100,000 formulas | 100,000/100,000 at input 25 and 35 | 100,000/100,000 at input 25 and 35 | 100,000 formula cells evaluated; cache output matched 100,000/100,000 |

The resident delta APIs execute actual recalculation and produce values that
match a fresh Excel full rebuild. This verifies result correctness for the
fixture; it does not prove that every mutation recalculates only the minimal
dependency subgraph. Formualizer's returned delta is recorded in the raw
`engine-runs.json` for each run. BetterOffice's `edit_cell` includes recalculation
in the mutation timing, so it is reported separately from a standalone
recalculate call.

For the scale cache path, a second full cache pass on already updated data
changes zero cache cells in the 10k observed run. The first stale-cache pass
changes 4,286 cells at 10k and 23,743 cells at 70k, and only the Calc worksheet
part is modified. Cache update counts are fixture-specific, not a general
incremental speed claim.
