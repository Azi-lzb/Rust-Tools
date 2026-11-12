# Formula performance

## Method

The 10k, 70k, and 100k fixtures use seven repeated formula groups and the same 100-row
input table. For each fixture, the source has stale formula caches (input
`Data!B2` changed from 1 to 25 without changing the old cached results). Each
engine performs a full calculation, extracts all target values, and saves an
output; the adapter also changes the input to 35 and validates the incremental
results. Formualizer's source-aware cache path is timed separately. Excel COM
creates oracle values outside the timed region.

Each run is a Release process. There is one warm-up and three measured runs,
each starting from the same unmodified input. Values below are medians of the
three measured rounds. All measured rounds passed calculation comparisons;
BetterOffice saves and Formualizer source-cache outputs passed cache/fidelity
verification. Formualizer resident-writer outputs were correctly calculated
in memory but failed typed-cache fidelity and are excluded as viable output.

## Median timings

| Formula cells | Engine/path | Load (ms) | Calculate (ms) | Save/serialize (ms) | End-to-end measured path (ms) | Peak RSS (MB) |
|---:|---|---:|---:|---:|---:|---:|
| 10,000 | BetterOffice full calc + save | 241.2 | 319.6 | 66.6 | 646.9 | 42.8* |
| 10,000 | Formualizer resident workbook | 76.8 | 56.6 | 129.6 | 272.2 | 42.8* |
| 10,000 | Formualizer source-cache update | — | included | included | 135.2 | 42.8* |
| 70,000 | BetterOffice full calc + save | 1,175.9 | 1,783.6 | 394.2 | 3,420.9 | 185.3* |
| 70,000 | Formualizer resident workbook | 547.1 | 423.0 | 950.5 | 1,959.8 | 185.3* |
| 70,000 | Formualizer source-cache update | — | included | included | 1,000.7 | 185.3* |
| 100,000 | BetterOffice full calc + save | 1,663.5 | 2,524.6 | 553.4 | 4,837.4 | 256.4* |
| 100,000 | Formualizer resident workbook | 723.4 | 569.1 | 1,311.4 | 2,692.1 | 256.4* |
| 100,000 | Formualizer source-cache update | — | included | included | 1,325.7 | 256.4* |

`*` Peak RSS is for the single Rust harness process that runs both engines and
the cache updater sequentially. It is not an isolated per-engine memory
measurement and must not be used to rank BetterOffice against Formualizer.
Raw timing CSVs are under `results/perf-10k-stale-memory/`,
`results/perf-70k-stale-final/`, and `results/perf-100k-stale/`.

## Interpretation

On these synthetic formula groups, Formualizer's calculation stage is faster
than BetterOffice's current two-pass `recalculate_all` baseline. At 70k,
calculation is about 4.2× faster (1,783.6 ms vs 423.0 ms); at 100k it is about
4.4× faster (2,524.6 ms vs 569.1 ms). The resident
Formualizer package serializer is slower than its calculator and fails typed
cache fidelity, so its 1,959.8 ms resident end-to-end time is not a valid
production win. The source-aware cache route both matches all tested caches and
preserves the source package apart from the formula worksheet; its total is
about 3.4× lower at 70k and 3.6× lower at 100k than BetterOffice's
open/calculate/extract/save path. This is
a useful targeted recalculation result, not a same-feature general workbook
save benchmark: the cache updater does not perform the adapter's input mutation
and incremental stage or rebuild the whole workbook.

No timing against Excel COM calculation or actual 1-4 business files was run.
No claim is made that these results predict real-workbook end-to-end speed.
