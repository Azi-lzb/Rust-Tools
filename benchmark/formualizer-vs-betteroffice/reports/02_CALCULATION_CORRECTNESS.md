# Calculation correctness

## Oracle and comparison

Microsoft Excel COM (`CalculateFullRebuild`) on Windows generated the expected
values for each fixture and input state. Each engine's formula target values
are compared with Excel COM, with numeric tolerance `max(1e-9, 1e-9 × |value|)`;
errors, text, and booleans are compared by type and value. Formula fixtures are
generated and calculated on the same Excel installation. The COM process only
provides oracle values; it does not perform either engine's calculation.

## Results

| Fixture | Formula targets | BetterOffice full / changed input | Formualizer full / changed input | Notes |
|---|---:|---|---|---|
| Basic mixed-function suite | 33 | 33/33, 33/33 | 33/33, 33/33 | Includes known errors and one deliberately unknown function; both resident paths match Excel for target values |
| Stale-cache suite with a supported replacement for the unknown function | 33 | 33/33, 33/33 | 33/33, 33/33 | Strict Formualizer cache update completes |
| 10k repeated formula fixture, stale caches | 10,000 | 10,000/10,000, 10,000/10,000 | 10,000/10,000, 10,000/10,000 | Seven formula groups |
| 70k repeated formula fixture, stale caches | 70,000 | 70,000/70,000, 70,000/70,000 | 70,000/70,000, 70,000/70,000 | Seven formula groups; one successful full correctness run and subsequent three-round performance validation |
| 100k repeated formula fixture, stale caches | 100,000 | 100,000/100,000, 100,000/100,000 | 100,000/100,000, 100,000/100,000 | Seven formula groups; one successful full correctness run and three-round performance validation |

No tested known-function target disagreed with Excel COM. This result is
limited to the listed functions and synthetic data. It is not evidence that
either engine has full Excel calculation compatibility or exceeds the other
on the actual rustTools business formula set.

## Coverage limitation behavior

The basic suite contains `FORMUALIZER_UNKNOWN_FN(1)`. Resident calculation
returns a value matching the Excel error result for this explicit case, but
source-aware strict cache mode reports incomplete known-function coverage and
refuses to publish a strict output. Lenient mode leaves that one cache
uncomputed: it matches 32/33 formula caches. This is correctly treated as
PARTIAL, not a successful full-cache calculation.

The strict supported-function fixture replaces that target with a known SUM
formula. Strict cache mode then evaluates all 33 formula cells and matches
Excel's caches 33/33.

## Evidence files

The synthetic input workbooks, formula manifests, and Excel COM oracle JSON are
under `fixtures/synthetic/`. Per-run formula results, mismatches, and reopened
package cache verification are saved under the ignored `results/` directory.
