# Local experiment outputs

Runtime results are ignored by Git to avoid committing bulky repeated
measurements. The reproducible input workbooks, formula manifests, oracle JSON,
and scripts are in `../fixtures/synthetic/`.

Key completed result groups:

- `synthetic/`, `valid-stale/`: small formula correctness, strict-cache, and
  Excel COM fidelity outputs.
- `perf-10k-stale-memory/`: 10,000 formulas, warm-up + three rounds, timing
  CSV and summary.
- `scale-70k-stale/`: 70,000 formula correctness and package-cache evidence.
- `perf-70k-stale-final/`: 70,000 formulas, warm-up + three rounds, timing
  CSV and summary.
- `scale-100k-stale/` and `perf-100k-stale/`: 100,000 formula results and
  performance.
- `linux228-stale/` and `linux228-offline/`: Rocky Linux 8.9/glibc 2.28
  execution evidence.
- `perf-two-pass-10k-indirect/` and `linux228-offline-two-pass-10k/`: repeated
  10,000-formula two-pass timings, Excel-oracle comparisons, and offline
  glibc 2.28 execution/package-cache evidence for the INDIRECT fixture.

Each timed scale run begins from the same stale-cache input copy. Keep the
fixture, oracle, and matching spec together when reproducing a result.
