# rustTools integration assessment

## Current recommendation

Keep BetterOffice as the rustTools production formula engine for now. Do not
replace it based on this experiment alone. The resident Formualizer API
calculated the tested values correctly but its `to_xlsx_bytes()` output
serialized numeric formula caches as text and rebuilt package parts that the
small fidelity fixture expects to preserve. This disqualifies that writer path
for current production workbooks.

Keep Formualizer's source-aware cache updater as an isolated experiment or
optional future route for a narrowly defined task: recalculate supported
formula caches in an existing XLSX without rebuilding the workbook. Its strict
mode matched all formulas on the supported 33-, 10k-, 70k-, and 100k-target
fixtures tested so far and retained other ZIP parts. It correctly refuses
unknown-function coverage in strict mode. Before any production use, run the
actual rustTools 1-4 workbook corpus, validate the exact formula set and
business-visible outputs, and verify the remaining OOXML features on those
files.

This experiment changed only `benchmark/formualizer-vs-betteroffice/`. It did
not change BetterOffice, menus, or any production calculation route.

## Linux glibc 2.28 evidence

The isolated Rust harness was built and executed on Rocky Linux 8.9 x86_64 in a
container with glibc 2.28. The actual build used Rocky's system rustc/cargo
1.92.0 (a separately installed Rust 1.99.1 was not used by this build command).
To compile with that system toolchain, this
experiment applies two semantics-preserving compatibility shims in vendored
BetterOffice transitive dependencies (three low-bit calls in its render crate
and one `yrs` match guard); it does not change the formula engine. The binary
uses the normal Rust dynamic runtime and links to the system `libc`,
`libgcc_s`, `libpthread`, `libm`, `libdl`, and `ld-linux-x86-64`; it does not
bundle/replace glibc. `file` reports an x86-64 dynamically linked ELF;
`readelf --version-info` shows the highest required GLIBC symbol is
`GLIBC_2.28`. It was copied/run in the target image and also executed with
container networking disabled. The stale-cache smoke fixture matched all 33
in-memory targets. Independent package checks matched BetterOffice 33/33 and
Formualizer source-cache 33/33; the resident writer's 7/33 typed-cache mismatch
was reproduced. Record this as `GLIBC228_PASS` for the isolated patched
experiment binary, not as a Kylin V10 or UOS V20 hardware certification. Those
systems remain `NOT_TESTED`.

This was a Linux build/run of the calculation harness, not a full rustTools deb
or production packaging test. The experiment binary has no .NET or Python
runtime dependency. It must still be packaged with ordinary target-system
glibc; it is not a statically self-contained ELF.

## Cost/benefit

| Option | Benefit evidenced here | Cost / unresolved risk | Recommendation |
|---|---|---|---|
| BetterOffice only | Correct outputs and good tested XLSX save fidelity; already integrated | Slower on the synthetic 10k/70k formula groups | Keep production baseline; measure real files before optimizing |
| Replace with Formualizer resident API | Lower calculation and resident-path wall time on synthetic tests | Current XLSX serializer produces wrong cached cell types and loses observed metadata | Do not adopt this writer path |
| BetterOffice plus Formualizer cache updater | Faster cache-only recalculation on tested known functions; preserves unrelated package members | Separate API/coverage gate; not general editing; real business formulas and files are untested | Worth a gated prototype only after 1-4 formula-corpus validation |

Formualizer's calculation and dirty-delta interfaces may also inform a future
BetterOffice optimization, but this experiment does not isolate an algorithmic
optimization or prove a safe implementation change. No production change is
recommended yet.

## Reproduction

See the experiment [`README.md`](../README.md), the scripts under
`fixtures/synthetic/`, and the per-round CSV/JSON in the ignored `results/`
directory. The generated XLSX fixtures and Excel COM oracle JSON are tracked as
synthetic reproducibility inputs; runtime output files stay local.
