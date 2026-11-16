# Optional resident Aspose engine — evaluation trial

## 2026-11-16 test-package delivery

Windows TEST-main packages now include independent, self-contained Cells and
Words workers. No installed .NET SDK is needed for the packaged apphost route.
Evaluation output is retained; Auto and Linux routes are unchanged. Current
full-suite results, packaging checks and recorded failed attempts are in
`docs/RELEASE_VALIDATION_20261116.md`. Statements below about not packaging or
updating dist describe the earlier dated trial, not the current test package.

## 2026-11-15 full-menu extension

Current opt-in menu coverage, interfaces, typed/structural edits, conversion
Workers, safety boundaries and measured results are documented in
`docs/ASPOSE_FULL_MENU_TRIAL_20261115.md`. That report supersedes the initial
capability limits below; earlier probe history is retained. Automatic and Linux
routes remain unchanged. Words uses a separate project and process. Evaluation
output remains intact. This task distributes no Worker binaries.

New regressions: `test_plan_operations.py`, `test_menu_matrix.py`,
`test_openpyxl_row_dimensions.py`, and `pytools-core/tests/aspose_modify_cli.rs`.
`benchmark_pipeline.py` compares the protected production Worker APIs and real
menus; `benchmark_jobs.py` measures a real 100-file batch and the production
named-validation calculation flow. Startup, warmups and API/E2E clocks are
separate. Raw outputs are retained and verified outside the clock. Read the
report's DELIVERY.json and PERFORMANCE.csv for actual results, including any
failed or incomplete cohort; scripts alone are not acceptance evidence.

The user accepted watermarked evaluation output on 2026-11-15. The extra
evaluation sheet and its active-tab change are expected in this phase, not a
business misalignment. No saved evaluation sheet is removed. The old
`--memory-save-probe` and its failed results remain evidence that pre-Save
deletion does not produce clean output with Aspose.Cells 26.10.

## Architecture and scope

`AsposeWorker.csproj` is an independent .NET 10 application, pinned to
Aspose.Cells 26.10.0. `../aspose-worker-host` is an independent Cargo workspace.
Its compatibility facade re-exports `CalcEngine`, `ModifyEngine`, typed
requests/operations and `AsposeWorkerBackend` from `crates/excel-aspose`.
The shared implementation now lives in Excel-Core. Business callers depend on
the traits, not a .NET command; the standalone probe still uses the same API.
Replacing the backend does not require changing those callers.

The original independent integration is the explicit `worker_case` executable.
It sends requests through the traits, verifies candidates and the source
snapshot, and continues in Rust to write the result manifest. A subsequent
main-project integration adds explicit calculation selection to menu 1-5 and
`summary-named-validation --calc-engine aspose`. Auto, ExcelEditPlan modification
routing and Linux routes retain their previous behavior. This is not a rollout
to all production business functions; see `docs/ASPOSE_CALC_INTEGRATION.md`.

The protocol is one JSON object per line on stdio. Startup emits READY with
protocol version, PID, engine version and budget. Requests carry `request_id`,
`run_id`, `method`, absolute `input` and unique `candidate` paths. Methods are
`Ping`, `Calculate` and `Modify`; stdout is protocol only, stderr is diagnostic
JSONL. Diagnostics do not log cell values. The Rust trial gives each batch a
unique run ID, UTC start/end timestamps and per-task result records.

Input support is XLS/XLSX, output XLSX only. Macros are rejected. Initially
supported modifications are explicit text/finite IEEE double/boolean/blank,
formula text, bold/italic/font/size/fill/number-format/wrap changes on a cell.
Coordinates are strict uppercase A1 within XLSX bounds. Dates, Excel error
values, unrecognized styles/operations and other formats are not silently
converted. This is not an arbitrary Aspose API/JSON execution gateway.

## Safety and lifetime

- Inputs are never saved in place. Candidates use CreateNew, not overwrite.
- All modification instructions are checked before writing a candidate.
- The host hashes the input before/after and verifies ZIP metadata, worksheet
  relationships and worksheet XML before accepting a candidate.
- The worker independently checks requested styles against serialized OOXML;
  memory-only style success is insufficient. A failed candidate can remain for
  diagnostics but is returned as failure, never a successful result.
- Returning a verified candidate is not permission to replace a business file.
  A production caller must still use its existing final validation/commit path.
- One worker handles one task at a time. Ping consumes no file budget. Every
  attempted Workbook load consumes one slot, including a failed load/operation.
  Slot 99's response has `draining=true`; the worker then exits. Rust receives
  that response, reaps it and starts a new worker only for the next task.
- A lost response, timeout or engine/limit error is not automatically replayed.
  Startup, blocked pipe writes and response waits have deadlines. Shutdown
  closes stdin, waits, then kills/reaps only the owned direct child if needed.
  This worker does not launch grandchildren; arbitrary descendant cleanup is
  not claimed. Windows workers use CREATE_NO_WINDOW.
- Inputs already containing an evaluation-like sheet with unknown ownership
  are rejected rather than risking a real user's identically named sheet.
  Outputs produced by this same live worker can be chained if their hashes
  still match. Such ownership is not carried across worker generations.

## Current evidence and limitation

Real tests use synthetic files only. The complex fixture has 2048 formulas,
styles and original active tab 1. Independent openpyxl/BIFF checks verify actual
files; evaluation sheets do not count as business data. `test_resident.py` checks
calculation against independent numeric expectations, typed values, styles,
coordinate rejection, chaining, conflicts and failure preservation. The real
100-task boundary uses a tiny synthetic XLSX and 100 distinct candidates: it is
not 100 real business files. It verifies 99 replies in one process and task 100
in a second process with a reset counter.

**Known compatibility gap, not fixed:** with the original xlwt-authored BIFF
fixture, styles at AA100/AA12 are correct in memory but lost during XLSX save.
File/stream saves and Cell/Range style APIs reproduce it. The same data saved
as XLS by an isolated Excel 16.0 session, and the XLSX control, retain those
styles. This establishes an input/engine compatibility difference; it does not
prove that all xlwt files are invalid or all XLS files are supported.
`test_style_guard.py` preserves this case and verifies OUTPUT_VALIDATION_FAILED
with the source unchanged. The failing raw files are retained. No fallback
engine, format-only renaming or OOXML repair hides this limitation.

The original assertion failures, CRLF-only protocol fixture failure and a
missing test-environment invocation are retained in the task/tool history.
The protocol fixture now writes LF explicitly; the exactly-one-request
assertions remain. Protocol tests have isolated temporary-directory and Python
defaults; explicit paths can be supplied for reproducible capture.

## Build and targeted tests

Use the portable SDK unpacked at
`F:/RustBuild/pytools/aspose-worker-sdk-10.0.404`; no system install is required.
Set DOTNET_ROOT to that directory. Restore from the existing NuGet cache, then
build with BaseIntermediateOutputPath and OutputPath under
`F:/RustBuild/pytools/aspose-worker/{obj,bin}`. Repository bin/obj, DLL/PDB and
test-output are ignored; licenses must never enter Git.

Before **every** Cargo invocation in main set
`CARGO_TARGET_DIR=F:/RustBuild/pytools`.

```text
cargo fmt --manifest-path tools/aspose-worker-host/Cargo.toml -- --check
cargo test --offline --manifest-path tools/aspose-worker-host/Cargo.toml --test protocol -- --test-threads=1
cargo build --offline --manifest-path tools/aspose-worker-host/Cargo.toml --bin worker_case
```

Protocol tests use Python standard-library stubs, not Office or Aspose. Set
ASPOSE_TEST_PYTHON and ASPOSE_TEST_ROOT to override their isolated defaults.
Real tests additionally need the existing openpyxl/xlrd/xlwt test environment:

```text
py -3.12 tools/AsposeWorker/test_resident.py NEW_OUTPUT_DIR DOTNET_EXE WORKER_DLL HOST_EXE EXCEL_AUTHORED_SYNTHETIC_XLS
py -3.12 tools/AsposeWorker/test_style_guard.py NEW_OUTPUT_DIR ORIGINAL_XLWT_XLS DOTNET_EXE WORKER_DLL
```

Excel was used only to independently author the synthetic XLS control in an
owned DispatchEx session; it is not a runtime dependency of the worker/host.
Omitting the Excel-authored fixture argument intentionally exercises the
original failing xlwt fixture, not a passing replacement. Every run uses a new
output directory and preserves all evaluation output. No full suite is run.

Performance capture alternates three resident/fresh pairs, five equal tasks per
batch (one Calculate, four Modify), same Excel-authored XLS and same Rust host.
The outer clock includes Rust process creation, .NET startup, first-use JIT,
hash/ZIP/XML/style validation, saving, result transport and shutdown. Host and
worker phase clocks are also saved. There is no subtraction or warmup in this
end-to-end trial. The Rust host is a debug build and CLR defaults are used;
these are not the earlier release/fixed-JIT pure-API benchmark numbers.

NOT_TESTED: licensed mode, clean output, restored active tab, production
deployment/packaging, Linux/WSL/Docker/Kylin, WPS, macros/XLSM, arbitrary Excel
features/formula functions, locale-dependent formats, concurrent clients,
unlimited process descendants, real business-file acceptance and crashes during
an existing production commit. No dist update, commit, push or other worktree
operation is performed.
