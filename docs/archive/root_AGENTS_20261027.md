# 历史参考：以下原规则已失效，当前规则见项目根 AGENTS.md。

# PyTools Rust project rules

This repository is the independent Rust migration of `py_statistics_tools`. It owns its business code, configuration, Cargo workspace, and local Excel-Core copies. It does not depend on Audit-Rust or the former combined workspace. The original Python source at `C:\Users\AZI\Desktop\VibeCode\4.pytools-main` is read-only.

## Project boundaries

- Keep PyTools business logic under `pytools-core/`. Excel primitives are maintained in this repository's own `crates/`; no shared path outside this repository, Git submodule, or automatic synchronization is allowed.
- Preserve the product scope and exclusions recorded in `pytools-core/AGENTS.md`, `MIGRATION_PLAN.md`, and `MIGRATION_INVENTORY.md`. Python in `probe/` and test helpers is comparison-only. The isolated openpyxl worker remains a narrow, existing-workbook modification compatibility path; it must not perform PyTools business logic.
- Preserve the local `config.xlsx` as user configuration. It is ignored by Git. Do not add real business workbooks, credentials, or local configuration to commits.
- Rules under `rules/` apply within this independent repository. Product-specific business semantics remain in the PyTools inventory and rules; generic Excel primitives may be changed here without coordinating with Main or Audit.

## Build and test authorization

- Do not run any command that may compile, build, execute, or install Rust code unless the current user request explicitly authorizes that class of operation. This includes `cargo build/check/test/run/clippy/bench/install`, `rustc`, and scripts or tools that invoke them.
- Implementing or reviewing code is not build authorization. `cargo fmt --check` and static inspection are allowed. Do not start a build from a run script.
- `run.bat` and its `run-pytools-dev.bat` compatibility alias launch only the existing `%~dp0dist\pytools.exe`; they never invoke Cargo and must fail clearly if the EXE is absent. `build-release.bat` is the separate, explicit Release build entry, uses `CARGO_TARGET_DIR=F:\RustBuild\pytools`, and copies successful outputs into this project's `dist` directory. `package.bat` packages existing `dist` outputs and must not build. All project BAT files resolve the source root through `%~dp0`; do not point them at Main or a Codex worktree.
- Never run `cargo clean` or remove target caches unless the user explicitly requests cache maintenance.

## Validation and delivery

- Record implementation, unit/smoke-test, Windows real-sample, and Linux/deb validation as separate states. Tests written but not run remain unverified.
- Keep new work incremental and tied to a real PyTools feature. Use the agreed manual feature-acceptance workflow; do not launch a full business audit, full Python parity run, or broad Excel-style comparison unless the user asks. Do not replace business migration with architecture work, broad rewrites, or repeated full-project testing. Do not remove existing business behavior without a replacement that has been tested.
- Keep Linux delivery requirements in force: x86_64, glibc 2.28 compatibility, offline debs, and no requirement for a user-installed Python, Go, Office, WPS, LibreOffice, or JVM. Optional external engines remain explicit and capability-specific.
