//! Fidelity conversion engine (former `lo-convert` sidecar).
//! Office format conversion uses the bundled pywin32 COM worker on Windows.
//! Microsoft Excel is preferred for ordinary Office files and WPS COM is used
//! when Excel is unavailable or for WPS-native formats. LibreOffice conversion
//! is paused and is not selected by production routing.
//!
//! The older LibreOffice batch implementation remains dormant for now.

use anyhow::{anyhow, bail, ensure, Context, Result};
use encoding_rs::GBK;
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId,
    ExcelEditPlan, OpenpyxlWorkerConfig, PreservationRequirement, RangeStyleMode, WorkbookFormat,
};
use excel_reader::{
    read_value_columns_with_iso_dates, read_workbook, CellValue as ReaderCellValue, WorkbookDump,
};
use std::{
    collections::{HashMap, HashSet},
    fs,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    time::{SystemTime, UNIX_EPOCH},
};
#[cfg(windows)]
use std::{
    io::{BufRead, BufReader, BufWriter, Write},
    process::{Child, ChildStdin, ChildStdout},
};
#[cfg(target_os = "linux")]
use std::{
    io::{BufRead, BufReader, BufWriter, Write},
    process::{Child, ChildStdin, ChildStdout},
};

/// Protocol version written by the bundled pywin32 worker. Bumped whenever the
/// JSON-lines request/response contract changes (v2 added ET/ETT and Word).
#[cfg(windows)]
const PYWIN32_PROTOCOL_VERSION: u32 = 3;

/// Lowest FileFormat WPS assigns to its own spreadsheet containers. Values
/// below it come from Excel (56 = XLS, 51 = XLSX, ...), which would mean the
/// `.et`/`.ett` extension was only a rename.
const WPS_NATIVE_FILE_FORMAT_MIN: i32 = 65_521;

/// State scoped to one user conversion request or batch.
pub struct ConversionRuntime {
    #[cfg(windows)]
    pywin32: Option<PythonExcelWorker>,
    #[cfg(target_os = "linux")]
    libreoffice_calculator: Option<LibreOfficeCalcWorker>,
    libreoffice: Option<LibreOfficeBatchSession>,
    /// Test hook: replaces the real `soffice` process execution.
    libreoffice_executor: Option<LibreOfficeExecutor>,
}

impl Default for ConversionRuntime {
    fn default() -> Self {
        Self {
            #[cfg(windows)]
            pywin32: None,
            #[cfg(target_os = "linux")]
            libreoffice_calculator: None,
            libreoffice: None,
            libreoffice_executor: None,
        }
    }
}

/// Outcome of one LibreOffice process execution. Tests can fake this; the
/// real runner spawns `soffice`, records the child PID and waits for exit.
pub(crate) struct LibreOfficeRunOutcome {
    pub process_id: u32,
    pub status_success: bool,
    pub stderr: String,
}

pub(crate) type LibreOfficeExecutor =
    Box<dyn FnMut(&[String], &Path) -> Result<LibreOfficeRunOutcome> + Send>;

#[cfg(target_os = "linux")]
struct LibreOfficeCalcWorker {
    child: Child,
    input: BufWriter<ChildStdin>,
    output: BufReader<ChildStdout>,
}

#[cfg(target_os = "linux")]
impl LibreOfficeCalcWorker {
    fn start() -> Result<Self> {
        let worker = libreoffice_calc_worker_path()?;
        let python = std::env::var_os("PYTOOLS_LO_PYTHON")
            .map(PathBuf::from)
            .unwrap_or_else(|| {
                let system_python = PathBuf::from("/usr/bin/python3");
                if system_python.is_file() {
                    system_python
                } else {
                    PathBuf::from("python3")
                }
            });
        let mut child = Command::new(&python)
            .arg("-u")
            .arg(&worker)
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::inherit())
            .spawn()
            .with_context(|| {
                format!(
                    "启动 LibreOffice 计算 worker 失败（Python：{}，worker：{}）",
                    python.display(),
                    worker.display()
                )
            })?;
        let input = child
            .stdin
            .take()
            .context("LibreOffice 计算 worker stdin 不可用")?;
        let output = child
            .stdout
            .take()
            .context("LibreOffice 计算 worker stdout 不可用")?;
        let mut worker = Self {
            child,
            input: BufWriter::new(input),
            output: BufReader::new(output),
        };
        let ready = worker.read_response("启动 LibreOffice 计算会话")?;
        ensure!(
            ready.get("ok").and_then(serde_json::Value::as_bool) == Some(true)
                && ready.get("ready").and_then(serde_json::Value::as_bool) == Some(true)
                && ready.get("protocol").and_then(serde_json::Value::as_u64) == Some(1),
            "LibreOffice 计算 worker 握手失败：{}",
            ready
                .get("error")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("worker 协议版本或就绪状态不匹配")
        );
        Ok(worker)
    }

    fn recalculate(&mut self, source_copy: &Path) -> Result<()> {
        serde_json::to_writer(
            &mut self.input,
            &serde_json::json!({
                "op": "recalculate",
                "source_copy": source_copy,
            }),
        )
        .context("写入 LibreOffice 重算请求失败")?;
        self.input
            .write_all(b"\n")
            .and_then(|()| self.input.flush())
            .context("发送 LibreOffice 重算请求失败")?;
        let response = self.read_response("读取 LibreOffice 重算结果")?;
        ensure!(
            response.get("ok").and_then(serde_json::Value::as_bool) == Some(true),
            "LibreOffice 重算失败（阶段：{}）：{}",
            response
                .get("stage")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("未知阶段"),
            response
                .get("error")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("worker 未提供错误详情")
        );
        ensure!(
            response
                .get("application")
                .and_then(serde_json::Value::as_str)
                == Some("LibreOffice Calc"),
            "LibreOffice 重算 worker 未确认计算引擎"
        );
        Ok(())
    }

    fn read_response(&mut self, action: &str) -> Result<serde_json::Value> {
        if let Some(status) = self
            .child
            .try_wait()
            .with_context(|| format!("检查 LibreOffice 计算 worker 进程失败：{action}"))?
        {
            bail!("LibreOffice 计算 worker 已退出（状态：{status}）：{action}");
        }
        let mut line = String::new();
        let bytes = self
            .output
            .read_line(&mut line)
            .with_context(|| format!("{action}失败"))?;
        ensure!(bytes > 0, "LibreOffice 计算 worker 未返回结果：{action}");
        serde_json::from_str(&line)
            .with_context(|| format!("LibreOffice 计算 worker 返回的结果不是有效 JSON：{action}"))
    }
}

#[cfg(target_os = "linux")]
impl Drop for LibreOfficeCalcWorker {
    fn drop(&mut self) {
        let _ = writeln!(self.input, "{{\"op\":\"shutdown\"}}");
        let _ = self.input.flush();
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(10);
        loop {
            match self.child.try_wait() {
                Ok(Some(_)) => return,
                Ok(None) if std::time::Instant::now() < deadline => {
                    std::thread::sleep(std::time::Duration::from_millis(25));
                }
                _ => {
                    let _ = self.child.kill();
                    let _ = self.child.wait();
                    return;
                }
            }
        }
    }
}

#[cfg(target_os = "linux")]
fn libreoffice_calc_worker_path() -> Result<PathBuf> {
    if let Some(path) = std::env::var_os("PYTOOLS_LO_CALC_WORKER") {
        let path = PathBuf::from(path);
        ensure!(
            path.is_file(),
            "PYTOOLS_LO_CALC_WORKER 指向的 worker 不存在：{}",
            path.display()
        );
        return Ok(path);
    }
    let executable = std::env::current_exe().context("定位当前 RustTools 程序失败")?;
    let executable_dir = executable
        .parent()
        .context("当前 RustTools 程序没有父目录")?;
    let manifest_worker = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("python")
        .join("libreoffice_calc_worker.py");
    let candidates = [
        executable_dir
            .join("python")
            .join("libreoffice_calc_worker.py"),
        executable_dir.join("../libexec/pytools/openpyxl/worker/libreoffice_calc_worker.py"),
        manifest_worker,
    ];
    candidates
        .iter()
        .find(|path| path.is_file())
        .cloned()
        .with_context(|| {
            format!(
                "未找到 LibreOffice 计算 worker；已检查 {:?}，可用 PYTOOLS_LO_CALC_WORKER 指定路径",
                candidates
            )
        })
}

/// One output file a fake or real runner is expected to have produced.
struct LibreOfficeJob {
    staged_input: PathBuf,
    staged_output: PathBuf,
}

/// Batch-scoped LibreOffice session: a shared isolated profile plus the
/// mapping of already-converted inputs to their staged candidates. The
/// session is created lazily on the first prepared batch and dropped with
/// the runtime; only processes started by this session are ever awaited,
/// and they exit on their own after `--convert-to` finishes.
pub(crate) struct LibreOfficeBatchSession {
    session_dir: PathBuf,
    profile: PathBuf,
    process_ids: Vec<u32>,
    /// `(input, target) -> staged candidate` for successful conversions.
    converted: HashMap<(PathBuf, String), PathBuf>,
    /// Extracted plain text of every Word source document, keyed by the
    /// caller's input path.
    source_texts: HashMap<PathBuf, String>,
    /// Extracted plain text of every converted candidate, keyed by
    /// `(input, target)` — the per-file temp output path is not knowable
    /// at batch-planning time.
    candidate_texts: HashMap<(PathBuf, String), String>,
    /// Every input the session took responsibility for, successful or not.
    covered: HashSet<(PathBuf, String)>,
    /// First fatal session error (engine missing, conversion failed).
    error: Option<String>,
}

impl Drop for LibreOfficeBatchSession {
    fn drop(&mut self) {
        // Processes were awaited inside the session; only temporary files
        // remain. Never terminate processes by name or PID here.
        let _ = fs::remove_dir_all(&self.session_dir);
    }
}

impl LibreOfficeBatchSession {
    fn create() -> Result<Self> {
        let session_dir = unique_lo_session_dir()?;
        let profile = session_dir.join("lo-profile");
        libreoffice_user_installation_arg(&profile)?;
        Ok(Self {
            session_dir,
            profile,
            process_ids: Vec::new(),
            converted: HashMap::new(),
            source_texts: HashMap::new(),
            candidate_texts: HashMap::new(),
            covered: HashSet::new(),
            error: None,
        })
    }

    fn run_libreoffice_processes(
        &mut self,
        executor: &mut Option<LibreOfficeExecutor>,
        stages: Vec<(String, PathBuf, Vec<LibreOfficeJob>)>,
    ) -> Result<()> {
        for (filter, outputs_dir, jobs) in stages {
            if jobs.is_empty() {
                continue;
            }
            fs::create_dir_all(&outputs_dir)
                .context("create LibreOffice session output directory")?;
            let input_paths = jobs
                .iter()
                .map(|job| job.staged_input.clone())
                .collect::<Vec<_>>();
            let mut arguments = vec![
                "--headless".to_owned(),
                "--nologo".to_owned(),
                "--nodefault".to_owned(),
                "--nolockcheck".to_owned(),
                libreoffice_user_installation_arg(&self.profile)?,
                "--convert-to".to_owned(),
                filter,
                "--outdir".to_owned(),
                outputs_dir.to_string_lossy().into_owned(),
            ];
            arguments.extend(
                input_paths
                    .iter()
                    .map(|path| path.to_string_lossy().into_owned()),
            );
            let outcome = if let Some(executor) = executor.as_mut() {
                executor(&arguments, &outputs_dir)?
            } else {
                run_soffice_process(&arguments)?
            };
            self.process_ids.push(outcome.process_id);
            // Process-level failure is fatal for the whole stage; a missing
            // per-file output stays a per-file failure resolved through
            // `libreoffice_batch_outcome` so the remaining files still work.
            ensure!(
                outcome.status_success,
                "LibreOffice conversion failed ({}): {}",
                outcome.process_id,
                outcome.stderr
            );
        }
        Ok(())
    }
}

fn run_soffice_process(arguments: &[String]) -> Result<LibreOfficeRunOutcome> {
    let mut command = Command::new("soffice");
    command
        .args(arguments)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    let mut child = command.spawn().with_context(|| {
        "该功能需要外部引擎: LibreOffice is unavailable; install soffice to use this conversion"
    })?;
    let process_id = child.id();
    let output = child
        .wait_with_output()
        .context("wait for LibreOffice conversion process")?;
    Ok(LibreOfficeRunOutcome {
        process_id,
        status_success: output.status.success(),
        stderr: String::from_utf8_lossy(&output.stderr).into_owned(),
    })
}

/// Resolve the effective fidelity engine exactly like
/// [`run_fidelity_conversion_with_runtime`] does, so the batch planner and
/// the per-file path cannot drift apart.
///
/// There is no Rust COM channel anymore: every Windows conversion goes
/// through the bundled pywin32 worker. `com` is kept as a compatibility
/// alias for `pywin32`.
pub(crate) fn select_fidelity_engine(engine: &str) -> &'static str {
    if engine == "auto" {
        "pywin32"
    } else if engine == "lo" {
        "lo"
    } else {
        "pywin32"
    }
}

fn select_conversion_engine(
    engine: &str,
    source_ext: &str,
    target: &str,
    is_windows: bool,
) -> &'static str {
    if engine != "auto" {
        return match engine {
            "lo" => "lo",
            "com" | "pywin32" => "pywin32",
            _ => "pywin32",
        };
    }
    // Linux：电子表格 .xls/.xlsm → .xlsx 由 LibreOffice 另存。2026-11-09 用
    // 11 个真实/合成 .xls 实测：公式文本与 COM 另存零差异、值零差异（布尔
    // 结果缓存表示除外），详见 docs/BETTEROFFICE_VENDOR_INTEGRATION.md §5.10。
    // Word 与 WPS 原生格式仍仅 Windows pywin32。
    if !is_windows && target == "xlsx" && matches!(source_ext, "xls" | "xlsm") {
        return "lo";
    }
    // Windows：Excel/WPS 表格与 Word/WPS 文档共用批量 pywin32 worker。
    "pywin32"
}

/// Whether this source would run through LibreOffice under the requested
/// engine, using the same source/target matrix the per-file path accepts.
pub(crate) fn libreoffice_batch_request(source_ext: &str, target: &str, engine: &str) -> bool {
    let _ = (source_ext, target, engine);
    // Office conversion is currently COM-only. Keep the batch session code
    // available as dormant infrastructure, but never schedule `soffice`.
    false
}

/// 已把模板校验公式复制进工作副本的阶段结果；组级批量编排在并行计算前用它
/// 记录每本副本的复制统计，无需重复扫描。
#[derive(Debug, Clone)]
pub(crate) struct PreparedValidationCopy {
    pub copied_formula_cells: usize,
    pub copied_template_sheets: Vec<String>,
    pub dependency_copy_engine: String,
}

/// 把模板的校验区域公式（含必要的辅助工作表）复制进工作副本。
///
/// 从 [`ConversionRuntime::prepare_named_validation_copy`] 抽出，供批量编排在
/// 并行阶段直接调用：只读模板、只写自己的副本，没有共享可变状态。
fn copy_template_regions(
    template: &Path,
    source_copy: &Path,
    copy_requests: &[ooxml_patch::CellRegionCopyRequest],
) -> Result<(ooxml_patch::CellRegionCopyReport, String)> {
    match ooxml_patch::copy_cell_region(template, source_copy, copy_requests) {
        Ok(copied) => Ok((copied, "OOXML".to_owned())),
        Err(primary) if format!("{primary:#}").contains("OPENPYXL_FALLBACK_REQUIRED") => {
            let dependency_sheets = ooxml_patch::missing_formula_dependency_sheets(
                template,
                source_copy,
                copy_requests,
            )
            .context("发现 OOXML 不支持的辅助表关系，但无法解析完整公式依赖")?;
            ensure!(
                !dependency_sheets.is_empty(),
                "OOXML 请求 Openpyxl 回退，但没有找到缺失的辅助工作表"
            );
            if let Err(fallback) = ConversionRuntime::import_validation_dependencies_openpyxl(
                template,
                source_copy,
                &dependency_sheets,
            ) {
                return Err(anyhow!(
                    "OOXML 定点复制辅助表失败：{primary:#}；受保护的 Openpyxl 回退也未通过：{fallback:#}"
                ));
            }
            let mut copied = ooxml_patch::copy_cell_region(template, source_copy, copy_requests)
                .with_context(|| {
                    format!(
                        "Openpyxl 已补装辅助表，但随后用 OOXML 复制校验区域公式失败（模板：{}）",
                        template.display()
                    )
                })?;
            copied.sheets_copied = dependency_sheets;
            Ok((copied, "Openpyxl 回退 + OOXML".to_owned()))
        }
        Err(error) => Err(error).with_context(|| {
            format!(
                "把模板校验公式复制到工作副本失败（模板：{}）",
                template.display()
            )
        }),
    }
}

impl ConversionRuntime {
    /// Run the dedicated 5-1 legacy XLS operation through Excel/WPS COM.
    /// Callers pass an isolated target candidate; the source is committed only
    /// after the worker succeeds and the caller rechecks its source snapshot.
    pub fn village_split_com(
        &mut self,
        current: &Path,
        previous: &Path,
        candidate: &Path,
    ) -> Result<serde_json::Value> {
        #[cfg(windows)]
        {
            ensure!(
                [current, previous, candidate].iter().all(|path| path
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"))),
                "5-1 COM 路由要求本期、上期和目标候选均为 .xls"
            );
            if self.pywin32.is_none() {
                self.pywin32 = Some(PythonExcelWorker::start()?);
            }
            self.pywin32.as_mut().expect("worker initialized").request(
                &serde_json::json!({
                    "op": "village_split_com",
                    "current": current,
                    "previous": previous,
                    "candidate": candidate,
                }),
                "5-1 Excel/WPS COM 拆分",
            )
        }
        #[cfg(not(windows))]
        {
            let _ = (current, previous, candidate);
            bail!("5-1 的 .xls 处理需要 Windows Excel/WPS COM")
        }
    }

    /// Run the dedicated 5-7 legacy XLS operation on an isolated current-file
    /// candidate. The caller owns snapshot validation and safe replacement.
    pub fn rural_loan_com(
        &mut self,
        current_candidate: &Path,
        previous: &Path,
    ) -> Result<serde_json::Value> {
        #[cfg(windows)]
        {
            ensure!(
                [current_candidate, previous].iter().all(|path| path
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"))),
                "5-7 COM 路由要求本期候选和上期工作簿均为 .xls"
            );
            if self.pywin32.is_none() {
                self.pywin32 = Some(PythonExcelWorker::start()?);
            }
            self.pywin32.as_mut().expect("worker initialized").request(
                &serde_json::json!({
                    "op": "rural_loan_com",
                    "current_candidate": current_candidate,
                    "previous": previous,
                }),
                "5-7 Excel/WPS COM 涉农贷款修正",
            )
        }
        #[cfg(not(windows))]
        {
            let _ = (current_candidate, previous);
            bail!("5-7 的 .xls 处理需要 Windows Excel/WPS COM")
        }
    }

    /// Apply Rust-planned 5-2 institution-name updates to an isolated legacy
    /// `.xls` candidate through the current Excel/WPS COM batch session.
    pub fn institution_normalize_com(
        &mut self,
        candidate: &Path,
        updates: &[(String, String, String)],
    ) -> Result<serde_json::Value> {
        #[cfg(windows)]
        {
            ensure!(
                candidate
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xls")),
                "5-2 COM 路由要求 .xls 候选"
            );
            ensure!(!updates.is_empty(), "5-2 COM 路由没有单元格更新");
            if self.pywin32.is_none() {
                self.pywin32 = Some(PythonExcelWorker::start()?);
            }
            let updates = updates
                .iter()
                .map(|(sheet, cell, value)| {
                    serde_json::json!({"sheet": sheet, "cell": cell, "value": value})
                })
                .collect::<Vec<_>>();
            self.pywin32.as_mut().expect("worker initialized").request(
                &serde_json::json!({
                    "op": "institution_normalize_com",
                    "candidate": candidate,
                    "updates": updates,
                }),
                "5-2 Excel/WPS COM 机构名称规范化",
            )
        }
        #[cfg(not(windows))]
        {
            let _ = (candidate, updates);
            bail!("5-2 的 .xls 处理需要 Windows Excel/WPS COM")
        }
    }

    /// Apply Rust-planned 2-1 red-fill duplicate marks to an isolated legacy
    /// `.xls` candidate through Excel/WPS COM.
    pub fn duplicate_mark_com(
        &mut self,
        candidate: &Path,
        marks: &[(String, String)],
    ) -> Result<serde_json::Value> {
        #[cfg(windows)]
        {
            ensure!(
                candidate
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xls")),
                "2-1 COM 路由要求 .xls 候选"
            );
            ensure!(!marks.is_empty(), "2-1 COM 路由没有重复标记");
            if self.pywin32.is_none() {
                self.pywin32 = Some(PythonExcelWorker::start()?);
            }
            let marks = marks
                .iter()
                .map(|(sheet, cell)| serde_json::json!({"sheet": sheet, "cell": cell}))
                .collect::<Vec<_>>();
            self.pywin32.as_mut().expect("worker initialized").request(
                &serde_json::json!({
                    "op": "duplicate_mark_com",
                    "candidate": candidate,
                    "marks": marks,
                }),
                "2-1 Excel/WPS COM 重复标记",
            )
        }
        #[cfg(not(windows))]
        {
            let _ = (candidate, marks);
            bail!("2-1 的 .xls 处理需要 Windows Excel/WPS COM")
        }
    }

    /// Delete Rust-planned whole rows on an isolated legacy `.xls` candidate
    /// through Excel/WPS COM. Shared by 2-2 (duplicate deletion) and 5-3
    /// (foreign-bank row removal).
    pub fn delete_rows_com(
        &mut self,
        candidate: &Path,
        deletions: &[(String, u32, u32)],
    ) -> Result<serde_json::Value> {
        #[cfg(windows)]
        {
            ensure!(
                candidate
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xls")),
                "删行 COM 路由要求 .xls 候选"
            );
            ensure!(!deletions.is_empty(), "删行 COM 路由没有删行计划");
            if self.pywin32.is_none() {
                self.pywin32 = Some(PythonExcelWorker::start()?);
            }
            let deletions = deletions
                .iter()
                .map(|(sheet, first_row, count)| {
                    serde_json::json!({"sheet": sheet, "first_row": first_row, "count": count})
                })
                .collect::<Vec<_>>();
            self.pywin32.as_mut().expect("worker initialized").request(
                &serde_json::json!({
                    "op": "delete_rows_com",
                    "candidate": candidate,
                    "deletions": deletions,
                }),
                "Excel/WPS COM 删行",
            )
        }
        #[cfg(not(windows))]
        {
            let _ = (candidate, deletions);
            bail!("删行的 .xls 处理需要 Windows Excel/WPS COM")
        }
    }

    /// Apply PyCode's 5-4 COM header algorithm to an isolated legacy `.xls`
    /// candidate through Excel/WPS COM.
    pub fn fx_header_com(&mut self, candidate: &Path) -> Result<serde_json::Value> {
        #[cfg(windows)]
        {
            ensure!(
                candidate
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xls")),
                "5-4 COM 路由要求 .xls 候选"
            );
            if self.pywin32.is_none() {
                self.pywin32 = Some(PythonExcelWorker::start()?);
            }
            self.pywin32.as_mut().expect("worker initialized").request(
                &serde_json::json!({"op": "fx_header_com", "candidate": candidate}),
                "5-4 Excel/WPS COM 外汇页眉",
            )
        }
        #[cfg(not(windows))]
        {
            let _ = candidate;
            bail!("5-4 的 .xls 处理需要 Windows Excel/WPS COM")
        }
    }

    /// Append typed 2-D rows to a sheet in an isolated legacy `.xls` target
    /// candidate through Excel/WPS COM (2-3 append to .xls target).
    pub fn append_rows_com(
        &mut self,
        candidate: &Path,
        sheet: &str,
        start_row: u32,
        rows: &[Vec<serde_json::Value>],
    ) -> Result<serde_json::Value> {
        #[cfg(windows)]
        {
            ensure!(
                candidate
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xls")),
                "追加 COM 路由要求 .xls 候选"
            );
            ensure!(!rows.is_empty(), "追加 COM 路由没有数据行");
            if self.pywin32.is_none() {
                self.pywin32 = Some(PythonExcelWorker::start()?);
            }
            self.pywin32.as_mut().expect("worker initialized").request(
                &serde_json::json!({
                    "op": "append_rows_com",
                    "candidate": candidate,
                    "sheet": sheet,
                    "start_row": start_row,
                    "rows": rows,
                }),
                "2-3 Excel/WPS COM 去重追加",
            )
        }
        #[cfg(not(windows))]
        {
            let _ = (candidate, sheet, start_row, rows);
            bail!("2-3 的 .xls 目标追加需要 Windows Excel/WPS COM")
        }
    }

    pub(crate) fn rename_sheets_in_candidate(
        &mut self,
        candidate: &Path,
        changes: &[(String, String)],
    ) -> Result<()> {
        #[cfg(windows)]
        {
            if self.pywin32.is_none() {
                self.pywin32 = Some(PythonExcelWorker::start()?);
            }
            let changes = changes.iter().map(|(old_name, new_name)| {
                serde_json::json!({"old_name": old_name, "new_name": new_name})
            }).collect::<Vec<_>>();
            let response = self.pywin32.as_mut().expect("worker initialized").request(
                &serde_json::json!({"op": "rename_sheets", "candidate": candidate.to_string_lossy(), "changes": changes}),
                "工作表改名",
            )?;
            ensure!(
                response["validation"]["verified"].as_bool() == Some(true),
                "工作表改名未通过回读校验"
            );
            Ok(())
        }
        #[cfg(not(windows))]
        {
            let _ = (candidate, changes);
            bail!("工作表改名及公式依赖更新依赖随包 openpyxl worker；该运行时目前只在 Windows 构建中提供")
        }
    }

    /// Rename a legacy BIFF `.xls` candidate through the dedicated Excel/WPS
    /// COM operation used by the old PyCode route. This is intentionally
    /// limited to 4-5 sheet renaming; it is not a general workbook editor.
    pub(crate) fn rename_sheets_com_in_candidate(
        &mut self,
        candidate: &Path,
        changes: &[(String, String)],
    ) -> Result<String> {
        #[cfg(windows)]
        {
            if self.pywin32.is_none() {
                self.pywin32 = Some(PythonExcelWorker::start()?);
            }
            let changes = changes
                .iter()
                .map(|(old_name, new_name)| {
                    serde_json::json!({"old_name": old_name, "new_name": new_name})
                })
                .collect::<Vec<_>>();
            let response = self.pywin32.as_mut().expect("worker initialized").request(
                &serde_json::json!({
                    "op": "rename_sheets_com",
                    "candidate": candidate.to_string_lossy(),
                    "changes": changes,
                }),
                "Excel/WPS COM 工作表改名",
            )?;
            ensure!(
                response["validation"]["verified"].as_bool() == Some(true),
                "Excel/WPS COM 工作表改名未通过回读校验"
            );
            Ok(response["application"]
                .as_str()
                .unwrap_or("Excel/WPS COM")
                .to_owned())
        }
        #[cfg(not(windows))]
        {
            let _ = (candidate, changes);
            bail!("BIFF .xls 工作表改名需要 Windows Excel/WPS COM")
        }
    }

    /// Copy template validation formulas into a disposable source-workbook copy,
    /// recalculate it through the platform calculation worker, and return
    /// matching formula results. One runtime reuses the same calculation
    /// session for every workbook in a batch.
    ///
    /// The formulas are copied here in Rust (OOXML), so the COM host only ever
    /// opens the caller's working copy. Opening the template through COM as
    /// well doubled the failure surface: the template usually lives in a user
    /// directory, and any state Excel dislikes there aborted the whole file.
    /// 校验区域重算的第一阶段：安全检查 + 把模板校验公式复制进工作副本。
    ///
    /// 组级批量编排会先对所有副本调用本方法（可并行：只写各自副本、只读模板），
    /// 再统一并行跑 betteroffice，最后只把拒绝项交给平台引擎。
    pub(crate) fn prepare_named_validation_copy(
        template: &Path,
        source_copy: &Path,
        regions: &[NamedValidationRegionRequest],
    ) -> Result<PreparedValidationCopy> {
        ensure!(!regions.is_empty(), "校验区域请求不能为空");
        ensure!(
            source_copy
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
                }),
            "校验区域重算需要 .xlsx/.xlsm 工作副本（当前 {}）",
            source_copy.display()
        );
        let reference_scan = ooxml_patch::scan_external_workbook_references(template)
            .with_context(|| format!("扫描模板外部引用失败：{}", template.display()))?;
        ensure!(
            reference_scan.is_clean(),
            "模板「{}」含外部工作簿引用，校验区域公式复制到工作副本后会变成悬空引用，\
             Excel 将拒绝打开。请在模板中删除外部链接并把公式改为引用本地工作表后重试。\n详情：{}",
            template.display(),
            reference_scan.describe()
        );
        let copy_requests = regions
            .iter()
            .map(|region| ooxml_patch::CellRegionCopyRequest {
                source_sheet: region.template_sheet.clone(),
                target_sheet: region.source_sheet.clone(),
                address: region.address.clone(),
            })
            .collect::<Vec<_>>();
        let (copied, dependency_copy_engine) =
            copy_template_regions(template, source_copy, &copy_requests)?;
        // shared formula 的 ref 起点归一必须发生在**任何引擎打开该副本之前**，
        // 且与路由无关：Excel/COM/LibreOffice 打开保存都会把 ref 起点归一到
        // 主格，而 betteroffice 不重写公式文本。若只修 COM 分支，两条路由的
        // 副本口径就不一致，公式文本比对会报"重算改变公式"的假阳性。
        // 该修复也顺带处理表名引号与矛盾冻结窗格，失败不阻断计算。
        ooxml_patch::repair_excel_com_open_issues(source_copy);
        Ok(PreparedValidationCopy {
            copied_formula_cells: copied.cells_with_formula,
            copied_template_sheets: copied.sheets_copied,
            dependency_copy_engine,
        })
    }

    /// 第二阶段：只做计算（副本的公式已由 [`Self::prepare_named_validation_copy`]
    /// 复制好）。引擎路由与回退语义与单文件路径完全一致。
    pub(crate) fn compute_named_validation_regions(
        &mut self,
        source_copy: &Path,
        regions: &[NamedValidationRegionRequest],
        prepared: &PreparedValidationCopy,
    ) -> Result<NamedValidationCalculation> {
        if std::env::var_os("PYTOOLS_CONVERT_ENGINES_DISABLED").is_some() {
            bail!("校验区域计算需要 Excel/WPS COM 或 LibreOffice；当前环境已禁用外部引擎");
        }
        self.compute_prepared_copy(source_copy, regions, prepared, None)
    }

    pub(crate) fn recalculate_named_validation_regions(
        &mut self,
        template: &Path,
        source_copy: &Path,
        regions: &[NamedValidationRegionRequest],
    ) -> Result<NamedValidationCalculation> {
        let prepared = Self::prepare_named_validation_copy(template, source_copy, regions)?;
        self.compute_named_validation_regions(source_copy, regions, &prepared)
    }

    /// 组级批量编排已用 betteroffice 算完的副本 → 回读缓存并提取校验区域。
    ///
    /// 与 [`Self::compute_prepared_copy`] 的 betteroffice 成功分支等价，只是
    /// 重算在批量阶段已完成，这里不再重算第二遍。
    pub(crate) fn calculation_from_batch_success(
        source_copy: &Path,
        regions: &[NamedValidationRegionRequest],
        prepared: &PreparedValidationCopy,
        outcome: &crate::bo_calc::BoCalcOutcome,
    ) -> Result<NamedValidationCalculation> {
        let workbook = read_workbook(source_copy)
            .context("betteroffice 批量重算后使用 Calamine 回读公式缓存失败")?;
        let cells = extract_named_validation_cells(&workbook, regions)?;
        Ok(NamedValidationCalculation {
            application: outcome.application.to_owned(),
            cells,
            copied_formula_cells: prepared.copied_formula_cells,
            copied_template_sheets: prepared.copied_template_sheets.clone(),
            dependency_copy_engine: prepared.dependency_copy_engine.clone(),
        })
    }

    /// 已复制好公式的副本 → 计算 + 回退。`forced_fallback` 用于组级批量编排：
    /// betteroffice 已在批量阶段拒绝过的副本直接带原因进平台引擎，避免重算两遍。
    pub(crate) fn compute_prepared_copy(
        &mut self,
        source_copy: &Path,
        regions: &[NamedValidationRegionRequest],
        prepared: &PreparedValidationCopy,
        forced_fallback: Option<String>,
    ) -> Result<NamedValidationCalculation> {
        let dependency_copy_engine = prepared.dependency_copy_engine.clone();
        let copied = ooxml_patch::CellRegionCopyReport {
            cells_with_formula: prepared.copied_formula_cells,
            sheets_copied: prepared.copied_template_sheets.clone(),
            ..ooxml_patch::CellRegionCopyReport::default()
        };
        if let Some(reason) = forced_fallback {
            // 批量阶段已判定 betteroffice 不可采用：直接交给平台引擎整簿重算。
            return self.recalculate_named_validation_regions_platform(
                source_copy,
                regions,
                copied,
                dependency_copy_engine,
                Some(reason),
            );
        }
        // 配置优先级：config.xlsx 全局配置（校验计算* 键）> 环境变量 > 默认。
        let bo_config = crate::bo_calc::BoCalcConfig::load(
            crate::config::default_config_workbook_path().as_deref(),
        );
        if bo_config.enabled {
            match crate::bo_calc::recalculate(source_copy, &bo_config) {
                Ok(outcome) => {
                    let workbook = read_workbook(source_copy)
                        .context("betteroffice 重算后使用 Calamine 回读公式缓存失败")?;
                    let cells = extract_named_validation_cells(&workbook, regions)?;
                    return Ok(NamedValidationCalculation {
                        application: outcome.application.to_owned(),
                        cells,
                        copied_formula_cells: prepared.copied_formula_cells,
                        copied_template_sheets: prepared.copied_template_sheets.clone(),
                        dependency_copy_engine,
                    });
                }
                Err(refusal) => {
                    let refusal_reason = format!("{:#}", refusal);
                    if !bo_config.fallback_available() {
                        bail!(
                            "betteroffice 校验区域重算不可采用且未配置可用回退：{refusal_reason}"
                        );
                    }
                    return self.recalculate_named_validation_regions_platform(
                        source_copy,
                        regions,
                        copied,
                        dependency_copy_engine,
                        Some(format!("betteroffice 回退（{refusal_reason}）")),
                    );
                }
            }
        }
        self.recalculate_named_validation_regions_platform(
            source_copy,
            regions,
            copied,
            dependency_copy_engine,
            None,
        )
    }

    /// The platform engine branch: Excel/WPS COM on Windows, LibreOffice on
    /// Linux. Also the fallback target when betteroffice refuses a workbook;
    /// `fallback_reason` then records why, so the run log keeps one trace.
    fn recalculate_named_validation_regions_platform(
        &mut self,
        source_copy: &Path,
        regions: &[NamedValidationRegionRequest],
        copied: ooxml_patch::CellRegionCopyReport,
        dependency_copy_engine: String,
        fallback_reason: Option<String>,
    ) -> Result<NamedValidationCalculation> {
        let decorate = |application: String| match &fallback_reason {
            Some(reason) => format!("{application}；{reason}"),
            None => application,
        };
        #[cfg(windows)]
        {
            // 表名引号/冻结窗格/shared formula ref 的修复已在
            // prepare_named_validation_copy 阶段对工作副本统一做过（路由无关），
            // 这里不再重复。
            if self.pywin32.is_none() {
                self.pywin32 = Some(PythonExcelWorker::start()?);
            }
            let calculation = self
                .pywin32
                .as_mut()
                .expect("pywin32 worker initialized")
                .recalculate_named_validation_regions(source_copy, regions)?;
            Ok(NamedValidationCalculation {
                application: decorate(calculation.application),
                cells: calculation.cells,
                copied_formula_cells: copied.cells_with_formula,
                copied_template_sheets: copied.sheets_copied,
                dependency_copy_engine,
            })
        }
        #[cfg(not(windows))]
        {
            #[cfg(target_os = "linux")]
            {
                if self.libreoffice_calculator.is_none() {
                    self.libreoffice_calculator = Some(LibreOfficeCalcWorker::start()?);
                }
                self.libreoffice_calculator
                    .as_mut()
                    .expect("LibreOffice calculator initialized")
                    .recalculate(source_copy)?;
                let workbook = read_workbook(source_copy)
                    .context("LibreOffice 重算后使用 Calamine 回读公式缓存失败")?;
                let cells = extract_named_validation_cells(&workbook, regions)?;
                Ok(NamedValidationCalculation {
                    application: decorate("LibreOffice Calc".to_owned()),
                    cells,
                    copied_formula_cells: copied.cells_with_formula,
                    copied_template_sheets: copied.sheets_copied,
                    dependency_copy_engine,
                })
            }
            #[cfg(not(target_os = "linux"))]
            {
                let _ = (copied, decorate);
                bail!("校验区域公式重算仅在 Windows（Excel/WPS COM）和 Linux（LibreOffice）提供")
            }
        }
    }

    /// Use the existing Openpyxl plan executor only when the narrow OOXML
    /// worksheet copier cannot carry an auxiliary sheet relationship. The
    /// target here is already an isolated working copy; the backend still
    /// creates and validates a separate candidate, and its merged-payload
    /// preflight refuses workbooks that Openpyxl would alter lossily.
    fn import_validation_dependencies_openpyxl(
        template: &Path,
        working_copy: &Path,
        dependency_sheets: &[String],
    ) -> Result<()> {
        let binary =
            std::env::current_exe().context("定位 rustTools 程序以解析 Openpyxl 运行时")?;
        let config = OpenpyxlWorkerConfig::resolve_runtime(&binary, "pytools")
            .context("解析随包 Openpyxl Worker 运行时")?;
        Self::import_validation_dependencies_with_config(
            template,
            working_copy,
            dependency_sheets,
            config,
        )
    }

    fn import_validation_dependencies_with_config(
        template: &Path,
        working_copy: &Path,
        dependency_sheets: &[String],
        config: OpenpyxlWorkerConfig,
    ) -> Result<()> {
        ensure!(
            !dependency_sheets.is_empty(),
            "Openpyxl 回退没有需要补装的辅助表"
        );
        ensure!(
            template.extension().and_then(|value| value.to_str()) == Some("xlsx")
                && working_copy.extension().and_then(|value| value.to_str()) == Some("xlsx"),
            "受保护的 Openpyxl 辅助表回退目前仅支持 .xlsx 模板和工作副本"
        );

        let template_dump = read_workbook(template).context("Openpyxl 回退读取模板工作簿")?;
        let target_profile = excel_editor::inspect_xlsx_profile(working_copy)
            .context("Openpyxl 回退检查工作副本 OOXML 结构")?;
        ensure!(
            target_profile.format == WorkbookFormat::Xlsx && target_profile.inventory_complete,
            "Openpyxl 回退要求工作副本为结构清单完整的 .xlsx 工作簿"
        );

        let mut operations = Vec::new();
        for sheet_name in dependency_sheets {
            let source_sheet = template_dump
                .sheets
                .iter()
                .find(|sheet| sheet.name.eq_ignore_ascii_case(sheet_name))
                .with_context(|| format!("Openpyxl 回退模板缺少辅助工作表 {sheet_name:?}"))?;
            let source_range = sheet_used_range(source_sheet);
            operations.push(EditOperation::CreateSheet {
                name: sheet_name.clone(),
            });
            operations.push(EditOperation::ImportRange {
                source_workbook: template.to_path_buf(),
                source_sheet: source_sheet.name.clone(),
                source_range,
                target_sheet: sheet_name.clone(),
                target_start: "A1".to_owned(),
                copy_values: true,
                copy_formulas: true,
                style_mode: RangeStyleMode::None,
                copy_comments: false,
                copy_merges: true,
                copy_row_heights: false,
                copy_column_widths: false,
                copy_row_hidden: false,
                copy_column_hidden: false,
                copy_outline_levels: false,
                copy_all_row_dimensions: false,
                copy_all_column_dimensions: false,
                materialize_blank_cells: false,
                hide_zero: false,
            });
        }

        let preserve: std::collections::BTreeSet<PreservationRequirement> = [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::ConditionalFormatting,
            PreservationRequirement::DataValidations,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::PageSettings,
            PreservationRequirement::HeaderFooterText,
            PreservationRequirement::Relationships,
            PreservationRequirement::SheetIdentityAndOrder,
            PreservationRequirement::FreezePanes,
            PreservationRequirement::AutoFilter,
        ]
        .into_iter()
        .collect();
        let request_id = format!(
            "validation-openpyxl-{}-{}",
            std::process::id(),
            SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos()
        );
        let output = unique_validation_fallback_output(working_copy)?;
        let plan = ExcelEditPlan {
            request_id,
            input: working_copy.to_path_buf(),
            output: output.clone(),
            format: WorkbookFormat::Xlsx,
            workbook: target_profile.clone(),
            operations,
            preserve: preserve.clone(),
            drop_allowed: Default::default(),
        };
        plan.validate().context("校验 Openpyxl 辅助表回退计划")?;

        let platform = crate::inplace_edit::platform_id();
        let capability = BackendCapability {
            engine: EngineId::OpenpyxlWorker,
            status: CapabilityStatus::Passed,
            format: WorkbookFormat::Xlsx,
            operations: plan.operations.iter().map(EditOperation::kind).collect(),
            supported_features: target_profile.features_present,
            preserved: preserve,
            platform: platform.clone(),
            evidence: Some(CapabilityEvidence {
                sample_id: "named-validation-auxiliary-sheet-fallback".to_owned(),
                report_path: "pytools-core/src/conversion_engine.rs".to_owned(),
                adapter_version: format!("pytools/{}", env!("CARGO_PKG_VERSION")),
                engine_version: "bundled openpyxl worker".to_owned(),
                platform: platform.clone(),
            }),
        };
        let original_hash = crate::file_commit::sha256_file(working_copy)?;
        let execution = match crate::edit_plan::apply_plan_with_openpyxl_config(
            &plan, capability, config, &platform,
        ) {
            Ok(execution) => execution,
            Err(error) => {
                let _ = fs::remove_file(&output);
                return Err(error).context("通过 Openpyxl 受保护回退复制模板辅助工作表");
            }
        };
        ensure!(
            execution.engine == EngineId::OpenpyxlWorker,
            "辅助工作表回退实际使用了非 Openpyxl 引擎：{}",
            execution.engine
        );
        let commit = match crate::file_commit::replace_target_from_candidate(
            &output,
            working_copy,
            &original_hash,
        ) {
            Ok(commit) => commit,
            Err(error) => {
                let _ = fs::remove_file(&output);
                return Err(error).context("将已验证的辅助工作表候选替换到隔离工作副本");
            }
        };
        let _ = fs::remove_file(&output);
        if let Some(backup) = commit.backup_path {
            bail!(
                "Openpyxl 回退已完成，但隔离工作副本的回滚备份未能清理：{}；{}",
                backup.display(),
                commit.cleanup_warning.unwrap_or_default()
            );
        }
        Ok(())
    }

    #[cfg(windows)]
    fn spreadsheet_pywin32(&mut self, input: &Path, output: &Path) -> Result<ComExecution> {
        if self.pywin32.is_none() {
            self.pywin32 = Some(PythonExcelWorker::start()?);
        }
        let converted = self
            .pywin32
            .as_mut()
            .expect("pywin32 worker initialized")
            .convert(input, output)?;
        Ok(ComExecution {
            source_format: converted
                .validation
                .get("source_file_format")
                .and_then(serde_json::Value::as_i64)
                .and_then(|value| i32::try_from(value).ok()),
            source_extension: converted
                .validation
                .get("source_extension")
                .and_then(serde_json::Value::as_str)
                .map(str::to_owned),
            bridge: "pywin32",
            application: Some(converted.application),
            process_id: converted.process_id,
            worker_process_id: Some(converted.worker_process_id),
            validation: Some(converted.validation),
        })
    }

    #[cfg(windows)]
    fn word_pywin32(&mut self, input: &Path, output: &Path) -> Result<ComExecution> {
        if self.pywin32.is_none() {
            self.pywin32 = Some(PythonExcelWorker::start()?);
        }
        let converted = self
            .pywin32
            .as_mut()
            .expect("pywin32 worker initialized")
            .convert_word(input, output)?;
        Ok(ComExecution {
            source_format: None,
            source_extension: None,
            bridge: "pywin32",
            application: Some(converted.application),
            process_id: converted.process_id,
            worker_process_id: Some(converted.worker_process_id),
            validation: Some(converted.validation),
        })
    }

    #[cfg(windows)]
    fn extract_word_text_pywin32(&mut self, input: &Path) -> Result<String> {
        if self.pywin32.is_none() {
            self.pywin32 = Some(PythonExcelWorker::start()?);
        }
        self.pywin32
            .as_mut()
            .expect("pywin32 worker initialized")
            .extract_word_text(input)
    }

    /// Stage and convert every requested file through one LibreOffice
    /// session: one shared profile, one `soffice` process per stage
    /// (main conversion; Word text extraction), zero per-file process
    /// starts. Inputs are copied to uniquely named disposable files first,
    /// so the session never touches caller-owned paths and same-stem
    /// outputs cannot collide. Called once per batch before the per-file
    /// loop; a no-op when the batch has no LibreOffice jobs.
    pub(crate) fn prepare_libreoffice_batch(&mut self, requests: &[(PathBuf, String)]) {
        if requests.is_empty() || self.libreoffice.is_some() {
            return;
        }
        if std::env::var_os("PYTOOLS_CONVERT_ENGINES_DISABLED").is_some() {
            // Keep the test override authoritative: no session, no process.
            return;
        }
        let unique_requests = {
            let mut seen = HashSet::new();
            requests
                .iter()
                .filter(|(input, target)| seen.insert((input.clone(), target.clone())))
                .cloned()
                .collect::<Vec<_>>()
        };
        let mut session = match LibreOfficeBatchSession::create() {
            Ok(session) => session,
            Err(error) => {
                // Session setup failure must surface per file, not switch engines.
                if let Some(mut runtime_session) = self.libreoffice.take() {
                    runtime_session.covered.extend(
                        unique_requests
                            .iter()
                            .map(|(input, target)| (input.clone(), target.clone())),
                    );
                    runtime_session.error = Some(format!("{error:#}"));
                    self.libreoffice = Some(runtime_session);
                }
                return;
            }
        };
        for (input, target) in &unique_requests {
            session.covered.insert((input.clone(), target.clone()));
        }
        let mut executor = self.libreoffice_executor.take();
        let outcome =
            Self::convert_batch_with_session(&mut session, &unique_requests, &mut executor);
        if let Err(error) = outcome {
            session.error = Some(format!("{error:#}"));
        }
        self.libreoffice_executor = executor;
        self.libreoffice = Some(session);
    }

    /// Test hook: replace the real `soffice` process runner with a fake.
    #[cfg(test)]
    fn set_libreoffice_executor(&mut self, executor: LibreOfficeExecutor) {
        self.libreoffice_executor = Some(executor);
    }

    /// Test hook: expose the session directory to assert its cleanup.
    #[cfg(test)]
    fn libreoffice_session_dir(&self) -> Option<PathBuf> {
        self.libreoffice
            .as_ref()
            .map(|session| session.session_dir.clone())
    }

    fn convert_batch_with_session(
        session: &mut LibreOfficeBatchSession,
        requests: &[(PathBuf, String)],
        executor: &mut Option<LibreOfficeExecutor>,
    ) -> Result<()> {
        let inputs_dir = session.session_dir.join("inputs");
        fs::create_dir_all(&inputs_dir).context("create LibreOffice session staging directory")?;
        let targets = requests
            .iter()
            .map(|(_, target)| target.clone())
            .collect::<HashSet<_>>();
        ensure!(
            targets.len() == 1,
            "a LibreOffice batch session handles exactly one conversion target"
        );
        let target = targets.into_iter().next().expect("single target checked");
        let filter = if target == "xlsx" {
            "xlsx:Calc MS Excel 2007 XML"
        } else {
            "docx:Office Open XML Text"
        };
        let source_exts = requests
            .iter()
            .map(|(input, _)| extension(input))
            .collect::<Result<HashSet<_>>>()?;
        ensure!(
            source_exts.len() == 1,
            "a LibreOffice batch session handles one source format"
        );
        let source_ext = source_exts
            .into_iter()
            .next()
            .expect("source format checked");

        // Stage disposable, uniquely named copies so same-stem inputs cannot
        // collide inside the shared outdir and the engine never sees the
        // caller's original file.
        let mut main_jobs = Vec::new();
        for (index, (input, request_target)) in requests.iter().enumerate() {
            let staged_input = inputs_dir.join(format!("lo-{index:06}.{source_ext}"));
            fs::copy(input, &staged_input).with_context(|| {
                format!(
                    "stage disposable copy for LibreOffice batch conversion from {}",
                    input.display()
                )
            })?;
            let staged_output =
                outputs_dir_for(session).join(format!("lo-{index:06}.{request_target}"));
            main_jobs.push(LibreOfficeJob {
                staged_input,
                staged_output,
            });
        }
        session.run_libreoffice_processes(
            executor,
            vec![(filter.to_owned(), outputs_dir_for(session), main_jobs)],
        )?;
        for (index, (input, request_target)) in requests.iter().enumerate() {
            let staged_output =
                outputs_dir_for(session).join(format!("lo-{index:06}.{request_target}"));
            if staged_output.is_file() && fs::metadata(&staged_output)?.len() > 0 {
                session
                    .converted
                    .insert((input.clone(), request_target.clone()), staged_output);
            }
        }

        if target == "docx" {
            // Word fidelity validation needs plain text of every source and
            // every converted candidate. One extra process extracts all of
            // them; copies keep the two kinds from colliding on stems.
            let texts_dir = session.session_dir.join("texts");
            let texts_in = texts_dir.join("in");
            let texts_out = texts_dir.join("out");
            fs::create_dir_all(&texts_in).context("create LibreOffice text staging directory")?;
            fs::create_dir_all(&texts_out)?;
            let mut text_jobs = Vec::new();
            let mut text_sources = Vec::new();
            for (index, (input, _)) in requests.iter().enumerate() {
                let staged_source = texts_in.join(format!("t-{index:06}.{source_ext}"));
                fs::copy(input, &staged_source).with_context(|| {
                    format!(
                        "stage source copy for LibreOffice text extraction from {}",
                        input.display()
                    )
                })?;
                text_jobs.push(LibreOfficeJob {
                    staged_output: texts_out.join(format!("t-{index:06}.txt")),
                    staged_input: staged_source,
                });
                text_sources.push((index, input.clone()));
            }
            let mut candidate_index = requests.len();
            let mut candidate_texts = Vec::new();
            for (input, request_target) in requests.iter() {
                let Some(candidate) = session
                    .converted
                    .get(&(input.clone(), request_target.clone()))
                else {
                    continue;
                };
                let staged_candidate = texts_in.join(format!("t-{candidate_index:06}.docx"));
                fs::copy(candidate, &staged_candidate).with_context(|| {
                    format!(
                        "stage converted copy for LibreOffice text extraction from {}",
                        candidate.display()
                    )
                })?;
                text_jobs.push(LibreOfficeJob {
                    staged_output: texts_out.join(format!("t-{candidate_index:06}.txt")),
                    staged_input: staged_candidate,
                });
                candidate_texts.push((
                    (input.clone(), request_target.clone()),
                    texts_out.join(format!("t-{candidate_index:06}.txt")),
                ));
                candidate_index += 1;
            }
            session.run_libreoffice_processes(
                executor,
                vec![("txt:Text".to_owned(), texts_out.clone(), text_jobs)],
            )?;
            for (index, input) in text_sources {
                let text_path = texts_out.join(format!("t-{index:06}.txt"));
                let text = fs::read_to_string(&text_path).with_context(|| {
                    format!("read staged Word source text {}", text_path.display())
                })?;
                session.source_texts.insert(input, text);
            }
            for ((input, request_target), text_path) in candidate_texts {
                let text = fs::read_to_string(&text_path).with_context(|| {
                    format!("read staged Word candidate text {}", text_path.display())
                })?;
                session
                    .candidate_texts
                    .insert((input, request_target), text);
            }
        }
        Ok(())
    }

    fn libreoffice_session(&self) -> Option<&LibreOfficeBatchSession> {
        self.libreoffice.as_ref()
    }

    /// Whether the batch session already took responsibility for this
    /// input (successfully or not). Covered inputs must never fall back to
    /// a second `soffice` process.
    pub(crate) fn libreoffice_covers(&self, input: &Path, target: &str) -> bool {
        self.libreoffice_session().is_some_and(|session| {
            session
                .covered
                .contains(&(input.to_path_buf(), target.to_owned()))
        })
    }

    pub(crate) fn libreoffice_source_text(&self, input: &Path) -> Option<String> {
        self.libreoffice_session()
            .and_then(|session| session.source_texts.get(input).cloned())
    }

    pub(crate) fn libreoffice_candidate_text(&self, input: &Path, target: &str) -> Option<String> {
        self.libreoffice_session().and_then(|session| {
            session
                .candidate_texts
                .get(&(input.to_path_buf(), target.to_owned()))
                .cloned()
        })
    }

    /// Batch outcome for a covered input: `Some(Ok(candidate))` when the
    /// session produced a candidate, `Some(Err(reason))` when the session
    /// covered the input but produced nothing (process failure or missing
    /// output), `None` when this input was never part of a batch session.
    pub(crate) fn libreoffice_batch_outcome(
        &self,
        input: &Path,
        target: &str,
    ) -> Option<std::result::Result<PathBuf, String>> {
        let session = self.libreoffice_session()?;
        let key = (input.to_path_buf(), target.to_owned());
        if !session.covered.contains(&key) {
            return None;
        }
        if let Some(error) = &session.error {
            return Some(Err(error.clone()));
        }
        match session.converted.get(&key) {
            Some(candidate) => Some(Ok(candidate.clone())),
            None => Some(Err(format!(
                "LibreOffice batch session did not produce an output for {}",
                input.display()
            ))),
        }
    }

    pub(crate) fn libreoffice_process_ids(&self) -> Vec<u32> {
        self.libreoffice_session()
            .map(|session| session.process_ids.clone())
            .unwrap_or_default()
    }
}

#[derive(Clone, Debug, serde::Serialize)]
pub(crate) struct NamedValidationRegionRequest {
    pub template_sheet: String,
    pub source_sheet: String,
    pub address: String,
}

#[derive(Clone, Debug)]
pub(crate) struct NamedValidationCellResult {
    pub sheet: String,
    pub address: String,
    pub value: String,
}

#[derive(Clone, Debug)]
pub(crate) struct NamedValidationCalculation {
    pub application: String,
    pub cells: Vec<NamedValidationCellResult>,
    /// Cells whose template formulas were copied in by the OOXML patcher
    /// before COM recalculated the working copy.
    pub copied_formula_cells: usize,
    /// Template sheets the patcher had to add because the copied formulas
    /// reference them and the working copy lacked them.
    pub copied_template_sheets: Vec<String>,
    /// Engine used to bring auxiliary formula dependencies into the disposable
    /// calculation copy before the platform calculator ran.
    pub dependency_copy_engine: String,
}

fn unique_validation_fallback_output(working_copy: &Path) -> Result<PathBuf> {
    let parent = working_copy
        .parent()
        .context("validation working copy has no parent directory")?;
    for attempt in 0..256u32 {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let output = parent.join(format!(
            ".named-validation-openpyxl-{}-{nonce}-{attempt}.xlsx",
            std::process::id()
        ));
        if !output.exists() {
            return Ok(output);
        }
    }
    bail!("could not reserve an Openpyxl validation fallback candidate")
}

fn sheet_used_range(sheet: &excel_reader::SheetDump) -> String {
    let mut max_row = 1u32;
    let mut max_column = 1u32;
    for address in sheet.values.keys().chain(sheet.formulas.keys()) {
        if let Some((row, column)) = parse_a1_coordinate(address) {
            max_row = max_row.max(row);
            max_column = max_column.max(column);
        }
    }
    for merge in &sheet.merges {
        for address in merge.split(':') {
            if let Some((row, column)) = parse_a1_coordinate(address) {
                max_row = max_row.max(row);
                max_column = max_column.max(column);
            }
        }
    }
    format!("A1:{}{}", excel_column_name(max_column), max_row)
}

fn parse_a1_coordinate(address: &str) -> Option<(u32, u32)> {
    let address = address.trim().replace('$', "");
    let split = address.find(|character: char| character.is_ascii_digit())?;
    let (column, row) = address.split_at(split);
    if column.is_empty() || !column.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return None;
    }
    let column = column.bytes().try_fold(0u32, |value, byte| {
        value
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
    })?;
    let row = row.parse::<u32>().ok()?;
    (row > 0 && column > 0).then_some((row, column))
}

fn excel_column_name(mut column: u32) -> String {
    let mut name = String::new();
    while column > 0 {
        let remainder = ((column - 1) % 26) as u8;
        name.insert(0, char::from(b'A' + remainder));
        column = (column - 1) / 26;
    }
    name
}

pub(crate) fn named_validation_calculation_engine() -> &'static str {
    #[cfg(windows)]
    {
        "Excel/WPS COM"
    }
    #[cfg(target_os = "linux")]
    {
        "LibreOffice Calc"
    }
    #[cfg(not(any(windows, target_os = "linux")))]
    {
        "unavailable"
    }
}

fn extract_named_validation_cells(
    workbook: &WorkbookDump,
    regions: &[NamedValidationRegionRequest],
) -> Result<Vec<NamedValidationCellResult>> {
    let mut matches = Vec::new();
    for region in regions {
        let sheet = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == region.source_sheet)
            .with_context(|| format!("LibreOffice 重算结果缺少工作表 {:?}", region.source_sheet))?;
        ensure!(
            sheet.formula_status == "read_by_calamine",
            "Calamine 无法读取工作表 {:?} 的公式位置：{}",
            region.source_sheet,
            sheet.formula_status
        );
        let bounds = parse_named_validation_region(&region.address).with_context(|| {
            format!(
                "不支持校验区域地址 {:?}（工作表 {:?}）",
                region.address, region.source_sheet
            )
        })?;
        for (address, formula) in &sheet.formulas {
            // Calamine's worksheet_formula API returns formula text without
            // Excel's leading '=' (unlike our JSON fixtures). Any non-empty
            // formula at a validation-region cell is eligible for keyword
            // extraction once its cached value is available.
            if formula.trim().is_empty() || !named_validation_cell_in_bounds(address, bounds) {
                continue;
            }
            let Some(value) = sheet
                .values
                .get(address)
                .filter(|value| {
                    value.get("type").and_then(serde_json::Value::as_str) == Some("string")
                })
                .and_then(|value| value.get("value"))
                .and_then(serde_json::Value::as_str)
            else {
                continue;
            };
            if ["错", "硬性", "软性", "校验失败"]
                .iter()
                .any(|keyword| value.contains(keyword))
            {
                matches.push(NamedValidationCellResult {
                    sheet: region.source_sheet.clone(),
                    address: address.clone(),
                    value: value.to_owned(),
                });
            }
        }
    }
    Ok(matches)
}

fn parse_named_validation_region(address: &str) -> Result<(u32, u32, u32, u32)> {
    let address = address
        .rsplit('!')
        .next()
        .unwrap_or(address)
        .replace('$', "");
    let (start, end) = address
        .split_once(':')
        .unwrap_or((address.as_str(), address.as_str()));
    let (start_row, start_column) = parse_named_validation_cell(start)?;
    let (end_row, end_column) = parse_named_validation_cell(end)?;
    Ok((
        start_row.min(end_row),
        start_column.min(end_column),
        start_row.max(end_row),
        start_column.max(end_column),
    ))
}

fn parse_named_validation_cell(address: &str) -> Result<(u32, u32)> {
    let split = address
        .find(|character: char| character.is_ascii_digit())
        .context("单元格地址缺少行号")?;
    let (column_letters, row_text) = address.split_at(split);
    ensure!(!column_letters.is_empty(), "单元格地址缺少列号");
    ensure!(
        column_letters
            .bytes()
            .all(|byte| byte.is_ascii_alphabetic()),
        "单元格列号格式非法"
    );
    ensure!(
        row_text.bytes().all(|byte| byte.is_ascii_digit()),
        "单元格行号格式非法"
    );
    let mut column = 0u32;
    for byte in column_letters.bytes() {
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .context("单元格列号超出支持范围")?;
    }
    let row = row_text.parse::<u32>().context("单元格行号格式非法")?;
    ensure!(row > 0 && column > 0, "单元格行号和列号必须从 1 开始");
    Ok((row, column))
}

fn named_validation_cell_in_bounds(
    address: &str,
    (min_row, min_column, max_row, max_column): (u32, u32, u32, u32),
) -> bool {
    parse_named_validation_cell(address).is_ok_and(|(row, column)| {
        (min_row..=max_row).contains(&row) && (min_column..=max_column).contains(&column)
    })
}

#[cfg(windows)]
struct PyWin32Conversion {
    application: String,
    process_id: Option<u32>,
    worker_process_id: u32,
    validation: serde_json::Value,
}

#[cfg(windows)]
struct PythonExcelWorker {
    child: Child,
    input: BufWriter<ChildStdin>,
    output: BufReader<ChildStdout>,
}

#[cfg(windows)]
impl PythonExcelWorker {
    fn start() -> Result<Self> {
        let (python, worker) = pywin32_worker_paths()?;
        let mut child = Command::new(&python)
            .arg("-I")
            .arg(&worker)
            .arg("--serve")
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::null())
            .spawn()
            .with_context(|| {
                format!(
                    "启动随包 pywin32 Excel 转换 Worker 失败（Python={}，Worker={}）",
                    python.display(),
                    worker.display()
                )
            })?;

        let startup = (|| -> Result<(BufWriter<ChildStdin>, BufReader<ChildStdout>)> {
            let input = BufWriter::new(
                child
                    .stdin
                    .take()
                    .context("pywin32 worker 未提供标准输入")?,
            );
            let mut output = BufReader::new(
                child
                    .stdout
                    .take()
                    .context("pywin32 worker 未提供标准输出")?,
            );
            let mut line = String::new();
            let bytes = output
                .read_line(&mut line)
                .context("读取 pywin32 worker 启动握手失败")?;
            ensure!(bytes > 0, "pywin32 worker 在启动握手前退出");
            let ready: serde_json::Value =
                serde_json::from_str(&line).context("pywin32 worker 启动握手不是有效 JSON")?;
            ensure!(
                ready.get("ready").and_then(serde_json::Value::as_bool) == Some(true)
                    && ready.get("protocol").and_then(serde_json::Value::as_u64)
                        == Some(u64::from(PYWIN32_PROTOCOL_VERSION)),
                "pywin32 worker 启动失败：{}",
                ready
                    .get("error")
                    .and_then(serde_json::Value::as_str)
                    .unwrap_or("协议版本不匹配")
            );
            Ok((input, output))
        })();

        match startup {
            Ok((input, output)) => Ok(Self {
                child,
                input,
                output,
            }),
            Err(error) => {
                let _ = child.kill();
                let _ = child.wait();
                Err(error)
            }
        }
    }

    /// Send one JSON request and read the single JSON response line.
    fn request(&mut self, request: &serde_json::Value, what: &str) -> Result<serde_json::Value> {
        if let Some(status) = self
            .child
            .try_wait()
            .context("检查 pywin32 worker 进程状态失败")?
        {
            bail!("pywin32 worker 已退出（状态：{status}）");
        }
        serde_json::to_writer(&mut self.input, request)
            .with_context(|| format!("写入 {what} 请求失败"))?;
        self.input
            .write_all(b"\n")
            .and_then(|()| self.input.flush())
            .with_context(|| format!("发送 {what} 请求失败"))?;

        let mut line = String::new();
        let bytes = self
            .output
            .read_line(&mut line)
            .with_context(|| format!("读取 {what} 结果失败"))?;
        ensure!(bytes > 0, "pywin32 worker 未返回{what}结果");
        let response: serde_json::Value =
            serde_json::from_str(&line).with_context(|| format!("{what}结果不是有效 JSON"))?;
        ensure!(
            response.get("ok").and_then(serde_json::Value::as_bool) == Some(true),
            "{what}失败（阶段：{}）：{}",
            response
                .get("stage")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("未知阶段"),
            response
                .get("error")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("worker 未提供错误详情")
        );
        Ok(response)
    }

    fn convert(&mut self, input: &Path, output: &Path) -> Result<PyWin32Conversion> {
        let response = self.request(
            &serde_json::json!({
                "op": "convert",
                "input": input.to_string_lossy(),
                "output": output.to_string_lossy(),
            }),
            "pywin32 Excel 转换",
        )?;
        let validation = response
            .get("validation")
            .cloned()
            .context("pywin32 worker 缺少工作簿校验结果")?;
        ensure!(
            validation
                .get("verified")
                .and_then(serde_json::Value::as_bool)
                == Some(true),
            "pywin32 worker 未报告工作簿校验通过"
        );
        let application = response
            .get("application")
            .and_then(serde_json::Value::as_str)
            .context("pywin32 worker 未报告 Excel/WPS COM 应用")?
            .to_owned();
        let process_id = response
            .get("process_id")
            .and_then(serde_json::Value::as_u64)
            .and_then(|value| u32::try_from(value).ok());
        Ok(PyWin32Conversion {
            application,
            process_id,
            worker_process_id: self.child.id(),
            validation,
        })
    }

    fn convert_word(&mut self, input: &Path, output: &Path) -> Result<PyWin32Conversion> {
        let response = self.request(
            &serde_json::json!({
                "op": "convert_word",
                "input": input.to_string_lossy(),
                "output": output.to_string_lossy(),
            }),
            "pywin32 Word 转换",
        )?;
        let application = response
            .get("application")
            .and_then(serde_json::Value::as_str)
            .context("pywin32 worker 未报告 Word/WPS Writer COM 应用")?
            .to_owned();
        let process_id = response
            .get("process_id")
            .and_then(serde_json::Value::as_u64)
            .and_then(|value| u32::try_from(value).ok());
        Ok(PyWin32Conversion {
            application,
            process_id,
            worker_process_id: self.child.id(),
            validation: response
                .get("validation")
                .cloned()
                .unwrap_or_else(|| serde_json::json!({"level": "docx_package_and_text_content"})),
        })
    }

    fn extract_word_text(&mut self, input: &Path) -> Result<String> {
        let response = self.request(
            &serde_json::json!({
                "op": "extract_word_text",
                "input": input.to_string_lossy(),
            }),
            "pywin32 Word 文本提取",
        )?;
        Ok(response
            .get("text")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .to_owned())
    }

    /// Ask the worker to recalculate one already-prepared working copy. The
    /// template is not involved: its formulas were copied in by
    /// [`ooxml_patch::copy_cell_region`] before this call.
    fn recalculate_named_validation_regions(
        &mut self,
        source_copy: &Path,
        regions: &[NamedValidationRegionRequest],
    ) -> Result<NamedValidationCalculation> {
        if let Some(status) = self
            .child
            .try_wait()
            .context("检查 pywin32 worker 进程状态失败")?
        {
            bail!("pywin32 worker 已退出（状态：{status}）");
        }
        let request = serde_json::json!({
            "op": "recalculate_named_validation",
            "source_copy": source_copy.to_string_lossy(),
            "regions": regions,
        });
        serde_json::to_writer(&mut self.input, &request)
            .context("写入校验区域 COM 重算请求失败")?;
        self.input
            .write_all(b"\n")
            .and_then(|()| self.input.flush())
            .context("发送校验区域 COM 重算请求失败")?;

        let mut line = String::new();
        let bytes = self
            .output
            .read_line(&mut line)
            .context("读取校验区域 COM 重算结果失败")?;
        ensure!(bytes > 0, "pywin32 worker 未返回校验区域重算结果");
        let response: serde_json::Value = serde_json::from_str(&line)
            .context("pywin32 worker 返回的校验区域结果不是有效 JSON")?;
        ensure!(
            response.get("ok").and_then(serde_json::Value::as_bool) == Some(true),
            "校验区域 COM 重算失败（阶段：{}）：{}",
            response
                .get("stage")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("未知阶段"),
            response
                .get("error")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("worker 未提供错误详情")
        );
        let application = response
            .get("application")
            .and_then(serde_json::Value::as_str)
            .context("校验区域 COM 结果未报告 Excel/WPS 应用")?
            .to_owned();
        let cells = response
            .get("cells")
            .and_then(serde_json::Value::as_array)
            .context("校验区域 COM 结果缺少单元格数组")?
            .iter()
            .map(|cell| {
                Ok(NamedValidationCellResult {
                    sheet: cell
                        .get("sheet")
                        .and_then(serde_json::Value::as_str)
                        .context("校验结果缺少工作表名")?
                        .to_owned(),
                    address: cell
                        .get("address")
                        .and_then(serde_json::Value::as_str)
                        .context("校验结果缺少单元格地址")?
                        .to_owned(),
                    value: cell
                        .get("value")
                        .and_then(serde_json::Value::as_str)
                        .context("校验结果缺少公式文本")?
                        .to_owned(),
                })
            })
            .collect::<Result<Vec<_>>>()?;
        Ok(NamedValidationCalculation {
            application,
            cells,
            copied_formula_cells: 0,
            copied_template_sheets: Vec::new(),
            dependency_copy_engine: "由调用方报告".to_owned(),
        })
    }
}

#[cfg(windows)]
impl Drop for PythonExcelWorker {
    fn drop(&mut self) {
        let _ = writeln!(self.input, "{{\"op\":\"shutdown\"}}");
        let _ = self.input.flush();
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(8);
        loop {
            match self.child.try_wait() {
                Ok(Some(_)) => return,
                Ok(None) if std::time::Instant::now() < deadline => {
                    std::thread::sleep(std::time::Duration::from_millis(25));
                }
                _ => {
                    let _ = self.child.kill();
                    let _ = self.child.wait();
                    return;
                }
            }
        }
    }
}

/// 与随包 worker `_ensure_application` 的默认引擎选择同一口径：设置了
/// `PYTOOLS_COM_PREFER_WPS` 时 COM 以 WPS 打开工作簿（WPS 能解析裸引用与
/// 矛盾窗格），此时不做 Excel 打开修复。
fn com_prefers_wps() -> bool {
    matches!(
        std::env::var("PYTOOLS_COM_PREFER_WPS")
            .unwrap_or_default()
            .trim()
            .to_ascii_lowercase()
            .as_str(),
        "1" | "true" | "yes" | "y"
    )
}

#[cfg(windows)]
fn pywin32_worker_paths() -> Result<(PathBuf, PathBuf)> {
    let current_exe = std::env::current_exe().context("定位当前 pytools.exe 失败")?;
    let exe_dir = current_exe
        .parent()
        .context("当前 pytools.exe 没有父目录")?;
    let packaged_dir = exe_dir.join("python");
    let manifest_root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .context("定位 pytools-core 仓库根目录失败")?
        .to_path_buf();
    let development_dir = manifest_root.join("dist").join("python");

    let python = std::env::var_os("PYTOOLS_COM_PYTHON")
        .map(PathBuf::from)
        .or_else(|| {
            [
                packaged_dir.join("python.exe"),
                development_dir.join("python.exe"),
            ]
            .into_iter()
            .find(|candidate| candidate.is_file())
        })
        .context("未找到随包 Python；请将 pywin32 运行时暂存到 dist\\python")?;
    let worker = std::env::var_os("PYTOOLS_COM_WORKER")
        .map(PathBuf::from)
        .or_else(|| {
            [
                packaged_dir.join("pywin32_excel_conversion_worker.py"),
                development_dir.join("pywin32_excel_conversion_worker.py"),
            ]
            .into_iter()
            .find(|candidate| candidate.is_file())
        })
        .context("未找到 pywin32 Excel 转换 Worker；请重新暂存 dist\\python")?;
    ensure!(
        python.is_file(),
        "pywin32 Python 解释器不存在：{}",
        python.display()
    );
    ensure!(
        worker.is_file(),
        "pywin32 Worker 不存在：{}",
        worker.display()
    );
    Ok((python, worker))
}

fn outputs_dir_for(session: &LibreOfficeBatchSession) -> PathBuf {
    session.session_dir.join("outputs")
}

fn unique_lo_session_dir() -> Result<PathBuf> {
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    for i in 0..100 {
        let path = std::env::temp_dir().join(format!(
            ".pytools-lo-session-{}-{:?}-{nanos}-{i}",
            std::process::id(),
            std::thread::current().id()
        ));
        if !path.exists() {
            fs::create_dir(&path)?;
            return Ok(path);
        }
    }
    bail!("cannot allocate LibreOffice session directory")
}

struct ComExecution {
    source_format: Option<i32>,
    /// Native extension reported by the worker for the opened source (used
    /// to cross-check ET/ETT source-format evidence).
    source_extension: Option<String>,
    bridge: &'static str,
    application: Option<String>,
    process_id: Option<u32>,
    worker_process_id: Option<u32>,
    validation: Option<serde_json::Value>,
}

/// Run one same-category fidelity conversion and retain the existing JSON
/// detail contract. The caller may reuse the runtime across a batch.
pub(crate) fn run_fidelity_conversion_with_runtime(
    input: &Path,
    output: &Path,
    target: &str,
    engine: &str,
    runtime: &mut ConversionRuntime,
) -> Result<String> {
    if std::env::var_os("PYTOOLS_CONVERT_ENGINES_DISABLED").is_some() {
        // Test override: simulate an environment with no external engine.
        bail!("该功能需要外部引擎: no conversion engine is available (PYTOOLS_CONVERT_ENGINES_DISABLED)");
    }
    ensure!(
        engine == "auto" || engine == "lo" || engine == "com" || engine == "pywin32",
        "engine must be auto, lo, com or pywin32"
    );
    // LibreOffice（lo）仅用于 Linux 电子表格 .xls/.xlsm → .xlsx（2026-11-09
    // 实测与 COM 另存一致，见 docs §5.10）；Word 与 WPS 原生格式仍仅 Windows。
    ensure!(
        engine != "lo" || target == "xlsx",
        "LibreOffice conversion is available for spreadsheet .xls/.xlsm -> .xlsx on Linux; Word targets require Windows pywin32"
    );
    ensure!(
        input.is_file(),
        "input file does not exist: {}",
        input.display()
    );
    let source_ext = extension(input)?;
    ensure!(
        (matches!(target, "xlsx")
            && matches!(
                source_ext.as_str(),
                "xls" | "xlsb" | "xlsm" | "et" | "ett" | "xlsx"
            ))
            || (target == "docx" && matches!(source_ext.as_str(), "doc" | "wps" | "wpt")),
        "unsupported conversion: .{source_ext} -> .{target}"
    );
    if !cfg!(windows)
        && (engine == "pywin32"
            || target == "docx"
            || matches!(source_ext.as_str(), "et" | "ett" | "xlsb"))
    {
        bail!(
            "Office format conversion on Linux is limited to LibreOffice spreadsheet .xls/.xlsm -> .xlsx; Word/WPS-native/.xlsb formats require Windows pywin32"
        );
    }
    if let Some(parent) = output.parent() {
        fs::create_dir_all(parent)?;
    }
    ensure!(
        !output.exists(),
        "refusing to overwrite existing output: {}",
        output.display()
    );

    let temp_guard = TempDirGuard(unique_temp_dir(
        output.parent().unwrap_or_else(|| Path::new(".")),
    )?);
    let temp_dir = &temp_guard.0;
    let selected = select_conversion_engine(engine, &source_ext, target, cfg!(windows));
    // Office rewrites relative external-workbook links relative to the
    // SaveAs destination. Keep a COM candidate beside its final output so
    // validating it and atomically moving it do not change the link base.
    // A candidate inside the disposable staging directory would leave links
    // pointing into that directory after publication (or fail formula checks
    // because Excel qualifies them with the temporary path).
    let converted = if selected == "pywin32" {
        unique_sibling_conversion_candidate(output, target)?
    } else {
        temp_dir.join(format!("{}.{}", file_stem(input), target))
    };
    let _candidate_guard =
        (selected == "pywin32").then(|| ConversionCandidateGuard(converted.clone()));
    if source_ext == "xlsx" && target == "xlsx" && selected == "lo" {
        ensure!(
            crate::conversion::is_ooxml_spreadsheet_package(input)?,
            "同扩展名 .xlsx 的非 OOXML 文件需要 Excel/WPS COM 识别真实格式并另存；LibreOffice 仅接受有效 OOXML 工作簿"
        );
    }
    let com_execution = match selected {
        "lo" => {
            match runtime.libreoffice_batch_outcome(input, target) {
                // The batch session covered this file: reuse its candidate
                // or fail with its reason — never start a second `soffice`
                // process for a covered input.
                Some(Ok(batch_candidate)) => {
                    fs::copy(&batch_candidate, &converted).with_context(|| {
                        format!(
                            "stage batch-converted candidate {}",
                            batch_candidate.display()
                        )
                    })?;
                }
                Some(Err(reason)) => bail!("{reason}"),
                None => run_libreoffice(input, temp_dir, target)?,
            }
            None
        }
        "pywin32" if cfg!(windows) => Some(run_pywin32(input, &converted, target, &source_ext, runtime)?),
        "pywin32" => bail!(
            "该功能需要外部引擎: Excel/Word COM conversion requires Windows with Microsoft Office or WPS installed"
        ),
        _ => unreachable!(),
    };
    let source_file_format = com_execution
        .as_ref()
        .and_then(|execution| execution.source_format);
    let source_extension = com_execution
        .as_ref()
        .and_then(|execution| execution.source_extension.clone());
    let candidate = if converted.exists() {
        converted
    } else {
        locate_converted(temp_dir, target)?
    };
    ensure!(
        candidate.is_file() && fs::metadata(&candidate)?.len() > 0,
        "conversion engine did not create a non-empty output file"
    );
    // The COM routes call CalculateFullRebuild before SaveAs. Keep this as a
    // follow-up request for Office/WPS when the delivered workbook is opened;
    // this OOXML flag alone does not create formula caches for Calamine now.
    if target == "xlsx" {
        crate::conversion::patch_full_calc_on_load(&candidate)?;
    }
    let validation = validate_conversion(
        input,
        &candidate,
        &source_ext,
        target,
        selected,
        temp_dir,
        source_file_format,
        source_extension.as_deref(),
        com_execution
            .as_ref()
            .and_then(|execution| execution.validation.as_ref()),
        runtime,
    )?;
    let validation = serde_json::from_str::<serde_json::Value>(&validation)
        .context("decode conversion validation details")?;
    let mut report = serde_json::json!({
        "status": "converted",
        "engine": selected,
        "input": input.display().to_string(),
        "output": output.display().to_string(),
        "validation": validation,
    });
    if let Some(execution) = com_execution {
        report["com_bridge"] = serde_json::json!(execution.bridge);
        if let Some(application) = execution.application {
            report["com_application"] = serde_json::json!(application);
        }
        if let Some(process_id) = execution.process_id {
            report["com_process_id"] = serde_json::json!(process_id);
        }
        if let Some(process_id) = execution.worker_process_id {
            report["com_worker_process_id"] = serde_json::json!(process_id);
        }
    }
    // Session evidence: the PIDs of the LibreOffice processes this batch
    // started prove one instance served the whole batch.
    let lo_process_ids = runtime.libreoffice_process_ids();
    if !lo_process_ids.is_empty() {
        report["lo_session"] = serde_json::json!({
            "process_ids": lo_process_ids,
            "shared_profile": true,
        });
    }
    fs::rename(&candidate, output)
        .with_context(|| format!("commit validated output {}", output.display()))?;
    Ok(report.to_string())
}

fn run_libreoffice(input: &Path, temp_dir: &Path, target: &str) -> Result<()> {
    let profile = temp_dir.join("lo-profile");
    let profile_argument = libreoffice_user_installation_arg(&profile)?;
    let mut command = Command::new("soffice");
    command
        .arg("--headless")
        .arg("--nologo")
        .arg("--nodefault")
        .arg("--nolockcheck")
        .arg(profile_argument)
        .arg("--convert-to")
        .arg(if target == "xlsx" {
            "xlsx:Calc MS Excel 2007 XML"
        } else {
            "docx:Office Open XML Text"
        })
        .arg("--outdir")
        .arg(temp_dir)
        .arg(input)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    let output = command.output().with_context(|| {
        "该功能需要外部引擎: LibreOffice is unavailable; install soffice to use this conversion"
    })?;
    ensure!(
        output.status.success(),
        "LibreOffice conversion failed ({}): {}",
        output.status,
        String::from_utf8_lossy(&output.stderr)
    );
    Ok(())
}

fn run_pywin32(
    input: &Path,
    output: &Path,
    target: &str,
    source_ext: &str,
    runtime: &mut ConversionRuntime,
) -> Result<ComExecution> {
    // Excel/WPS COM resolves relative paths against *its own* working
    // directory, not ours, so every path handed to the worker must be
    // absolute. `std::path::absolute` keeps the drive-letter form (unlike
    // canonicalize's `\\?\` prefix, which Office rejects).
    let absolute_output = std::path::absolute(output)
        .with_context(|| format!("resolve absolute output path {}", output.display()))?;
    let output_parent = absolute_output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(output_parent)?;
    let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    // Keep the private COM input beside the original so relative external
    // workbook links resolve from the same directory as they do in PyCode.
    // The automation host still receives only a disposable copy, never the
    // caller's source file.
    let source_parent = input.parent().unwrap_or_else(|| Path::new("."));
    let com_input = source_parent.join(format!(
        ".pytools-convert-source-{}-{nonce}.{source_ext}",
        std::process::id()
    ));
    // Some WPS COM failure paths can alter the workbook passed to Open even
    // when ReadOnly=true. External automation must never receive the caller's
    // source path; it only opens this disposable copy.
    fs::copy(input, &com_input).with_context(|| {
        format!(
            "stage a private source copy for Office/WPS COM conversion from {}",
            input.display()
        )
    })?;
    // Excel 会拒绝含全角括号表名裸引用/矛盾冻结窗格的 .xlsx/.xlsm；WPS 宽容。
    // 仅当 COM 将以 Excel 打开（非 WPS 优先）时修复这份一次性私有副本；
    // .xls/.et 等二进制容器与 Word 文档不在该修复范围内。
    if matches!(source_ext, "xlsx" | "xlsm") && target == "xlsx" && !com_prefers_wps() {
        ooxml_patch::repair_excel_com_open_issues(&com_input);
    }
    let result = match target {
        "xlsx" => {
            #[cfg(windows)]
            let result = runtime.spreadsheet_pywin32(&com_input, &absolute_output);
            #[cfg(not(windows))]
            let result = {
                let _ = runtime;
                Err(anyhow!("COM conversion is only available on Windows"))
            };
            result
        }
        "docx" => {
            #[cfg(windows)]
            let result = runtime.word_pywin32(&com_input, &absolute_output);
            #[cfg(not(windows))]
            let result = {
                let _ = runtime;
                Err(anyhow!("COM conversion is only available on Windows"))
            };
            result
        }
        other => {
            let _ = fs::remove_file(&com_input);
            bail!("unsupported COM conversion target {other}");
        }
    };
    let _ = fs::remove_file(&com_input);
    result
}

fn validate_conversion(
    input: &Path,
    output: &Path,
    source_ext: &str,
    target: &str,
    engine: &str,
    temp_dir: &Path,
    source_file_format: Option<i32>,
    source_extension: Option<&str>,
    external_validation: Option<&serde_json::Value>,
    runtime: &mut ConversionRuntime,
) -> Result<String> {
    if target == "docx" {
        let bytes = fs::read(output)?;
        ensure!(
            bytes.starts_with(b"PK\x03\x04"),
            "converted DOCX is not a valid OOXML ZIP package"
        );
        let (source_text, output_text) = match engine {
            "lo" => {
                // Batch sessions pre-extracted both texts with one extra
                // process; fall back to a direct extraction otherwise.
                let extracted_source = runtime.libreoffice_source_text(input);
                let extracted_output = runtime.libreoffice_candidate_text(input, target);
                (
                    match extracted_source {
                        Some(text) => text,
                        None => extract_word_text_lo(input, &temp_dir.join("source-text"))?,
                    },
                    match extracted_output {
                        Some(text) => text,
                        None => extract_word_text_lo(output, &temp_dir.join("output-text"))?,
                    },
                )
            }
            "pywin32" if cfg!(windows) => {
                // Same batch-owned Word/WPS instance that produced the
                // candidate also extracts both texts - no extra process.
                #[cfg(windows)]
                {
                    let source_text = runtime.extract_word_text_pywin32(input)?;
                    let output_text = runtime.extract_word_text_pywin32(output)?;
                    (source_text, output_text)
                }
                #[cfg(not(windows))]
                {
                    unreachable!("COM conversion is only available on Windows")
                }
            }
            _ => bail!("unsupported text-validation engine {engine:?}"),
        };
        ensure!(
            normalize_document_text(&source_text) == normalize_document_text(&output_text),
            "converted DOCX text content differs from the source document"
        );
        // Keep the worker's source evidence in the final report so callers can
        // tell which native format (.doc/.wps/.wpt) actually produced the DOCX.
        let mut validation = serde_json::json!({"level": "docx_package_and_text_content"});
        if let Some(source_extension) = external_validation
            .and_then(|value| value.get("source_extension"))
            .cloned()
        {
            validation["source_extension"] = source_extension;
        }
        return Ok(validation.to_string());
    }
    let converted = read_comparable_workbook(output)
        .with_context(|| format!("validate converted workbook {}", output.display()))?;
    ensure!(
        !converted.sheets.is_empty(),
        "converted workbook has no worksheets"
    );
    if source_ext == "csv" {
        validate_csv_conversion(input, &converted)?;
        return Ok(
            "{\"level\":\"one_worksheet_and_csv_values\",\"numeric_tolerance\":1e-15}".into(),
        );
    }
    if let Some(validation) = external_validation {
        let verified = validation
            .get("verified")
            .and_then(serde_json::Value::as_bool)
            == Some(true);
        if matches!(source_ext, "xls" | "xlsb" | "xlsm" | "xlsx") {
            ensure!(
                engine == "pywin32" && verified,
                "pywin32 external validation evidence is invalid for this conversion"
            );
            return serde_json::to_string(validation)
                .context("encode pywin32 Excel preservation validation");
        }
        if matches!(source_ext, "et" | "ett") {
            // The worker compared the WPS-side source snapshot with the
            // reopened XLSX candidate; only accept it together with the
            // native WPS FileFormat evidence recorded during that pass.
            ensure!(
                engine == "pywin32" && verified,
                "pywin32 external validation evidence is invalid for this conversion"
            );
            let native_format =
                source_file_format.is_some_and(|format| format >= WPS_NATIVE_FILE_FORMAT_MIN);
            ensure!(
                native_format
                    && source_extension == Some(source_ext)
                    && validation
                        .get("source_format_verified")
                        .and_then(serde_json::Value::as_bool)
                        == Some(true)
                    && validation
                        .get("source_extension")
                        .and_then(serde_json::Value::as_str)
                        == Some(source_ext),
                "WPS-native source-format evidence is missing or inconsistent"
            );
            return serde_json::to_string(validation)
                .context("encode WPS-native Excel preservation validation");
        }
    }
    let source = read_comparable_workbook(input).with_context(|| {
        format!(
            "read source workbook for preservation validation {}",
            input.display()
        )
    })?;
    compare_workbooks(&source, &converted)?;
    Ok(
        "{\"level\":\"worksheet_order_names_values_and_formula_text\",\"numeric_tolerance\":1e-15}"
            .into(),
    )
}

fn extract_word_text_lo(input: &Path, output_dir: &Path) -> Result<String> {
    fs::create_dir_all(output_dir)?;
    let profile = output_dir.join("lo-profile");
    let profile_argument = libreoffice_user_installation_arg(&profile)?;
    let command = Command::new("soffice")
        .arg("--headless")
        .arg("--nologo")
        .arg("--nodefault")
        .arg("--nolockcheck")
        .arg(profile_argument)
        .arg("--convert-to")
        .arg("txt:Text")
        .arg("--outdir")
        .arg(output_dir)
        .arg(input)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .with_context(|| {
            "该功能需要外部引擎: LibreOffice is required to validate converted Word text"
        })?;
    ensure!(
        command.status.success(),
        "LibreOffice text extraction failed: {}",
        String::from_utf8_lossy(&command.stderr)
    );
    let text_path = output_dir.join(format!("{}.txt", file_stem(input)));
    fs::read_to_string(&text_path)
        .with_context(|| format!("read extracted Word text {}", text_path.display()))
}

fn normalize_document_text(text: &str) -> String {
    text.replace("\r\n", "\n")
        .replace('\r', "\n")
        .trim_end_matches(['\n', '\u{000c}'])
        .to_owned()
}

fn read_comparable_workbook(path: &Path) -> Result<WorkbookDump> {
    let mut workbook = read_workbook(path)?;
    let has_date_serial = workbook.sheets.iter().any(|sheet| {
        sheet.values.values().any(|value| {
            value.get("type").and_then(serde_json::Value::as_str) == Some("datetime_serial")
        })
    });
    if !has_date_serial {
        return Ok(workbook);
    }
    let max_columns = workbook
        .sheets
        .iter()
        .flat_map(|sheet| {
            sheet
                .values
                .keys()
                .filter_map(|address| column_number(address))
        })
        .max()
        .unwrap_or(1)
        .max(1);
    for sheet in read_value_columns_with_iso_dates(path, max_columns)? {
        let Some(target) = workbook
            .sheets
            .iter_mut()
            .find(|candidate| candidate.name == sheet.name)
        else {
            continue;
        };
        for (address, value) in sheet.values {
            if let ReaderCellValue::DateTimeIso(value) = value {
                target.values.insert(
                    address,
                    serde_json::json!({"type":"datetime_iso","value":value}),
                );
            }
        }
    }
    Ok(workbook)
}

fn compare_workbooks(source: &WorkbookDump, converted: &WorkbookDump) -> Result<()> {
    let source_names = source
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    let converted_names = converted
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    ensure!(
        source_names == converted_names,
        "conversion changed worksheet count, names or order"
    );
    let sheet_name_refs = source_names
        .iter()
        .map(|name| name.as_str())
        .collect::<Vec<_>>();
    for (source_sheet, output_sheet) in source.sheets.iter().zip(&converted.sheets) {
        ensure!(
            source_sheet.formulas.len() == output_sheet.formulas.len(),
            "conversion changed formula count in worksheet {:?}",
            source_sheet.name
        );
        for (address, formula) in &source_sheet.formulas {
            let Some(output_formula) = output_sheet.formulas.get(address) else {
                bail!(
                    "conversion removed formula {}!{}",
                    source_sheet.name,
                    address
                );
            };
            // 与 Rust-Audit 同口径：只容忍引擎重存会做的规范化。
            ensure!(
                excel_reader::formula_text_equivalent_after_recalculation(
                    formula,
                    output_formula,
                    &sheet_name_refs,
                ),
                "conversion changed formula text at {}!{}: {:?} -> {:?}",
                source_sheet.name,
                address,
                formula,
                output_formula
            );
        }
        let source_cells = source_sheet
            .values
            .keys()
            .filter(|address| !source_sheet.formulas.contains_key(*address))
            .collect::<std::collections::BTreeSet<_>>();
        let output_cells = output_sheet
            .values
            .keys()
            .filter(|address| !output_sheet.formulas.contains_key(*address))
            .collect::<std::collections::BTreeSet<_>>();
        ensure!(
            source_cells == output_cells,
            "conversion changed non-formula data cell coordinates in worksheet {:?}",
            source_sheet.name
        );
        for address in source_cells {
            let source_value = &source_sheet.values[address];
            let output_value = &output_sheet.values[address];
            ensure!(
                cell_values_equal(source_value, output_value),
                "conversion changed stored value at {}!{}: {} -> {}",
                source_sheet.name,
                address,
                source_value,
                output_value
            );
        }
    }
    Ok(())
}

fn cell_values_equal(left: &serde_json::Value, right: &serde_json::Value) -> bool {
    let numeric = |kind: &str| matches!(kind, "number" | "integer" | "datetime_serial");
    let left_kind = cell_kind(left);
    let right_kind = cell_kind(right);
    if numeric(left_kind) && numeric(right_kind) {
        if left_kind == "integer" && right_kind == "integer" {
            return left.get("value").and_then(serde_json::Value::as_i64)
                == right.get("value").and_then(serde_json::Value::as_i64);
        }
        let Some(left) = left.get("value").and_then(serde_json::Value::as_f64) else {
            return left == right;
        };
        let Some(right) = right.get("value").and_then(serde_json::Value::as_f64) else {
            return false;
        };
        return left == right
            || (left.is_finite()
                && right.is_finite()
                && left != 0.0
                && right != 0.0
                && (left - right).abs() <= 1e-15 * left.abs().max(right.abs()));
    }
    left == right
}

fn cell_kind(value: &serde_json::Value) -> &str {
    value
        .get("type")
        .and_then(serde_json::Value::as_str)
        .unwrap_or_default()
}

fn column_number(address: &str) -> Option<usize> {
    let split = address.find(|character: char| character.is_ascii_digit())?;
    address[..split].bytes().try_fold(0usize, |value, byte| {
        value
            .checked_mul(26)?
            .checked_add((byte.to_ascii_uppercase().checked_sub(b'A')? + 1) as usize)
    })
}

fn locate_converted(temp: &Path, target: &str) -> Result<PathBuf> {
    let mut entries = fs::read_dir(temp)?
        .filter_map(|entry| entry.ok().map(|e| e.path()))
        .filter(|p| {
            p.extension()
                .and_then(|e| e.to_str())
                .is_some_and(|e| e.eq_ignore_ascii_case(target))
        })
        .collect::<Vec<_>>();
    entries.sort();
    entries
        .into_iter()
        .next()
        .ok_or_else(|| anyhow!("conversion produced no .{target} output"))
}

fn unique_temp_dir(parent: &Path) -> Result<PathBuf> {
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    for i in 0..100 {
        let path = parent.join(format!(
            ".pytools-convert-{}-{:?}-{nanos}-{i}",
            std::process::id(),
            std::thread::current().id()
        ));
        if !path.exists() {
            fs::create_dir(&path)?;
            return Ok(path);
        }
    }
    bail!("cannot allocate conversion work directory")
}

fn unique_sibling_conversion_candidate(output: &Path, target: &str) -> Result<PathBuf> {
    let parent = output.parent().unwrap_or_else(|| Path::new("."));
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    for index in 0..100 {
        let path = parent.join(format!(
            ".pytools-convert-candidate-{}-{nanos}-{index}.{target}",
            std::process::id()
        ));
        if !path.exists() {
            return Ok(path);
        }
    }
    bail!("cannot allocate sibling conversion candidate")
}

struct ConversionCandidateGuard(PathBuf);

impl Drop for ConversionCandidateGuard {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.0);
    }
}

fn extension(path: &Path) -> Result<String> {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase())
        .ok_or_else(|| anyhow!("file has no extension: {}", path.display()))
}

fn file_stem(path: &Path) -> String {
    path.file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "converted".into())
}

/// Pre-seed an isolated LibreOffice profile so headless `--convert-to`
/// always recalculates formulas when loading Excel workbooks.
///
/// LO decides whether to recalculate on load from
/// `/org.openoffice.Office.Calc/Formula/Load/OOXMLRecalcMode`
/// (0 = always recalculate, 1 = never, 2 = prompt). The shipped default can
/// vary by build and locale, which would make conversion output depend on the
/// machine — so the isolated profile pins it to "always". Calamine (and 1-1 /
/// 1-2 / 1-3) read the cached formula results from the converted file; a
/// deterministic recalculation is what keeps those caches trustworthy.
fn seed_lo_recalc_profile(profile: &Path) {
    let registry = profile.join("user").join("registrymodifications.xcu");
    if registry.exists() {
        // The profile has been launched before; LO owns this file now.
        return;
    }
    let content = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        "\n",
        r#"<oor:items xmlns:oor="http://openoffice.org/2001/registry""#,
        r#" xmlns:xs="http://www.w3.org/2001/XMLSchema""#,
        r#" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">"#,
        "\n",
        r#" <item oor:path="/org.openoffice.Office.Calc/Formula/Load">"#,
        r#"<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>"#,
        "\n",
        r#"</oor:items>"#,
        "\n",
    );
    let user_dir = registry.parent().expect("profile user dir has a parent");
    if fs::create_dir_all(user_dir).is_err() {
        return;
    }
    if let Err(error) = fs::write(&registry, content) {
        eprintln!(
            "{}",
            crate::run_logging::terminal_line(
                crate::run_logging::Level::Warn,
                &format!(
                    "预置 LibreOffice 重算配置失败（不影响转换，按 LO 默认设置执行）：{error:#}"
                ),
            )
        );
    }
}

fn libreoffice_user_installation_arg(profile: &Path) -> Result<String> {
    fs::create_dir_all(profile).context("create LibreOffice isolated profile directory")?;
    seed_lo_recalc_profile(profile);
    Ok(format!("-env:UserInstallation={}", file_uri(profile)?))
}

fn file_uri(path: &Path) -> Result<String> {
    ensure!(
        path.is_absolute(),
        "LibreOffice profile path must be absolute"
    );
    let path = path.to_string_lossy().replace('\\', "/");
    #[cfg(windows)]
    let path = if path.starts_with('/') {
        path
    } else {
        format!("/{path}")
    };
    let mut encoded = String::with_capacity(path.len());
    for byte in path.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'/' | b':' | b'-' | b'.' | b'_' | b'~') {
            encoded.push(byte as char);
        } else {
            encoded.push_str(&format!("%{byte:02X}"));
        }
    }
    Ok(format!("file://{encoded}"))
}

fn validate_csv_conversion(input: &Path, converted: &WorkbookDump) -> Result<()> {
    let bytes = fs::read(input).with_context(|| format!("read CSV {}", input.display()))?;
    let bytes = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(&bytes);
    let decoded = match std::str::from_utf8(bytes) {
        Ok(text) => text.to_owned(),
        Err(_) => {
            let (text, _, had_errors) = GBK.decode(bytes);
            ensure!(
                !had_errors,
                "CSV is neither valid UTF-8 nor valid GBK; cannot validate converted values"
            );
            text.into_owned()
        }
    };
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(decoded.as_bytes());
    let rows = reader
        .records()
        .collect::<std::result::Result<Vec<_>, _>>()
        .context("parse CSV for conversion validation")?;
    ensure!(
        converted.sheets.len() == 1,
        "CSV conversion must produce exactly one worksheet"
    );
    let sheet = &converted.sheets[0];
    let mut expected_cells = std::collections::BTreeSet::new();
    for (row_index, row) in rows.iter().enumerate() {
        for (column_index, raw) in row.iter().enumerate() {
            if raw.is_empty() {
                continue;
            }
            let address = format!("{}{}", csv_column_name(column_index + 1), row_index + 1);
            expected_cells.insert(address.clone());
            if let Some(formula) = sheet.formulas.get(&address) {
                ensure!(
                    raw.starts_with('=')
                        && raw.trim_start_matches('=').trim()
                            == formula.trim_start_matches('=').trim(),
                    "CSV conversion changed formula text at {}!{address}",
                    sheet.name
                );
                continue;
            }
            let value = sheet
                .values
                .get(&address)
                .with_context(|| format!("CSV conversion lost data at {}!{address}", sheet.name))?;
            ensure!(
                csv_text_matches_probe(raw, value),
                "CSV conversion changed data at {}!{address}: source={raw:?}, output={value}",
                sheet.name
            );
        }
    }
    for address in sheet.values.keys().chain(sheet.formulas.keys()) {
        let non_empty = sheet
            .values
            .get(address)
            .and_then(|value| value.get("value"))
            .is_some_and(|value| {
                !value.is_null() && value.as_str().is_none_or(|text| !text.is_empty())
            });
        if non_empty || sheet.formulas.contains_key(address) {
            ensure!(
                expected_cells.contains(address),
                "CSV conversion introduced data at {}!{address}",
                sheet.name
            );
        }
    }
    Ok(())
}

fn csv_text_matches_probe(source: &str, actual: &serde_json::Value) -> bool {
    let Some(kind) = actual.get("type").and_then(serde_json::Value::as_str) else {
        return false;
    };
    let Some(value) = actual.get("value") else {
        return false;
    };
    match kind {
        "string" | "error" => value.as_str() == Some(source),
        "number" | "integer" | "datetime_serial" => {
            let Some(left) = source.trim().parse::<f64>().ok() else {
                return false;
            };
            let Some(right) = value.as_f64() else {
                return false;
            };
            left.is_finite()
                && right.is_finite()
                && (left - right).abs() <= 1e-15 * left.abs().max(right.abs()).max(1.0)
        }
        "boolean" => {
            let source = source.trim();
            match value.as_bool() {
                Some(true) => source.eq_ignore_ascii_case("true") || source == "1",
                Some(false) => source.eq_ignore_ascii_case("false") || source == "0",
                None => false,
            }
        }
        "datetime_iso" => {
            let source = source.trim().replace('/', "-").replace(' ', "T");
            let actual = value.as_str().unwrap_or_default();
            source == actual
                || (source.len() == 10 && actual.starts_with(&source))
                || chrono::NaiveDate::parse_from_str(&source, "%Y-%m-%d")
                    .ok()
                    .is_some_and(|date| actual.starts_with(&date.format("%Y-%m-%d").to_string()))
        }
        _ => false,
    }
}

fn csv_column_name(mut column: usize) -> String {
    let mut name = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        name.insert(0, (b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    name
}

struct TempDirGuard(PathBuf);

impl Drop for TempDirGuard {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[cfg(test)]
mod tests {
    #[cfg(target_os = "linux")]
    use super::{
        extract_named_validation_cells, parse_named_validation_region, NamedValidationRegionRequest,
    };
    use super::{
        file_uri, libreoffice_batch_request, libreoffice_user_installation_arg,
        select_conversion_engine, select_fidelity_engine, unique_sibling_conversion_candidate,
        ConversionRuntime, LibreOfficeExecutor, LibreOfficeRunOutcome,
    };
    use anyhow::Result;
    use excel_editor::OpenpyxlWorkerConfig;
    #[cfg(target_os = "linux")]
    use excel_reader::{SheetDump, WorkbookDump};
    use ooxml_patch::{copy_cell_region, CellRegionCopyRequest};
    #[cfg(target_os = "linux")]
    use std::collections::BTreeMap;
    use std::fs::{self, File};
    use std::io::{Read, Write};
    use std::path::{Path, PathBuf};
    use std::sync::atomic::{AtomicU32, Ordering};
    use std::sync::{Arc, Mutex};
    use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

    fn fake_input(dir: &Path, name: &str) -> PathBuf {
        let path = dir.join(name);
        fs::write(&path, b"synthetic conversion payload").expect("write synthetic input");
        path
    }

    fn unique_test_root(tag: &str) -> Result<PathBuf> {
        let root = std::env::temp_dir().join(format!(
            "pytools-lo-session-{tag}-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        fs::create_dir_all(&root)?;
        Ok(root)
    }

    fn add_formula_sheet_drawing_relation(path: &Path) -> Result<()> {
        let staged = path.with_file_name(format!(
            "{}.staged.xlsx",
            path.file_stem().unwrap().to_string_lossy()
        ));
        let mut source = ZipArchive::new(File::open(path)?)?;
        let mut target = ZipWriter::new(File::create(&staged)?);
        for index in 0..source.len() {
            let mut entry = source.by_index(index)?;
            let name = entry.name().to_owned();
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes)?;
            if name == "xl/worksheets/sheet2.xml" {
                let mut xml = String::from_utf8(bytes)?;
                if !xml.contains("xmlns:r=") {
                    xml = xml.replacen(
                        "<worksheet ",
                        "<worksheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
                        1,
                    );
                }
                xml = xml.replace(
                    "</worksheet>",
                    "<drawing r:id=\"rIdFallbackTest\"/></worksheet>",
                );
                bytes = xml.into_bytes();
            } else if name == "[Content_Types].xml" {
                let mut xml = String::from_utf8(bytes)?;
                xml = xml.replace(
                    "</Types>",
                    "<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/></Types>",
                );
                bytes = xml.into_bytes();
            }
            target.start_file(name, SimpleFileOptions::default())?;
            target.write_all(&bytes)?;
        }
        target.start_file(
            "xl/worksheets/_rels/sheet2.xml.rels",
            SimpleFileOptions::default(),
        )?;
        target.write_all(
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFallbackTest" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
        )?;
        target.start_file("xl/drawings/drawing1.xml", SimpleFileOptions::default())?;
        target.write_all(
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>"#,
        )?;
        target.finish()?;
        fs::remove_file(path)?;
        fs::rename(staged, path)?;
        Ok(())
    }

    #[test]
    #[ignore = "requires the packaged Openpyxl runtime; run explicitly for fallback verification"]
    fn named_validation_uses_protected_openpyxl_fallback_for_unsupported_helper_relations(
    ) -> Result<()> {
        let root = unique_test_root("named-validation-openpyxl-fallback")?;
        let template = root.join("template.xlsx");
        let working_copy = root.join("source-copy.xlsx");

        let mut template_book = rust_xlsxwriter::Workbook::new();
        let validation = template_book.add_worksheet();
        validation.set_name("模板")?;
        validation.write_formula(1, 1, "='参照表'!D2")?;
        let helper = template_book.add_worksheet();
        helper.set_name("参照表")?;
        helper.write_formula(1, 3, "='基础参数'!A1*2")?;
        let parameters = template_book.add_worksheet();
        parameters.set_name("基础参数")?;
        parameters.write_number(0, 0, 21.0)?;
        template_book.save(&template)?;
        add_formula_sheet_drawing_relation(&template)?;

        let mut source_book = rust_xlsxwriter::Workbook::new();
        source_book.add_worksheet().set_name("检查表")?;
        source_book.save(&working_copy)?;
        let requests = [CellRegionCopyRequest {
            source_sheet: "模板".to_owned(),
            target_sheet: "检查表".to_owned(),
            address: "B2".to_owned(),
        }];

        let direct_error = copy_cell_region(&template, &working_copy, &requests)
            .expect_err("unsupported drawing relationship should select fallback");
        assert!(
            format!("{direct_error:#}").contains("OPENPYXL_FALLBACK_REQUIRED"),
            "{direct_error:#}"
        );
        let dependencies =
            ooxml_patch::missing_formula_dependency_sheets(&template, &working_copy, &requests)?;
        assert_eq!(dependencies.len(), 2);

        let dist_python = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../dist/python");
        let config = OpenpyxlWorkerConfig::new(
            dist_python.join("python.exe"),
            dist_python.join("openpyxl_worker.py"),
        )?;
        ConversionRuntime::import_validation_dependencies_with_config(
            &template,
            &working_copy,
            &dependencies,
            config,
        )?;
        let copied = copy_cell_region(&template, &working_copy, &requests)?;
        assert_eq!(copied.cells_with_formula, 1);
        let output = super::read_workbook(&working_copy)?;
        let names = output
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>();
        assert!(names.contains(&"参照表"));
        assert!(names.contains(&"基础参数"));
        assert_eq!(
            output.sheets[0].formulas.get("B2").map(String::as_str),
            Some("'参照表'!D2")
        );
        assert_eq!(
            output
                .sheets
                .iter()
                .find(|sheet| sheet.name == "参照表")
                .and_then(|sheet| sheet.formulas.get("D2"))
                .map(String::as_str),
            Some("'基础参数'!A1*2")
        );
        drop(output);
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn office_candidate_uses_the_final_output_directory_and_is_unique() -> Result<()> {
        let root = unique_test_root("office-candidate")?;
        let output = root.join("delivered.xlsx");
        let first = unique_sibling_conversion_candidate(&output, "xlsx")?;
        let second = unique_sibling_conversion_candidate(&output, "xlsx")?;

        assert_eq!(first.parent(), output.parent());
        assert_eq!(
            first.extension().and_then(|value| value.to_str()),
            Some("xlsx")
        );
        assert_ne!(first, second);
        assert!(!first.exists());
        assert!(!second.exists());

        fs::remove_dir_all(root)?;
        Ok(())
    }

    /// A fake `soffice` runner: records every invocation, derives the output
    /// extension from `--convert-to`, and materializes outputs for the input
    /// positions listed in `produce_outputs` (0-based positions of the
    /// trailing input file arguments). `fail_call` aborts one call; a true
    /// `fail_status` reports a non-zero exit instead.
    fn fake_executor(
        produce_per_call: &[&[usize]],
        fail_call: Option<usize>,
        fail_status: bool,
    ) -> (
        LibreOfficeExecutor,
        Arc<Mutex<Vec<Vec<String>>>>,
        Arc<AtomicU32>,
    ) {
        let calls: Arc<Mutex<Vec<Vec<String>>>> = Arc::new(Mutex::new(Vec::new()));
        let counter = Arc::new(AtomicU32::new(0));
        let produce: Vec<Vec<usize>> = produce_per_call
            .iter()
            .map(|positions| positions.to_vec())
            .collect();
        let call_counter = counter.clone();
        let calls_for_closure = calls.clone();
        let executor: LibreOfficeExecutor =
            Box::new(move |arguments: &[String], outputs_dir: &Path| {
                let call_index = call_counter.fetch_add(1, Ordering::SeqCst) as usize;
                calls_for_closure.lock().unwrap().push(arguments.to_vec());
                if let Some(fail_at) = fail_call {
                    if call_index == fail_at {
                        return Err(anyhow::anyhow!("fake soffice spawn failure"));
                    }
                }
                let filter = arguments
                    .iter()
                    .position(|argument| argument == "--convert-to")
                    .and_then(|index| arguments.get(index + 1))
                    .map(|value| value.to_owned())
                    .unwrap_or_default();
                let output_ext = filter.split(':').next().unwrap_or_default().to_owned();
                let inputs = arguments
                    .iter()
                    .skip_while(|argument| *argument != "--outdir")
                    .skip(2)
                    .collect::<Vec<_>>();
                let produced = produce
                    .get(call_index)
                    .cloned()
                    .unwrap_or_else(|| (0..inputs.len()).collect());
                for (index, input) in inputs.iter().enumerate() {
                    if !produced.contains(&index) {
                        continue;
                    }
                    let stem = Path::new(input)
                        .file_stem()
                        .map(|stem| stem.to_string_lossy().into_owned())
                        .unwrap_or_default();
                    let output = outputs_dir.join(format!("{stem}.{output_ext}"));
                    fs::write(&output, b"fake conversion output")?;
                }
                if fail_status {
                    return Ok(LibreOfficeRunOutcome {
                        process_id: 424242 + call_index as u32,
                        status_success: false,
                        stderr: "fake soffice failure".to_owned(),
                    });
                }
                Ok(LibreOfficeRunOutcome {
                    process_id: 424242 + call_index as u32,
                    status_success: true,
                    stderr: String::new(),
                })
            });
        (executor, calls, counter)
    }

    #[test]
    fn lo_batch_converts_all_files_in_one_process() -> Result<()> {
        let root = unique_test_root("ok")?;
        let inputs: Vec<PathBuf> = (0..3)
            .map(|index| fake_input(&root, &format!("batch-{index}.xls")))
            .collect();
        let mut runtime = ConversionRuntime::default();
        let (executor, calls, _counter) = fake_executor(&[&[0, 1, 2]], None, false);
        runtime.set_libreoffice_executor(executor);
        let requests = inputs
            .iter()
            .map(|input| (input.clone(), "xlsx".to_owned()))
            .collect::<Vec<_>>();
        runtime.prepare_libreoffice_batch(&requests);

        // Exactly one soffice process served the whole batch.
        assert_eq!(calls.lock().unwrap().len(), 1);
        assert_eq!(runtime.libreoffice_process_ids().len(), 1);
        let arguments = calls.lock().unwrap()[0].clone();
        assert_eq!(
            arguments
                .iter()
                .filter(|argument| argument.starts_with("-env:UserInstallation="))
                .count(),
            1,
            "one shared isolated profile"
        );
        assert_eq!(
            arguments
                .iter()
                .filter(|argument| argument.contains("inputs") && argument.ends_with(".xls"))
                .count(),
            3,
            "all inputs staged into the single invocation"
        );

        for input in &inputs {
            match runtime.libreoffice_batch_outcome(input, "xlsx") {
                Some(Ok(candidate)) => assert!(candidate.is_file()),
                other => panic!("expected candidate for {input:?}: {other:?}"),
            }
            assert!(runtime.libreoffice_covers(input, "xlsx"));
        }
        let session_dir = runtime.libreoffice_session_dir().expect("session created");
        assert!(session_dir.is_dir());
        drop(runtime);
        assert!(
            !session_dir.exists(),
            "the session directory must be cleaned up on drop"
        );
        fs::remove_dir_all(&root)?;
        Ok(())
    }

    #[test]
    fn lo_batch_startup_failure_fails_all_covered_files_without_retry() -> Result<()> {
        let root = unique_test_root("spawn-fail")?;
        let inputs: Vec<PathBuf> = (0..3)
            .map(|index| fake_input(&root, &format!("fail-{index}.xls")))
            .collect();
        let mut runtime = ConversionRuntime::default();
        let (executor, calls, _counter) = fake_executor(&[&[]], Some(0), false);
        runtime.set_libreoffice_executor(executor);
        let requests = inputs
            .iter()
            .map(|input| (input.clone(), "xlsx".to_owned()))
            .collect::<Vec<_>>();
        runtime.prepare_libreoffice_batch(&requests);

        // One attempted start; no per-file retries afterwards.
        assert_eq!(calls.lock().unwrap().len(), 1);
        for input in &inputs {
            let outcome = runtime.libreoffice_batch_outcome(input, "xlsx");
            assert!(
                matches!(&outcome, Some(Err(message)) if message.contains("fake soffice spawn failure")),
                "startup failure must surface per file: {outcome:?}"
            );
        }
        drop(runtime);
        fs::remove_dir_all(&root)?;
        Ok(())
    }

    #[test]
    fn lo_batch_process_failure_marks_all_files_failed() -> Result<()> {
        let root = unique_test_root("status-fail")?;
        let inputs: Vec<PathBuf> = (0..2)
            .map(|index| fake_input(&root, &format!("status-{index}.xls")))
            .collect();
        let mut runtime = ConversionRuntime::default();
        let (executor, _calls, _counter) = fake_executor(&[&[]], None, true);
        runtime.set_libreoffice_executor(executor);
        let requests = inputs
            .iter()
            .map(|input| (input.clone(), "xlsx".to_owned()))
            .collect::<Vec<_>>();
        runtime.prepare_libreoffice_batch(&requests);
        for input in &inputs {
            assert!(matches!(
                runtime.libreoffice_batch_outcome(input, "xlsx"),
                Some(Err(message)) if message.contains("fake soffice failure")
            ));
        }
        drop(runtime);
        fs::remove_dir_all(&root)?;
        Ok(())
    }

    #[test]
    fn lo_batch_missing_output_fails_only_that_file() -> Result<()> {
        let root = unique_test_root("partial")?;
        let inputs: Vec<PathBuf> = (0..3)
            .map(|index| fake_input(&root, &format!("partial-{index}.xls")))
            .collect();
        let mut runtime = ConversionRuntime::default();
        // Only the first two inputs receive an output from the fake process.
        let (executor, calls, _counter) = fake_executor(&[&[0, 1]], None, false);
        runtime.set_libreoffice_executor(executor);
        let requests = inputs
            .iter()
            .map(|input| (input.clone(), "xlsx".to_owned()))
            .collect::<Vec<_>>();
        runtime.prepare_libreoffice_batch(&requests);
        assert_eq!(calls.lock().unwrap().len(), 1, "still one process");
        assert!(matches!(
            runtime.libreoffice_batch_outcome(&inputs[0], "xlsx"),
            Some(Ok(_))
        ));
        assert!(matches!(
            runtime.libreoffice_batch_outcome(&inputs[1], "xlsx"),
            Some(Ok(_))
        ));
        assert!(matches!(
            runtime.libreoffice_batch_outcome(&inputs[2], "xlsx"),
            Some(Err(message)) if message.contains("did not produce an output")
        ));
        drop(runtime);
        fs::remove_dir_all(&root)?;
        Ok(())
    }

    #[test]
    fn lo_word_batch_uses_two_processes_and_preextracts_texts() -> Result<()> {
        let root = unique_test_root("word")?;
        let inputs: Vec<PathBuf> = (0..3)
            .map(|index| fake_input(&root, &format!("word-{index}.doc")))
            .collect();
        let mut runtime = ConversionRuntime::default();
        let (executor, calls, _counter) =
            fake_executor(&[&[0, 1, 2], &[0, 1, 2, 3, 4, 5]], None, false);
        runtime.set_libreoffice_executor(executor);
        let requests = inputs
            .iter()
            .map(|input| (input.clone(), "docx".to_owned()))
            .collect::<Vec<_>>();
        runtime.prepare_libreoffice_batch(&requests);

        // One process for the docx conversion, one for all text extraction.
        assert_eq!(calls.lock().unwrap().len(), 2);
        assert_eq!(runtime.libreoffice_process_ids().len(), 2);
        let filters = calls
            .lock()
            .unwrap()
            .iter()
            .map(|arguments| {
                arguments
                    .iter()
                    .position(|argument| argument == "--convert-to")
                    .and_then(|index| arguments.get(index + 1))
                    .cloned()
                    .unwrap_or_default()
            })
            .collect::<Vec<_>>();
        assert_eq!(filters, vec!["docx:Office Open XML Text", "txt:Text"]);
        for input in &inputs {
            assert!(runtime.libreoffice_source_text(input).is_some());
            assert!(runtime.libreoffice_candidate_text(input, "docx").is_some());
        }
        drop(runtime);
        fs::remove_dir_all(&root)?;
        Ok(())
    }

    #[test]
    fn libreoffice_batch_request_matrix_matches_per_file_routing() {
        assert_eq!(select_fidelity_engine("auto"), "pywin32");
        assert_eq!(select_fidelity_engine("lo"), "lo");
        assert_eq!(select_fidelity_engine("com"), "pywin32");
        assert_eq!(select_fidelity_engine("pywin32"), "pywin32");
        assert!(!libreoffice_batch_request("xls", "xlsx", "lo"));
        assert!(!libreoffice_batch_request("xlsb", "xlsx", "lo"));
        assert!(!libreoffice_batch_request("xlsm", "xlsx", "lo"));
        assert!(!libreoffice_batch_request("doc", "docx", "lo"));
        // WPS-native containers stay per-file: only COM can prove the source
        // really is a native WPS file, so an LO result would be rejected.
        assert!(!libreoffice_batch_request("et", "xlsx", "lo"));
        assert!(!libreoffice_batch_request("ett", "xlsx", "lo"));
        assert!(!libreoffice_batch_request("wps", "docx", "lo"));
        assert!(!libreoffice_batch_request("wpt", "docx", "lo"));
        // Excel format normalization remains COM-only as well.
        assert!(!libreoffice_batch_request("xlsx", "xlsx", "lo"));
        assert!(!libreoffice_batch_request("xls", "xlsx", "com"));
        assert!(!libreoffice_batch_request("xls", "xlsx", "auto"));
    }

    #[cfg(target_os = "linux")]
    #[test]
    fn named_validation_result_extraction_uses_only_formula_text_in_region() -> Result<()> {
        let workbook = WorkbookDump {
            schema_version: 1,
            reader: "calamine",
            elapsed_ms: 0,
            sheets: vec![SheetDump {
                name: "检查表".to_owned(),
                values: BTreeMap::from([
                    (
                        "B2".to_owned(),
                        serde_json::json!({"type": "string", "value": "错|校验项|差异"}),
                    ),
                    (
                        "B4".to_owned(),
                        serde_json::json!({"type": "string", "value": "硬性|校验项|差异"}),
                    ),
                    (
                        "B8".to_owned(),
                        serde_json::json!({"type": "string", "value": "软性|超出区域"}),
                    ),
                    (
                        "C3".to_owned(),
                        serde_json::json!({"type": "string", "value": "固定说明，不是公式"}),
                    ),
                ]),
                formulas: BTreeMap::from([
                    ("B2".to_owned(), "=A2".to_owned()),
                    ("B4".to_owned(), "=A4".to_owned()),
                    ("B8".to_owned(), "=A8".to_owned()),
                ]),
                formula_status: "read_by_calamine".to_owned(),
                merges: Vec::new(),
                merge_status: "read_by_calamine".to_owned(),
                comments: BTreeMap::new(),
                comments_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
                conditional_formatting: Vec::new(),
                conditional_formatting_status: "read_by_namespace_aware_ooxml_supplement"
                    .to_owned(),
                data_validations: Vec::new(),
                data_validations_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            }],
        };
        let regions = [NamedValidationRegionRequest {
            template_sheet: "模板".to_owned(),
            source_sheet: "检查表".to_owned(),
            address: "$B$2:$B$4".to_owned(),
        }];

        let cells = extract_named_validation_cells(&workbook, &regions)?;
        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0].address, "B2");
        assert!(cells.iter().all(|cell| cell.sheet == "检查表"));
        assert!(!cells.iter().any(|cell| cell.address == "B8"));
        Ok(())
    }

    #[cfg(target_os = "linux")]
    #[test]
    fn named_validation_region_parser_accepts_single_cells_and_sheet_qualified_ranges() -> Result<()>
    {
        assert_eq!(
            parse_named_validation_region("'检查 表'!$B$2:$D$4")?,
            (2, 2, 4, 4)
        );
        assert_eq!(parse_named_validation_region("C3")?, (3, 3, 3, 3));
        assert!(parse_named_validation_region("A1,B2").is_err());
        Ok(())
    }

    #[cfg(target_os = "linux")]
    #[test]
    #[ignore = "requires LibreOffice Calc and python3-uno; run explicitly on Linux"]
    fn libreoffice_recalculates_validation_formula_cache_on_a_working_copy() -> Result<()> {
        let root = unique_test_root("named-validation-lo")?;
        let template = root.join("template.xlsx");
        let source_copy = root.join("source-copy.xlsx");

        let mut template_workbook = rust_xlsxwriter::Workbook::new();
        let template_sheet = template_workbook.add_worksheet();
        template_sheet.set_name("模板")?;
        template_sheet.write_formula(1, 1, "=IF(A2>0,\"错|校验项|差异\",\"\")")?;
        template_workbook.save(&template)?;

        let mut source_workbook = rust_xlsxwriter::Workbook::new();
        let source_sheet = source_workbook.add_worksheet();
        source_sheet.set_name("检查表")?;
        source_sheet.write_number(1, 0, 1)?;
        source_workbook.save(&source_copy)?;

        let regions = [NamedValidationRegionRequest {
            template_sheet: "模板".to_owned(),
            source_sheet: "检查表".to_owned(),
            address: "B2:B2".to_owned(),
        }];
        let mut runtime = ConversionRuntime::default();
        let calculation =
            runtime.recalculate_named_validation_regions(&template, &source_copy, &regions)?;

        assert_eq!(calculation.application, "LibreOffice Calc");
        assert_eq!(calculation.copied_formula_cells, 1);
        assert_eq!(calculation.cells.len(), 1);
        assert_eq!(calculation.cells[0].sheet, "检查表");
        assert_eq!(calculation.cells[0].address, "B2");
        assert!(calculation.cells[0].value.contains("错|校验项|差异"));
        drop(runtime);
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn automatic_windows_conversion_uses_the_packaged_pywin32_worker() {
        for source in ["xls", "xlsb", "xlsm", "et", "ett", "xlsx"] {
            assert_eq!(
                select_conversion_engine("auto", source, "xlsx", true),
                "pywin32",
                "auto on Windows must route .{source} through pywin32"
            );
        }
        for source in ["doc", "wps", "wpt"] {
            assert_eq!(
                select_conversion_engine("auto", source, "docx", true),
                "pywin32",
                "auto on Windows must route .{source} through pywin32"
            );
        }
        // Linux：电子表格 .xls/.xlsm → .xlsx 走 LibreOffice（2026-11-09 实测
        // 与 COM 另存一致）。其余格式仍要求 Windows pywin32。
        assert_eq!(select_conversion_engine("auto", "xls", "xlsx", false), "lo");
        assert_eq!(
            select_conversion_engine("auto", "xlsm", "xlsx", false),
            "lo"
        );
        // 同扩展名 .xlsx 规范化与 .xlsb/.et 仍不在 Linux LO 路由内。
        assert_eq!(
            select_conversion_engine("auto", "xlsx", "xlsx", false),
            "pywin32"
        );
        assert_eq!(
            select_conversion_engine("auto", "xlsb", "xlsx", false),
            "pywin32"
        );
        assert_eq!(
            select_conversion_engine("auto", "doc", "docx", false),
            "pywin32"
        );
        // 显式 lo 在 Windows 上仍被选中（由调用方给出），显式 com 保持 pywin32。
        assert_eq!(select_conversion_engine("lo", "xls", "xlsx", true), "lo");
        // The retired Rust COM channel is only kept as a CLI alias.
        assert_eq!(
            select_conversion_engine("com", "xls", "xlsx", true),
            "pywin32"
        );
        assert_eq!(
            select_conversion_engine("pywin32", "xls", "xlsx", true),
            "pywin32"
        );
    }

    #[test]
    fn libreoffice_profile_path_is_encoded_as_a_file_uri() {
        let root = std::env::temp_dir().join("pytools engine profile # 中文");
        let uri = file_uri(&root).unwrap();
        assert!(uri.starts_with("file://"));
        assert!(uri.contains("%20"));
        assert!(uri.contains("%23"));
        assert!(!uri.contains(' '));
    }

    #[test]
    fn libreoffice_user_installation_argument_creates_profile() {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let profile = std::env::temp_dir().join(format!(
            "pytools_engine_profile_{}_{nonce}",
            std::process::id()
        ));
        let _ = std::fs::remove_dir_all(&profile);
        let argument = libreoffice_user_installation_arg(&profile).unwrap();
        assert!(profile.is_dir());
        assert!(argument.starts_with("-env:UserInstallation=file://"));
        std::fs::remove_dir_all(profile).ok();
    }
}
