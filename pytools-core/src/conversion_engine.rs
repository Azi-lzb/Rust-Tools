//! Fidelity conversion engine (former `lo-convert` sidecar).
//! Windows spreadsheet conversion uses Rust COM on an isolated STA thread;
//! Word retains its PowerShell COM bridge. LibreOffice uses the local
//! `soffice` CLI. No separate `lo-convert` executable is required.
//!
//! LibreOffice runs as a batch-scoped session: one shared isolated profile
//! and one `soffice` process per conversion stage handle the whole batch
//! (inputs are staged as uniquely named disposable copies). Direct
//! single-file calls without a prepared batch keep the historical
//! one-process-per-file behavior.

use anyhow::{anyhow, bail, ensure, Context, Result};
use encoding_rs::GBK;
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

/// State scoped to one user conversion request or batch.
pub(crate) struct ConversionRuntime {
    #[cfg(windows)]
    spreadsheet_com: Option<crate::conversion_com::ComWorker>,
    libreoffice: Option<LibreOfficeBatchSession>,
    /// Test hook: replaces the real `soffice` process execution.
    libreoffice_executor: Option<LibreOfficeExecutor>,
}

impl Default for ConversionRuntime {
    fn default() -> Self {
        Self {
            #[cfg(windows)]
            spreadsheet_com: None,
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
pub(crate) fn select_fidelity_engine(engine: &str) -> &'static str {
    if engine == "auto" {
        if cfg!(windows) {
            "com"
        } else {
            "lo"
        }
    } else if engine == "lo" {
        "lo"
    } else {
        "com"
    }
}

/// Whether this source would run through LibreOffice under the requested
/// engine, using the same source/target matrix the per-file path accepts.
pub(crate) fn libreoffice_batch_request(source_ext: &str, target: &str, engine: &str) -> bool {
    if select_fidelity_engine(engine) != "lo" {
        return false;
    }
    // .et stays on the per-file path: its conversion result is rejected
    // after the fact by the ET source-format evidence check in
    // `conversion.rs`, so batching it cannot succeed anyway.
    (target == "xlsx" && matches!(source_ext, "xls" | "xlsb" | "xlsm"))
        || (target == "docx" && source_ext == "doc")
}

impl ConversionRuntime {
    #[cfg(windows)]
    fn spreadsheet_com(
        &mut self,
        input: &Path,
        output: &Path,
        source_ext: &str,
    ) -> Result<ComExecution> {
        if self.spreadsheet_com.is_none() {
            self.spreadsheet_com = Some(crate::conversion_com::ComWorker::start()?);
        }
        let converted = self.spreadsheet_com
            .as_mut()
            .expect("spreadsheet COM worker initialized")
            .convert(input, output, source_ext)?;
        Ok(ComExecution {
            source_format: converted.source_format,
            bridge: "rust",
            application: Some(converted.application),
            process_id: Some(converted.process_id),
        })
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
            session
                .covered
                .insert((input.clone(), target.clone()));
        }
        let mut executor = self.libreoffice_executor.take();
        let outcome = Self::convert_batch_with_session(&mut session, &unique_requests, &mut executor);
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
        let source_ext = source_exts.into_iter().next().expect("source format checked");

        // Stage disposable, uniquely named copies so same-stem inputs cannot
        // collide inside the shared outdir and the engine never sees the
        // caller's original file.
        let mut main_jobs = Vec::new();
        for (index, (input, request_target)) in requests.iter().enumerate() {
            let staged_input =
                inputs_dir.join(format!("lo-{index:06}.{source_ext}"));
            fs::copy(input, &staged_input).with_context(|| {
                format!(
                    "stage disposable copy for LibreOffice batch conversion from {}",
                    input.display()
                )
            })?;
            let staged_output = outputs_dir_for(session).join(format!(
                "lo-{index:06}.{request_target}"
            ));
            main_jobs.push(LibreOfficeJob {
                staged_input,
                staged_output,
            });
        }
        session.run_libreoffice_processes(
            executor,
            vec![(
                filter.to_owned(),
                outputs_dir_for(session),
                main_jobs,
            )],
        )?;
        for (index, (input, request_target)) in requests.iter().enumerate() {
            let staged_output = outputs_dir_for(session).join(format!(
                "lo-{index:06}.{request_target}"
            ));
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
                let Some(candidate) = session.converted.get(&(input.clone(), request_target.clone())) else {
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
        self.libreoffice_session()
            .is_some_and(|session| session.covered.contains(&(input.to_path_buf(), target.to_owned())))
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
    bridge: &'static str,
    application: Option<&'static str>,
    process_id: Option<u32>,
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
        engine == "auto" || engine == "lo" || engine == "com",
        "engine must be auto, lo or com"
    );
    ensure!(
        input.is_file(),
        "input file does not exist: {}",
        input.display()
    );
    let source_ext = extension(input)?;
    ensure!(
        (matches!(target, "xlsx")
            && matches!(source_ext.as_str(), "xls" | "xlsb" | "xlsm" | "et"))
            || (target == "docx" && source_ext == "doc"),
        "unsupported conversion: .{source_ext} -> .{target}"
    );
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
    let converted = temp_dir.join(format!("{}.{}", file_stem(input), target));
    let selected = if engine == "auto" {
        if cfg!(windows) {
            "com"
        } else {
            "lo"
        }
    } else {
        engine
    };
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
        "com" if cfg!(windows) => Some(run_com(input, &converted, target, &source_ext, runtime)?),
        "com" => bail!(
            "该功能需要外部引擎: COM conversion is only available on Windows; install LibreOffice or use --engine lo"
        ),
        _ => unreachable!(),
    };
    let source_file_format = com_execution.as_ref().and_then(|execution| execution.source_format);
    let candidate = if converted.exists() {
        converted
    } else {
        locate_converted(temp_dir, target)?
    };
    ensure!(
        candidate.is_file() && fs::metadata(&candidate)?.len() > 0,
        "conversion engine did not create a non-empty output file"
    );
    let validation = validate_conversion(
        input,
        &candidate,
        &source_ext,
        target,
        selected,
        temp_dir,
        source_file_format,
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

fn run_com(
    input: &Path,
    output: &Path,
    target: &str,
    source_ext: &str,
    runtime: &mut ConversionRuntime,
) -> Result<ComExecution> {
    let output_parent = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(output_parent)?;
    let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    let com_input = output_parent.join(format!(
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
    if target == "xlsx" {
        #[cfg(windows)]
        let result = runtime.spreadsheet_com(&com_input, output, source_ext);
        #[cfg(not(windows))]
        let result = {
            let _ = runtime;
            Err(anyhow!("COM conversion is only available on Windows"))
        };
        let _ = fs::remove_file(&com_input);
        return result;
    }
    let app_names: &[&str] = if target == "docx" {
        &[
            "Word.Application",
            "kwps.Application",
            "KWPS.Application",
            "wps.Application",
        ]
    } else if source_ext == "et" {
        // ET is a WPS-native workbook format. Do not route it through Excel
        // COM, which may open a renamed XLS and make the extension look valid.
        &["KET.Application"]
    } else {
        &["Excel.Application", "ket.Application", "KET.Application"]
    };
    let format_code = if target == "docx" { 16 } else { 51 };
    let mut script = String::from("$ErrorActionPreference='Stop';\n");
    script.push_str("$app=$null; $appName=''; $lastError='';\n");
    for app_name in app_names {
        script.push_str(&format!(
            "if ($null -eq $app) {{ try {{ $app=New-Object -ComObject '{}'; $appName='{}' }} catch {{ $lastError=$_.Exception.Message }} }}\n",
            app_name, app_name
        ));
    }
    script.push_str("if ($null -eq $app) { throw \"No supported Office/WPS COM engine is available: $lastError\" }\n");
    script.push_str("$doc=$null; $book=$null;\ntry {\n$app.Visible=$false; ");
    script.push_str(com_display_alerts_setting(target));
    script.push('\n');
    if target == "docx" {
        script.push_str(&format!("$doc=$app.Documents.Open('{}',$false,$true); try {{ $doc.SaveAs2('{}',{}) }} catch {{ $doc.SaveAs('{}',{}) }};\n", ps_quote(&com_input.to_string_lossy()), ps_quote(&output.to_string_lossy()), format_code, ps_quote(&output.to_string_lossy()), format_code));
    } else {
        script.push_str(&format!(
            "$book=$app.Workbooks.Open('{}',0,$true);\n",
            ps_quote(&com_input.to_string_lossy())
        ));
        if source_ext == "et" {
            // Inspect the format WPS actually opened before publishing an
            // XLSX. A CFB/XLS file renamed to .et reports 56; native ET is
            // identified by KET as 65521.
            script.push_str(
                "$sourceFormat=[int]$book.FileFormat; if ($sourceFormat -ne 65521) { throw \"Input has .et extension but KET opened FileFormat $sourceFormat; expected native ET FileFormat 65521\" }; [Console]::Out.WriteLine(\"PYTOOLS_SOURCE_FILE_FORMAT=$sourceFormat\");\n",
            );
        }
        script.push_str(&format!(
            "$book.SaveAs('{}',{});\n",
            ps_quote(&output.to_string_lossy()),
            format_code
        ));
    }
    script.push_str("} finally { if ($null -ne $doc) { $doc.Close($false) }; if ($null -ne $book) { $book.Close($false) }; if ($null -ne $app) { $app.Quit() } }\n");
    let script_path = output.with_extension("pytools-convert.ps1");
    let result = (|| {
        write_powershell_script(&script_path, &script)?;
        Command::new("powershell.exe")
            .args([
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
            ])
            .arg(&script_path)
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .output()
            .with_context(|| "start Office/WPS COM automation")
    })();
    let _ = fs::remove_file(script_path);
    let _ = fs::remove_file(com_input);
    let result = result?;
    if !result.status.success() {
        bail!(
            "Office/WPS COM conversion failed: {}",
            String::from_utf8_lossy(&result.stderr)
        );
    }
    if source_ext == "et" {
        let file_format = String::from_utf8_lossy(&result.stdout)
            .lines()
            .find_map(|line| line.trim().strip_prefix("PYTOOLS_SOURCE_FILE_FORMAT="))
            .context("KET did not return source FileFormat evidence for .et input")?
            .parse::<i32>()
            .context("KET returned invalid source FileFormat evidence")?;
        ensure!(
            file_format == 65521,
            "KET opened .et input as FileFormat {file_format}; expected native ET FileFormat 65521"
        );
        Ok(ComExecution {
            source_format: Some(file_format),
            bridge: "powershell",
            application: None,
            process_id: None,
        })
    } else {
        Ok(ComExecution {
            source_format: None,
            bridge: "powershell",
            application: None,
            process_id: None,
        })
    }
}

fn com_display_alerts_setting(target: &str) -> &'static str {
    if target == "docx" {
        // Word expects the WdAlertLevel enum (wdAlertsNone = 0), not a Boolean.
        "$app.DisplayAlerts=0;"
    } else {
        "$app.DisplayAlerts=$false;"
    }
}

fn write_powershell_script(path: &Path, script: &str) -> Result<()> {
    let mut encoded = Vec::with_capacity(script.len() + 3);
    encoded.extend_from_slice(&[0xEF, 0xBB, 0xBF]);
    encoded.extend_from_slice(script.as_bytes());
    fs::write(path, encoded).with_context(|| format!("write PowerShell script {}", path.display()))
}

fn validate_conversion(
    input: &Path,
    output: &Path,
    source_ext: &str,
    target: &str,
    engine: &str,
    temp_dir: &Path,
    source_file_format: Option<i32>,
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
            },
            "com" if cfg!(windows) => (
                extract_word_text_com(input, &temp_dir.join("source.txt"))?,
                extract_word_text_com(output, &temp_dir.join("output.txt"))?,
            ),
            _ => bail!("unsupported text-validation engine {engine:?}"),
        };
        ensure!(
            normalize_document_text(&source_text) == normalize_document_text(&output_text),
            "converted DOCX text content differs from the source document"
        );
        return Ok("{\"level\":\"docx_package_and_text_content\"}".into());
    }
    let converted = read_comparable_workbook(output)
        .with_context(|| format!("validate converted workbook {}", output.display()))?;
    ensure!(
        !converted.sheets.is_empty(),
        "converted workbook has no worksheets"
    );
    if source_ext == "csv" {
        validate_csv_conversion(input, &converted)?;
        return Ok("{\"level\":\"one_worksheet_and_csv_values\",\"numeric_tolerance\":1e-15}".into());
    }
    let source_path = if source_ext == "et" {
        // ET is a CFB container, not necessarily an OOXML ZIP. Convert the
        // original source separately with the selected engine, then compare
        // that workbook's sheets, values and formulas with the candidate.
        // Never disguise the source by changing its extension and asking
        // Calamine to read it as XLSX.
        let probe_dir = temp_dir.join("source-probe");
        fs::create_dir_all(&probe_dir)?;
        match engine {
            "lo" => {
                run_libreoffice(input, &probe_dir, target)
                    .context("convert ET source for preservation validation with LibreOffice")?;
                locate_converted(&probe_dir, target)
                    .context("locate separately converted ET source for validation")?
            }
            "com" if cfg!(windows) => {
                let probe = probe_dir.join(format!("source.{target}"));
                let probe_file_format = run_com(input, &probe, target, source_ext, runtime)
                    .context("convert ET source for preservation validation with COM")?
                    .source_format;
                ensure!(
                    source_file_format == probe_file_format,
                    "ET source format verification differed between candidate and preservation probe"
                );
                ensure!(
                    probe.is_file() && fs::metadata(&probe)?.len() > 0,
                    "COM did not create a non-empty ET source validation workbook"
                );
                probe
            }
            _ => bail!("unsupported ET source-validation engine {engine:?}"),
        }
    } else {
        input.to_path_buf()
    };
    let source = read_comparable_workbook(&source_path).with_context(|| {
        format!(
            "read source workbook for preservation validation {}",
            input.display()
        )
    })?;
    compare_workbooks(&source, &converted)?;
    if source_ext == "et" {
        let (source_format_verified, source_format, source_file_format) =
            match (engine, source_file_format) {
                ("com", Some(65521)) => (true, "et", 65521),
                _ => (false, "unknown", 0),
            };
        Ok(format!("{{\"level\":\"worksheet_order_names_values_and_formula_text\",\"source_probe\":\"separate_conversion_using_selected_engine\",\"numeric_tolerance\":1e-15,\"source_format_verified\":{source_format_verified},\"source_format\":\"{source_format}\",\"source_file_format\":{source_file_format}}}"))
    } else {
        Ok(
            "{\"level\":\"worksheet_order_names_values_and_formula_text\",\"numeric_tolerance\":1e-15}"
                .into(),
        )
    }
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
        .with_context(|| "该功能需要外部引擎: LibreOffice is required to validate converted Word text")?;
    ensure!(
        command.status.success(),
        "LibreOffice text extraction failed: {}",
        String::from_utf8_lossy(&command.stderr)
    );
    let text_path = output_dir.join(format!("{}.txt", file_stem(input)));
    fs::read_to_string(&text_path)
        .with_context(|| format!("read extracted Word text {}", text_path.display()))
}

fn extract_word_text_com(input: &Path, text_output: &Path) -> Result<String> {
    if let Some(parent) = text_output.parent() {
        fs::create_dir_all(parent)?;
    }
    let script_path = text_output.with_extension("extract-word-text.ps1");
    let script = format!(
        "$ErrorActionPreference='Stop';\n$app=$null; $appName=''; $lastError='';\nforeach ($name in @('Word.Application','kwps.Application','KWPS.Application','wps.Application')) {{ if ($null -eq $app) {{ try {{ $app=New-Object -ComObject $name; $appName=$name }} catch {{ $lastError=$_.Exception.Message }} }} }}\nif ($null -eq $app) {{ throw \"No supported Office/WPS COM engine is available: $lastError\" }}\n$doc=$null;\ntry {{ $app.Visible=$false; $app.DisplayAlerts=0; $doc=$app.Documents.Open('{}',$false,$true); $text=$doc.Content.Text; [System.IO.File]::WriteAllText('{}',$text,[System.Text.UTF8Encoding]::new($false)) }} finally {{ if ($null -ne $doc) {{ $doc.Close($false) }}; if ($null -ne $app) {{ $app.Quit() }} }}\n",
        ps_quote(&input.to_string_lossy()),
        ps_quote(&text_output.to_string_lossy())
    );
    write_powershell_script(&script_path, &script)?;
    let result = Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
        ])
        .arg(&script_path)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .with_context(|| "start Word COM text validation")?;
    let _ = fs::remove_file(script_path);
    ensure!(
        result.status.success(),
        "Word COM text extraction failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    fs::read_to_string(text_output)
        .with_context(|| format!("read extracted Word text {}", text_output.display()))
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
            ensure!(
                formula.trim_start_matches('=').trim()
                    == output_formula.trim_start_matches('=').trim(),
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

fn libreoffice_user_installation_arg(profile: &Path) -> Result<String> {
    fs::create_dir_all(profile).context("create LibreOffice isolated profile directory")?;
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
        if byte.is_ascii_alphanumeric() || matches!(byte, b'/' | b':' | b'-' | b'.' | b'_' | b'~')
        {
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

fn ps_quote(value: &str) -> String {
    value.replace('\'', "''")
}

struct TempDirGuard(PathBuf);

impl Drop for TempDirGuard {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[cfg(test)]
mod tests {
    use super::{
        com_display_alerts_setting, file_uri, libreoffice_batch_request,
        libreoffice_user_installation_arg, select_fidelity_engine, ConversionRuntime,
        LibreOfficeExecutor, LibreOfficeRunOutcome,
    };
    use anyhow::Result;
    use std::fs;
    use std::path::{Path, PathBuf};
    use std::sync::atomic::{AtomicU32, Ordering};
    use std::sync::{Arc, Mutex};

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
        assert_eq!(
            select_fidelity_engine("auto"),
            if cfg!(windows) { "com" } else { "lo" }
        );
        assert_eq!(select_fidelity_engine("lo"), "lo");
        assert_eq!(select_fidelity_engine("com"), "com");
        assert!(libreoffice_batch_request("xls", "xlsx", "lo"));
        assert!(libreoffice_batch_request("xlsb", "xlsx", "lo"));
        assert!(libreoffice_batch_request("xlsm", "xlsx", "lo"));
        assert!(libreoffice_batch_request("doc", "docx", "lo"));
        // ET stays per-file: its LO result is rejected afterwards anyway.
        assert!(!libreoffice_batch_request("et", "xlsx", "lo"));
        // Same-category copies and other engines never batch.
        assert!(!libreoffice_batch_request("xlsx", "xlsx", "lo"));
        assert!(!libreoffice_batch_request("xls", "xlsx", "com"));
        assert!(!libreoffice_batch_request("xls", "xlsx", "auto") || !cfg!(windows));
    }

    #[test]
    fn com_display_alerts_setting_uses_host_application_type() {
        assert_eq!(com_display_alerts_setting("docx"), "$app.DisplayAlerts=0;");
        assert_eq!(
            com_display_alerts_setting("xlsx"),
            "$app.DisplayAlerts=$false;"
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
