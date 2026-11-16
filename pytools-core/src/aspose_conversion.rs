//! Explicit external conversion route; never selected automatically.
use anyhow::{ensure, Context, Result};
use excel_aspose::{AsposeWorkerBackend, WorkbookConverterEngine, WorkbookRequest, WorkerConfig};
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};

pub(crate) struct AsposeConversionSession {
    engine: Box<dyn WorkbookConverterEngine + Send>,
    run_id: String,
    diagnostics: PathBuf,
}
impl AsposeConversionSession {
    fn new() -> Result<Self> {
        crate::aspose_edit::ModifyMode::Aspose.preflight()?;
        let stamp = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let run_id = format!("aspose-convert-{}-{stamp}", std::process::id());
        let diagnostics = crate::run_logging::default_log_dir()
            .join("aspose-convert")
            .join(&run_id);
        let (executable, arguments) = crate::aspose_integration::worker_command()?;
        let engine = Box::new(AsposeWorkerBackend::new(WorkerConfig {
            executable,
            arguments,
            diagnostics_dir: diagnostics.clone(),
            startup_timeout: Duration::from_secs(30),
            request_timeout: Duration::from_secs(180),
        }));
        Ok(Self {
            engine,
            run_id,
            diagnostics,
        })
    }
    fn convert(&mut self, input: &Path, output: &Path) -> Result<String> {
        let started = Instant::now();
        let input = fs::canonicalize(input)?;
        ensure!(
            input
                .extension()
                .is_some_and(|e| e.eq_ignore_ascii_case("xlsx") || e.eq_ignore_ascii_case("xls")),
            "Aspose 转换仅启用 XLS/XLSX → XLSX；其他格式未验证"
        );
        ensure!(!output.exists(), "拒绝覆盖转换目标");
        let hash = crate::file_commit::sha256_file(&input)?;
        let source = excel_reader::read_workbook(&input)?;
        let parent = output
            .parent()
            .filter(|p| !p.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        fs::create_dir_all(parent)?;
        let parent = fs::canonicalize(parent)?;
        let stamp = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let scratch = parent.join(format!(".aspose-convert-{}-{stamp}", std::process::id()));
        fs::create_dir(&scratch)?;
        struct Cleanup(PathBuf);
        impl Drop for Cleanup {
            fn drop(&mut self) {
                let _ = fs::remove_file(self.0.join("candidate.xlsx"));
                let _ = fs::remove_dir(&self.0);
            }
        }
        let _cleanup = Cleanup(scratch.clone());
        let candidate = scratch.join("candidate.xlsx");
        let result = self
            .engine
            .convert(&WorkbookRequest {
                run_id: self.run_id.clone(),
                input: input.clone(),
                candidate: candidate.clone(),
            })
            .context("Aspose 转换失败，未回退其他引擎")?;
        let validation_started = Instant::now();
        let saved = excel_reader::read_workbook(&candidate)?;
        let array_results = if input
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
        {
            let original = excel_editor::ArrayFormulaResults::read(&input)?;
            ensure!(
                original == excel_editor::ArrayFormulaResults::read(&candidate)?,
                "Aspose 转换改变数组公式结果范围"
            );
            original
        } else {
            excel_editor::ArrayFormulaResults::default()
        };
        crate::aspose_integration::validate_recalculated_workbook_with_arrays(
            &source,
            &saved,
            &array_results,
        )?;
        if input
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
        {
            use excel_editor::{ExcelEditPlan, PreservationRequirement as P, WorkbookFormat};
            // Formula caches are expected to change during conversion. Literal
            // values/types were checked above; all other semantics retain the
            // existing plan-aware validator, including names and structure.
            let plan = ExcelEditPlan {
                request_id: self.run_id.clone(),
                input: input.clone(),
                output: candidate.clone(),
                format: WorkbookFormat::Xlsx,
                workbook: excel_editor::inspect_xlsx_profile(&input)?,
                operations: Vec::new(),
                preserve: [
                    P::FormulaText,
                    P::CommentsAndPositions,
                    P::MergedRanges,
                    P::DefinedNames,
                    P::PageSettings,
                    P::HeaderFooterText,
                    P::Relationships,
                    P::SheetIdentityAndOrder,
                    P::ConditionalFormatting,
                    P::DataValidations,
                    P::FreezePanes,
                    P::AutoFilter,
                ]
                .into_iter()
                .collect(),
                drop_allowed: Default::default(),
            };
            excel_editor::OpenpyxlXlsxValidator
                .validate_with_evaluation_sheet(&plan, &candidate)?;
        }
        crate::file_commit::verify_planning_snapshot(&input, Some(&hash))?;
        let validation_ms = validation_started.elapsed().as_secs_f64() * 1000.0;
        let commit = crate::file_commit::create_target_from_candidate(&candidate, output)?;
        let report = serde_json::json!({"status":"converted","engine":"aspose","run_id":self.run_id,"evaluation_output":true,"result":result,"validation":{"verified":true,"level":"source_values_types_formula_text_merges_comments_and_sheet_order","source_sha256":hash,"elapsed_ms":validation_ms},"commit":commit,"end_to_end_ms":started.elapsed().as_secs_f64()*1000.0});
        fs::create_dir_all(&self.diagnostics)?;
        let mut log = fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(self.diagnostics.join("host.jsonl"))?;
        writeln!(log, "{report}")?;
        Ok(serde_json::to_string(&report)?)
    }
}
pub(crate) fn convert(
    input: &Path,
    output: &Path,
    runtime: &mut crate::conversion_engine::ConversionRuntime,
) -> Result<String> {
    if runtime.aspose_converter.is_none() {
        runtime.aspose_converter = Some(AsposeConversionSession::new()?);
    }
    runtime
        .aspose_converter
        .as_mut()
        .context("missing conversion session")?
        .convert(input, output)
}
