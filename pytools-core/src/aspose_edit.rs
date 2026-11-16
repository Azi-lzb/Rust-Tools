//! Explicit, scoped modification selection. Aspose is never an automatic route.
use anyhow::{bail, ensure, Context, Result};
use excel_aspose::{
    AsposeWorkerBackend, CellStyle, ModifyEngine, ModifyRequest, WorkbookOperation,
    WorkbookRequest, WorkerConfig,
};
use excel_editor::{
    BackendCapability, BackendFailure, CellValue, EditBackend, EditOperation, EngineId,
    ExcelEditPlan, OutputValidator, WorkbookFeature, WorkbookFormat,
};
use std::{
    cell::RefCell,
    fs,
    io::Write,
    marker::PhantomData,
    path::{Path, PathBuf},
    rc::Rc,
    time::{Duration, SystemTime, UNIX_EPOCH},
};

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum ModifyMode {
    #[default]
    Auto,
    Openpyxl,
    Aspose,
}
impl ModifyMode {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Openpyxl => "openpyxl",
            Self::Aspose => "aspose",
        }
    }
    pub fn parse(value: &str) -> Result<Self> {
        match value {
            "auto" => Ok(Self::Auto),
            "openpyxl" => Ok(Self::Openpyxl),
            "aspose" => Ok(Self::Aspose),
            _ => bail!("--modify-engine 仅支持 auto、openpyxl、aspose"),
        }
    }
    pub fn preflight(self) -> Result<()> {
        if self == Self::Aspose {
            ensure!(
                cfg!(any(windows, target_os = "linux")),
                "Aspose Worker 仅支持 Windows 和 Linux；未回退其他引擎"
            );
            ensure!(
                std::env::var_os("PYTOOLS_CONVERT_ENGINES_DISABLED").is_none(),
                "外部引擎已禁用；未启动 Aspose"
            );
            crate::aspose_integration::worker_command()?;
        }
        Ok(())
    }
}
#[derive(Default)]
struct ContextState {
    mode: ModifyMode,
    run_id: String,
    engine: Option<AsposeWorkerBackend>,
}
thread_local! { static STATE: RefCell<ContextState> = RefCell::new(ContextState::default()); }
/// CLI jobs run on one thread. No process environment is changed, including on cancellation.
pub struct ModifyExecutionScope {
    previous: Option<ContextState>,
    _thread: PhantomData<Rc<()>>,
}
impl ModifyExecutionScope {
    pub fn enter(mode: ModifyMode) -> Result<Self> {
        mode.preflight()?;
        let previous = STATE.with(|s| {
            let run_id = s.borrow().run_id.clone();
            std::mem::replace(
                &mut *s.borrow_mut(),
                ContextState {
                    mode,
                    run_id,
                    ..Default::default()
                },
            )
        });
        Ok(Self {
            previous: Some(previous),
            _thread: PhantomData,
        })
    }
}
impl Drop for ModifyExecutionScope {
    fn drop(&mut self) {
        if let Some(previous) = self.previous.take() {
            STATE.with(|s| *s.borrow_mut() = previous);
        }
    }
}
pub fn bind_run_id(run_id: &str) {
    STATE.with(|s| s.borrow_mut().run_id = run_id.to_owned());
}
pub fn route_override() -> Option<crate::excel_edit_router::RouteConfig> {
    use crate::excel_edit_router::{RouteConfig, RouteMode};
    STATE.with(|s| match s.borrow().mode {
        ModifyMode::Auto => None,
        ModifyMode::Openpyxl => Some(RouteConfig::new(RouteMode::OpenpyxlOnly)),
        ModifyMode::Aspose => Some(RouteConfig::new(RouteMode::AsposeOnly)),
    })
}
pub fn requested_mode() -> ModifyMode {
    STATE.with(|s| s.borrow().mode)
}
/// Explicit modifiers in these menus have an XLSX-only contract. Auto retains
/// legacy per-feature COM handling; selecting Openpyxl never implies COM.
pub fn xlsx_only_selection() -> bool {
    aspose_selected()
        || route_override()
            .is_some_and(|c| c.mode == crate::excel_edit_router::RouteMode::OpenpyxlOnly)
}
pub fn aspose_selected() -> bool {
    route_override()
        .or_else(|| crate::excel_edit_router::RouteConfig::from_environment().ok())
        .is_some_and(|c| c.mode == crate::excel_edit_router::RouteMode::AsposeOnly)
}

fn operations(plan: &ExcelEditPlan) -> Result<Vec<WorkbookOperation>> {
    plan.operations
        .iter()
        .map(|op| {
            if let EditOperation::SetCellValue { value, .. } = op {
                match value {
                    CellValue::Decimal(raw) => {
                        let n = raw.parse::<f64>().context("Aspose 数值文本无效")?;
                        let digits: String = raw
                            .split(['e', 'E'])
                            .next()
                            .unwrap_or_default()
                            .chars()
                            .filter(char::is_ascii_digit)
                            .collect();
                        ensure!(
                            n.is_finite() && (digits.trim_matches('0').len() <= 15
                                || (n.fract() != 0.0 && n.to_string() == raw.trim())),
                            "Aspose 拒绝无法往返保留的数字或超过 15 位有效数字的整数；长标识请使用文本"
                        );
                        return Ok(WorkbookOperation::EditPlanOperation {
                            operation: serde_json::to_value(op)?,
                        });
                    }
                    CellValue::DateTime { serial, .. } => {
                        ensure!(
                            serial.parse::<f64>().is_ok_and(f64::is_finite),
                            "Aspose 日期序列无效"
                        );
                        return Ok(WorkbookOperation::EditPlanOperation {
                            operation: serde_json::to_value(op)?,
                        });
                    }
                    CellValue::Error(error) => {
                        ensure!(
                            ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"]
                                .contains(&error.as_str()),
                            "Aspose 未验证此 Excel 错误类型"
                        );
                        return Ok(WorkbookOperation::EditPlanOperation {
                            operation: serde_json::to_value(op)?,
                        });
                    }
                    _ => {}
                }
            }
            Ok(match op {
                EditOperation::SetCellValue { sheet, cell, value } => WorkbookOperation::SetValue {
                    sheet: sheet.clone(),
                    cell: cell.clone(),
                    value: match value {
                        CellValue::Text(v) => excel_aspose::CellValue::Text(v.clone()),
                        CellValue::Boolean(v) => excel_aspose::CellValue::Boolean(*v),
                        CellValue::Blank => excel_aspose::CellValue::Blank,
                        // Typed numeric/date/error writes are validated and
                        // returned through the plan primitive above.
                        _ => unreachable!("typed value handled above"),
                    },
                },
                EditOperation::SetFormulaText {
                    sheet,
                    cell,
                    formula,
                } => WorkbookOperation::SetFormula {
                    sheet: sheet.clone(),
                    cell: cell.clone(),
                    formula: if formula.starts_with('=') {
                        formula.clone()
                    } else {
                        format!("={formula}")
                    },
                },
                EditOperation::SetCellFillColor { sheet, cell, argb } => {
                    ensure!(
                        argb.len() == 6
                            || (argb.len() == 8 && argb[..2].eq_ignore_ascii_case("FF")),
                        "Aspose 暂只支持 RGB 或不透明 FF AARRGGBB 填色"
                    );
                    WorkbookOperation::SetStyle {
                        sheet: sheet.clone(),
                        cell: cell.clone(),
                        style: CellStyle {
                            fill_rgb: Some(if argb.len() == 6 {
                                argb.clone()
                            } else {
                                argb[2..].to_owned()
                            }),
                            ..Default::default()
                        },
                    }
                }
                EditOperation::ReplaceHeaderText {
                    sheet,
                    search,
                    replacement,
                } => WorkbookOperation::ReplaceHeaderText {
                    sheet: sheet.clone(),
                    search: search.clone(),
                    replacement: replacement.clone(),
                },
                _ => WorkbookOperation::EditPlanOperation {
                    operation: serde_json::to_value(op)?,
                },
            })
        })
        .collect()
}
pub(crate) fn ensure_supported(plan: &ExcelEditPlan) -> Result<()> {
    ensure!(
        cfg!(any(windows, target_os = "linux")),
        "Aspose Worker 仅支持 Windows 和 Linux；未回退其他引擎"
    );
    plan.validate()?;
    ensure!(
        plan.format == WorkbookFormat::Xlsx
            && plan
                .input
                .extension()
                .is_some_and(|e| e.eq_ignore_ascii_case("xlsx")),
        "Aspose 修改接入仅支持 XLSX；XLS/XLSM 未启用，未回退 COM"
    );
    ensure!(
        plan.workbook.inventory_complete,
        "Aspose 拒绝未完成结构清点的工作簿"
    );
    for f in &plan.workbook.features_present {
        ensure!(
            !matches!(
                f,
                WorkbookFeature::Macros
                    | WorkbookFeature::ExternalLinks
                    | WorkbookFeature::PivotTables
                    | WorkbookFeature::Slicers
                    | WorkbookFeature::ActiveX
                    | WorkbookFeature::DigitalSignatures
                    | WorkbookFeature::UnknownExtensions
                    | WorkbookFeature::CellMetadata
                    | WorkbookFeature::SheetProtection
            ),
            "Aspose 尚未验证结构 {f:?}，拒绝修改"
        );
    }
    operations(plan)?;
    Ok(())
}
pub(crate) struct AsposeEditBackend {
    capability: BackendCapability,
}
impl AsposeEditBackend {
    pub(crate) fn new(capability: BackendCapability) -> Result<Self> {
        ensure!(
            capability.engine == EngineId::AsposeWorker,
            "incorrect Aspose capability"
        );
        Ok(Self { capability })
    }
    fn apply_inner(&self, plan: &ExcelEditPlan) -> Result<()> {
        let parent = fs::canonicalize(plan.output.parent().context("candidate has no parent")?)?;
        let stamp = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let scratch = parent.join(format!(".aspose-worker-{}-{stamp}", std::process::id()));
        fs::create_dir(&scratch).context("reserve private Aspose candidate directory")?;
        let candidate = scratch.join("candidate.xlsx");
        // Only our unique worker path is cleaned, including failed Save/verification.
        struct Cleanup(PathBuf);
        impl Drop for Cleanup {
            fn drop(&mut self) {
                let _ = fs::remove_file(self.0.join("candidate.xlsx"));
                let _ = fs::remove_dir(&self.0);
            }
        }
        let _cleanup = Cleanup(scratch);
        let ops = operations(plan)?;
        STATE.with(|state| -> Result<()> {
            let mut state = state.borrow_mut();
            if state.run_id.is_empty() { state.run_id = plan.request_id.clone(); }
            let directory_key: String = state.run_id.chars().map(|c| if c.is_ascii_alphanumeric() || matches!(c, '-' | '_') { c } else { '_' }).collect();
            let diagnostics = crate::run_logging::default_log_dir().join("aspose-modify").join(directory_key);
            if state.engine.is_none() {
                ModifyMode::Aspose.preflight()?;
                let (executable, arguments) = crate::aspose_integration::worker_command()?;
                state.engine = Some(AsposeWorkerBackend::new(WorkerConfig { executable, arguments, diagnostics_dir: diagnostics.clone(), startup_timeout: Duration::from_secs(30), request_timeout: Duration::from_secs(180) }));
            }
            let request = ModifyRequest { workbook: WorkbookRequest { run_id: state.run_id.clone(), input: plan.input.clone(), candidate }, operations: ops };
            let outcome = state.engine.as_mut().unwrap().modify(&request);
            let result = outcome.as_ref().map_err(|e| anyhow::anyhow!("Aspose 修改失败：{e}"));
            fs::create_dir_all(&diagnostics)?;
            let mut log = fs::OpenOptions::new().create(true).append(true).open(diagnostics.join("host.jsonl"))?;
            let event = serde_json::json!({"run_id":state.run_id,"request_id":plan.request_id,"timestamp":chrono::Utc::now().to_rfc3339(),"stage":"modify_candidate","engine":"aspose_worker","input":plan.input,"output":plan.output,"result":outcome.as_ref().ok(),"error":outcome.as_ref().err().map(ToString::to_string)});
            writeln!(log, "{event}")?;
            let result = result?;
            // ExcelModifier owns/reserves plan.output; never ask Worker to overwrite it.
            let mut output = fs::OpenOptions::new().write(true).truncate(true).open(&plan.output)?;
            std::io::copy(&mut fs::File::open(&result.path)?, &mut output)?;
            output.sync_all()?;
            Ok(())
        })
    }
}
impl EditBackend for AsposeEditBackend {
    fn capability(&self) -> &BackendCapability {
        &self.capability
    }
    fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure> {
        ensure_supported(plan)
            .map_err(|e| BackendFailure::unsupported_before_write(format!("{e:#}")))?;
        self.apply_inner(plan)
            .map_err(|e| BackendFailure::failed(format!("{e:#}")))
    }
    fn validate_output(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        AsposePlanValidator.validate(plan, candidate_output)
    }
}
pub(crate) struct AsposePlanValidator;
impl OutputValidator for AsposePlanValidator {
    fn validate(&self, plan: &ExcelEditPlan, output: &Path) -> Result<()> {
        excel_editor::OpenpyxlXlsxValidator.validate_with_evaluation_sheet(plan, output)
    }
    fn validate_with_notes(&self, plan: &ExcelEditPlan, output: &Path) -> Result<Vec<String>> {
        excel_editor::OpenpyxlXlsxValidator.validate_with_evaluation_sheet_and_notes(plan, output)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn scope_restores_selection_without_environment_mutation() -> Result<()> {
        let prior = std::env::var_os(crate::excel_edit_router::ROUTE_ENV);
        assert!(route_override().is_none());
        {
            let _scope = ModifyExecutionScope::enter(ModifyMode::Openpyxl)?;
            assert_eq!(
                route_override().unwrap().mode,
                crate::excel_edit_router::RouteMode::OpenpyxlOnly
            );
            {
                let _auto = ModifyExecutionScope::enter(ModifyMode::Auto)?;
                assert!(route_override().is_none());
            }
            assert!(route_override().is_some());
        }
        assert!(route_override().is_none());
        assert_eq!(prior, std::env::var_os(crate::excel_edit_router::ROUTE_ENV));
        Ok(())
    }
}
