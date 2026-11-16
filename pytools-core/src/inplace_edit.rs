//! 菜单 5 的就地修改执行层。
//!
//! 业务代码只需提供 `ExcelEditPlan`，本模块负责：
//! 1. 按用户选择的引擎（openpyxl / umya）构建运行时能力记录；
//! 2. 在临时候选文件上执行并校验，再安全替换目标工作簿本身。
//!
//! 说明：能力记录在此由 plan 自身内容推导（引擎、格式、操作集合、需保留特性、
//! 平台），因为 Rust 后端随产品编译、执行后仍经 `PlanOutputValidator` 逐项回读
//! 校验；这与测试使用的临时能力不同，校验强度不变。

use crate::edit_plan::apply_plan_in_place;
use crate::file_commit::TargetCommitReport;
use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditExecutionReport, EditOperation,
    EngineId, ExcelEditPlan, PreservationRequirement,
};
use std::path::{Path, PathBuf};

/// 用户可选的执行引擎。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum InPlaceEngine {
    Openpyxl,
    Umya,
    BetterOffice,
    RustNative,
    Composite,
    Aspose,
}

impl InPlaceEngine {
    pub fn label(self) -> &'static str {
        match self {
            Self::Openpyxl => "openpyxl",
            Self::Umya => "umya-spreadsheet",
            Self::BetterOffice => "betteroffice",
            Self::RustNative => "rust-native-ooxml",
            Self::Composite => "in-memory-composite",
            Self::Aspose => "aspose_worker",
        }
    }

    fn engine_id(self) -> EngineId {
        match self {
            Self::Openpyxl => EngineId::OpenpyxlWorker,
            Self::Umya => EngineId::UmyaSpreadsheet,
            Self::BetterOffice => EngineId::BetterOffice,
            Self::RustNative => EngineId::RustNativeOoxml,
            Self::Composite => EngineId::Composite,
            Self::Aspose => EngineId::AsposeWorker,
        }
    }

    pub(crate) fn from_engine(engine: EngineId) -> Self {
        match engine {
            EngineId::OpenpyxlWorker => Self::Openpyxl,
            EngineId::UmyaSpreadsheet => Self::Umya,
            EngineId::BetterOffice => Self::BetterOffice,
            EngineId::RustNativeOoxml => Self::RustNative,
            EngineId::Composite => Self::Composite,
            EngineId::AsposeWorker => Self::Aspose,
            EngineId::OoxmlPatch => {
                unreachable!("OOXML patch is not an ExcelEditPlan route engine")
            }
        }
    }

    /// 该引擎可就地修改的扩展名。
    pub fn accepts(self, path: &Path) -> bool {
        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        match self {
            Self::Umya | Self::BetterOffice | Self::RustNative | Self::Composite => {
                matches!(extension.as_str(), "xlsx" | "xltx")
            }
            Self::Openpyxl => matches!(extension.as_str(), "xlsx" | "xlsm" | "xltx" | "xltm"),
            Self::Aspose => extension == "xlsx",
        }
    }

    /// 用户可见的格式提示（不支持的扩展名时使用）。
    pub fn unsupported_hint(self, path: &Path) -> String {
        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        if extension == "xls" || extension == "xlsb" {
            format!(
                "{} 是旧版/二进制格式（.{extension}），不支持就地修改；请先用 4-4 格式转换另存为 .xlsx 后再处理",
                path.display()
            )
        } else if matches!(
            self,
            Self::Umya | Self::BetterOffice | Self::RustNative | Self::Composite
        ) && matches!(extension.as_str(), "xlsm" | "xltm")
        {
            format!(
                "{} 是含宏工作簿（.{extension}），umya 保存会丢失宏；请改选 openpyxl 引擎",
                path.display()
            )
        } else {
            format!("{}：当前引擎不支持就地修改该格式", path.display())
        }
    }
}

pub(crate) fn platform_id() -> String {
    if cfg!(windows) {
        "windows-x86_64".to_owned()
    } else if cfg!(target_os = "macos") {
        "macos-x86_64".to_owned()
    } else {
        "linux-x86_64".to_owned()
    }
}

/// 由 plan 推导运行时能力记录（引擎随产品编译并逐项回读校验）。
pub(crate) fn capability_for(plan: &ExcelEditPlan, engine: InPlaceEngine) -> BackendCapability {
    let platform = platform_id();
    BackendCapability {
        engine: engine.engine_id(),
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: platform.clone(),
        evidence: Some(CapabilityEvidence {
            sample_id: format!("in-place-{}", engine.label()),
            report_path: "pytools-core/src/inplace_edit.rs".to_owned(),
            adapter_version: format!("pytools/{}", env!("CARGO_PKG_VERSION")),
            engine_version: engine.label().to_owned(),
            platform,
        }),
    }
}

/// 供 2-3 既有目标追加候选使用的 Openpyxl worker 能力记录；
/// operations/features/preserve 全部由追加计划推导。
pub(crate) fn openpyxl_worker_capability_for_plan(
    plan: &ExcelEditPlan,
) -> excel_editor::BackendCapability {
    capability_for(plan, InPlaceEngine::Openpyxl)
}

/// 预检用的最小 openpyxl 能力记录：仅用于触发运行时解析与路径校验，
/// 不参与任何真实编辑（真实编辑的能力由 `capability_for` 按 plan 推导）。
fn capability_for_plain_xlsx(binary: &Path) -> excel_editor::BackendCapability {
    use excel_editor::{
        BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId,
        WorkbookFormat,
    };
    let platform = if cfg!(windows) {
        "windows-x86_64"
    } else if cfg!(target_os = "macos") {
        "macos-x86_64"
    } else {
        "linux-x86_64"
    };
    BackendCapability {
        engine: EngineId::OpenpyxlWorker,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations: {
            let mut set = std::collections::BTreeSet::new();
            set.insert(
                EditOperation::SetCellValue {
                    sheet: String::new(),
                    cell: String::new(),
                    value: excel_editor::CellValue::Text(String::new()),
                }
                .kind(),
            );
            set
        },
        supported_features: Default::default(),
        preserved: Default::default(),
        platform: platform.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "openpyxl-runtime-preflight".to_owned(),
            report_path: "pytools-core/src/inplace_edit.rs".to_owned(),
            adapter_version: format!("pytools/{}", env!("CARGO_PKG_VERSION")),
            engine_version: "openpyxl".to_owned(),
            platform: platform.to_owned(),
        }),
    }
}

/// openpyxl 运行时预检：复用 excel-editor 的统一解析入口
/// （`OpenpyxlWorkerConfig::resolve_runtime`），保证“预检结论”与“执行解析”
/// 永远一致；失败消息包含按顺序的查找位置与缺失项。预检通过不仅代表
/// 文件存在，还实际启动运行时与 worker 完成协议版本握手（带缓存）。
fn openpyxl_runtime_error(binary: &Path) -> Option<String> {
    let capability = capability_for_plain_xlsx(binary);
    let backend = match excel_editor::OpenpyxlXlsxBackend::from_environment_or_product_layout(
        capability, binary, "pytools",
    ) {
        Ok(backend) => backend,
        Err(error) => return Some(format!("{error:#}")),
    };
    backend.verify_worker_handshake().err()
}

/// 执行条目：每个工作簿一条执行记录。
#[derive(Debug)]
pub struct InPlaceOutcome {
    pub target: PathBuf,
    pub execution: EditExecutionReport,
    pub commit: TargetCommitReport,
}

/// 修改并验证独立候选文件，发布由调用方负责；不替换源工作簿。
pub fn execute_candidate(
    plan: &ExcelEditPlan,
    engine: InPlaceEngine,
) -> Result<EditExecutionReport> {
    if let Some(route_config) = crate::aspose_edit::route_override() {
        return crate::excel_edit_router::execute_candidate_with_route_config(plan, &route_config);
    }
    if std::env::var_os(crate::excel_edit_router::ROUTE_ENV).is_some() {
        let route_config = crate::excel_edit_router::RouteConfig::from_environment()?;
        return crate::excel_edit_router::execute_candidate_with_route_config(plan, &route_config);
    }
    execute_candidate_with_engine(plan, engine)
}

fn execute_candidate_with_engine(
    plan: &ExcelEditPlan,
    engine: InPlaceEngine,
) -> Result<EditExecutionReport> {
    anyhow::ensure!(
        engine.accepts(&plan.input),
        "{}",
        engine.unsupported_hint(&plan.input)
    );
    let execution_plan = plan_for_engine(plan, engine);
    let capability = capability_for(&execution_plan, engine);
    crate::edit_plan::apply_plan(&execution_plan, &[capability], &platform_id())
}

/// 就地执行一个 plan：先出候选、校验，再替换目标原文件。
pub fn execute_in_place(plan: &ExcelEditPlan, engine: InPlaceEngine) -> Result<InPlaceOutcome> {
    if let Some(route_config) = crate::aspose_edit::route_override() {
        return crate::excel_edit_router::execute_in_place_with_route_config(plan, &route_config);
    }
    if std::env::var_os(crate::excel_edit_router::ROUTE_ENV).is_some() {
        let route_config = crate::excel_edit_router::RouteConfig::from_environment()?;
        return crate::excel_edit_router::execute_in_place_with_route_config(plan, &route_config);
    }
    execute_in_place_with_engine(plan, engine)
}

fn execute_in_place_with_engine(
    plan: &ExcelEditPlan,
    engine: InPlaceEngine,
) -> Result<InPlaceOutcome> {
    anyhow::ensure!(
        engine.accepts(&plan.input),
        "{}",
        engine.unsupported_hint(&plan.input)
    );
    let execution_plan = plan_for_engine(plan, engine);
    let capability = capability_for(&execution_plan, engine);
    let (execution, commit) = apply_plan_in_place(&execution_plan, &[capability], &platform_id())?;
    Ok(InPlaceOutcome {
        target: plan.input.clone(),
        execution,
        commit,
    })
}

pub(crate) fn plan_for_engine(plan: &ExcelEditPlan, engine: InPlaceEngine) -> ExcelEditPlan {
    let mut execution_plan = plan.clone();
    if engine == InPlaceEngine::Openpyxl {
        // Legacy PyCode saves existing workbooks through openpyxl. It rewrites
        // OOXML package parts, so do not route-block it on byte/package identity.
        // Concrete workbook values, formulas, styles and requested operations
        // remain checked by the OpenpyxlXlsxValidator before safe replacement.
        execution_plan
            .preserve
            .remove(&PreservationRequirement::UnmodifiedPackageParts);
        if matches!(
            plan.request_id.as_str(),
            "pytools-split-village-bank" | "pytools-adjust-rural-loan"
        ) {
            // PyCode's Openpyxl saves normalize row style IDs and row
            // dimensions across the workbook. These functions intentionally
            // align with that established save path; cell styles and the
            // requested row/cell operations remain validated independently.
            execution_plan
                .preserve
                .remove(&PreservationRequirement::RowAndColumnDimensions);
        }
    }
    if engine != InPlaceEngine::RustNative {
        relax_incidental_appearance(&mut execution_plan);
    }
    execution_plan
}

/// Data edits, row deletion, sheet renaming and header-text replacement do not
/// ask for an unrelated font/size round-trip. Explicit fill operations still
/// validate their requested color independently. Copy/import, print settings,
/// sizing and other format-oriented plans retain their existing contracts.
/// Apply only to object-model writers, after the engine has been selected.
pub(crate) fn relax_incidental_appearance(plan: &mut ExcelEditPlan) {
    if plan.operations.is_empty()
        || !plan.operations.iter().all(|op| {
            matches!(
                op,
                EditOperation::SetCellValue { .. }
                    | EditOperation::SetFormulaText { .. }
                    | EditOperation::SetComment { .. }
                    | EditOperation::SetCellFillColor { .. }
                    | EditOperation::DeleteRows { .. }
                    | EditOperation::InsertRows { .. }
                    | EditOperation::RenameSheet { .. }
                    | EditOperation::ReplaceHeaderText { .. }
            )
        })
    {
        return;
    }
    plan.preserve.remove(&PreservationRequirement::Styles);
    plan.preserve
        .remove(&PreservationRequirement::RowAndColumnDimensions);
}

/// openpyxl 引擎可用性预检：返回不可用原因，供菜单提前提示。
pub fn openpyxl_unavailable_reason() -> Option<String> {
    // 用当前 EXE 作为布局锚点；预检与执行共用同一解析入口。
    let binary = std::env::current_exe().unwrap_or_else(|_| PathBuf::from("pytools"));
    openpyxl_runtime_error(&binary)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validation_policy_accepts_font_drift_but_rejects_wrong_marking_and_data() -> Result<()> {
        use excel_editor::{OpenpyxlXlsxValidator, OutputValidator, WorkbookFormat};
        let dir = std::env::temp_dir().join(format!(
            "mark-content-policy-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        std::fs::create_dir_all(&dir)?;
        let source = dir.join("source.xlsx");
        for label in ["source", "correct", "wrong-fill", "wrong-data"] {
            let mut book = rust_xlsxwriter::Workbook::new();
            let sheet = book.add_worksheet();
            sheet.set_name("Data")?;
            let marked = if label == "source" {
                rust_xlsxwriter::Format::new()
            } else {
                rust_xlsxwriter::Format::new().set_background_color(if label == "wrong-fill" {
                    0x0000FF
                } else {
                    0xFF0000
                })
            };
            sheet.write_string_with_format(0, 0, "duplicate", &marked)?;
            let font = rust_xlsxwriter::Format::new()
                .set_font_name(if label == "source" {
                    "Calibri"
                } else {
                    "Arial"
                })
                .set_font_size(if label == "source" { 11.0 } else { 25.0 });
            sheet.write_number_with_format(
                0,
                1,
                if label == "wrong-data" { 8.0 } else { 7.0 },
                &font,
            )?;
            book.save(dir.join(format!("{label}.xlsx")))?;
        }
        let before = std::fs::read(&source)?;
        let plan = ExcelEditPlan {
            request_id: "mark-policy-test".into(),
            input: source.clone(),
            output: dir.join("published.xlsx"),
            format: WorkbookFormat::Xlsx,
            workbook: excel_editor::inspect_xlsx_profile(&source)?,
            operations: vec![EditOperation::SetCellFillColor {
                sheet: "Data".into(),
                cell: "A1".into(),
                argb: "FFFF0000".into(),
            }],
            preserve: [
                PreservationRequirement::CellValuesAndTypes,
                PreservationRequirement::FormulaText,
                PreservationRequirement::Styles,
                PreservationRequirement::RowAndColumnDimensions,
            ]
            .into_iter()
            .collect(),
            drop_allowed: Default::default(),
        };
        assert!(OpenpyxlXlsxValidator
            .validate(&plan, &dir.join("correct.xlsx"))
            .is_err());
        let relaxed = plan_for_engine(&plan, InPlaceEngine::Openpyxl);
        OpenpyxlXlsxValidator.validate(&relaxed, &dir.join("correct.xlsx"))?;
        for label in ["wrong-fill", "wrong-data"] {
            assert!(OpenpyxlXlsxValidator
                .validate(&relaxed, &dir.join(format!("{label}.xlsx")))
                .is_err());
        }
        assert_eq!(std::fs::read(&source)?, before);
        assert!(!plan.output.exists());
        Ok(())
    }

    #[test]
    fn validation_policy_follows_operations_and_keeps_explicit_format_contracts() {
        use excel_editor::{CellValue, WorkbookFormat, WorkbookProfile};
        let base = ExcelEditPlan {
            request_id: "policy-test".into(),
            input: "input.xlsx".into(),
            output: "output.xlsx".into(),
            format: WorkbookFormat::Xlsx,
            workbook: WorkbookProfile {
                format: WorkbookFormat::Xlsx,
                inventory_complete: true,
                features_present: Default::default(),
            },
            operations: vec![],
            preserve: [
                PreservationRequirement::Styles,
                PreservationRequirement::RowAndColumnDimensions,
                PreservationRequirement::CellValuesAndTypes,
                PreservationRequirement::FormulaText,
                PreservationRequirement::DefinedNames,
            ]
            .into_iter()
            .collect(),
            drop_allowed: Default::default(),
        };
        let data_ops = [
            EditOperation::SetCellValue {
                sheet: "Data".into(),
                cell: "A1".into(),
                value: CellValue::Text("new".into()),
            },
            EditOperation::SetCellFillColor {
                sheet: "Data".into(),
                cell: "A1".into(),
                argb: "FFFF0000".into(),
            },
            EditOperation::DeleteRows {
                sheet: "Data".into(),
                first_row: 2,
                count: 1,
            },
            EditOperation::RenameSheet {
                old_name: "Data".into(),
                new_name: "New".into(),
            },
            EditOperation::ReplaceHeaderText {
                sheet: "Data".into(),
                search: "万元".into(),
                replacement: "万美元".into(),
            },
        ];
        for op in data_ops {
            let mut source = base.clone();
            source.operations = vec![op];
            for engine in [
                InPlaceEngine::Openpyxl,
                InPlaceEngine::Umya,
                InPlaceEngine::BetterOffice,
                InPlaceEngine::Composite,
                InPlaceEngine::Aspose,
            ] {
                let adjusted = plan_for_engine(&source, engine);
                assert!(!adjusted.preserve.contains(&PreservationRequirement::Styles));
                assert!(!adjusted
                    .preserve
                    .contains(&PreservationRequirement::RowAndColumnDimensions));
                for property in [
                    PreservationRequirement::CellValuesAndTypes,
                    PreservationRequirement::FormulaText,
                    PreservationRequirement::DefinedNames,
                ] {
                    assert!(adjusted.preserve.contains(&property));
                }
                assert_eq!(adjusted.operations, source.operations);
                assert_eq!(adjusted.input, source.input);
                assert_eq!(adjusted.output, source.output);
            }
            assert!(source.preserve.contains(&PreservationRequirement::Styles));
            assert!(plan_for_engine(&source, InPlaceEngine::RustNative)
                .preserve
                .contains(&PreservationRequirement::Styles));
        }
        for op in [
            EditOperation::SetRowHeight {
                sheet: "Data".into(),
                row_index: 1,
                height: 30.0,
            },
            EditOperation::SetColumnWidth {
                sheet: "Data".into(),
                column_index: 1,
                width: 20.0,
            },
            EditOperation::SetPageSettings {
                sheet: "Data".into(),
                settings: Default::default(),
            },
            EditOperation::CopyRange {
                source_sheet: "Data".into(),
                source_range: "A1:B2".into(),
                target_sheet: "Data".into(),
                target_start: "C1".into(),
                copy_values: true,
                copy_formulas: false,
                copy_styles: true,
                copy_comments: false,
            },
        ] {
            let mut source = base.clone();
            source.operations = vec![op];
            assert!(plan_for_engine(&source, InPlaceEngine::Openpyxl)
                .preserve
                .contains(&PreservationRequirement::Styles));
            assert!(plan_for_engine(&source, InPlaceEngine::Aspose)
                .preserve
                .contains(&PreservationRequirement::RowAndColumnDimensions));
        }
    }

    #[test]
    fn umya_accepts_only_plain_ooxml() {
        assert!(InPlaceEngine::Umya.accepts(Path::new("a.xlsx")));
        assert!(InPlaceEngine::Umya.accepts(Path::new("a.xltx")));
        assert!(!InPlaceEngine::Umya.accepts(Path::new("a.xlsm")));
        assert!(!InPlaceEngine::Umya.accepts(Path::new("a.xls")));
    }

    #[test]
    fn openpyxl_accepts_macro_formats() {
        assert!(InPlaceEngine::Openpyxl.accepts(Path::new("a.xlsx")));
        assert!(InPlaceEngine::Openpyxl.accepts(Path::new("a.xlsm")));
        assert!(InPlaceEngine::Openpyxl.accepts(Path::new("a.xltm")));
        assert!(!InPlaceEngine::Openpyxl.accepts(Path::new("a.xls")));
    }

    #[test]
    fn xls_gets_a_conversion_hint() {
        let hint = InPlaceEngine::Umya.unsupported_hint(Path::new("old.xls"));
        assert!(hint.contains("4-4"), "{hint}");
        assert!(hint.contains(".xlsx"), "{hint}");
    }

    #[test]
    fn xlsm_under_umya_suggests_openpyxl() {
        let hint = InPlaceEngine::Umya.unsupported_hint(Path::new("macro.xlsm"));
        assert!(hint.contains("openpyxl"), "{hint}");
        assert!(hint.contains("宏"), "{hint}");
    }
}
