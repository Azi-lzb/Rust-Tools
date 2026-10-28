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
    EngineId, ExcelEditPlan,
};
use std::path::{Path, PathBuf};

/// 用户可选的执行引擎。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum InPlaceEngine {
    Openpyxl,
    Umya,
}

impl InPlaceEngine {
    pub fn label(self) -> &'static str {
        match self {
            Self::Openpyxl => "openpyxl",
            Self::Umya => "umya-spreadsheet",
        }
    }

    fn engine_id(self) -> EngineId {
        match self {
            Self::Openpyxl => EngineId::OpenpyxlWorker,
            Self::Umya => EngineId::UmyaSpreadsheet,
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
            Self::Umya => matches!(extension.as_str(), "xlsx" | "xltx"),
            Self::Openpyxl => matches!(extension.as_str(), "xlsx" | "xlsm" | "xltx" | "xltm"),
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
        } else if self == Self::Umya && matches!(extension.as_str(), "xlsm" | "xltm") {
            format!(
                "{} 是含宏工作簿（.{extension}），umya 保存会丢失宏；请改选 openpyxl 引擎",
                path.display()
            )
        } else {
            format!("{}：当前引擎不支持就地修改该格式", path.display())
        }
    }
}

fn platform_id() -> String {
    if cfg!(windows) {
        "windows-x86_64".to_owned()
    } else if cfg!(target_os = "macos") {
        "macos-x86_64".to_owned()
    } else {
        "linux-x86_64".to_owned()
    }
}

/// 由 plan 推导运行时能力记录（引擎随产品编译并逐项回读校验）。
fn capability_for(plan: &ExcelEditPlan, engine: InPlaceEngine) -> BackendCapability {
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
            set.insert(EditOperation::SetCellValue {
                sheet: String::new(),
                cell: String::new(),
                value: excel_editor::CellValue::Text(String::new()),
            }.kind());
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
        capability,
        binary,
        "pytools",
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

/// 就地执行一个 plan：先出候选、校验，再替换目标原文件。
pub fn execute_in_place(plan: &ExcelEditPlan, engine: InPlaceEngine) -> Result<InPlaceOutcome> {
    anyhow::ensure!(
        engine.accepts(&plan.input),
        "{}",
        engine.unsupported_hint(&plan.input)
    );
    let capability = capability_for(plan, engine);
    let (execution, commit) = apply_plan_in_place(plan, &[capability], &platform_id())?;
    Ok(InPlaceOutcome {
        target: plan.input.clone(),
        execution,
        commit,
    })
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
