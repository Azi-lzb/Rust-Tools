//! Plan serialization and evidence-gated ExcelModifier execution.

use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, EditBackend, EditExecutionReport, EditOperation, EngineId, ExcelEditPlan,
    ExcelModifier, OoxmlHeaderTextBackend, OoxmlHeaderTextValidator, OpenpyxlXlsxBackend,
    UmyaXlsxBackend, UmyaXlsxValidator,
};
use serde::Deserialize;
use crate::file_commit::TargetCommitReport;
use std::{
    fs,
    path::{Path, PathBuf},
    sync::Arc,
};

#[derive(Deserialize)]
#[serde(untagged)]
enum CapabilityDocument {
    Many(Vec<BackendCapability>),
    One(BackendCapability),
}

/// Apply an edit plan only through backend capability records containing
/// matching preservation evidence. An empty or unverified capability list is
/// rejected by ExcelModifier before any candidate is written.
pub fn apply_plan_file(
    plan_path: impl AsRef<Path>,
    capability_path: impl AsRef<Path>,
    platform: &str,
) -> Result<EditExecutionReport> {
    let plan_path = plan_path.as_ref();
    let capability_path = capability_path.as_ref();
    let plan: ExcelEditPlan = serde_json::from_slice(
        &fs::read(plan_path).with_context(|| format!("read edit plan {}", plan_path.display()))?,
    )
    .with_context(|| format!("decode edit plan {}", plan_path.display()))?;
    let document: CapabilityDocument = serde_json::from_slice(
        &fs::read(capability_path)
            .with_context(|| format!("read backend capability {}", capability_path.display()))?,
    )
    .with_context(|| format!("decode backend capability {}", capability_path.display()))?;
    let capabilities = match document {
        CapabilityDocument::Many(values) => values,
        CapabilityDocument::One(value) => vec![value],
    };
    apply_plan(&plan, &capabilities, platform)
}

/// 就地执行：把 plan 的编辑结果写回目标工作簿本身。
///
/// 安全语义：先在临时候选文件上执行并验证（复用 ExcelModifier 的既有校验），
/// 再用带规划期 sha256 的替换提交覆盖目标；若目标在规划后被改动则拒绝写入。
/// 会话中断时目标保持原样，不会出现半写状态。
pub fn apply_plan_in_place(
    plan: &ExcelEditPlan,
    capabilities: &[BackendCapability],
    platform: &str,
) -> Result<(EditExecutionReport, TargetCommitReport)> {
    let target = plan.input.clone();
    let original_hash = crate::file_commit::sha256_file(&target)?;
    // 候选文件放在目标同目录，保证同卷重命名语义。
    let candidate = unique_in_place_candidate(&target)?;
    let mut candidate_plan = plan.clone();
    candidate_plan.output = candidate.clone();

    let result = apply_plan(&candidate_plan, capabilities, platform);
    let execution = match result {
        Ok(execution) => execution,
        Err(error) => {
            let _ = fs::remove_file(&candidate);
            return Err(error);
        }
    };
    let commit = match crate::file_commit::replace_target_from_candidate(
        &candidate,
        &target,
        &original_hash,
    ) {
        Ok(commit) => commit,
        Err(error) => {
            let _ = fs::remove_file(&candidate);
            return Err(error).with_context(|| {
                format!("就地提交修改结果到 {} 失败", target.display())
            });
        }
    };
    let _ = fs::remove_file(&candidate);
    Ok((execution, commit))
}

fn unique_in_place_candidate(target: &Path) -> Result<PathBuf> {
    let parent = target
        .parent()
        .context("in-place edit target has no parent directory")?;
    for _ in 0..256 {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)?
            .as_nanos();
        let candidate = parent.join(format!(
            ".pytools-inplace-{}-{nonce}.xlsx",
            std::process::id()
        ));
        if !candidate.exists() {
            return Ok(candidate);
        }
    }
    anyhow::bail!("cannot allocate an in-place candidate beside {}", target.display())
}

pub fn apply_plan_with_capability_file(
    plan: &ExcelEditPlan,
    capability_path: impl AsRef<Path>,
    platform: &str,
) -> Result<EditExecutionReport> {
    let capability_path = capability_path.as_ref();
    let document: CapabilityDocument = serde_json::from_slice(
        &fs::read(capability_path)
            .with_context(|| format!("read backend capability {}", capability_path.display()))?,
    )
    .with_context(|| format!("decode backend capability {}", capability_path.display()))?;
    let capabilities = match document {
        CapabilityDocument::Many(values) => values,
        CapabilityDocument::One(value) => vec![value],
    };
    apply_plan(plan, &capabilities, platform)
}

fn apply_plan(
    plan: &ExcelEditPlan,
    capabilities: &[BackendCapability],
    platform: &str,
) -> Result<EditExecutionReport> {
    let mut backends: Vec<Arc<dyn EditBackend>> = Vec::new();
    for capability in capabilities.iter().cloned() {
        match capability.engine {
            EngineId::UmyaSpreadsheet => backends.push(Arc::new(UmyaXlsxBackend::new(capability)?)),
            EngineId::OoxmlPatch => {
                backends.push(Arc::new(OoxmlHeaderTextBackend::new(capability)?))
            }
            EngineId::OpenpyxlWorker => {
                backends.push(Arc::new(OpenpyxlXlsxBackend::for_product_binary(
                    capability,
                    std::env::current_exe()
                        .context("resolve pytools executable for bundled openpyxl runtime")?,
                    "pytools",
                )?))
            }
            other => anyhow::bail!("no runtime adapter is registered for backend {other}"),
        }
    }
    let modifier = ExcelModifier::new(platform, backends, Arc::new(PlanOutputValidator))?;
    modifier.execute(plan).map_err(anyhow::Error::from)
}

struct PlanOutputValidator;

impl excel_editor::OutputValidator for PlanOutputValidator {
    fn validate(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        if plan
            .operations
            .iter()
            .all(|operation| matches!(operation, EditOperation::ReplaceHeaderText { .. }))
        {
            OoxmlHeaderTextValidator.validate(plan, candidate_output)
        } else {
            UmyaXlsxValidator.validate(plan, candidate_output)
        }
    }
}
