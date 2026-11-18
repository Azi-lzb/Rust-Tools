//! Plan serialization and evidence-gated ExcelModifier execution.

use crate::file_commit::TargetCommitReport;
use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, EditBackend, EditExecutionReport, EditOperation, EngineId, ExcelEditPlan,
    ExcelModifier, OoxmlHeaderTextBackend, OoxmlHeaderTextValidator, OpenpyxlWorkerConfig,
    OpenpyxlXlsxBackend, UmyaXlsxBackend, UmyaXlsxValidator,
};
use serde::Deserialize;
use std::{
    fs,
    path::{Path, PathBuf},
    sync::Arc,
};

use crate::betteroffice_edit::BetterOfficeAdapter;

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
    let plan: ExcelEditPlan = serde_json::from_slice(
        &fs::read(plan_path).with_context(|| format!("read edit plan {}", plan_path.display()))?,
    )
    .with_context(|| format!("decode edit plan {}", plan_path.display()))?;
    apply_plan_with_capability_file(&plan, capability_path, platform)
}

/// 就地执行：把 plan 的编辑结果写回目标工作簿本身。
///
/// 安全语义：先在临时候选文件上执行并验证（复用 ExcelModifier 的既有校验），
/// 再用规划期 sha256 决定覆盖目标还是另存结果副本。
/// 已验证候选遇到源文件变化时另存冲突副本，提交报告返回实际结果路径。
pub fn apply_plan_in_place(
    plan: &ExcelEditPlan,
    capabilities: &[BackendCapability],
    platform: &str,
) -> Result<(EditExecutionReport, TargetCommitReport)> {
    let target = plan.input.clone();
    let original_hash = crate::file_commit::sha256_file(&target)?;
    // Validate and edit an immutable input copy. Comparing the candidate against
    // the user's concurrently edited live file would reject a valid old-snapshot
    // result before the conflict-copy policy could run.
    let snapshot = crate::file_commit::snapshot_workbook(&target, &original_hash)?;
    // 候选文件放在目标同目录，保证同卷重命名语义。
    let candidate = unique_in_place_candidate(&target)?;
    let mut candidate_plan = plan.clone();
    candidate_plan.input = snapshot.path().to_path_buf();
    candidate_plan.output = candidate.clone();

    let result = apply_plan(&candidate_plan, capabilities, platform);
    let execution = match result {
        Ok(execution) => execution,
        Err(error) => {
            let _ = fs::remove_file(&candidate);
            return Err(error);
        }
    };
    let commit =
        match crate::file_commit::commit_validated_candidate(&candidate, &target, &original_hash) {
            Ok(commit) => commit,
            Err(error) => {
                let _ = fs::remove_file(&candidate);
                return Err(error)
                    .with_context(|| format!("就地提交修改结果到 {} 失败", target.display()));
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
    anyhow::bail!(
        "cannot allocate an in-place candidate beside {}",
        target.display()
    )
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
    anyhow::ensure!(
        !capabilities.iter().any(|capability| matches!(
            capability.engine,
            EngineId::AsposeWorker | EngineId::BetterOffice | EngineId::Composite
        )),
        "修改能力文件不再接受 Aspose 或 BetterOffice 修改后端，请选用 OOXML、Openpyxl 或 UMYA"
    );
    apply_plan(plan, &capabilities, platform)
}

pub(crate) fn apply_plan(
    plan: &ExcelEditPlan,
    capabilities: &[BackendCapability],
    platform: &str,
) -> Result<EditExecutionReport> {
    let mut backends: Vec<Arc<dyn EditBackend>> = Vec::new();
    for capability in capabilities.iter().cloned() {
        match capability.engine {
            EngineId::AsposeWorker => backends.push(Arc::new(
                crate::aspose_edit::AsposeEditBackend::new(capability)?,
            )),
            EngineId::UmyaSpreadsheet => backends.push(Arc::new(UmyaXlsxBackend::new(capability)?)),
            EngineId::OoxmlPatch => {
                backends.push(Arc::new(OoxmlHeaderTextBackend::new(capability)?))
            }
            EngineId::OpenpyxlWorker => backends.push(Arc::new(
                OpenpyxlXlsxBackend::for_product_binary(
                    capability,
                    std::env::current_exe()
                        .context("resolve pytools executable for bundled openpyxl runtime")?,
                    "pytools",
                )?
                .with_sheet_verification(
                    crate::result_validation::current()
                        != crate::result_validation::ResultValidation::Off,
                ),
            )),
            EngineId::RustNativeOoxml => backends.push(Arc::new(
                excel_editor::RustNativeOoxmlBackend::new(capability)?,
            )),
            EngineId::BetterOffice => {
                backends.push(Arc::new(BetterOfficeAdapter::new(capability)?))
            }
            EngineId::Composite => backends.push(Arc::new(
                crate::excel_edit_router::CompositeAdapter::new(capability)?,
            )),
            other => anyhow::bail!("no runtime adapter is registered for backend {other}"),
        }
    }
    let aspose = capabilities
        .iter()
        .any(|c| c.engine == EngineId::AsposeWorker);
    anyhow::ensure!(
        !aspose || capabilities.len() == 1,
        "Aspose explicit route cannot mix fallback engines"
    );
    let validator: Arc<dyn excel_editor::OutputValidator> =
        if aspose && crate::result_validation::check_content() {
            Arc::new(crate::aspose_edit::AsposePlanValidator)
        } else {
            Arc::new(PlanOutputValidator)
        };
    let mut modifier = ExcelModifier::new(platform, backends, validator)?
        .with_validation_failure_observer(Arc::new(crate::review_output::ReviewPublisher));
    if !crate::result_validation::check_content() {
        modifier = modifier
            .without_backend_output_validation()
            .with_failed_candidate_retention();
    }
    modifier.execute(plan).map_err(anyhow::Error::from)
}

/// Run a single protected plan with an explicitly resolved Openpyxl runtime.
/// This narrow entry point lets internal calculation fallbacks reuse the same
/// ExcelModifier validation and commit path without relying on capability JSON.
pub(crate) fn apply_plan_with_openpyxl_config(
    plan: &ExcelEditPlan,
    capability: BackendCapability,
    config: OpenpyxlWorkerConfig,
    platform: &str,
) -> Result<EditExecutionReport> {
    let backend = OpenpyxlXlsxBackend::new(capability, config)?.with_sheet_verification(
        crate::result_validation::current() != crate::result_validation::ResultValidation::Off,
    );
    let mut modifier = ExcelModifier::new(
        platform,
        vec![Arc::new(backend)],
        Arc::new(PlanOutputValidator),
    )?;
    if !crate::result_validation::check_content() {
        modifier = modifier
            .without_backend_output_validation()
            .with_failed_candidate_retention();
    }
    modifier.execute(plan).map_err(anyhow::Error::from)
}

struct PlanOutputValidator;

impl excel_editor::OutputValidator for PlanOutputValidator {
    fn validate_for_engine(
        &self,
        engine: EngineId,
        plan: &ExcelEditPlan,
        candidate_output: &Path,
    ) -> Result<Vec<String>> {
        // The worker already reopened and validated the final saved package.
        if engine != EngineId::OpenpyxlWorker {
            self.validate(plan, candidate_output)?;
        }
        Ok(Vec::new())
    }

    fn validate(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        if !crate::result_validation::check_content() {
            return crate::result_validation::check_edit_plan(plan, candidate_output);
        }
        if plan
            .operations
            .iter()
            .all(|operation| matches!(operation, EditOperation::ReplaceHeaderText { .. }))
        {
            if plan
                .preserve
                .contains(&excel_editor::PreservationRequirement::UnmodifiedPackageParts)
            {
                OoxmlHeaderTextValidator.validate(plan, candidate_output)
            } else {
                excel_editor::OpenpyxlXlsxValidator.validate(plan, candidate_output)
            }
        } else {
            UmyaXlsxValidator.validate(plan, candidate_output)
        }
    }
}
