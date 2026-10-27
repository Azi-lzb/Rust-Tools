//! Plan serialization and evidence-gated ExcelModifier execution.

use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, EditBackend, EditExecutionReport, EditOperation, EngineId, ExcelEditPlan,
    ExcelModifier, OoxmlHeaderTextBackend, OoxmlHeaderTextValidator, OpenpyxlXlsxBackend,
    UmyaXlsxBackend, UmyaXlsxValidator,
};
use serde::Deserialize;
use std::{fs, path::Path, sync::Arc};

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
