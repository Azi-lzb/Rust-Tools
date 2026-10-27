//! Narrow OOXML Patch backend for changing existing worksheet header text.

use crate::{
    BackendCapability, BackendFailure, EditBackend, EditOperation, EditOperationKind,
    ExcelEditPlan, OutputValidator, WorkbookFormat,
};
use anyhow::{bail, Result};
use ooxml_patch::{
    preflight_header_text_patch, replace_header_text, validate_header_text_patch, HeaderTextUpdate,
};
use std::path::Path;

pub struct OoxmlHeaderTextBackend {
    capability: BackendCapability,
}

impl OoxmlHeaderTextBackend {
    /// The caller supplies capability evidence; this constructor does not
    /// promote an unverified engine operation to passed.
    pub fn new(capability: BackendCapability) -> Result<Self> {
        if capability.engine != crate::EngineId::OoxmlPatch {
            bail!("header text backend must be registered as ooxml_patch");
        }
        if !capability
            .operations
            .contains(&EditOperationKind::ReplaceHeaderText)
        {
            bail!("header text backend capability omits ReplaceHeaderText");
        }
        Ok(Self { capability })
    }
}

impl EditBackend for OoxmlHeaderTextBackend {
    fn capability(&self) -> &BackendCapability {
        &self.capability
    }

    fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure> {
        if plan.format != WorkbookFormat::Xlsx {
            return Err(BackendFailure::unsupported_before_write(
                "OOXML header text patch currently accepts only XLSX",
            ));
        }
        let updates = updates_from_plan(plan)
            .map_err(|error| BackendFailure::unsupported_before_write(error.to_string()))?;
        preflight_header_text_patch(&plan.input, &updates)
            .map_err(|error| BackendFailure::failed(format!("header patch preflight: {error}")))?;
        replace_header_text(&plan.input, &plan.output, &updates)
            .map_err(|error| BackendFailure::failed(format!("write header patch: {error}")))?;
        Ok(())
    }
}

pub struct OoxmlHeaderTextValidator;

impl OutputValidator for OoxmlHeaderTextValidator {
    fn validate(&self, original_plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        let updates = updates_from_plan(original_plan)?;
        validate_header_text_patch(&original_plan.input, candidate_output, &updates)
    }
}

fn updates_from_plan(plan: &ExcelEditPlan) -> Result<Vec<HeaderTextUpdate>> {
    plan.operations
        .iter()
        .map(|operation| match operation {
            EditOperation::ReplaceHeaderText {
                sheet,
                search,
                replacement,
            } => Ok(HeaderTextUpdate {
                sheet_name: sheet.clone(),
                search: search.clone(),
                replacement: replacement.clone(),
            }),
            other => bail!(
                "OOXML header text backend cannot execute {:?}",
                other.kind()
            ),
        })
        .collect()
}
