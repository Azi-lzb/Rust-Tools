//! In-memory route selection and the unified ExcelEditPlan execution entry.
//!
//! Route selection only inspects the already-built plan. It never reads a
//! workbook or probes an engine; adapters and the existing protected
//! candidate/commit pipeline own those steps.

use crate::{
    betteroffice_edit::{supported_operation as betteroffice_supports, BetterOfficeAdapter},
    inplace_edit::{capability_for, platform_id, InPlaceEngine},
};
use anyhow::{bail, ensure, Context, Result};
use excel_editor::{
    BackendCapability, BackendFailure, EditBackend, EditOperation, EngineId, ExcelEditPlan,
    OutputValidator, PreservationRequirement, WorkbookFeature, WorkbookFormat,
};
use std::{fs, io::Write, path::Path};

pub const ROUTE_ENV: &str = "PYTOOLS_EXCEL_EDIT_ROUTE";

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum RouteMode {
    #[default]
    OpenpyxlOnly,
    UmyaOnly,
    BetterOfficeOnly,
    RustNativeOnly,
    Auto,
}

impl RouteMode {
    pub const ALL: [Self; 5] = [
        Self::OpenpyxlOnly,
        Self::UmyaOnly,
        Self::BetterOfficeOnly,
        Self::RustNativeOnly,
        Self::Auto,
    ];

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::OpenpyxlOnly => "OpenpyxlOnly",
            Self::UmyaOnly => "UMYAOnly",
            Self::BetterOfficeOnly => "BetterOfficeOnly",
            Self::RustNativeOnly => "RustNativeOnly",
            Self::Auto => "Auto",
        }
    }

    pub const fn menu_number(self) -> &'static str {
        match self {
            Self::OpenpyxlOnly => "1",
            Self::UmyaOnly => "2",
            Self::BetterOfficeOnly => "3",
            Self::RustNativeOnly => "4",
            Self::Auto => "5",
        }
    }
}

impl std::str::FromStr for RouteMode {
    type Err = anyhow::Error;

    fn from_str(value: &str) -> Result<Self> {
        let normalized = value
            .trim()
            .chars()
            .filter(|character| *character != '_' && *character != '-')
            .flat_map(char::to_lowercase)
            .collect::<String>();
        match normalized.as_str() {
            "openpyxlonly" | "openpyxl" => Ok(Self::OpenpyxlOnly),
            "umyaonly" | "umya" => Ok(Self::UmyaOnly),
            "betterofficeonly" | "betteroffice" => Ok(Self::BetterOfficeOnly),
            "rustnativeonly" | "rustnative" | "native" => Ok(Self::RustNativeOnly),
            "auto" => Ok(Self::Auto),
            _ => bail!(
                "invalid {ROUTE_ENV}={value:?}; expected OpenpyxlOnly, UMYAOnly, BetterOfficeOnly, RustNativeOnly, or Auto"
            ),
        }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RouteConfig {
    pub mode: RouteMode,
}

impl RouteConfig {
    pub const fn new(mode: RouteMode) -> Self {
        Self { mode }
    }

    /// Missing configuration deliberately retains the original Worker route.
    pub fn from_environment() -> Result<Self> {
        match std::env::var(ROUTE_ENV) {
            Ok(value) => Ok(Self::new(value.parse()?)),
            Err(std::env::VarError::NotPresent) => Ok(Self::default()),
            Err(error) => Err(error).context("read Excel edit route environment variable"),
        }
    }

    /// Set the route for this process and child commands launched by it.
    pub fn set_for_current_process(self) {
        std::env::set_var(ROUTE_ENV, self.mode.as_str());
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RouteSelection {
    pub mode: RouteMode,
    pub engine: EngineId,
    pub fallback_to_openpyxl: bool,
    pub explanation: &'static str,
}

impl RouteSelection {
    fn engines(&self) -> Vec<EngineId> {
        let mut engines = vec![self.engine];
        if self.fallback_to_openpyxl && self.engine != EngineId::OpenpyxlWorker {
            engines.push(EngineId::OpenpyxlWorker);
        }
        engines
    }
}

/// Choose an engine using only the plan's declared operations and profile.
pub fn route_plan(plan: &ExcelEditPlan, config: &RouteConfig) -> Result<RouteSelection> {
    plan.validate()
        .context("validate ExcelEditPlan before route selection")?;
    let (engine, fallback_to_openpyxl, explanation) = match config.mode {
        RouteMode::OpenpyxlOnly => (EngineId::OpenpyxlWorker, false, "manual OpenpyxlOnly route"),
        RouteMode::UmyaOnly => {
            ensure_umya_plan(plan)?;
            (EngineId::UmyaSpreadsheet, false, "manual UMYAOnly route")
        }
        RouteMode::BetterOfficeOnly => {
            ensure_betteroffice_plan(plan)?;
            (
                EngineId::BetterOffice,
                false,
                "manual BetterOfficeOnly route",
            )
        }
        RouteMode::RustNativeOnly => {
            ensure_native_plan(plan)?;
            (
                EngineId::RustNativeOoxml,
                false,
                "manual RustNativeOnly route",
            )
        }
        RouteMode::Auto => auto_route(plan)?,
    };

    if engine == EngineId::UmyaSpreadsheet && has_comment_edit(plan) {
        bail!("UMYA route is blocked for comment edits; select OpenpyxlOnly or Auto");
    }
    if engine == EngineId::UmyaSpreadsheet {
        ensure_umya_plan(plan)?;
    }
    if engine == EngineId::BetterOffice {
        ensure_betteroffice_plan(plan)?;
    }
    if engine == EngineId::RustNativeOoxml {
        ensure_native_plan(plan)?;
    }
    Ok(RouteSelection {
        mode: config.mode,
        engine,
        fallback_to_openpyxl,
        explanation,
    })
}

fn auto_route(plan: &ExcelEditPlan) -> Result<(EngineId, bool, &'static str)> {
    if !plan.workbook.inventory_complete {
        return Ok((
            EngineId::OpenpyxlWorker,
            false,
            "incomplete workbook inventory requires Openpyxl",
        ));
    }
    if has_comment_edit(plan) || has_complex_preservation_features(plan) {
        return Ok((
            EngineId::OpenpyxlWorker,
            false,
            "comments or complex workbook objects require Openpyxl",
        ));
    }

    let has_structure = plan.operations.iter().any(is_structural_edit);
    if has_structure
        && (plan
            .workbook
            .features_present
            .contains(&WorkbookFeature::CalcChain)
            || plan
                .workbook
                .features_present
                .contains(&WorkbookFeature::ExternalLinks))
    {
        return Ok((
            EngineId::OpenpyxlWorker,
            false,
            "structural edits with calcChain or external links require Openpyxl",
        ));
    }

    if contains_umya_style_or_dimension(plan) {
        return Ok((
            EngineId::UmyaSpreadsheet,
            true,
            "style or dimension edits route to UMYA",
        ));
    }

    let native_only = plan
        .operations
        .iter()
        .all(|operation| native_supports(operation));
    if native_only {
        return Ok((
            EngineId::RustNativeOoxml,
            true,
            "cell values and OOXML print/name properties route to RustNative",
        ));
    }

    let betteroffice_only = plan
        .operations
        .iter()
        .all(|operation| betteroffice_supports(operation.kind()));
    if betteroffice_only {
        return Ok((
            EngineId::BetterOffice,
            true,
            "formula or row/column structure edits route to BetterOffice",
        ));
    }

    let native_exclusive_count = plan
        .operations
        .iter()
        .filter(|operation| native_supports(operation) && !betteroffice_supports(operation.kind()))
        .count();
    let betteroffice_exclusive_count = plan
        .operations
        .iter()
        .filter(|operation| betteroffice_supports(operation.kind()) && !native_supports(operation))
        .count();
    if native_exclusive_count > 0
        && betteroffice_exclusive_count > 0
        && plan
            .operations
            .iter()
            .all(|operation| native_supports(operation) || betteroffice_supports(operation.kind()))
    {
        return Ok((
            EngineId::Composite,
            true,
            "mixed RustNative and BetterOffice edits use an in-memory composite",
        ));
    }

    if has_structure {
        return Ok((
            EngineId::BetterOffice,
            true,
            "row/column structure edits route to BetterOffice",
        ));
    }
    if plan
        .operations
        .iter()
        .all(|operation| matches!(operation, EditOperation::SetCellValue { .. }))
    {
        return Ok((
            EngineId::RustNativeOoxml,
            true,
            "cell-value edits route to RustNative",
        ));
    }
    Ok((
        EngineId::OpenpyxlWorker,
        false,
        "operation mix is outside the narrow Rust adapter scope",
    ))
}

fn has_comment_edit(plan: &ExcelEditPlan) -> bool {
    plan.operations
        .iter()
        .any(|operation| matches!(operation, EditOperation::SetComment { .. }))
}

fn is_structural_edit(operation: &EditOperation) -> bool {
    matches!(
        operation,
        EditOperation::InsertRows { .. }
            | EditOperation::DeleteRows { .. }
            | EditOperation::InsertColumns { .. }
            | EditOperation::DeleteColumns { .. }
    )
}

fn has_complex_preservation_features(plan: &ExcelEditPlan) -> bool {
    use WorkbookFeature::*;
    // Every ordinary OOXML workbook contains .rels package parts. Their
    // presence alone is not a complex feature; output validators separately
    // enforce PreservationRequirement::Relationships.
    [
        Comments,
        Charts,
        Drawings,
        Images,
        ExternalLinks,
        Macros,
        PivotTables,
        Slicers,
        ActiveX,
        DigitalSignatures,
        UnknownExtensions,
        CellMetadata,
    ]
    .iter()
    .any(|feature| plan.workbook.features_present.contains(feature))
}

fn contains_umya_style_or_dimension(plan: &ExcelEditPlan) -> bool {
    plan.operations.iter().any(|operation| {
        matches!(
            operation,
            EditOperation::SetCellFillColor { .. }
                | EditOperation::SetRowHeight { .. }
                | EditOperation::SetColumnWidth { .. }
                | EditOperation::CopyRange {
                    copy_styles: true,
                    ..
                }
                | EditOperation::ImportRange {
                    style_mode: excel_editor::RangeStyleMode::Basic
                        | excel_editor::RangeStyleMode::Full,
                    copy_merges: true,
                    ..
                }
        )
    })
}

fn native_supports(operation: &EditOperation) -> bool {
    match operation {
        EditOperation::SetCellValue { value, .. } => {
            !matches!(value, excel_editor::CellValue::DateTime { .. })
        }
        EditOperation::SetPageSettings { .. }
        | EditOperation::SetMargins { .. }
        | EditOperation::SetDefinedName { .. }
        | EditOperation::SetWorkbookCalculationProperties { .. }
        | EditOperation::ReplaceHeaderText { .. } => true,
        _ => false,
    }
}

fn ensure_native_plan(plan: &ExcelEditPlan) -> Result<()> {
    ensure!(
        plan.format == WorkbookFormat::Xlsx,
        "RustNativeOnly accepts XLSX workbooks"
    );
    ensure!(
        plan.workbook.inventory_complete,
        "RustNativeOnly refuses workbooks with an incomplete OOXML inventory"
    );
    ensure!(
        plan.operations.iter().all(native_supports),
        "RustNativeOnly cannot execute every operation in this plan"
    );
    Ok(())
}

fn ensure_umya_plan(plan: &ExcelEditPlan) -> Result<()> {
    ensure!(
        plan.format == WorkbookFormat::Xlsx,
        "UMYAOnly accepts XLSX workbooks"
    );
    ensure!(
        plan.workbook.inventory_complete,
        "UMYAOnly refuses workbooks with an incomplete OOXML inventory"
    );
    ensure!(
        !has_complex_preservation_features(plan),
        "UMYAOnly refuses workbooks with comments, charts, external links, macros, or other complex objects"
    );
    Ok(())
}

fn ensure_betteroffice_plan(plan: &ExcelEditPlan) -> Result<()> {
    ensure!(
        plan.format == WorkbookFormat::Xlsx,
        "BetterOfficeOnly accepts XLSX workbooks"
    );
    ensure!(
        plan.workbook.inventory_complete,
        "BetterOfficeOnly refuses workbooks with an incomplete OOXML inventory"
    );
    if plan
        .workbook
        .features_present
        .contains(&WorkbookFeature::ExternalLinks)
    {
        bail!("BetterOffice refuses workbooks containing external links");
    }
    ensure!(
        !has_complex_preservation_features(plan),
        "BetterOffice refuses workbooks with comments, charts, macros, or other complex objects"
    );
    if plan.operations.iter().any(is_structural_edit)
        && plan
            .workbook
            .features_present
            .contains(&WorkbookFeature::CalcChain)
    {
        bail!("BetterOffice refuses row/column edits when calcChain is present");
    }
    ensure!(
        plan.operations
            .iter()
            .all(|operation| betteroffice_supports(operation.kind())),
        "BetterOfficeOnly cannot execute every operation in this plan"
    );
    Ok(())
}

/// Execute a plan against the same workbook path through the candidate,
/// validation, source-hash check, and safe replacement flow.
pub fn execute_plan(
    plan: &ExcelEditPlan,
    file_path: &Path,
    route_config: &RouteConfig,
) -> Result<()> {
    ensure!(
        plan.input == file_path,
        "ExcelEditPlan input path does not match execute_plan file_path"
    );
    execute_in_place_with_route_config(plan, route_config)?;
    Ok(())
}

pub fn execute_candidate_with_route_config(
    plan: &ExcelEditPlan,
    route_config: &RouteConfig,
) -> Result<excel_editor::EditExecutionReport> {
    let selection = route_plan(plan, route_config)?;
    let engines = selection.engines();
    let execution_plan = execution_plan_for(plan, route_config, selection.engine);
    let capabilities = engines
        .iter()
        .map(|engine| capability_for(&execution_plan, InPlaceEngine::from_engine(*engine)))
        .collect::<Vec<_>>();
    let execution = crate::edit_plan::apply_plan(&execution_plan, &capabilities, &platform_id())?;
    emit_route_result(&execution, route_config, &plan.input);
    Ok(execution)
}

pub(crate) fn execute_in_place_with_route_config(
    plan: &ExcelEditPlan,
    route_config: &RouteConfig,
) -> Result<crate::inplace_edit::InPlaceOutcome> {
    let selection = route_plan(plan, route_config)?;
    let engines = selection.engines();
    let execution_plan = execution_plan_for(plan, route_config, selection.engine);
    let capabilities = engines
        .iter()
        .map(|engine| capability_for(&execution_plan, InPlaceEngine::from_engine(*engine)))
        .collect::<Vec<_>>();
    let (execution, commit) =
        crate::edit_plan::apply_plan_in_place(&execution_plan, &capabilities, &platform_id())?;
    let outcome = crate::inplace_edit::InPlaceOutcome {
        target: plan.input.clone(),
        execution,
        commit,
    };
    emit_route_result(&outcome.execution, route_config, &plan.input);
    Ok(outcome)
}

fn execution_plan_for(
    plan: &ExcelEditPlan,
    route_config: &RouteConfig,
    engine: EngineId,
) -> ExcelEditPlan {
    let mut execution_plan = if engine == EngineId::OpenpyxlWorker {
        crate::inplace_edit::plan_for_engine(plan, InPlaceEngine::Openpyxl)
    } else {
        plan.clone()
    };
    if route_config.mode == RouteMode::Auto
        || matches!(
            engine,
            EngineId::UmyaSpreadsheet | EngineId::BetterOffice | EngineId::Composite
        )
    {
        // These engines rewrite package XML instead of preserving every ZIP
        // member byte-for-byte. Concrete values, formulas, styles, relations,
        // and requested edits remain covered by their validators.
        execution_plan
            .preserve
            .remove(&PreservationRequirement::UnmodifiedPackageParts);
    }
    execution_plan
}

fn emit_route_result(
    execution: &excel_editor::EditExecutionReport,
    config: &RouteConfig,
    file_path: &Path,
) {
    let timestamp = chrono::Local::now().format("%Y-%m-%d %H:%M:%S");
    eprintln!(
        "[{timestamp}] [INFO] ExcelEditPlan 路由：模式 {}，引擎 {}，文件 {}，耗时 {} ms",
        config.mode.as_str(),
        execution.engine,
        crate::run_logging::portable_file_name(file_path),
        execution.elapsed_ms
    );
}

pub(crate) struct CompositeAdapter {
    capability: BackendCapability,
}

impl CompositeAdapter {
    pub(crate) fn new(capability: BackendCapability) -> Result<Self> {
        ensure!(
            capability.engine == EngineId::Composite,
            "composite adapter requires a composite capability"
        );
        Ok(Self { capability })
    }

    fn apply_bytes(source: &[u8], plan: &ExcelEditPlan) -> Result<Vec<u8>> {
        let mut betteroffice_plan = plan.clone();
        betteroffice_plan.operations.clear();
        let mut native_plan = plan.clone();
        native_plan.operations.clear();
        for operation in &plan.operations {
            if betteroffice_supports(operation.kind()) {
                betteroffice_plan.operations.push(operation.clone());
            } else if native_supports(operation) {
                native_plan.operations.push(operation.clone());
            } else {
                bail!("composite adapter cannot execute {:?}", operation.kind());
            }
        }
        ensure!(
            !betteroffice_plan.operations.is_empty() && !native_plan.operations.is_empty(),
            "composite adapter requires both BetterOffice and RustNative operations"
        );
        let intermediate = BetterOfficeAdapter::apply_bytes(source, &betteroffice_plan)
            .context("apply BetterOffice operations in memory")?;
        excel_editor::RustNativeOoxmlBackend::apply_bytes(&intermediate, &native_plan)
            .context("apply RustNative OOXML operations to in-memory package")
    }
}

impl EditBackend for CompositeAdapter {
    fn capability(&self) -> &BackendCapability {
        &self.capability
    }

    fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure> {
        if plan.format != WorkbookFormat::Xlsx
            || !plan.workbook.inventory_complete
            || has_complex_preservation_features(plan)
        {
            return Err(BackendFailure::unsupported_before_write(
                "in-memory composite accepts simple XLSX packages without complex workbook objects",
            ));
        }
        let source = fs::read(&plan.input)
            .map_err(|error| BackendFailure::failed(format!("read composite input: {error}")))?;
        let candidate = Self::apply_bytes(&source, plan).map_err(|error| {
            BackendFailure::unsupported_before_write(format!("composite preflight/edit: {error:#}"))
        })?;
        let mut output = fs::OpenOptions::new()
            .write(true)
            .truncate(true)
            .open(&plan.output)
            .map_err(|error| {
                BackendFailure::failed(format!("open composite candidate: {error}"))
            })?;
        output
            .write_all(&candidate)
            .and_then(|_| output.sync_all())
            .map_err(|error| BackendFailure::failed(format!("write composite candidate: {error}")))
    }

    fn validate_output(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        CompositeValidator.validate(plan, candidate_output)
    }
}

struct CompositeValidator;

impl OutputValidator for CompositeValidator {
    fn validate(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        excel_editor::NativeOoxmlValidator.validate(plan, candidate_output)?;
        excel_editor::UmyaXlsxValidator.validate(plan, candidate_output)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeSet;

    fn minimal_plan(operations: Vec<EditOperation>) -> ExcelEditPlan {
        ExcelEditPlan {
            request_id: "route-test".to_owned(),
            input: "source.xlsx".into(),
            output: "target.xlsx".into(),
            format: WorkbookFormat::Xlsx,
            workbook: excel_editor::WorkbookProfile {
                format: WorkbookFormat::Xlsx,
                inventory_complete: true,
                features_present: BTreeSet::new(),
            },
            operations,
            preserve: BTreeSet::new(),
            drop_allowed: Default::default(),
        }
    }

    #[test]
    fn default_route_keeps_openpyxl_worker() {
        let plan = minimal_plan(vec![EditOperation::SetCellValue {
            sheet: "Sheet1".to_owned(),
            cell: "A1".to_owned(),
            value: excel_editor::CellValue::Text("x".to_owned()),
        }]);
        let selection = route_plan(&plan, &RouteConfig::default()).unwrap();
        assert_eq!(selection.engine, EngineId::OpenpyxlWorker);
        assert!(!selection.fallback_to_openpyxl);
    }

    #[test]
    fn auto_uses_rustnative_for_simple_cell_value() {
        let plan = minimal_plan(vec![EditOperation::SetCellValue {
            sheet: "Sheet1".to_owned(),
            cell: "A1".to_owned(),
            value: excel_editor::CellValue::Text("x".to_owned()),
        }]);
        let selection = route_plan(&plan, &RouteConfig::new(RouteMode::Auto)).unwrap();
        assert_eq!(selection.engine, EngineId::RustNativeOoxml);
        assert!(selection.fallback_to_openpyxl);
    }

    #[test]
    fn umya_manual_route_rejects_comment_edits() {
        let plan = minimal_plan(vec![EditOperation::SetComment {
            sheet: "Sheet1".to_owned(),
            cell: "A1".to_owned(),
            text: Some("comment".to_owned()),
            author: None,
        }]);
        assert!(route_plan(&plan, &RouteConfig::new(RouteMode::UmyaOnly)).is_err());
    }

    #[test]
    fn ordinary_ooxml_relationship_parts_do_not_block_rewrite_engines() {
        let mut plan = minimal_plan(vec![EditOperation::SetCellValue {
            sheet: "Sheet1".to_owned(),
            cell: "A1".to_owned(),
            value: excel_editor::CellValue::Text("x".to_owned()),
        }]);
        plan.workbook
            .features_present
            .insert(WorkbookFeature::Relationships);

        for mode in [RouteMode::UmyaOnly, RouteMode::BetterOfficeOnly] {
            assert!(
                route_plan(&plan, &RouteConfig::new(mode)).is_ok(),
                "ordinary relationship parts should not block {mode:?}"
            );
        }
    }

    #[test]
    fn manual_rewrite_routes_refuse_incomplete_workbook_inventory() {
        let mut plan = minimal_plan(vec![EditOperation::SetCellValue {
            sheet: "Sheet1".to_owned(),
            cell: "A1".to_owned(),
            value: excel_editor::CellValue::Text("x".to_owned()),
        }]);
        plan.workbook.inventory_complete = false;

        for mode in [
            RouteMode::UmyaOnly,
            RouteMode::BetterOfficeOnly,
            RouteMode::RustNativeOnly,
        ] {
            assert!(
                route_plan(&plan, &RouteConfig::new(mode)).is_err(),
                "incomplete workbook inventory should block {mode:?}"
            );
        }
    }

    #[test]
    fn auto_routes_structural_edits_with_calcchain_to_openpyxl() {
        let mut plan = minimal_plan(vec![EditOperation::InsertRows {
            sheet: "Sheet1".to_owned(),
            first_row: 2,
            count: 1,
        }]);
        plan.workbook
            .features_present
            .insert(WorkbookFeature::CalcChain);
        let selection = route_plan(&plan, &RouteConfig::new(RouteMode::Auto)).unwrap();
        assert_eq!(selection.engine, EngineId::OpenpyxlWorker);
        assert!(!selection.fallback_to_openpyxl);
    }

    #[test]
    fn auto_uses_memory_composite_for_mixed_native_and_structural_edits() {
        let plan = minimal_plan(vec![
            EditOperation::SetMargins {
                sheet: "Sheet1".to_owned(),
                margins: excel_editor::PageMargins {
                    left: Some(0.5),
                    ..Default::default()
                },
            },
            EditOperation::InsertRows {
                sheet: "Sheet1".to_owned(),
                first_row: 2,
                count: 1,
            },
        ]);
        let selection = route_plan(&plan, &RouteConfig::new(RouteMode::Auto)).unwrap();
        assert_eq!(selection.engine, EngineId::Composite);
        assert!(selection.fallback_to_openpyxl);
    }
}
