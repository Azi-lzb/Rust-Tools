//! Isolated Python/openpyxl adapter for batch edits to existing XLSX files.
//!
//! The adapter is opt-in and evidence gated. It launches one bundled Python
//! process per workbook plan; the Python runtime and worker are never linked
//! into the Rust process.

use crate::{
    BackendCapability, BackendFailure, EditBackend, EditOperation, EditOperationKind,
    ExcelEditPlan, OutputValidator, PreservationRequirement, RangeStyleMode, WorkbookFormat,
};
use anyhow::{bail, Context, Result};
use excel_reader::{read_workbook, SheetDump};
use serde::Deserialize;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{self, File, OpenOptions},
    io::{Read, Write},
    path::{Path, PathBuf},
    process::{Command, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    thread,
    time::{Duration, Instant},
};

static SCRATCH_SEQUENCE: AtomicU64 = AtomicU64::new(0);
const DEFAULT_TIMEOUT: Duration = Duration::from_secs(120);
const SPREADSHEET_MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

#[derive(Clone, Debug)]
pub struct OpenpyxlWorkerConfig {
    pub python_executable: PathBuf,
    pub worker_script: PathBuf,
    pub timeout: Duration,
}

impl OpenpyxlWorkerConfig {
    pub fn new(
        python_executable: impl Into<PathBuf>,
        worker_script: impl Into<PathBuf>,
    ) -> Result<Self> {
        let config = Self {
            python_executable: python_executable.into(),
            worker_script: worker_script.into(),
            timeout: DEFAULT_TIMEOUT,
        };
        config.validate_paths()?;
        Ok(config)
    }

    pub fn with_timeout(mut self, timeout: Duration) -> Result<Self> {
        anyhow::ensure!(
            !timeout.is_zero(),
            "openpyxl worker timeout must be positive"
        );
        self.timeout = timeout;
        Ok(self)
    }

    /// Resolve the isolated runtime layout installed with a product binary.
    /// For `/usr/bin/pytools`, this returns paths under
    /// `/usr/libexec/pytools/openpyxl/`; no PATH lookup is performed.
    pub fn for_product_binary(binary: impl AsRef<Path>, product: &str) -> Option<Self> {
        if product.is_empty()
            || !product
                .bytes()
                .all(|byte| byte.is_ascii_lowercase() || byte.is_ascii_digit() || byte == b'-')
        {
            return None;
        }
        let binary = binary.as_ref();
        let installation_root = binary.parent()?.parent()?;
        let base = installation_root
            .join("libexec")
            .join(product)
            .join("openpyxl");
        Some(Self {
            python_executable: base.join("runtime/bin/python3"),
            worker_script: base.join("worker/openpyxl_worker.py"),
            timeout: DEFAULT_TIMEOUT,
        })
    }

    fn validate_paths(&self) -> Result<()> {
        anyhow::ensure!(
            self.python_executable.is_absolute(),
            "Python executable path must be absolute"
        );
        anyhow::ensure!(
            self.worker_script.is_absolute(),
            "openpyxl worker path must be absolute"
        );
        anyhow::ensure!(
            self.timeout > Duration::ZERO,
            "openpyxl worker timeout must be positive"
        );
        Ok(())
    }

    fn available(&self) -> bool {
        if self.validate_paths().is_err()
            || !self.python_executable.is_file()
            || !self.worker_script.is_file()
        {
            return false;
        }
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            return fs::metadata(&self.python_executable)
                .map(|metadata| metadata.permissions().mode() & 0o111 != 0)
                .unwrap_or(false);
        }
        #[cfg(not(unix))]
        true
    }
}

pub struct OpenpyxlXlsxBackend {
    capability: BackendCapability,
    config: Option<OpenpyxlWorkerConfig>,
    unavailable_reason: Option<String>,
}

impl OpenpyxlXlsxBackend {
    pub fn new(capability: BackendCapability, config: OpenpyxlWorkerConfig) -> Result<Self> {
        Self::build(capability, Some(config))
    }

    /// Register the candidate against the product's bundled runtime paths.
    /// Missing files mark this registration unsupported, allowing the router
    /// to select another independently qualified backend.
    pub fn for_product_binary(
        capability: BackendCapability,
        binary: impl AsRef<Path>,
        product: &str,
    ) -> Result<Self> {
        let config = OpenpyxlWorkerConfig::for_product_binary(binary, product);
        Self::build(capability, config)
    }

    fn build(
        mut capability: BackendCapability,
        config: Option<OpenpyxlWorkerConfig>,
    ) -> Result<Self> {
        anyhow::ensure!(
            capability.engine == crate::EngineId::OpenpyxlWorker,
            "openpyxl backend requires an openpyxl_worker capability"
        );
        anyhow::ensure!(
            capability.format == WorkbookFormat::Xlsx,
            "openpyxl worker supports XLSX only"
        );

        capability
            .operations
            .retain(|operation| supported_operation(*operation));
        // The Rust validator checks these semantic contracts independently.
        // The worker also verifies requested edits, but that does not qualify
        // preservation requirements Rust cannot inspect here.
        capability.preserved.retain(|requirement| {
            matches!(
                requirement,
                PreservationRequirement::CellValuesAndTypes
                    | PreservationRequirement::FormulaText
                    | PreservationRequirement::CommentsAndPositions
                    | PreservationRequirement::Styles
                    | PreservationRequirement::MergedRanges
                    | PreservationRequirement::SheetIdentityAndOrder
                    | PreservationRequirement::ConditionalFormatting
                    | PreservationRequirement::DataValidations
                    | PreservationRequirement::DefinedNames
                    | PreservationRequirement::RowAndColumnDimensions
                    | PreservationRequirement::PageSettings
                    | PreservationRequirement::HeaderFooterText
                    | PreservationRequirement::Relationships
                    | PreservationRequirement::FreezePanes
                    | PreservationRequirement::AutoFilter
            )
        });

        let unavailable_reason = match &config {
            Some(config) if config.available() => None,
            Some(_) => Some(
                "bundled Python/openpyxl worker paths are missing or not executable".to_owned(),
            ),
            None => Some("bundled Python/openpyxl worker paths could not be resolved".to_owned()),
        };
        if unavailable_reason.is_some() {
            capability.status = crate::CapabilityStatus::Unsupported;
        }
        Ok(Self {
            capability,
            config,
            unavailable_reason,
        })
    }
}

impl EditBackend for OpenpyxlXlsxBackend {
    fn capability(&self) -> &BackendCapability {
        &self.capability
    }

    fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure> {
        let Some(config) = self.config.as_ref().filter(|config| config.available()) else {
            return Err(BackendFailure::unsupported_before_write(
                self.unavailable_reason
                    .as_deref()
                    .unwrap_or("openpyxl worker is unavailable"),
            ));
        };
        if plan.format != WorkbookFormat::Xlsx {
            return Err(BackendFailure::unsupported_before_write(
                "openpyxl worker accepts XLSX only",
            ));
        }
        if !plan.input.is_file() {
            return Err(BackendFailure::failed(format!(
                "input workbook is missing: {}",
                plan.input.display()
            )));
        }
        let observed = crate::inspect_xlsx_profile(&plan.input).map_err(|error| {
            BackendFailure::failed(format!(
                "inspect source workbook before openpyxl: {error:#}"
            ))
        })?;
        if !observed.inventory_complete {
            return Err(BackendFailure::unsupported_before_write(
                "source workbook contains unverified OOXML parts",
            ));
        }
        if observed.features_present != plan.workbook.features_present
            || observed.format != plan.workbook.format
        {
            return Err(BackendFailure::unsupported_before_write(
                "ExcelEditPlan workbook profile differs from the source package inventory",
            ));
        }
        match merged_non_anchor_payloads(&plan.input, None) {
            Ok(payloads) => {
                if let Some((sheet, cell)) = offending_target_merged_payload(plan, payloads) {
                    return Err(BackendFailure::unsupported_before_write(format!(
                        "openpyxl cannot preserve stored content in merged non-anchor cell {sheet}!{cell}"
                    )));
                }
            }
            Err(error) => {
                return Err(BackendFailure::failed(format!(
                    "inspect merged-cell payloads before openpyxl: {error:#}"
                )))
            }
        }
        if !plan
            .operations
            .iter()
            .all(|operation| supported_operation(operation.kind()))
        {
            return Err(BackendFailure::unsupported_before_write(
                "openpyxl worker does not support the complete operation batch",
            ));
        }
        for operation in &plan.operations {
            if let EditOperation::ImportRange {
                source_workbook,
                source_sheet,
                source_range,
                ..
            } = operation
            {
                let profile = crate::inspect_xlsx_profile(source_workbook).map_err(|error| {
                    BackendFailure::failed(format!("inspect ImportRange source: {error:#}"))
                })?;
                if !profile.inventory_complete
                    || !profile
                        .features_present
                        .is_subset(&self.capability.supported_features)
                {
                    return Err(BackendFailure::unsupported_before_write(
                        "ImportRange source contains workbook features outside the openpyxl capability evidence",
                    ));
                }
                match merged_non_anchor_payloads(
                    source_workbook,
                    Some((source_sheet, source_range)),
                ) {
                    Ok(payloads) => {
                        if let Some((sheet, cell)) = payloads.into_iter().next() {
                            return Err(BackendFailure::unsupported_before_write(format!(
                                "openpyxl cannot preserve ImportRange source content in merged non-anchor cell {sheet}!{cell}"
                            )));
                        }
                    }
                    Err(error) => {
                        return Err(BackendFailure::failed(format!(
                            "inspect ImportRange merged-cell payloads before openpyxl: {error:#}"
                        )))
                    }
                }
            }
            if let EditOperation::ImportRange {
                copy_row_heights,
                copy_column_widths,
                copy_row_hidden,
                copy_column_hidden,
                copy_outline_levels,
                copy_all_row_dimensions,
                copy_all_column_dimensions,
                ..
            } = operation
            {
                if *copy_row_heights
                    || *copy_column_widths
                    || *copy_row_hidden
                    || *copy_column_hidden
                    || *copy_outline_levels
                    || *copy_all_row_dimensions
                    || *copy_all_column_dimensions
                {
                    return Err(BackendFailure::unsupported_before_write(
                        "ImportRange dimension-copy flags are not enabled until Rust readback validation covers them",
                    ));
                }
            }
        }

        let parent = plan
            .output
            .parent()
            .filter(|path| !path.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let request_file = match ScratchFile::create(parent, "openpyxl-request") {
            Ok(file) => file,
            Err(error) => {
                return Err(BackendFailure::failed(format!(
                    "create worker request file: {error:#}"
                )))
            }
        };
        let response_file = match ScratchFile::create(parent, "openpyxl-response") {
            Ok(file) => file,
            Err(error) => {
                return Err(BackendFailure::failed(format!(
                    "create worker response file: {error:#}"
                )))
            }
        };
        let stdout_file = match ScratchFile::create(parent, "openpyxl-stdout") {
            Ok(file) => file,
            Err(error) => {
                return Err(BackendFailure::failed(format!(
                    "create worker stdout file: {error:#}"
                )))
            }
        };
        let stderr_file = match ScratchFile::create(parent, "openpyxl-stderr") {
            Ok(file) => file,
            Err(error) => {
                return Err(BackendFailure::failed(format!(
                    "create worker stderr file: {error:#}"
                )))
            }
        };
        let request = serde_json::json!({
            "schema_version": 1,
            "request_id": plan.request_id,
            "format": "xlsx",
            "input": plan.input,
            "output": plan.output,
            "plan": plan,
        });
        let request_bytes = serde_json::to_vec(&request).map_err(|error| {
            BackendFailure::failed(format!("serialize openpyxl request: {error}"))
        })?;
        fs::write(request_file.path(), request_bytes)
            .map_err(|error| BackendFailure::failed(format!("write openpyxl request: {error}")))?;

        let stdout = File::create(stdout_file.path()).map_err(|error| {
            BackendFailure::failed(format!("open worker stdout capture: {error}"))
        })?;
        let stderr = File::create(stderr_file.path()).map_err(|error| {
            BackendFailure::failed(format!("open worker stderr capture: {error}"))
        })?;
        let mut child = Command::new(&config.python_executable)
            .arg("-I")
            .arg(&config.worker_script)
            .arg("--request")
            .arg(request_file.path())
            .arg("--response")
            .arg(response_file.path())
            .stdin(Stdio::null())
            .stdout(Stdio::from(stdout))
            .stderr(Stdio::from(stderr))
            .spawn()
            .map_err(|error| {
                BackendFailure::unsupported_before_write(format!(
                    "start bundled openpyxl worker: {error}"
                ))
            })?;
        let started = Instant::now();
        let status = loop {
            match child.try_wait() {
                Ok(Some(status)) => break status,
                Ok(None) if started.elapsed() < config.timeout => {
                    thread::sleep(Duration::from_millis(10))
                }
                Ok(None) => {
                    let _ = child.kill();
                    let _ = child.wait();
                    return Err(BackendFailure::failed(format!(
                        "openpyxl worker timed out after {} ms; candidate discarded and fallback stopped",
                        config.timeout.as_millis()
                    )));
                }
                Err(error) => {
                    let _ = child.kill();
                    let _ = child.wait();
                    return Err(BackendFailure::failed(format!(
                        "wait for openpyxl worker: {error}"
                    )));
                }
            }
        };
        let response_bytes = fs::read(response_file.path()).unwrap_or_default();
        let response: WorkerResponse =
            serde_json::from_slice(&response_bytes).map_err(|error| {
                let diagnostic = read_diagnostic(stderr_file.path());
                BackendFailure::failed(format!(
                    "invalid openpyxl worker response (exit {:?}): {error}; {diagnostic}",
                    status.code()
                ))
            })?;
        if response.schema_version != 1
            || response.request_id.as_deref() != Some(plan.request_id.as_str())
        {
            return Err(BackendFailure::failed(
                "openpyxl worker response version/request_id mismatch",
            ));
        }
        if response.status == "unsupported_before_write" {
            let error = response
                .error
                .map(|error| format!("{}: {}", error.kind, error.message))
                .unwrap_or_else(|| "worker rejected operation before writing".to_owned());
            return Err(BackendFailure::unsupported_before_write(error));
        }
        if response.status != "applied" || !status.success() {
            let error = response
                .error
                .map(|error| format!("{}: {}", error.kind, error.message))
                .unwrap_or_else(|| format!("worker exited with {:?}", status.code()));
            let diagnostic = read_diagnostic(stderr_file.path());
            return Err(BackendFailure::failed(format!(
                "openpyxl worker failed: {error}; {diagnostic}"
            )));
        }
        if response.operation_count != Some(plan.operations.len()) {
            return Err(BackendFailure::failed(
                "openpyxl worker applied-operation count mismatch",
            ));
        }
        if response.verified_operation_count != Some(plan.operations.len()) {
            return Err(BackendFailure::failed(
                "openpyxl worker verified-operation count mismatch",
            ));
        }
        if response.input_sha256_before.is_none()
            || response.input_sha256_before != response.input_sha256_after
            || response.output_sha256.as_deref().map_or(true, |hash| {
                hash.len() != 64 || !hash.bytes().all(|byte| byte.is_ascii_hexdigit())
            })
        {
            return Err(BackendFailure::failed(
                "openpyxl worker integrity diagnostics are incomplete or inconsistent",
            ));
        }
        Ok(())
    }

    fn validate_output(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        OpenpyxlXlsxValidator.validate(plan, candidate_output)
    }
}

fn supported_operation(operation: EditOperationKind) -> bool {
    matches!(
        operation,
        EditOperationKind::SetCellValue
            | EditOperationKind::SetFormulaText
            | EditOperationKind::SetComment
            | EditOperationKind::SetColumnWidth
            | EditOperationKind::SetRowHeight
            | EditOperationKind::RenameSheet
            | EditOperationKind::CreateSheet
            | EditOperationKind::CreateSheetAt
            | EditOperationKind::SetFreezePanes
            | EditOperationKind::SetAutoFilter
            | EditOperationKind::DeleteSheet
            | EditOperationKind::SetCellFillColor
            | EditOperationKind::CopyRange
            | EditOperationKind::ImportRange
            | EditOperationKind::SetWorkbookCalculationProperties
            | EditOperationKind::SetPageSettings
    )
}

pub struct OpenpyxlXlsxValidator;

impl OutputValidator for OpenpyxlXlsxValidator {
    fn validate(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        let source =
            read_workbook(&plan.input).context("read source workbook for openpyxl validation")?;
        let output = read_workbook(candidate_output)
            .context("reopen openpyxl candidate with Calamine/OOXML reader")?;
        let profile = crate::inspect_xlsx_profile(candidate_output)?;
        anyhow::ensure!(
            profile.inventory_complete,
            "openpyxl output contains unknown OOXML parts"
        );
        anyhow::ensure!(
            !output.sheets.is_empty(),
            "openpyxl output contains no worksheets"
        );
        validate_sheet_names(&source.sheets, &output.sheets, &plan.operations)?;
        validate_requested_operations(&source.sheets, &output.sheets, &plan.operations)?;
        validate_calculation_properties(candidate_output, &plan.operations)?;
        if plan.operations.iter().any(needs_umya_readback)
            || plan.preserve.contains(&PreservationRequirement::Styles)
        {
            validate_umya_readback(plan, candidate_output)?;
        }
        validate_preserved_values(&source.sheets, &output.sheets, plan)?;
        validate_preserved_formulas(&source.sheets, &output.sheets, plan)?;
        validate_preserved_comments(&source.sheets, &output.sheets, plan)?;
        validate_preserved_merges(&source.sheets, &output.sheets, plan)?;
        validate_preserved_ooxml_metadata(&plan.input, candidate_output, plan)?;
        validate_requested_sheet_layout(plan, candidate_output)?;
        Ok(())
    }
}

#[derive(Clone, Debug, Eq, Ord, PartialEq, PartialOrd)]
struct CanonicalXmlElement {
    name: String,
    attributes: Vec<(String, String, String)>,
    text: Option<String>,
    children: Vec<Self>,
}

#[derive(Clone, Debug, Default, Eq, PartialEq)]
struct WorkbookPreservationSnapshot {
    sheet_order: Vec<String>,
    defined_names: Vec<CanonicalDefinedName>,
    dimensions: Vec<(
        String,
        Vec<CanonicalXmlElement>,
        Vec<(String, Vec<(String, String)>)>,
    )>,
    page_settings: Vec<(String, Vec<CanonicalXmlElement>)>,
    header_footers: Vec<(String, Option<CanonicalXmlElement>)>,
    conditional_formatting: Vec<(String, Vec<CanonicalXmlElement>)>,
    data_validations: Vec<(String, Option<CanonicalXmlElement>)>,
    relationships: Vec<(String, Vec<CanonicalXmlElement>)>,
    freeze_panes: Vec<(String, Option<CanonicalXmlElement>)>,
    auto_filters: Vec<(String, Option<CanonicalXmlElement>)>,
}

#[derive(Clone, Debug, Eq, Ord, PartialEq, PartialOrd)]
struct CanonicalDefinedName {
    /// The owning worksheet for a local name, resolved from localSheetId.
    /// `None` denotes a workbook-scoped name.
    scope: Option<String>,
    /// The full definedName XML except for localSheetId, which is an order
    /// dependent index rather than a worksheet identity.
    element: CanonicalXmlElement,
}

type DimensionAttributeKey = (String, String);
type DimensionAttributes = BTreeMap<DimensionAttributeKey, String>;
type ColumnDimensionMap = BTreeMap<u32, DimensionAttributes>;
type RowDimensionMap = BTreeMap<u32, DimensionAttributes>;

fn canonical_column_dimensions(
    elements: &[CanonicalXmlElement],
    modified_columns: &BTreeSet<u32>,
) -> Result<ColumnDimensionMap> {
    let mut dimensions = ColumnDimensionMap::new();
    for element in elements {
        let attribute = |name: &str| {
            element
                .attributes
                .iter()
                .find(|(namespace, attribute_name, _)| {
                    namespace.is_empty() && attribute_name == name
                })
                .map(|(_, _, value)| value.as_str())
        };
        let min = attribute("min")
            .context("worksheet column dimension is missing its min attribute")?
            .parse::<u32>()
            .context("parse worksheet column dimension min")?;
        let max = attribute("max")
            .context("worksheet column dimension is missing its max attribute")?
            .parse::<u32>()
            .context("parse worksheet column dimension max")?;
        anyhow::ensure!(
            min > 0 && min <= max && max <= 16_384,
            "worksheet column dimension range {min}..={max} is invalid"
        );

        let attributes = element
            .attributes
            .iter()
            .filter(|(namespace, name, _)| {
                !(namespace.is_empty() && matches!(name.as_str(), "min" | "max"))
            })
            .map(|(namespace, name, value)| ((namespace.clone(), name.clone()), value.clone()))
            .collect::<DimensionAttributes>();
        for column in min..=max {
            let mut effective_attributes = attributes.clone();
            if modified_columns.contains(&column) {
                // The requested width is checked independently by Rust readback.
                // Preserve all other properties on the same column.
                effective_attributes.remove(&(String::new(), "width".to_owned()));
                effective_attributes.remove(&(String::new(), "customWidth".to_owned()));
            }
            if effective_attributes.is_empty() {
                continue;
            }
            if let Some(existing) = dimensions.get(&column) {
                anyhow::ensure!(
                    existing == &effective_attributes,
                    "worksheet column dimension ranges overlap with conflicting attributes at column {column}"
                );
            } else {
                dimensions.insert(column, effective_attributes);
            }
        }
    }
    Ok(dimensions)
}

fn canonical_row_dimensions(
    elements: &[(String, Vec<(String, String)>)],
    modified_rows: &BTreeSet<u32>,
) -> Result<RowDimensionMap> {
    let mut dimensions = RowDimensionMap::new();
    for (row_id, attributes) in elements {
        let row = row_id
            .parse::<u32>()
            .with_context(|| format!("parse worksheet row dimension index {row_id:?}"))?;
        anyhow::ensure!(
            (1..=1_048_576).contains(&row),
            "worksheet row dimension index {row} is out of range"
        );
        let attributes = attributes
            .iter()
            .filter(|(name, _)| name != "r")
            .map(|(name, value)| ((String::new(), name.clone()), value.clone()))
            .collect::<DimensionAttributes>();
        let mut attributes = attributes;
        if modified_rows.contains(&row) {
            // The requested height is checked independently by Rust readback.
            attributes.remove(&(String::new(), "ht".to_owned()));
            attributes.remove(&(String::new(), "customHeight".to_owned()));
        }
        if attributes.is_empty() {
            continue;
        }
        anyhow::ensure!(
            dimensions.insert(row, attributes).is_none(),
            "duplicate worksheet row dimension index {row}"
        );
    }
    Ok(dimensions)
}

fn shift_row_indices_for_insert(
    indices: BTreeSet<u32>,
    first_row: u32,
    count: u32,
) -> Result<BTreeSet<u32>> {
    indices
        .into_iter()
        .map(|row| {
            let shifted = if row >= first_row {
                row.checked_add(count)
                    .context("row dimension index overflows after row insertion")?
            } else {
                row
            };
            anyhow::ensure!(
                shifted <= 1_048_576,
                "row dimension index exceeds the worksheet limit after row insertion"
            );
            Ok(shifted)
        })
        .collect()
}

fn shift_row_indices_for_delete(
    indices: BTreeSet<u32>,
    first_row: u32,
    count: u32,
) -> Result<BTreeSet<u32>> {
    let last_deleted = first_row
        .checked_add(
            count
                .checked_sub(1)
                .context("row deletion count must not be zero")?,
        )
        .context("row deletion range overflows")?;
    indices
        .into_iter()
        .filter_map(|row| {
            if (first_row..=last_deleted).contains(&row) {
                None
            } else if row > last_deleted {
                Some(
                    row.checked_sub(count)
                        .context("row dimension index underflows after row deletion"),
                )
            } else {
                Some(Ok(row))
            }
        })
        .collect()
}

fn insert_row_dimensions(
    rows: RowDimensionMap,
    first_row: u32,
    count: u32,
) -> Result<RowDimensionMap> {
    let mut shifted = RowDimensionMap::new();
    for (row, attributes) in rows {
        let shifted_row = if row >= first_row {
            row.checked_add(count)
                .context("row dimension index overflows after row insertion")?
        } else {
            row
        };
        anyhow::ensure!(
            shifted_row <= 1_048_576,
            "row dimension index exceeds the worksheet limit after row insertion"
        );
        anyhow::ensure!(
            shifted.insert(shifted_row, attributes).is_none(),
            "row dimension collision after row insertion"
        );
    }
    Ok(shifted)
}

fn delete_row_dimensions(
    rows: RowDimensionMap,
    first_row: u32,
    count: u32,
) -> Result<RowDimensionMap> {
    let last_deleted = first_row
        .checked_add(
            count
                .checked_sub(1)
                .context("row deletion count must not be zero")?,
        )
        .context("row deletion range overflows")?;
    let mut shifted = RowDimensionMap::new();
    for (row, attributes) in rows {
        if (first_row..=last_deleted).contains(&row) {
            continue;
        }
        let shifted_row = if row > last_deleted {
            row.checked_sub(count)
                .context("row dimension index underflows after row deletion")?
        } else {
            row
        };
        anyhow::ensure!(
            shifted.insert(shifted_row, attributes).is_none(),
            "row dimension collision after row deletion"
        );
    }
    Ok(shifted)
}

fn expected_row_dimensions_after_plan(
    mut rows: RowDimensionMap,
    source_sheet: &str,
    output_sheet: &str,
    operations: &[EditOperation],
) -> Result<(RowDimensionMap, BTreeSet<u32>)> {
    let sheet_names = [source_sheet, output_sheet];
    let mut modified_rows = BTreeSet::new();
    for operation in operations {
        match operation {
            EditOperation::SetRowHeight {
                sheet, row_index, ..
            } if sheet_names.contains(&sheet.as_str()) => {
                modified_rows.insert(*row_index);
                if let Some(attributes) = rows.get_mut(row_index) {
                    attributes.remove(&(String::new(), "ht".to_owned()));
                    attributes.remove(&(String::new(), "customHeight".to_owned()));
                    if attributes.is_empty() {
                        rows.remove(row_index);
                    }
                }
            }
            EditOperation::InsertRows {
                sheet,
                first_row,
                count,
            } if sheet_names.contains(&sheet.as_str()) => {
                rows = insert_row_dimensions(rows, *first_row, *count)?;
                modified_rows = shift_row_indices_for_insert(modified_rows, *first_row, *count)?;
            }
            EditOperation::DeleteRows {
                sheet,
                first_row,
                count,
            } if sheet_names.contains(&sheet.as_str()) => {
                rows = delete_row_dimensions(rows, *first_row, *count)?;
                modified_rows = shift_row_indices_for_delete(modified_rows, *first_row, *count)?;
            }
            _ => {}
        }
    }
    Ok((rows, modified_rows))
}

fn row_dimension_attributes_match(
    expected: &DimensionAttributes,
    actual: &DimensionAttributes,
) -> bool {
    if expected == actual {
        return true;
    }

    // openpyxl adds customHeight="1" when serializing an existing explicit
    // height. Accept only that observed normalization when the numeric height
    // and every other effective dimension attribute are unchanged.
    let height = (String::new(), "ht".to_owned());
    let custom_height = (String::new(), "customHeight".to_owned());
    if expected.get(&height).is_some()
        && expected.get(&height) == actual.get(&height)
        && !expected.contains_key(&custom_height)
        && actual.get(&custom_height).is_some_and(|value| value == "1")
    {
        let mut normalized_actual = actual.clone();
        normalized_actual.remove(&custom_height);
        return expected == &normalized_actual;
    }
    false
}

fn row_dimension_maps_match(expected: &RowDimensionMap, actual: &RowDimensionMap) -> bool {
    let indices = expected
        .keys()
        .chain(actual.keys())
        .copied()
        .collect::<BTreeSet<_>>();
    indices.into_iter().all(|row| {
        let expected = expected.get(&row).cloned().unwrap_or_default();
        let actual = actual.get(&row).cloned().unwrap_or_default();
        row_dimension_attributes_match(&expected, &actual)
    })
}

fn validate_preserved_dimensions(
    source: &WorkbookPreservationSnapshot,
    output: &WorkbookPreservationSnapshot,
    plan: &ExcelEditPlan,
) -> Result<()> {
    let (retained_sheets, deleted_sheets) =
        planned_sheet_provenance(&source.sheet_order, &plan.operations)?;

    for (source_sheet, source_columns, source_rows) in &source.dimensions {
        let no_modified_dimensions = BTreeSet::new();
        let original_columns =
            canonical_column_dimensions(source_columns, &no_modified_dimensions)?;
        let original_rows = canonical_row_dimensions(source_rows, &no_modified_dimensions)?;
        if let Some(output_sheet) = retained_sheets.get(source_sheet) {
            let (output_columns, output_rows) = output
                .dimensions
                .iter()
                .find(|(name, _, _)| name == output_sheet)
                .map(|(_, columns, rows)| (columns, rows))
                .with_context(|| {
                    format!("openpyxl output is missing worksheet dimensions for {output_sheet:?}")
                })?;
            let dimension_sheet_names = [source_sheet.as_str(), output_sheet.as_str()];
            let modified_columns = plan
                .operations
                .iter()
                .filter_map(|operation| match operation {
                    EditOperation::SetColumnWidth {
                        sheet,
                        column_index,
                        ..
                    } if dimension_sheet_names.contains(&sheet.as_str()) => Some(*column_index),
                    _ => None,
                })
                .collect::<BTreeSet<_>>();
            let expected_columns = canonical_column_dimensions(source_columns, &modified_columns)?;
            let actual_columns = canonical_column_dimensions(output_columns, &modified_columns)?;
            let (expected_rows, modified_rows) = expected_row_dimensions_after_plan(
                original_rows.clone(),
                source_sheet,
                output_sheet,
                &plan.operations,
            )?;
            let actual_rows = canonical_row_dimensions(output_rows, &modified_rows)?;
            anyhow::ensure!(
                expected_columns == actual_columns,
                "openpyxl changed preserved column dimensions on worksheet {source_sheet:?}"
            );
            anyhow::ensure!(
                row_dimension_maps_match(&expected_rows, &actual_rows),
                "openpyxl changed preserved row dimensions on worksheet {source_sheet:?}"
            );
            continue;
        }

        if let Some(deleted_as) = deleted_sheets.get(source_sheet) {
            if !original_columns.is_empty() || !original_rows.is_empty() {
                let may_drop = plan.drop_allowed.get(deleted_as).is_some_and(|allowed| {
                    allowed.contains(&PreservationRequirement::RowAndColumnDimensions)
                });
                anyhow::ensure!(
                    may_drop,
                    "openpyxl would remove row or column dimensions from deleted worksheet {source_sheet:?} without an explicit RowAndColumnDimensions drop allowance"
                );
            }
            continue;
        }

        anyhow::bail!(
            "source worksheet {source_sheet:?} has no planned output identity for dimension validation"
        );
    }
    Ok(())
}

fn validate_preserved_ooxml_metadata(
    input_path: &Path,
    output_path: &Path,
    plan: &ExcelEditPlan,
) -> Result<()> {
    let checked = [
        PreservationRequirement::DefinedNames,
        PreservationRequirement::RowAndColumnDimensions,
        PreservationRequirement::PageSettings,
        PreservationRequirement::HeaderFooterText,
        PreservationRequirement::Relationships,
        PreservationRequirement::ConditionalFormatting,
        PreservationRequirement::DataValidations,
        PreservationRequirement::FreezePanes,
        PreservationRequirement::AutoFilter,
    ];
    if !checked
        .iter()
        .any(|requirement| plan.preserve.contains(requirement))
    {
        return Ok(());
    }

    let source = workbook_preservation_snapshot(input_path).map_err(|error| {
        anyhow::anyhow!(
            "read source OOXML metadata for openpyxl preservation validation: {error:#}"
        )
    })?;
    let output = workbook_preservation_snapshot(output_path).map_err(|error| {
        anyhow::anyhow!(
            "read output OOXML metadata for openpyxl preservation validation: {error:#}"
        )
    })?;

    if plan
        .preserve
        .contains(&PreservationRequirement::DefinedNames)
    {
        validate_preserved_defined_names(&source, &output, plan)?;
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::RowAndColumnDimensions)
    {
        validate_preserved_dimensions(&source, &output, plan)?;
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::PageSettings)
    {
        anyhow::ensure!(
            per_sheet_entries_match(&source.page_settings, &output.page_settings),
            "openpyxl changed worksheet page settings"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::HeaderFooterText)
    {
        anyhow::ensure!(
            per_sheet_entries_match(&source.header_footers, &output.header_footers),
            "openpyxl changed worksheet header/footer settings"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::Relationships)
    {
        anyhow::ensure!(
            source.relationships == output.relationships,
            "openpyxl changed OOXML relationships"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::ConditionalFormatting)
    {
        anyhow::ensure!(
            per_sheet_entries_match(
                &source.conditional_formatting,
                &output.conditional_formatting
            ),
            "openpyxl changed conditional formatting rules"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::DataValidations)
    {
        anyhow::ensure!(
            per_sheet_entries_match(&source.data_validations, &output.data_validations),
            "openpyxl changed data validations"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::FreezePanes)
    {
        anyhow::ensure!(
            preserved_sheet_metadata_matches(
                &source.freeze_panes,
                &output.freeze_panes,
                &plan.operations,
                EditOperationKind::SetFreezePanes,
            ),
            "openpyxl changed freeze-pane metadata on a preserved worksheet"
        );
    }
    if plan.preserve.contains(&PreservationRequirement::AutoFilter) {
        anyhow::ensure!(
            preserved_sheet_metadata_matches(
                &source.auto_filters,
                &output.auto_filters,
                &plan.operations,
                EditOperationKind::SetAutoFilter,
            ),
            "openpyxl changed AutoFilter metadata on a preserved worksheet"
        );
    }
    Ok(())
}

/// Worksheet order changes are legal edit-plan semantics (MoveSheet), so the
/// workbook-order per-sheet snapshot lists compare through sheet names:
/// every source worksheet must appear with identical content, position must
/// not matter. Content within one worksheet keeps its exact-order
/// comparison; sheet-order correctness is owned by `validate_sheet_names`.
fn per_sheet_entries_match<T: PartialEq>(source: &[(String, T)], output: &[(String, T)]) -> bool {
    source.len() == output.len()
        && source.iter().all(|(name, expected)| {
            output
                .iter()
                .find(|(output_name, _)| output_name == name)
                .is_some_and(|(_, actual)| actual == expected)
        })
}

fn preserved_sheet_metadata_matches(
    source: &[(String, Option<CanonicalXmlElement>)],
    output: &[(String, Option<CanonicalXmlElement>)],
    operations: &[EditOperation],
    operation_kind: EditOperationKind,
) -> bool {
    let (renamed, deleted) = sheet_name_changes(operations);
    source.iter().all(|(source_name, metadata)| {
        if deleted.contains(source_name) {
            return true;
        }
        let output_name = renamed.get(source_name).unwrap_or(source_name);
        let modified_by_plan =
            operations
                .iter()
                .any(|operation| match (operation_kind, operation) {
                    (
                        EditOperationKind::SetFreezePanes,
                        EditOperation::SetFreezePanes { sheet, .. },
                    )
                    | (
                        EditOperationKind::SetAutoFilter,
                        EditOperation::SetAutoFilter { sheet, .. },
                    ) => sheet == source_name || sheet == output_name,
                    _ => false,
                });
        if modified_by_plan {
            return true;
        }
        output
            .iter()
            .find(|(name, _)| name == output_name)
            .is_some_and(|(_, actual)| actual == metadata)
    })
}

fn validate_requested_sheet_layout(plan: &ExcelEditPlan, output_path: &Path) -> Result<()> {
    if !plan.operations.iter().any(|operation| {
        matches!(
            operation,
            EditOperation::SetFreezePanes { .. } | EditOperation::SetAutoFilter { .. }
        )
    }) {
        return Ok(());
    }
    let output = workbook_preservation_snapshot(output_path)?;
    for operation in &plan.operations {
        match operation {
            EditOperation::SetFreezePanes { sheet, cell } => {
                let actual = output
                    .freeze_panes
                    .iter()
                    .find(|(name, _)| name == sheet)
                    .and_then(|(_, pane)| pane.as_ref())
                    .and_then(|pane| {
                        pane.attributes.iter().find(|(namespace, name, _)| {
                            namespace.is_empty() && name == "topLeftCell"
                        })
                    })
                    .map(|(_, _, value)| normalize_address(value));
                let expected = cell.as_deref().map(normalize_address);
                anyhow::ensure!(
                    actual == expected,
                    "freeze panes readback mismatch for worksheet {sheet:?}: expected {expected:?}, got {actual:?}"
                );
            }
            EditOperation::SetAutoFilter { sheet, range } => {
                let actual = output
                    .auto_filters
                    .iter()
                    .find(|(name, _)| name == sheet)
                    .and_then(|(_, filter)| filter.as_ref())
                    .and_then(|filter| {
                        filter
                            .attributes
                            .iter()
                            .find(|(namespace, name, _)| namespace.is_empty() && name == "ref")
                    })
                    .map(|(_, _, value)| normalize_range(value));
                let expected = range.as_deref().map(normalize_range);
                anyhow::ensure!(
                    actual == expected,
                    "AutoFilter readback mismatch for worksheet {sheet:?}: expected {expected:?}, got {actual:?}"
                );
            }
            _ => {}
        }
    }
    Ok(())
}

fn workbook_preservation_snapshot(path: &Path) -> Result<WorkbookPreservationSnapshot> {
    const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const DOCUMENT_REL_NS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

    let mut archive = zip::ZipArchive::new(File::open(path)?)?;
    let workbook_xml = read_zip_part(&mut archive, "xl/workbook.xml")?;
    let workbook_document = roxmltree::Document::parse(&workbook_xml)?;
    let workbook_relationships_xml = read_zip_part(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let workbook_relationships_document = roxmltree::Document::parse(&workbook_relationships_xml)?;
    let relationship_targets = workbook_relationships_document
        .root_element()
        .children()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
        })
        .filter_map(|node| {
            let id = node.attribute("Id")?;
            let target = node.attribute("Target")?;
            let target_mode = node.attribute("TargetMode").unwrap_or_default();
            Some((id.to_owned(), (target.to_owned(), target_mode.to_owned())))
        })
        .collect::<BTreeMap<_, _>>();

    let mut sheet_parts = Vec::new();
    for node in workbook_document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "sheet"
    }) {
        let name = node
            .attribute("name")
            .context("workbook sheet is missing its name")?;
        let relationship_id = node
            .attributes()
            .find(|attribute| {
                attribute.name() == "id" && attribute.namespace() == Some(DOCUMENT_REL_NS)
            })
            .map(|attribute| attribute.value())
            .context("workbook sheet is missing its relationship id")?;
        let (target, target_mode) = relationship_targets
            .get(relationship_id)
            .with_context(|| format!("worksheet relationship {relationship_id:?} is missing"))?;
        anyhow::ensure!(
            target_mode.is_empty(),
            "worksheet relationship {relationship_id:?} is external"
        );
        sheet_parts.push((name.to_owned(), resolve_opc_target("xl", target)?));
    }

    let mut snapshot = WorkbookPreservationSnapshot::default();
    snapshot.sheet_order = sheet_parts
        .iter()
        .map(|(sheet_name, _)| sheet_name.clone())
        .collect();
    if let Some(defined_names) = workbook_document.descendants().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "definedNames"
    }) {
        snapshot.defined_names = canonical_defined_names(defined_names, &snapshot.sheet_order)?;
    }

    for (sheet_name, part_name) in sheet_parts {
        let xml = read_zip_part(&mut archive, &part_name)
            .with_context(|| format!("read worksheet part {part_name}"))?;
        let document = roxmltree::Document::parse(&xml)
            .with_context(|| format!("parse worksheet part {part_name}"))?;
        let root = document.root_element();

        let column_dimensions = root
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "cols")
            .map(|node| {
                node.children()
                    .filter(|child| child.is_element() && child.tag_name().name() == "col")
                    .map(canonical_xml_element)
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        let row_dimensions = root
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "sheetData")
            .map(|node| {
                node.children()
                    .filter(|child| child.is_element() && child.tag_name().name() == "row")
                    .map(|row| {
                        let row_id = row.attribute("r").unwrap_or_default().to_owned();
                        let attributes = row
                            .attributes()
                            .filter(|attribute| {
                                matches!(
                                    attribute.name(),
                                    "r" | "ht"
                                        | "hidden"
                                        | "customHeight"
                                        | "outlineLevel"
                                        | "collapsed"
                                        | "s"
                                        | "customFormat"
                                        | "thickTop"
                                        | "thickBot"
                                )
                            })
                            .map(|attribute| {
                                (attribute.name().to_owned(), attribute.value().to_owned())
                            })
                            .collect::<BTreeMap<_, _>>()
                            .into_iter()
                            .collect();
                        (row_id, attributes)
                    })
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        snapshot
            .dimensions
            .push((sheet_name.clone(), column_dimensions, row_dimensions));

        let page_settings = canonical_page_settings(root);
        snapshot
            .page_settings
            .push((sheet_name.clone(), page_settings));

        let pane = root
            .descendants()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "pane"
            })
            .map(canonical_xml_element);
        snapshot.freeze_panes.push((sheet_name.clone(), pane));
        let auto_filter = root
            .children()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "autoFilter"
            })
            .map(canonical_xml_element);
        snapshot
            .auto_filters
            .push((sheet_name.clone(), auto_filter));

        let header_footer = root
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "headerFooter")
            .map(canonical_xml_element);
        snapshot
            .header_footers
            .push((sheet_name.clone(), header_footer));

        let conditional_formatting = root
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "conditionalFormatting")
            .map(canonical_xml_element)
            .collect();
        snapshot
            .conditional_formatting
            .push((sheet_name.clone(), conditional_formatting));

        let data_validations = root
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "dataValidations")
            .map(canonical_xml_element);
        snapshot
            .data_validations
            .push((sheet_name, data_validations));
    }

    let relationship_parts = archive
        .file_names()
        .filter(|name| name.ends_with(".rels"))
        .map(str::to_owned)
        .collect::<Vec<_>>();
    for part_name in relationship_parts {
        let xml = read_zip_part(&mut archive, &part_name)?;
        let document = roxmltree::Document::parse(&xml)
            .with_context(|| format!("parse relationship part {part_name}"))?;
        anyhow::ensure!(
            document.root_element().tag_name().namespace() == Some(PACKAGE_REL_NS),
            "relationship part {part_name} has an unexpected root namespace"
        );
        let mut entries = document
            .root_element()
            .children()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                    && node.tag_name().name() == "Relationship"
            })
            .map(canonical_xml_element)
            .collect::<Vec<_>>();
        entries.sort();
        snapshot.relationships.push((part_name, entries));
    }
    snapshot
        .relationships
        .sort_by(|left, right| left.0.cmp(&right.0));
    Ok(snapshot)
}

fn read_zip_part(archive: &mut zip::ZipArchive<File>, part_name: &str) -> Result<String> {
    let mut part = archive.by_name(part_name)?;
    let mut xml = String::new();
    part.read_to_string(&mut xml)
        .with_context(|| format!("read OOXML part {part_name}"))?;
    Ok(xml)
}

fn canonical_page_settings(root: roxmltree::Node<'_, '_>) -> Vec<CanonicalXmlElement> {
    let mut page_settings = Vec::new();
    for child in root.children().filter(|node| node.is_element()) {
        if matches!(
            child.tag_name().name(),
            "pageMargins" | "pageSetup" | "printOptions" | "rowBreaks" | "colBreaks"
        ) {
            page_settings.push(canonical_xml_element(child));
        } else if child.tag_name().name() == "sheetPr" {
            if let Some(setup_properties) = child
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == "pageSetUpPr")
            {
                // An empty pageSetUpPr carries only schema defaults. Umya
                // reads fitToPage as false and omits the empty element again;
                // normalize that representation while retaining strict
                // per-sheet checks for every explicit attribute/child.
                let has_explicit_content = setup_properties.attributes().next().is_some()
                    || setup_properties.children().any(|node| node.is_element());
                if has_explicit_content {
                    page_settings.push(canonical_xml_element(setup_properties));
                }
            }
        }
    }
    page_settings
}

fn canonical_xml_element(node: roxmltree::Node<'_, '_>) -> CanonicalXmlElement {
    let mut attributes = node
        .attributes()
        .map(|attribute| {
            (
                attribute.namespace().unwrap_or_default().to_owned(),
                attribute.name().to_owned(),
                attribute.value().to_owned(),
            )
        })
        .collect::<Vec<_>>();
    attributes.sort();
    let children = node
        .children()
        .filter(|child| child.is_element())
        .map(canonical_xml_element)
        .collect();
    let text = node
        .children()
        .filter(|child| child.is_text())
        .filter_map(|child| child.text())
        .collect::<String>();
    CanonicalXmlElement {
        name: format!(
            "{}:{}",
            node.tag_name().namespace().unwrap_or_default(),
            node.tag_name().name()
        ),
        attributes,
        text: (!text.trim().is_empty()).then_some(text),
        children,
    }
}

fn canonical_defined_names(
    node: roxmltree::Node<'_, '_>,
    sheet_order: &[String],
) -> Result<Vec<CanonicalDefinedName>> {
    let mut names = Vec::new();
    for defined_name in node.children().filter(|child| child.is_element()) {
        let scope = defined_name
            .attribute("localSheetId")
            .map(|value| {
                let index = value.parse::<usize>().with_context(|| {
                    format!("definedName localSheetId {value:?} is not a valid index")
                })?;
                sheet_order.get(index).cloned().with_context(|| {
                    format!(
                        "definedName localSheetId {index} is outside the {}-sheet workbook",
                        sheet_order.len()
                    )
                })
            })
            .transpose()?;
        let mut element = canonical_xml_element(defined_name);
        element
            .attributes
            .retain(|(namespace, name, _)| !(namespace.is_empty() && name == "localSheetId"));
        names.push(CanonicalDefinedName { scope, element });
    }
    names.sort();
    Ok(names)
}

fn validate_preserved_defined_names(
    source: &WorkbookPreservationSnapshot,
    output: &WorkbookPreservationSnapshot,
    plan: &ExcelEditPlan,
) -> Result<()> {
    let (retained_sheets, deleted_sheets) =
        planned_sheet_provenance(&source.sheet_order, &plan.operations)?;
    let retained_output_sheets: BTreeSet<String> = retained_sheets.values().cloned().collect();
    let overridden_print_areas: BTreeSet<String> = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::SetPageSettings { sheet, settings } if settings.print_area.is_some() => {
                Some(final_sheet_name(plan, sheet))
            }
            _ => None,
        })
        .collect();
    let explicit_auto_filters: BTreeMap<String, Option<String>> = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::SetAutoFilter { sheet, range } => {
                Some((final_sheet_name(plan, sheet), range.clone()))
            }
            _ => None,
        })
        .collect();
    let mut required_filter_database_counts: BTreeMap<String, usize> = explicit_auto_filters
        .iter()
        .map(|(sheet, range)| (sheet.clone(), usize::from(range.is_some())))
        .collect();

    let mut expected = Vec::with_capacity(source.defined_names.len());
    for defined_name in &source.defined_names {
        let Some(source_scope) = &defined_name.scope else {
            expected.push(defined_name.clone());
            continue;
        };
        if let Some(final_scope) = retained_sheets.get(source_scope) {
            if overridden_print_areas.contains(final_scope)
                && is_print_area_name(&defined_name.element)
            {
                // The requested print area is checked independently against
                // the resulting workbook, while all other names stay strict.
                continue;
            }
            let explicit_filter_update = explicit_auto_filters.contains_key(final_scope);
            let preserved_filter_scope =
                plan.preserve.contains(&PreservationRequirement::AutoFilter)
                    && retained_output_sheets.contains(final_scope);
            if is_filter_database_name(&defined_name.element) {
                if explicit_filter_update {
                    // The requested AutoFilter range is independently checked
                    // against the saved worksheet and its generated name.
                    continue;
                }
                if preserved_filter_scope
                    && filter_database_matches_auto_filter(defined_name, source)
                {
                    *required_filter_database_counts
                        .entry(final_scope.clone())
                        .or_default() += 1;
                    continue;
                }
            }
            expected.push(CanonicalDefinedName {
                scope: Some(final_scope.clone()),
                element: defined_name.element.clone(),
            });
            continue;
        }
        if let Some(deleted_as) = deleted_sheets.get(source_scope) {
            let may_drop = plan
                .drop_allowed
                .get(deleted_as)
                .is_some_and(|allowed| allowed.contains(&PreservationRequirement::DefinedNames));
            anyhow::ensure!(
                may_drop,
                "openpyxl would remove defined names scoped to deleted worksheet {source_scope:?} without an explicit DefinedNames drop allowance"
            );
            continue;
        }
        anyhow::bail!(
            "source definedName scope {source_scope:?} is not present in the planned worksheet identities"
        );
    }

    let mut actual: Vec<_> = output
        .defined_names
        .iter()
        .filter(|defined_name| {
            !defined_name.scope.as_ref().is_some_and(|scope| {
                overridden_print_areas.contains(scope) && is_print_area_name(&defined_name.element)
            })
        })
        .cloned()
        .collect();
    let mut actual_filter_database_counts: BTreeMap<String, usize> = BTreeMap::new();
    actual.retain(|defined_name| {
        let Some(scope) = &defined_name.scope else {
            return true;
        };
        let validated_scope = explicit_auto_filters.contains_key(scope)
            || (plan.preserve.contains(&PreservationRequirement::AutoFilter)
                && retained_output_sheets.contains(scope));
        if is_filter_database_name(&defined_name.element)
            && validated_scope
            && filter_database_matches_auto_filter(defined_name, output)
        {
            *actual_filter_database_counts
                .entry(scope.clone())
                .or_default() += 1;
            false
        } else {
            true
        }
    });
    for (scope, expected_count) in &required_filter_database_counts {
        let actual_count = actual_filter_database_counts
            .get(scope)
            .copied()
            .unwrap_or_default();
        anyhow::ensure!(
            actual_count == *expected_count,
            "openpyxl changed the validated AutoFilter defined-name count; expected={expected_count}, actual={actual_count}"
        );
    }
    for actual_count in actual_filter_database_counts.values() {
        anyhow::ensure!(
            *actual_count <= 1,
            "openpyxl output contains duplicate AutoFilter defined names for a worksheet"
        );
    }
    expected.sort();
    actual.sort();
    let mut unmatched_output = actual;
    let mut missing_count = 0;
    for expected_name in &expected {
        if let Some(index) = unmatched_output
            .iter()
            .position(|output_name| output_name == expected_name)
        {
            unmatched_output.remove(index);
        } else {
            missing_count += 1;
        }
    }
    let unexpected_count = unmatched_output.len();
    anyhow::ensure!(
        missing_count == 0 && unexpected_count == 0,
        "openpyxl changed workbook defined names after worksheet-scope normalization; missing={missing_count}, added_or_changed={unexpected_count}"
    );
    Ok(())
}

fn is_print_area_name(element: &CanonicalXmlElement) -> bool {
    element.attributes.iter().any(|(namespace, name, value)| {
        namespace.is_empty() && name == "name" && value.eq_ignore_ascii_case("_xlnm.Print_Area")
    })
}

fn is_filter_database_name(element: &CanonicalXmlElement) -> bool {
    element.attributes.iter().any(|(namespace, name, value)| {
        namespace.is_empty()
            && name == "name"
            && value.eq_ignore_ascii_case("_xlnm._FilterDatabase")
    })
}

fn filter_database_matches_auto_filter(
    defined_name: &CanonicalDefinedName,
    workbook: &WorkbookPreservationSnapshot,
) -> bool {
    let Some(scope) = defined_name.scope.as_deref() else {
        return false;
    };
    let Some(auto_filter) = workbook
        .auto_filters
        .iter()
        .find(|(sheet, _)| sheet == scope)
        .and_then(|(_, filter)| filter.as_ref())
    else {
        return false;
    };
    let Some(auto_filter_range) = auto_filter
        .attributes
        .iter()
        .find(|(namespace, name, _)| namespace.is_empty() && name == "ref")
        .map(|(_, _, value)| value)
    else {
        return false;
    };
    let Some(formula) = defined_name.element.text.as_deref() else {
        return false;
    };
    let Some((formula_sheet, formula_range)) = formula.split_once('!') else {
        return false;
    };
    let formula_sheet = formula_sheet.trim().trim_start_matches('=');
    let formula_sheet = if formula_sheet.starts_with('\'') && formula_sheet.ends_with('\'') {
        formula_sheet[1..formula_sheet.len() - 1].replace("''", "'")
    } else {
        formula_sheet.to_owned()
    };
    !formula_sheet.contains('[')
        && !formula_sheet.contains(']')
        && formula_sheet.eq_ignore_ascii_case(scope)
        && normalize_range(formula_range) == normalize_range(auto_filter_range)
}

fn planned_sheet_provenance(
    source_sheet_order: &[String],
    operations: &[EditOperation],
) -> Result<(BTreeMap<String, String>, BTreeMap<String, String>)> {
    #[derive(Clone)]
    struct PlannedSheet {
        name: String,
        source_name: Option<String>,
    }

    let mut sheets: Vec<_> = source_sheet_order
        .iter()
        .map(|name| PlannedSheet {
            name: name.clone(),
            source_name: Some(name.clone()),
        })
        .collect();
    let mut deleted = BTreeMap::new();
    for operation in operations {
        match operation {
            EditOperation::RenameSheet { old_name, new_name } => {
                let sheet = sheets
                    .iter_mut()
                    .find(|sheet| sheet.name == *old_name)
                    .with_context(|| format!("worksheet {old_name:?} is missing in edit plan"))?;
                sheet.name.clone_from(new_name);
            }
            EditOperation::CreateSheet { name } => sheets.push(PlannedSheet {
                name: name.clone(),
                source_name: None,
            }),
            EditOperation::CreateSheetAt { name, index } => {
                anyhow::ensure!(
                    *index <= sheets.len(),
                    "worksheet insertion index {index} exceeds planned sheet count {}",
                    sheets.len()
                );
                sheets.insert(
                    *index,
                    PlannedSheet {
                        name: name.clone(),
                        source_name: None,
                    },
                );
            }
            EditOperation::DeleteSheet { name } => {
                let index = sheets
                    .iter()
                    .position(|sheet| sheet.name == *name)
                    .with_context(|| format!("worksheet {name:?} is missing in edit plan"))?;
                let removed = sheets.remove(index);
                if let Some(source_name) = removed.source_name {
                    deleted.insert(source_name, name.clone());
                }
            }
            _ => {}
        }
    }
    let retained = sheets
        .into_iter()
        .filter_map(|sheet| {
            sheet
                .source_name
                .map(|source_name| (source_name, sheet.name))
        })
        .collect();
    Ok((retained, deleted))
}

fn resolve_opc_target(base: &str, target: &str) -> Result<String> {
    anyhow::ensure!(
        !target.contains(':') && !target.contains('?') && !target.contains('#'),
        "unsupported worksheet relationship target {target:?}"
    );
    let mut segments = if target.starts_with('/') {
        Vec::new()
    } else {
        base.split('/').map(str::to_owned).collect()
    };
    for segment in target.trim_start_matches('/').split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                anyhow::ensure!(
                    segments.pop().is_some(),
                    "relationship target escapes package root"
                );
            }
            value => segments.push(value.to_owned()),
        }
    }
    anyhow::ensure!(!segments.is_empty(), "empty worksheet relationship target");
    Ok(segments.join("/"))
}

fn validate_sheet_names(
    source: &[SheetDump],
    output: &[SheetDump],
    operations: &[EditOperation],
) -> Result<()> {
    let mut expected: Vec<String> = source.iter().map(|sheet| sheet.name.clone()).collect();
    for operation in operations {
        match operation {
            EditOperation::RenameSheet { old_name, new_name } => {
                let name = expected
                    .iter_mut()
                    .find(|name| *name == old_name)
                    .with_context(|| {
                        format!("source sheet {old_name:?} missing from original workbook")
                    })?;
                *name = new_name.clone();
            }
            EditOperation::CreateSheet { name } => expected.push(name.clone()),
            EditOperation::CreateSheetAt { name, index } => {
                anyhow::ensure!(
                    *index <= expected.len(),
                    "worksheet insertion index {index} exceeds expected sheet count {}",
                    expected.len()
                );
                expected.insert(*index, name.clone());
            }
            EditOperation::DeleteSheet { name } => expected.retain(|candidate| candidate != name),
            EditOperation::MoveSheet { sheet, new_index } => {
                let current = expected
                    .iter()
                    .position(|name| name == sheet)
                    .with_context(|| {
                        format!("source sheet {sheet:?} missing from original workbook")
                    })?;
                anyhow::ensure!(
                    *new_index < expected.len(),
                    "move target index {new_index} exceeds expected sheet count {}",
                    expected.len()
                );
                let moved = expected.remove(current);
                expected.insert(*new_index, moved);
            }
            _ => {}
        }
    }
    let actual: Vec<String> = output.iter().map(|sheet| sheet.name.clone()).collect();
    anyhow::ensure!(
        actual == expected,
        "sheet identity/order mismatch: expected {expected:?}, got {actual:?}"
    );
    Ok(())
}

fn validate_requested_operations(
    source_sheets: &[SheetDump],
    sheets: &[SheetDump],
    operations: &[EditOperation],
) -> Result<()> {
    let by_name: BTreeMap<_, _> = sheets
        .iter()
        .map(|sheet| (sheet.name.as_str(), sheet))
        .collect();
    for (operation_index, operation) in operations.iter().enumerate() {
        match operation {
            EditOperation::SetCellValue { sheet, cell, value } => {
                let sheet = by_name
                    .get(sheet.as_str())
                    .with_context(|| format!("output sheet {sheet:?} missing"))?;
                let actual = sheet.values.get(&normalize_address(cell));
                validate_value(value, actual).with_context(|| format!("verify {cell}"))?;
            }
            EditOperation::SetFormulaText {
                sheet,
                cell,
                formula,
            } => {
                let actual = by_name
                    .get(sheet.as_str())
                    .and_then(|sheet| sheet.formulas.get(&normalize_address(cell)))
                    .with_context(|| {
                        format!("formula {sheet}!{cell} missing after openpyxl edit")
                    })?;
                anyhow::ensure!(
                    actual.trim_start_matches('=').trim() == formula.trim_start_matches('=').trim(),
                    "formula mismatch at {sheet}!{cell}: {actual:?}"
                );
            }
            EditOperation::SetComment {
                sheet, cell, text, ..
            } => {
                let actual = by_name
                    .get(sheet.as_str())
                    .and_then(|sheet| sheet.comments.get(&normalize_address(cell)));
                match (text, actual) {
                    (None, None) => {}
                    (Some(expected), Some(actual)) => anyhow::ensure!(
                        expected == actual,
                        "comment text mismatch at {sheet}!{cell}"
                    ),
                    _ => bail!("comment presence mismatch at {sheet}!{cell}"),
                }
            }
            EditOperation::ReplaceHeaderText { .. } => {
                bail!("ReplaceHeaderText is intentionally routed to the OOXML Patch backend")
            }
            EditOperation::CopyRange {
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                copy_values,
                copy_formulas,
                copy_comments,
                ..
            } => validate_range_copy(
                source_sheets,
                sheets,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                *copy_values,
                *copy_formulas,
                *copy_comments,
                &operations[..operation_index],
            )?,
            EditOperation::ImportRange {
                source_workbook,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                copy_values,
                copy_formulas,
                copy_comments,
                ..
            } => {
                let imported = read_workbook(source_workbook).with_context(|| {
                    format!(
                        "read ImportRange source {} for output validation",
                        source_workbook.display()
                    )
                })?;
                validate_range_copy(
                    &imported.sheets,
                    sheets,
                    source_sheet,
                    source_range,
                    target_sheet,
                    target_start,
                    *copy_values,
                    *copy_formulas,
                    *copy_comments,
                    &[],
                )?;
            }
            _ => {}
        }
    }
    Ok(())
}

fn validate_calculation_properties(
    candidate_output: &Path,
    operations: &[EditOperation],
) -> Result<()> {
    let Some((full_calc_on_load, force_full_calc)) = operations.iter().find_map(|operation| {
        if let EditOperation::SetWorkbookCalculationProperties {
            full_calc_on_load,
            force_full_calc,
        } = operation
        {
            Some((*full_calc_on_load, *force_full_calc))
        } else {
            None
        }
    }) else {
        return Ok(());
    };
    let actual = ooxml_patch::read_workbook_calculation_properties(candidate_output)?;
    if let Some(expected) = full_calc_on_load {
        anyhow::ensure!(
            actual.full_calc_on_load == Some(expected),
            "fullCalcOnLoad readback mismatch: expected {expected}, got {:?}",
            actual.full_calc_on_load
        );
    }
    if let Some(expected) = force_full_calc {
        anyhow::ensure!(
            actual.force_full_calc == Some(expected),
            "forceFullCalc readback mismatch: expected {expected}, got {:?}",
            actual.force_full_calc
        );
    }
    Ok(())
}

fn needs_umya_readback(operation: &EditOperation) -> bool {
    matches!(
        operation,
        EditOperation::SetColumnWidth { .. }
            | EditOperation::SetRowHeight { .. }
            | EditOperation::SetCellFillColor { .. }
            | EditOperation::SetPageSettings { .. }
            | EditOperation::SetCellValue {
                value: crate::CellValue::DateTime { .. },
                ..
            }
            | EditOperation::CopyRange {
                copy_styles: true,
                ..
            }
            | EditOperation::ImportRange {
                style_mode: RangeStyleMode::Basic | RangeStyleMode::Full,
                ..
            }
    )
}

fn validate_umya_readback(plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
    let source = if plan.preserve.contains(&PreservationRequirement::Styles) {
        Some(crate::open_xlsx(&plan.input).context("open source workbook for style readback")?)
    } else {
        None
    };
    let output = crate::open_xlsx(candidate_output)
        .context("open openpyxl output with umya for independent Rust readback")?;

    if let Some(source) = &source {
        validate_unmodified_styles(plan, source, &output)?;
    }

    let expected_styles = expected_import_styles(plan)?;
    for ((sheet_name, address), expected) in expected_styles {
        let actual = style_at(&output, &sheet_name, &address)?;
        anyhow::ensure!(
            styles_semantically_equal(&actual, &expected),
            "Rust style readback mismatch at {sheet_name}!{address}: expected {expected:?}, got {actual:?}"
        );
    }
    validate_copy_style_results(&output, &plan.operations)?;

    let mut row_heights = BTreeMap::new();
    let mut column_widths = BTreeMap::new();
    for operation in &plan.operations {
        match operation {
            EditOperation::SetRowHeight {
                sheet,
                row_index,
                height,
            } => {
                row_heights.insert((sheet.as_str(), *row_index), *height);
            }
            EditOperation::SetColumnWidth {
                sheet,
                column_index,
                width,
            } => {
                column_widths.insert((sheet.as_str(), *column_index), *width);
            }
            _ => {}
        }
    }
    for ((sheet_name, row_index), expected) in row_heights {
        let row = output
            .sheet_by_name(sheet_name)
            .with_context(|| format!("row-height output sheet {sheet_name:?} missing"))?
            .row_dimension(row_index)
            .with_context(|| format!("row {row_index} missing after openpyxl edit"))?;
        anyhow::ensure!(
            (row.height() - expected).abs() <= 1e-10,
            "Rust row-height readback mismatch at {sheet_name}!{row_index}"
        );
    }
    for ((sheet_name, column_index), expected) in column_widths {
        let column = output
            .sheet_by_name(sheet_name)
            .with_context(|| format!("column-width output sheet {sheet_name:?} missing"))?
            .column_dimension_by_number(column_index)
            .with_context(|| format!("column {column_index} missing after openpyxl edit"))?;
        anyhow::ensure!(
            (column.width() - expected).abs() <= 1e-10,
            "Rust column-width readback mismatch at {sheet_name} column {column_index}"
        );
    }

    for operation in &plan.operations {
        match operation {
            EditOperation::SetCellFillColor { sheet, cell, argb } => {
                let expected = normalize_argb_color(argb)?;
                let actual = output
                    .sheet_by_name(sheet)
                    .with_context(|| format!("fill output sheet {sheet:?} missing"))?
                    .get_cell(cell.as_str())
                    .and_then(|cell| cell.style().fill())
                    .and_then(|fill| fill.pattern_fill())
                    .and_then(|fill| fill.foreground_color())
                    .map(|color| color.argb_str().to_ascii_uppercase());
                anyhow::ensure!(
                    actual.as_deref() == Some(expected.as_str()),
                    "Rust fill-color readback mismatch at {sheet}!{cell}"
                );
            }
            EditOperation::SetCellValue {
                sheet,
                cell,
                value: crate::CellValue::DateTime { number_format, .. },
            } => {
                let actual = output
                    .sheet_by_name(sheet)
                    .with_context(|| format!("date-style output sheet {sheet:?} missing"))?
                    .get_cell(cell.as_str())
                    .and_then(|cell| {
                        cell.style()
                            .number_format()
                            .map(|format| format.format_code().to_owned())
                    });
                anyhow::ensure!(
                    actual.as_deref() == Some(number_format.as_str()),
                    "Rust date number-format readback mismatch at {sheet}!{cell}"
                );
            }
            EditOperation::SetPageSettings { sheet, settings } => {
                let worksheet = output
                    .sheet_by_name(sheet)
                    .with_context(|| format!("page-settings output sheet {sheet:?} missing"))?;
                let page = worksheet.page_setup();
                if let Some(orientation) = settings.orientation {
                    let matches = match orientation {
                        crate::PageOrientation::Portrait => matches!(
                            page.orientation(),
                            umya_spreadsheet::OrientationValues::Portrait
                        ),
                        crate::PageOrientation::Landscape => matches!(
                            page.orientation(),
                            umya_spreadsheet::OrientationValues::Landscape
                        ),
                    };
                    anyhow::ensure!(
                        matches,
                        "Rust page orientation readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.fit_to_page {
                    anyhow::ensure!(
                        page.fit_to_page() == value,
                        "Rust fit-to-page readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.fit_to_width {
                    anyhow::ensure!(
                        page.fit_to_width() == u32::from(value),
                        "Rust fit-to-width readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.fit_to_height {
                    anyhow::ensure!(
                        page.fit_to_height() == u32::from(value),
                        "Rust fit-to-height readback mismatch for {sheet:?}"
                    );
                }
                if settings.clear_scale {
                    anyhow::ensure!(
                        page.scale() == 0,
                        "Rust scale-clear readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.paper_size {
                    anyhow::ensure!(
                        page.paper_size() == u32::from(value),
                        "Rust paper-size readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.horizontal_centered {
                    anyhow::ensure!(
                        worksheet.print_options().horizontal_centered() == value,
                        "Rust horizontal-centering readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.vertical_centered {
                    anyhow::ensure!(
                        worksheet.print_options().vertical_centered() == value,
                        "Rust vertical-centering readback mismatch for {sheet:?}"
                    );
                }
                if let Some(expected) = &settings.footer {
                    let actual = worksheet.header_footer().odd_footer().value();
                    anyhow::ensure!(
                        actual == expected || actual == format!("&C{expected}"),
                        "Rust footer readback mismatch for {sheet:?}: expected {expected:?}, got {actual:?}"
                    );
                }
                if let Some(expected) = &settings.print_area {
                    let actual = read_local_print_area(candidate_output, sheet)?
                        .map(|value| canonical_print_area(&value));
                    anyhow::ensure!(
                        actual.as_deref() == Some(canonical_print_area(expected).as_str()),
                        "Rust print-area readback mismatch for {sheet:?}: expected {:?}, got {actual:?}",
                        canonical_print_area(expected)
                    );
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn validate_unmodified_styles(
    plan: &ExcelEditPlan,
    source: &umya_spreadsheet::Workbook,
    output: &umya_spreadsheet::Workbook,
) -> Result<()> {
    let deleted_sheets: BTreeSet<&str> = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::DeleteSheet { name } => Some(name.as_str()),
            _ => None,
        })
        .collect();
    for source_sheet in source.sheet_collection() {
        let source_name = source_sheet.name();
        if deleted_sheets.contains(source_name) {
            continue;
        }
        let output_name = final_sheet_name(plan, source_name);
        let output_sheet = output
            .sheet_by_name(&output_name)
            .with_context(|| format!("style-preservation output sheet {output_name:?} missing"))?;
        let row_map = plan_row_map(plan, source_name)?;
        for source_cell in source_sheet.cells() {
            let address = source_cell.coordinate().to_string();
            let output_address = match &row_map {
                Some(map) => match map.map_address(&address)? {
                    Some(mapped) => mapped,
                    // The plan deletes the row owning this source cell; its
                    // removal is the requested semantics, not a loss.
                    None => continue,
                },
                None => address.clone(),
            };
            if style_is_modified_by_plan(plan, source_name, &output_name, &output_address)? {
                continue;
            }
            let actual = output_sheet
                .cell(output_address.as_str())
                .map(|cell| cell.style().clone())
                .unwrap_or_default();
            // openpyxl normalizes non-anchor cells in merged ranges when it
            // loads/saves a workbook. Validate only border sides on the
            // merged range perimeter here; non-border style comes from the
            // anchor, which is checked independently above. The perimeter
            // check compares shared coordinates, so it only applies when the
            // row mapping leaves this address in place.
            if output_address == address {
                if let Some(equal) =
                    merged_member_effective_borders_equal(source_sheet, output_sheet, &address)
                {
                    anyhow::ensure!(
                        equal,
                        "Rust unmodified merged-cell border readback mismatch at {output_name}!{address}"
                    );
                    continue;
                }
            }
            anyhow::ensure!(
                styles_semantically_equal(&actual, source_cell.style()),
                "Rust unmodified-style readback mismatch at {output_name}!{address}"
            );
        }
    }
    Ok(())
}

fn merged_member_effective_borders_equal(
    source_sheet: &umya_spreadsheet::Worksheet,
    output_sheet: &umya_spreadsheet::Worksheet,
    address: &str,
) -> Option<bool> {
    let Ok((row, column)) = parse_cell(address) else {
        return None;
    };
    let merged = source_sheet.merge_cells().iter().find_map(|merged| {
        let Ok((start_row, start_column, end_row, end_column)) = parse_range(&merged.range())
        else {
            return None;
        };
        (row >= start_row
            && row <= end_row
            && column >= start_column
            && column <= end_column
            && (row != start_row || column != start_column))
            .then_some((start_row, start_column, end_row, end_column))
    })?;

    let anchor_address = format!("{}{}", column_letters(merged.1), merged.0);
    let bottom_right_address = format!("{}{}", column_letters(merged.3), merged.2);
    let source_anchor = source_sheet
        .cell(anchor_address.as_str())
        .map(|cell| cell.style().clone())
        .unwrap_or_default();
    let output_anchor = output_sheet
        .cell(anchor_address.as_str())
        .map(|cell| cell.style().clone())
        .unwrap_or_default();
    let source_bottom_right = source_sheet
        .cell(bottom_right_address.as_str())
        .map(|cell| cell.style().clone())
        .unwrap_or_default();
    let output_bottom_right = output_sheet
        .cell(bottom_right_address.as_str())
        .map(|cell| cell.style().clone())
        .unwrap_or_default();
    let source_anchor = source_anchor.borders().cloned().unwrap_or_default();
    let output_anchor = output_anchor.borders().cloned().unwrap_or_default();
    let source_bottom_right = source_bottom_right.borders().cloned().unwrap_or_default();
    let output_bottom_right = output_bottom_right.borders().cloned().unwrap_or_default();
    let source_bottom = merged_border_side(source_anchor.bottom(), source_bottom_right.bottom());
    let output_bottom = merged_border_side(output_anchor.bottom(), output_bottom_right.bottom());
    let source_right = merged_border_side(source_anchor.right(), source_bottom_right.right());
    let output_right = merged_border_side(output_anchor.right(), output_bottom_right.right());

    // openpyxl builds merged-cell border semantics from the top-left anchor,
    // then fills missing bottom/right edges from the bottom-right cell. An
    // existing anchor side takes precedence over the bottom-right side.
    Some(
        (row != merged.0 || source_anchor.top() == output_anchor.top())
            && (column != merged.1 || source_anchor.left() == output_anchor.left())
            && (row != merged.2 || source_bottom == output_bottom)
            && (column != merged.3 || source_right == output_right),
    )
}

fn merged_border_side<'a>(
    anchor: &'a umya_spreadsheet::Border,
    bottom_right: &'a umya_spreadsheet::Border,
) -> &'a umya_spreadsheet::Border {
    if anchor == &umya_spreadsheet::Border::default() {
        bottom_right
    } else {
        anchor
    }
}

fn final_sheet_name(plan: &ExcelEditPlan, source_name: &str) -> String {
    let mut name = source_name.to_owned();
    for operation in &plan.operations {
        if let EditOperation::RenameSheet { old_name, new_name } = operation {
            if old_name == &name {
                name.clone_from(new_name);
            }
        }
    }
    name
}

/// Sequential source-row → output-row mapping for one worksheet, derived by
/// simulating the plan's DeleteRows/InsertRows operations in application
/// order. Plans author row operations before the cell edits that target the
/// final coordinates; a cell edit authored before a row operation is not
/// remapped.
struct PlanRowMap {
    rows: BTreeMap<u32, Option<u32>>,
    beyond_from: u32,
    beyond_shift: i64,
}

impl PlanRowMap {
    fn output_row(&self, source_row: u32) -> Option<u32> {
        if source_row >= self.beyond_from {
            Some((source_row as i64 + self.beyond_shift) as u32)
        } else {
            self.rows.get(&source_row).copied().flatten()
        }
    }

    /// `Ok(None)` means the plan deletes the row owning this source address.
    fn map_address(&self, address: &str) -> Result<Option<String>> {
        let (row, column) = parse_cell(address)?;
        Ok(self
            .output_row(row)
            .map(|row| format!("{}{}", column_letters(column), row)))
    }
}

/// Build the per-sheet row mapping for the unmodified-readback comparisons.
/// Returns `None` when the plan has no row operations for the worksheet so
/// callers keep the identity fast path.
fn plan_row_map(plan: &ExcelEditPlan, source_sheet: &str) -> Result<Option<PlanRowMap>> {
    let matches_sheet =
        |sheet: &str| sheet == source_sheet || sheet == final_sheet_name(plan, source_sheet);
    let row_operations: Vec<(&u32, &u32, bool)> = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::DeleteRows {
                sheet,
                first_row,
                count,
            } if matches_sheet(sheet) => Some((first_row, count, true)),
            EditOperation::InsertRows {
                sheet,
                first_row,
                count,
            } if matches_sheet(sheet) => Some((first_row, count, false)),
            _ => None,
        })
        .collect();
    if row_operations.is_empty() {
        return Ok(None);
    }
    // rows[i] models current worksheet row i+1: Some(source row) or None for
    // a row the plan inserted. Source rows below the simulated window shift
    // uniformly and resolve through `beyond_shift`.
    let mut rows: Vec<Option<u32>> = Vec::new();
    let mut next_source_row: u32 = 1;
    for (first_row, count, delete) in row_operations {
        let first = *first_row as usize;
        let count = *count as usize;
        if delete {
            let end = first
                .checked_add(count)
                .and_then(|end| end.checked_sub(1))
                .context("row deletion range overflows")?;
            while rows.len() < end {
                rows.push(Some(next_source_row));
                next_source_row += 1;
            }
            rows.drain(first - 1..first - 1 + count);
        } else {
            while rows.len() < first - 1 {
                rows.push(Some(next_source_row));
                next_source_row += 1;
            }
            rows.splice(first - 1..first - 1, std::iter::repeat(None).take(count));
        }
    }
    let highest_placed = next_source_row - 1;
    let mut mapped = BTreeMap::new();
    for (index, entry) in rows.iter().enumerate() {
        if let Some(source_row) = entry {
            mapped.insert(*source_row, Some(index as u32 + 1));
        }
    }
    for source_row in 1..=highest_placed {
        mapped.entry(source_row).or_insert(None);
    }
    let beyond_shift = rows.len() as i64 - i64::from(highest_placed);
    Ok(Some(PlanRowMap {
        rows: mapped,
        beyond_from: highest_placed + 1,
        beyond_shift,
    }))
}

fn style_is_modified_by_plan(
    plan: &ExcelEditPlan,
    source_sheet_name: &str,
    output_sheet_name: &str,
    address: &str,
) -> Result<bool> {
    let sheet_matches = |sheet: &str| sheet == source_sheet_name || sheet == output_sheet_name;
    let address = normalize_address(address);
    for operation in &plan.operations {
        match operation {
            EditOperation::SetCellFillColor { sheet, cell, .. }
                if sheet_matches(sheet) && normalize_address(cell) == address =>
            {
                return Ok(true);
            }
            EditOperation::SetCellValue {
                sheet,
                cell,
                value: crate::CellValue::DateTime { .. },
            } if sheet_matches(sheet) && normalize_address(cell) == address => return Ok(true),
            EditOperation::CopyRange {
                target_sheet,
                source_range,
                target_start,
                copy_styles: true,
                ..
            } if sheet_matches(target_sheet)
                && target_range_contains(source_range, target_start, &address)? =>
            {
                return Ok(true);
            }
            EditOperation::ImportRange {
                target_sheet,
                source_range,
                target_start,
                style_mode: RangeStyleMode::Basic | RangeStyleMode::Full,
                ..
            } if sheet_matches(target_sheet)
                && target_range_contains(source_range, target_start, &address)? =>
            {
                return Ok(true);
            }
            _ => {}
        }
    }
    Ok(false)
}

fn target_range_contains(source_range: &str, target_start: &str, address: &str) -> Result<bool> {
    let (source_min_row, source_min_column, source_max_row, source_max_column) =
        parse_range(source_range)?;
    let (target_row, target_column) = parse_cell(target_start)?;
    let (row, column) = parse_cell(address)?;
    let target_max_row = target_row + source_max_row - source_min_row;
    let target_max_column = target_column + source_max_column - source_min_column;
    Ok((target_row..=target_max_row).contains(&row)
        && (target_column..=target_max_column).contains(&column))
}

fn read_local_print_area(path: &Path, sheet_name: &str) -> Result<Option<String>> {
    const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let mut archive = zip::ZipArchive::new(File::open(path)?)?;
    let mut workbook_xml = String::new();
    archive
        .by_name("xl/workbook.xml")?
        .read_to_string(&mut workbook_xml)?;
    let document = roxmltree::Document::parse(&workbook_xml)?;
    let sheets = document
        .descendants()
        .filter(|node| {
            node.tag_name().namespace() == Some(MAIN_NS) && node.tag_name().name() == "sheet"
        })
        .collect::<Vec<_>>();
    let sheet_index = sheets
        .iter()
        .position(|node| node.attribute("name") == Some(sheet_name))
        .context("find print-area worksheet in workbook.xml")?;
    let sheet_index = sheet_index.to_string();

    Ok(document
        .descendants()
        .find(|node| {
            node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "definedName"
                && node
                    .attribute("name")
                    .is_some_and(|name| name.eq_ignore_ascii_case("_xlnm.Print_Area"))
                && node.attribute("localSheetId") == Some(sheet_index.as_str())
        })
        .and_then(|node| node.text())
        .map(str::to_owned))
}

fn expected_import_styles(
    plan: &ExcelEditPlan,
) -> Result<BTreeMap<(String, String), umya_spreadsheet::Style>> {
    let mut styles = BTreeMap::new();
    let mut normalized_import_sources: BTreeMap<PathBuf, NormalizedImportStyleSource> =
        BTreeMap::new();
    for operation in &plan.operations {
        let EditOperation::ImportRange {
            source_workbook: import_path,
            source_sheet,
            source_range,
            target_sheet,
            target_start,
            style_mode,
            ..
        } = operation
        else {
            continue;
        };
        if *style_mode == RangeStyleMode::None {
            continue;
        }
        if !normalized_import_sources.contains_key(import_path) {
            let normalized = ScratchFile::create(&std::env::temp_dir(), "style-source-xlsx")?;
            let probe_cells = materialize_implicit_style_zero(import_path, normalized.path())?;
            normalized_import_sources.insert(
                import_path.clone(),
                NormalizedImportStyleSource {
                    file: normalized,
                    probe_cells,
                },
            );
        }
        let normalized_source = normalized_import_sources
            .get(import_path)
            .context("normalized ImportRange style source missing")?;
        let normalized_path = normalized_source.file.path();
        let imported = crate::open_xlsx(normalized_path)
            .with_context(|| format!("open ImportRange styles from {}", import_path.display()))?;
        let (min_row, min_col, max_row, max_col) = parse_range(source_range)?;
        let (target_row, target_col) = parse_cell(target_start)?;
        for row in min_row..=max_row {
            for column in min_col..=max_col {
                let source_cell = format!("{}{}", column_letters(column), row);
                let target_cell = format!(
                    "{}{}",
                    column_letters(target_col + column - min_col),
                    target_row + row - min_row
                );
                let worksheet = imported.sheet_by_name(source_sheet).with_context(|| {
                    format!("ImportRange source sheet {source_sheet:?} missing")
                })?;
                let style = if worksheet.cell(source_cell.as_str()).is_none() {
                    let probe = normalized_source
                        .probe_cells
                        .get(source_sheet)
                        .with_context(|| {
                            format!("XF0 probe for source sheet {source_sheet:?} missing")
                        })?;
                    style_at(&imported, source_sheet, probe)?
                } else {
                    style_at(&imported, source_sheet, &source_cell)?
                };
                styles.insert(
                    (target_sheet.clone(), target_cell),
                    project_import_style(&style, *style_mode),
                );
            }
        }
    }
    Ok(styles)
}

/// Umya applies a cell style only when a worksheet cell has an explicit `s`
/// attribute. OOXML defines an omitted `s` as style index 0, which can itself
/// contain non-default properties. Materialize that implicit index only in a
/// temporary readback copy so expected ImportRange styles include XF0. For
/// absent source cells the probe reads the direct cell style used by the
/// openpyxl worker; row/column dimensions remain separate worksheet objects.
struct NormalizedImportStyleSource {
    file: ScratchFile,
    probe_cells: BTreeMap<String, String>,
}

fn materialize_implicit_style_zero(
    source: &Path,
    output: &Path,
) -> Result<BTreeMap<String, String>> {
    let worksheet_parts = worksheet_parts_by_name(source)?;
    let mut input = zip::ZipArchive::new(File::open(source)?)?;
    let mut writer = zip::ZipWriter::new(File::create(output)?);
    let mut probe_cells = BTreeMap::new();
    for index in 0..input.len() {
        let mut entry = input.by_index(index)?;
        let name = entry.name().to_owned();
        let is_directory = entry.is_dir();
        let compression = entry.compression();
        let options = zip::write::SimpleFileOptions::default().compression_method(compression);

        if is_directory {
            writer.add_directory(name, options)?;
            continue;
        }

        writer.start_file(name.clone(), options)?;
        if worksheet_parts.values().any(|part| part == &name) {
            let mut xml = Vec::new();
            entry.read_to_end(&mut xml)?;
            let xml = std::str::from_utf8(&xml)
                .with_context(|| format!("worksheet part {name} is not UTF-8"))?;
            let normalized = materialize_implicit_style_zero_cells(xml)?;
            let (normalized, probe) = append_xf_zero_probe(&normalized)?;
            let sheet_name = worksheet_parts
                .iter()
                .find_map(|(sheet, part)| (part == &name).then_some(sheet.clone()))
                .context("worksheet part has no workbook sheet identity")?;
            probe_cells.insert(sheet_name, probe);
            writer.write_all(normalized.as_bytes())?;
        } else {
            std::io::copy(&mut entry, &mut writer)?;
        }
    }
    writer.finish()?;
    Ok(probe_cells)
}

fn worksheet_parts_by_name(source: &Path) -> Result<BTreeMap<String, String>> {
    const DOCUMENT_REL_NS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    let mut archive = zip::ZipArchive::new(File::open(source)?)?;
    let workbook_xml = read_zip_part(&mut archive, "xl/workbook.xml")?;
    let relationships_xml = read_zip_part(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let workbook = roxmltree::Document::parse(&workbook_xml)?;
    let relationships = roxmltree::Document::parse(&relationships_xml)?;
    let targets = relationships
        .root_element()
        .children()
        .filter(|node| node.is_element() && node.tag_name().namespace() == Some(PACKAGE_REL_NS))
        .filter_map(|node| {
            Some((
                node.attribute("Id")?.to_owned(),
                (
                    node.attribute("Target")?.to_owned(),
                    node.attribute("TargetMode").unwrap_or_default().to_owned(),
                ),
            ))
        })
        .collect::<BTreeMap<_, _>>();
    let mut parts = BTreeMap::new();
    for sheet in workbook.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "sheet"
    }) {
        let name = sheet.attribute("name").context("sheet name missing")?;
        let id = sheet
            .attributes()
            .find(|attribute| {
                attribute.name() == "id" && attribute.namespace() == Some(DOCUMENT_REL_NS)
            })
            .map(|attribute| attribute.value())
            .context("sheet relationship id missing")?;
        let (target, mode) = targets.get(id).context("sheet relationship missing")?;
        anyhow::ensure!(
            mode.is_empty(),
            "external worksheet relationship is unsupported"
        );
        parts.insert(name.to_owned(), resolve_opc_target("xl", target)?);
    }
    Ok(parts)
}

/// Openpyxl's normal worksheet loader removes every non-anchor cell in a
/// merged range. Check target workbooks globally; for ImportRange sources,
/// check only the selected sheet and source range whose payload will be read.
fn merged_non_anchor_payloads(
    source: &Path,
    scope: Option<(&str, &str)>,
) -> Result<Vec<(String, String)>> {
    let parts = worksheet_parts_by_name(source)?;
    if let Some((source_sheet, source_range)) = scope {
        anyhow::ensure!(
            parts
                .keys()
                .any(|name| name.eq_ignore_ascii_case(source_sheet)),
            "ImportRange source sheet {source_sheet:?} is missing"
        );
        parse_range(source_range).context("parse ImportRange source range")?;
    }
    let mut archive = zip::ZipArchive::new(File::open(source)?)?;
    let mut candidates = Vec::new();
    for (sheet, part) in &parts {
        if scope.is_some_and(|(source_sheet, _)| !sheet.eq_ignore_ascii_case(source_sheet)) {
            continue;
        }
        let xml = read_zip_part(&mut archive, part)?;
        let sheet_candidates = non_anchor_payload_cells(sheet, &xml)?;
        if let Some((source_sheet, source_range)) = scope {
            candidates.extend(merged_payload_cells_in_range(
                sheet_candidates,
                source_sheet,
                source_range,
            )?);
        } else {
            candidates.extend(sheet_candidates);
        }
    }
    if candidates.is_empty() {
        return Ok(Vec::new());
    }

    // Use the Rust reader to distinguish actual values from empty styled cells
    // that happen to contain an empty <v> or <is> element.
    let workbook = read_workbook(source).context("read cells identified in merged ranges")?;
    let mut payloads = Vec::new();
    for (sheet_name, address) in candidates {
        let Some(sheet) = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == sheet_name)
        else {
            continue;
        };
        let address = normalize_address(&address);
        if sheet.values.contains_key(&address) || sheet.formulas.contains_key(&address) {
            payloads.push((sheet_name, address));
        }
    }
    Ok(payloads)
}

/// First merged non-anchor payload of the target workbook the plan still has
/// to protect: worksheets with an explicit `MergedNonAnchorPayloads` allowance
/// may drop payloads outside every operation target rectangle (openpyxl
/// rewrite loss, baseline-equivalent), while payloads inside a rectangle the
/// plan writes into stay protected, as do payloads on worksheets without the
/// allowance.
fn offending_target_merged_payload(
    plan: &ExcelEditPlan,
    payloads: Vec<(String, String)>,
) -> Option<(String, String)> {
    let target_rectangles: Vec<(&str, &str, &str)> = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::CopyRange {
                target_sheet,
                source_range,
                target_start,
                ..
            }
            | EditOperation::ImportRange {
                target_sheet,
                source_range,
                target_start,
                ..
            } => Some((
                target_sheet.as_str(),
                source_range.as_str(),
                target_start.as_str(),
            )),
            _ => None,
        })
        .collect();
    for (sheet, address) in payloads {
        let allowance = plan.drop_allowed.get(&sheet).is_some_and(|allowed| {
            allowed.contains(&PreservationRequirement::MergedNonAnchorPayloads)
        });
        if !allowance {
            return Some((sheet, address));
        }
        let inside_operation_rectangle =
            target_rectangles
                .iter()
                .any(|(target_sheet, source_range, target_start)| {
                    *target_sheet == sheet.as_str()
                        && target_range_contains(source_range, target_start, &address)
                            .unwrap_or(false)
                });
        if inside_operation_rectangle {
            return Some((sheet, address));
        }
    }
    None
}

fn merged_payload_cells_in_range(
    candidates: Vec<(String, String)>,
    source_sheet: &str,
    source_range: &str,
) -> Result<Vec<(String, String)>> {
    let (range_row_1, range_col_1, range_row_2, range_col_2) = parse_range(source_range)?;
    let min_row = range_row_1.min(range_row_2);
    let max_row = range_row_1.max(range_row_2);
    let min_col = range_col_1.min(range_col_2);
    let max_col = range_col_1.max(range_col_2);
    let mut in_range = Vec::new();
    for (sheet, address) in candidates {
        if !sheet.eq_ignore_ascii_case(source_sheet) {
            continue;
        }
        let (row, column) = parse_cell(&address)?;
        if row >= min_row && row <= max_row && column >= min_col && column <= max_col {
            in_range.push((sheet, address));
        }
    }
    Ok(in_range)
}

fn non_anchor_payload_cells(sheet: &str, xml: &str) -> Result<Vec<(String, String)>> {
    let document = roxmltree::Document::parse(xml)?;
    let root = document.root_element();
    anyhow::ensure!(
        root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && root.tag_name().name() == "worksheet",
        "worksheet part has an unexpected root element"
    );
    let merged_ranges = root
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "mergeCell"
        })
        .map(|node| {
            let reference = node
                .attribute("ref")
                .context("merged range has no reference")?;
            let (row1, col1, row2, col2) = parse_range(reference)?;
            Ok((
                row1.min(row2),
                col1.min(col2),
                row1.max(row2),
                col1.max(col2),
            ))
        })
        .collect::<Result<Vec<_>>>()?;
    if merged_ranges.is_empty() {
        return Ok(Vec::new());
    }

    let mut candidates = Vec::new();
    for cell in root.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "c"
    }) {
        let Some(address) = cell.attribute("r") else {
            continue;
        };
        let (row, column) = parse_cell(address)?;
        let Some((min_row, min_col, _, _)) = merged_ranges.iter().find(|range| {
            row >= range.0 && row <= range.2 && column >= range.1 && column <= range.3
        }) else {
            continue;
        };
        if row == *min_row && column == *min_col {
            continue;
        }
        let has_payload = cell.children().any(|child| {
            if !child.is_element() || child.tag_name().namespace() != Some(SPREADSHEET_MAIN_NS) {
                return false;
            }
            match child.tag_name().name() {
                "f" => true,
                "v" => child.text().is_some_and(|text| !text.is_empty()),
                "is" => child.descendants().any(|text| {
                    text.is_element()
                        && text.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                        && text.tag_name().name() == "t"
                        && text.text().is_some_and(|text| !text.is_empty())
                }),
                _ => false,
            }
        });
        if has_payload {
            candidates.push((sheet.to_owned(), normalize_address(address)));
        }
    }
    Ok(candidates)
}

fn append_xf_zero_probe(xml: &str) -> Result<(String, String)> {
    let document = roxmltree::Document::parse(xml)?;
    let sheet_data = document
        .descendants()
        .find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "sheetData"
        })
        .context("worksheet has no sheetData")?;
    let rows = sheet_data
        .children()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "row"
        })
        .collect::<Vec<_>>();
    let used = rows
        .iter()
        .filter_map(|row| row.attribute("r")?.parse::<u32>().ok())
        .collect::<BTreeSet<_>>();
    let probe_row = (1..=1_048_576)
        .rev()
        .find(|row| !used.contains(row))
        .context("worksheet has no free row for temporary XF0 probe")?;
    let probe = format!("A{probe_row}");
    let insertion = if let Some(row) = rows.iter().find(|row| {
        row.attribute("r")
            .and_then(|value| value.parse::<u32>().ok())
            .is_some_and(|number| number > probe_row)
    }) {
        row.range().start
    } else {
        let range = sheet_data.range();
        let fragment = &xml[range.start..range.end];
        range.start
            + fragment
                .rfind("</")
                .context("sheetData is not an explicit XML element")?
    };
    let injected = format!(
        "<row xmlns=\"{SPREADSHEET_MAIN_NS}\" r=\"{probe_row}\"><c r=\"{probe}\" s=\"0\"/></row>"
    );
    let mut normalized = xml.to_owned();
    normalized.insert_str(insertion, &injected);
    Ok((normalized, probe))
}

fn materialize_implicit_style_zero_cells(xml: &str) -> Result<String> {
    let document = roxmltree::Document::parse(xml)?;
    let root = document.root_element();
    if root.tag_name().namespace() != Some(SPREADSHEET_MAIN_NS)
        || root.tag_name().name() != "worksheet"
    {
        return Ok(xml.to_owned());
    }

    let mut insertions = Vec::new();
    let bytes = xml.as_bytes();
    for cell in document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "c"
            && node.attribute("s").is_none()
    }) {
        let start = cell.range().start;
        let opening = &bytes[start..];
        let tag_length =
            xml_start_tag_length(opening).context("find worksheet cell start-tag boundary")?;
        let tag_end = start + tag_length;
        let mut insertion = tag_end - 1; // Before the closing '>'.
        let mut previous = insertion;
        while previous > start && bytes[previous - 1].is_ascii_whitespace() {
            previous -= 1;
        }
        if previous > start && bytes[previous - 1] == b'/' {
            insertion = previous - 1; // Preserve the `/>` delimiter.
        }
        insertions.push(insertion);
    }

    if insertions.is_empty() {
        return Ok(xml.to_owned());
    }
    insertions.sort_unstable_by(|left, right| right.cmp(left));
    let mut patched = bytes.to_vec();
    for insertion in insertions {
        patched.splice(insertion..insertion, b" s=\"0\"".iter().copied());
    }
    String::from_utf8(patched).context("patched worksheet XML is not UTF-8")
}

fn xml_start_tag_length(bytes: &[u8]) -> Option<usize> {
    let mut quote = None;
    for (index, byte) in bytes.iter().copied().enumerate() {
        match quote {
            Some(delimiter) if byte == delimiter => quote = None,
            Some(_) => {}
            None if byte == b'\'' || byte == b'"' => quote = Some(byte),
            None if byte == b'>' => return Some(index + 1),
            None => {}
        }
    }
    None
}

fn validate_copy_style_results(
    output: &umya_spreadsheet::Workbook,
    operations: &[EditOperation],
) -> Result<()> {
    for operation in operations {
        let EditOperation::CopyRange {
            source_sheet,
            source_range,
            target_sheet,
            target_start,
            copy_styles: true,
            ..
        } = operation
        else {
            continue;
        };
        let (min_row, min_col, max_row, max_col) = parse_range(source_range)?;
        let (target_row, target_col) = parse_cell(target_start)?;
        for row in min_row..=max_row {
            for column in min_col..=max_col {
                let source_cell = format!("{}{}", column_letters(column), row);
                let target_cell = format!(
                    "{}{}",
                    column_letters(target_col + column - min_col),
                    target_row + row - min_row
                );
                let source_style = style_at(output, source_sheet, &source_cell)?;
                let target_style = style_at(output, target_sheet, &target_cell)?;
                anyhow::ensure!(
                    source_style == target_style,
                    "Rust CopyRange style readback mismatch at {target_sheet}!{target_cell}"
                );
            }
        }
    }
    Ok(())
}

fn style_at(
    workbook: &umya_spreadsheet::Workbook,
    sheet: &str,
    address: &str,
) -> Result<umya_spreadsheet::Style> {
    let worksheet = workbook
        .sheet_by_name(sheet)
        .with_context(|| format!("style worksheet {sheet:?} missing"))?;
    Ok(worksheet
        .get_cell(address)
        .map(|cell| cell.style().clone())
        .unwrap_or_default())
}

fn project_import_style(
    source: &umya_spreadsheet::Style,
    mode: RangeStyleMode,
) -> umya_spreadsheet::Style {
    let mut target = umya_spreadsheet::Style::default();
    if let Some(value) = source.font() {
        target.set_font(value.clone());
    }
    if let Some(value) = source.fill() {
        target.set_fill(value.clone());
    }
    if let Some(value) = source.alignment() {
        target.set_alignment(value.clone());
    }
    if let Some(value) = source.number_format() {
        target.set_number_format(value.clone());
    }
    if mode == RangeStyleMode::Full {
        if let Some(value) = source.borders() {
            target.set_borders(value.clone());
        }
        if let Some(value) = source.protection() {
            target.set_protection(value.clone());
        }
    }
    target
}

fn styles_semantically_equal(
    actual: &umya_spreadsheet::Style,
    expected: &umya_spreadsheet::Style,
) -> bool {
    fn equal_defaulted<T>(actual: Option<&T>, expected: Option<&T>) -> bool
    where
        T: Clone + Default + PartialEq,
    {
        actual.cloned().unwrap_or_default() == expected.cloned().unwrap_or_default()
    }

    fn fill_is_visually_empty(fill: &umya_spreadsheet::Fill) -> bool {
        fill.gradient_fill().is_none()
            && fill.pattern_fill().is_none_or(|pattern| {
                pattern.pattern_type() == &umya_spreadsheet::PatternValues::None
                    && pattern.foreground_color().is_none()
                    && pattern.background_color().is_none()
            })
    }

    fn equal_fill(
        actual: Option<&umya_spreadsheet::Fill>,
        expected: Option<&umya_spreadsheet::Fill>,
    ) -> bool {
        let actual = actual.cloned().unwrap_or_default();
        let expected = expected.cloned().unwrap_or_default();

        if actual.gradient_fill() != expected.gradient_fill() {
            return false;
        }

        match (actual.pattern_fill(), expected.pattern_fill()) {
            (Some(actual), Some(expected)) => {
                actual.pattern_type() == expected.pattern_type()
                    && actual.foreground_color() == expected.foreground_color()
                    && actual.background_color() == expected.background_color()
            }
            (None, None) => true,
            (None, Some(_)) => fill_is_visually_empty(&expected),
            (Some(_), None) => fill_is_visually_empty(&actual),
        }
    }

    equal_defaulted(actual.font(), expected.font())
        && equal_fill(actual.fill(), expected.fill())
        && equal_defaulted(actual.borders(), expected.borders())
        && equal_defaulted(actual.alignment(), expected.alignment())
        && equal_defaulted(actual.protection(), expected.protection())
        && actual
            .number_format()
            .map(|format| format.format_code())
            .unwrap_or("General")
            == expected
                .number_format()
                .map(|format| format.format_code())
                .unwrap_or("General")
}

fn normalize_argb_color(value: &str) -> Result<String> {
    let mut value = value.trim().trim_start_matches('#').to_ascii_uppercase();
    if value.len() == 6 {
        value.insert_str(0, "FF");
    }
    anyhow::ensure!(
        value.len() == 8 && value.bytes().all(|byte| byte.is_ascii_hexdigit()),
        "fill color must be RGB or ARGB hexadecimal"
    );
    Ok(value)
}

fn canonical_print_area(value: &str) -> String {
    value
        .rsplit('!')
        .next()
        .unwrap_or(value)
        .replace('$', "")
        .replace('\'', "")
        .to_ascii_uppercase()
}

fn validate_range_copy(
    source_sheets: &[SheetDump],
    output_sheets: &[SheetDump],
    source_sheet_name: &str,
    source_range: &str,
    target_sheet_name: &str,
    target_start: &str,
    copy_values: bool,
    copy_formulas: bool,
    copy_comments: bool,
    preceding_operations: &[EditOperation],
) -> Result<()> {
    let source = source_sheets
        .iter()
        .find(|sheet| sheet.name == source_sheet_name)
        .with_context(|| format!("range source sheet {source_sheet_name:?} missing"))?;
    let target = output_sheets
        .iter()
        .find(|sheet| sheet.name == target_sheet_name)
        .with_context(|| format!("range target sheet {target_sheet_name:?} missing"))?;
    let (min_row, min_col, max_row, max_col) = parse_range(source_range)?;
    let (target_row, target_col) = parse_cell(target_start)?;
    for row in min_row..=max_row {
        for column in min_col..=max_col {
            let source_address = format!("{}{}", column_letters(column), row);
            let target_address = format!(
                "{}{}",
                column_letters(target_col + column - min_col),
                target_row + row - min_row
            );
            let mut planned_value = None;
            let mut planned_formula = None;
            let mut planned_comment = None;
            for operation in preceding_operations {
                match operation {
                    EditOperation::SetCellValue { sheet, cell, value }
                        if sheet == source_sheet_name
                            && normalize_address(cell) == source_address =>
                    {
                        planned_value = Some(value.clone());
                        planned_formula = Some(None);
                    }
                    EditOperation::SetFormulaText {
                        sheet,
                        cell,
                        formula,
                    } if sheet == source_sheet_name
                        && normalize_address(cell) == source_address =>
                    {
                        planned_value = Some(crate::CellValue::Blank);
                        planned_formula = Some(Some(formula.clone()));
                    }
                    EditOperation::SetComment {
                        sheet, cell, text, ..
                    } if sheet == source_sheet_name
                        && normalize_address(cell) == source_address =>
                    {
                        planned_comment = Some(text.clone());
                    }
                    _ => {}
                }
            }
            let source_formula = match &planned_formula {
                Some(formula) => formula.as_ref(),
                None => source.formulas.get(&source_address),
            };
            if copy_formulas && source_formula.is_some() {
                anyhow::ensure!(
                    target.formulas.contains_key(&target_address),
                    "copied formula is missing at {target_sheet_name}!{target_address}"
                );
            } else if copy_values {
                let actual = target.values.get(&target_address);
                if let Some(expected) = planned_value {
                    validate_planned_copy_value(&expected, actual).with_context(|| {
                        format!("copied cell value/type mismatch at {target_sheet_name}!{target_address}")
                    })?;
                } else {
                    let expected = source.values.get(&source_address);
                    anyhow::ensure!(
                        match (expected, actual) {
                            (None, None) => true,
                            (Some(expected), Some(actual)) =>
                                json_value_equal(expected, Some(actual)),
                            _ => false,
                        },
                        "copied cell value/type mismatch at {target_sheet_name}!{target_address}"
                    );
                }
            }
            if copy_comments {
                let expected = match &planned_comment {
                    Some(comment) => comment.as_deref(),
                    None => source.comments.get(&source_address).map(String::as_str),
                };
                let actual = target.comments.get(&target_address);
                anyhow::ensure!(
                    expected == actual.map(String::as_str),
                    "copied comment mismatch at {target_sheet_name}!{target_address}"
                );
            }
        }
    }
    Ok(())
}

fn validate_planned_copy_value(
    expected: &crate::CellValue,
    actual: Option<&serde_json::Value>,
) -> Result<()> {
    use crate::CellValue;
    match expected {
        CellValue::Blank => anyhow::ensure!(actual.is_none(), "expected blank copied cell"),
        CellValue::Text(value) => anyhow::ensure!(
            actual
                .and_then(|value| value.get("value"))
                .and_then(serde_json::Value::as_str)
                == Some(value),
            "copied text value mismatch"
        ),
        CellValue::Boolean(value) => anyhow::ensure!(
            actual
                .and_then(|value| value.get("value"))
                .and_then(serde_json::Value::as_bool)
                == Some(*value),
            "copied boolean value mismatch"
        ),
        CellValue::Error(value) => anyhow::ensure!(
            actual
                .and_then(|value| value.get("value"))
                .and_then(serde_json::Value::as_str)
                == Some(value),
            "copied error value mismatch"
        ),
        CellValue::Decimal(value) | CellValue::DateTime { serial: value, .. } => {
            let expected = value.parse::<f64>().context("parse copied numeric value")?;
            let actual = actual
                .and_then(|value| value.get("value"))
                .and_then(serde_json::Value::as_f64)
                .context("copied numeric value is missing")?;
            let scale = expected.abs().max(actual.abs()).max(1.0);
            anyhow::ensure!(
                (expected - actual).abs() <= scale * 1e-15,
                "copied numeric value mismatch"
            );
        }
    }
    Ok(())
}

fn validate_value(expected: &crate::CellValue, actual: Option<&serde_json::Value>) -> Result<()> {
    use crate::CellValue;
    match expected {
        CellValue::Blank => anyhow::ensure!(actual.is_none(), "expected blank cell"),
        CellValue::Text(value) => anyhow::ensure!(
            actual
                .and_then(|v| v.get("value"))
                .and_then(serde_json::Value::as_str)
                == Some(value),
            "text value mismatch"
        ),
        CellValue::Boolean(value) => anyhow::ensure!(
            actual
                .and_then(|v| v.get("value"))
                .and_then(serde_json::Value::as_bool)
                == Some(*value),
            "boolean value mismatch"
        ),
        CellValue::Error(value) => anyhow::ensure!(
            actual
                .and_then(|v| v.get("value"))
                .and_then(serde_json::Value::as_str)
                == Some(value),
            "error value mismatch"
        ),
        CellValue::Decimal(value) | CellValue::DateTime { serial: value, .. } => {
            let expected = value
                .parse::<f64>()
                .context("parse expected numeric cell")?;
            let actual = actual
                .and_then(|v| v.get("value"))
                .and_then(serde_json::Value::as_f64)
                .with_context(|| "numeric cell missing or has another type")?;
            let scale = expected.abs().max(actual.abs()).max(1.0);
            anyhow::ensure!(
                (expected - actual).abs() <= scale * 1e-15,
                "numeric cell differs beyond 1e-15 relative tolerance"
            );
        }
    }
    Ok(())
}

/// The merge range covering a non-anchor cell, if any.
fn non_anchor_merge_containing(sheet: &SheetDump, address: &str) -> Option<String> {
    let (row, column) = parse_cell(address).ok()?;
    sheet
        .merges
        .iter()
        .find(|merge| {
            let Ok((start_row, start_column, end_row, end_column)) = parse_range(merge) else {
                return false;
            };
            row >= start_row
                && row <= end_row
                && column >= start_column
                && column <= end_column
                && (row != start_row || column != start_column)
        })
        .cloned()
}

/// Baseline-equivalence exemption for merged non-anchor payloads: on a
/// worksheet with an explicit `MergedNonAnchorPayloads` allowance, a source
/// cell hidden under a merged range whose range string is unchanged may
/// surface as empty after the openpyxl rewrite, matching the legacy Python
/// baseline. The strict whole-sheet merge comparison stays in
/// `validate_preserved_merges`; this only exempts the emptied payload cells.
fn merged_non_anchor_payload_loss_exempt(
    plan: &ExcelEditPlan,
    source_sheet: &SheetDump,
    output_sheet: &SheetDump,
    address: &str,
) -> bool {
    plan.drop_allowed
        .get(&source_sheet.name)
        .is_some_and(|allowed| allowed.contains(&PreservationRequirement::MergedNonAnchorPayloads))
        && non_anchor_merge_containing(source_sheet, address)
            .is_some_and(|range| output_sheet.merges.contains(&range))
}

fn validate_preserved_values(
    source: &[SheetDump],
    output: &[SheetDump],
    plan: &ExcelEditPlan,
) -> Result<()> {
    if !plan
        .preserve
        .contains(&PreservationRequirement::CellValuesAndTypes)
    {
        return Ok(());
    }
    let renamed: BTreeMap<String, String> = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::RenameSheet { old_name, new_name } => {
                Some((old_name.clone(), new_name.clone()))
            }
            _ => None,
        })
        .collect();
    let deleted: BTreeSet<String> = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::DeleteSheet { name } => Some(name.clone()),
            _ => None,
        })
        .collect();
    for old_sheet in source {
        if deleted.contains(&old_sheet.name) {
            continue;
        }
        let output_name = renamed.get(&old_sheet.name).unwrap_or(&old_sheet.name);
        let Some(new_sheet) = output.iter().find(|sheet| &sheet.name == output_name) else {
            continue;
        };
        let row_map = plan_row_map(plan, &old_sheet.name)?;
        for (address, value) in &old_sheet.values {
            let output_address = match &row_map {
                Some(map) => match map.map_address(address)? {
                    Some(mapped) => mapped,
                    // The plan deletes the row owning this source cell.
                    None => continue,
                },
                None => address.clone(),
            };
            if cell_value_is_modified(&old_sheet.name, &output_address, &plan.operations) {
                continue;
            }
            let actual = new_sheet.values.get(&output_address);
            if actual.is_none()
                && output_address == *address
                && merged_non_anchor_payload_loss_exempt(plan, old_sheet, new_sheet, address)
            {
                continue;
            }
            anyhow::ensure!(
                json_value_equal(value, actual),
                "unmodified cell value changed at {}!{}: expected {:?}, got {:?}",
                old_sheet.name,
                address,
                value,
                actual
            );
        }
    }
    Ok(())
}

fn validate_preserved_formulas(
    source: &[SheetDump],
    output: &[SheetDump],
    plan: &ExcelEditPlan,
) -> Result<()> {
    if !plan
        .preserve
        .contains(&PreservationRequirement::FormulaText)
    {
        return Ok(());
    }
    let (renamed, deleted) = sheet_name_changes(&plan.operations);
    for old_sheet in source {
        if deleted.contains(&old_sheet.name) {
            continue;
        }
        let output_name = renamed.get(&old_sheet.name).unwrap_or(&old_sheet.name);
        let Some(new_sheet) = output.iter().find(|sheet| &sheet.name == output_name) else {
            continue;
        };
        let row_map = plan_row_map(plan, &old_sheet.name)?;
        for (address, formula) in &old_sheet.formulas {
            let output_address = match &row_map {
                Some(map) => match map.map_address(address)? {
                    Some(mapped) => mapped,
                    // The plan deletes the row owning this source formula.
                    None => continue,
                },
                None => address.clone(),
            };
            if output_address != *address {
                // Row insertion/deletion rewrites the references inside
                // shifted formulas; the engine-neutral validator cannot
                // recompute that rewrite, so shifted formulas are not
                // text-compared here. They stay covered by the engine's own
                // formula adjustment plus the requesting side's readback.
                continue;
            }
            if cell_formula_is_modified(&old_sheet.name, &output_address, &plan.operations) {
                continue;
            }
            if !new_sheet.formulas.contains_key(&output_address)
                && output_address == *address
                && merged_non_anchor_payload_loss_exempt(plan, old_sheet, new_sheet, address)
            {
                continue;
            }
            anyhow::ensure!(
                new_sheet.formulas.get(&output_address) == Some(formula),
                "unmodified formula changed at {}!{}",
                old_sheet.name,
                address
            );
        }
    }
    Ok(())
}

fn validate_preserved_comments(
    source: &[SheetDump],
    output: &[SheetDump],
    plan: &ExcelEditPlan,
) -> Result<()> {
    if !plan
        .preserve
        .contains(&PreservationRequirement::CommentsAndPositions)
    {
        return Ok(());
    }
    let (renamed, deleted) = sheet_name_changes(&plan.operations);
    for old_sheet in source {
        if deleted.contains(&old_sheet.name) {
            continue;
        }
        let output_name = renamed.get(&old_sheet.name).unwrap_or(&old_sheet.name);
        let Some(new_sheet) = output.iter().find(|sheet| &sheet.name == output_name) else {
            continue;
        };
        let row_map = plan_row_map(plan, &old_sheet.name)?;
        for (address, text) in &old_sheet.comments {
            let output_address = match &row_map {
                Some(map) => match map.map_address(address)? {
                    Some(mapped) => mapped,
                    // The plan deletes the row owning this source comment.
                    None => continue,
                },
                None => address.clone(),
            };
            if cell_comment_is_modified(&old_sheet.name, &output_address, &plan.operations) {
                continue;
            }
            anyhow::ensure!(
                new_sheet.comments.get(&output_address) == Some(text),
                "unmodified comment moved or changed at {}!{}",
                old_sheet.name,
                address
            );
        }
    }
    Ok(())
}

fn validate_preserved_merges(
    source: &[SheetDump],
    output: &[SheetDump],
    plan: &ExcelEditPlan,
) -> Result<()> {
    if !plan
        .preserve
        .contains(&PreservationRequirement::MergedRanges)
    {
        return Ok(());
    }
    let (renamed, deleted) = sheet_name_changes(&plan.operations);
    let mut expected: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    for old_sheet in source {
        if deleted.contains(&old_sheet.name) {
            continue;
        }
        let output_name = renamed.get(&old_sheet.name).unwrap_or(&old_sheet.name);
        expected.insert(
            output_name.clone(),
            old_sheet
                .merges
                .iter()
                .map(|range| normalize_range(range))
                .collect(),
        );
    }
    for operation in &plan.operations {
        match operation {
            EditOperation::CreateSheet { name } => {
                expected.entry(name.clone()).or_default();
            }
            EditOperation::CreateSheetAt { name, .. } => {
                expected.entry(name.clone()).or_default();
            }
            _ => {}
        }
    }
    for operation in &plan.operations {
        let EditOperation::ImportRange {
            source_workbook,
            source_sheet,
            source_range,
            target_sheet,
            target_start,
            copy_merges: true,
            ..
        } = operation
        else {
            continue;
        };
        let imported = read_workbook(source_workbook).with_context(|| {
            format!(
                "read ImportRange source {} for merge validation",
                source_workbook.display()
            )
        })?;
        let imported_sheet = imported
            .sheets
            .iter()
            .find(|sheet| sheet.name == *source_sheet)
            .with_context(|| format!("ImportRange source sheet {source_sheet:?} missing"))?;
        let (min_row, min_col, max_row, max_col) = parse_range(source_range)?;
        let (target_row, target_col) = parse_cell(target_start)?;
        let target_merges = expected.entry(target_sheet.clone()).or_default();
        for merge in &imported_sheet.merges {
            let (first_row, first_col, last_row, last_col) = parse_range(merge)?;
            let intersects = first_row <= max_row
                && last_row >= min_row
                && first_col <= max_col
                && last_col >= min_col;
            if !intersects {
                continue;
            }
            anyhow::ensure!(
                first_row >= min_row
                    && last_row <= max_row
                    && first_col >= min_col
                    && last_col <= max_col,
                "ImportRange cuts through merged range {merge:?}"
            );
            let row_offset = target_row as i64 - min_row as i64;
            let column_offset = target_col as i64 - min_col as i64;
            let translated = format!(
                "{}{}:{}{}",
                column_letters((first_col as i64 + column_offset) as u32),
                (first_row as i64 + row_offset),
                column_letters((last_col as i64 + column_offset) as u32),
                (last_row as i64 + row_offset),
            );
            target_merges.insert(normalize_range(&translated));
        }
    }
    for sheet in output {
        let actual: BTreeSet<_> = sheet
            .merges
            .iter()
            .map(|range| normalize_range(range))
            .collect();
        let expected_ranges = expected.get(&sheet.name).cloned().unwrap_or_default();
        anyhow::ensure!(
            expected_ranges == actual,
            "merged ranges changed on sheet {:?}: expected {:?}, got {:?}",
            sheet.name,
            expected_ranges,
            actual
        );
    }
    Ok(())
}

fn sheet_name_changes(
    operations: &[EditOperation],
) -> (BTreeMap<String, String>, BTreeSet<String>) {
    let renamed = operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::RenameSheet { old_name, new_name } => {
                Some((old_name.clone(), new_name.clone()))
            }
            _ => None,
        })
        .collect();
    let deleted = operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::DeleteSheet { name } => Some(name.clone()),
            _ => None,
        })
        .collect();
    (renamed, deleted)
}

fn normalize_range(value: &str) -> String {
    value.replace('$', "").to_ascii_uppercase()
}

fn column_letters(mut column: u32) -> String {
    let mut result = String::new();
    while column > 0 {
        let remainder = ((column - 1) % 26) as u8;
        result.push((b'A' + remainder) as char);
        column = (column - 1) / 26;
    }
    result.chars().rev().collect()
}

fn cell_value_is_modified(sheet: &str, address: &str, operations: &[EditOperation]) -> bool {
    operations.iter().any(|operation| match operation {
        EditOperation::SetCellValue {
            sheet: target_sheet,
            cell,
            ..
        }
        | EditOperation::SetFormulaText {
            sheet: target_sheet,
            cell,
            ..
        } => target_sheet == sheet && normalize_address(cell) == address,
        EditOperation::CopyRange {
            target_sheet,
            target_start,
            source_range,
            copy_values,
            copy_formulas,
            ..
        }
        | EditOperation::ImportRange {
            target_sheet,
            target_start,
            source_range,
            copy_values,
            copy_formulas,
            ..
        } => {
            (*copy_values || *copy_formulas)
                && target_sheet == sheet
                && address_in_copied_range(address, target_start, source_range)
        }
        _ => false,
    })
}

fn cell_formula_is_modified(sheet: &str, address: &str, operations: &[EditOperation]) -> bool {
    operations.iter().any(|operation| match operation {
        EditOperation::SetCellValue {
            sheet: target_sheet,
            cell,
            ..
        }
        | EditOperation::SetFormulaText {
            sheet: target_sheet,
            cell,
            ..
        } => target_sheet == sheet && normalize_address(cell) == address,
        EditOperation::CopyRange {
            target_sheet,
            target_start,
            source_range,
            copy_values,
            copy_formulas,
            ..
        }
        | EditOperation::ImportRange {
            target_sheet,
            target_start,
            source_range,
            copy_values,
            copy_formulas,
            ..
        } => {
            (*copy_values || *copy_formulas)
                && target_sheet == sheet
                && address_in_copied_range(address, target_start, source_range)
        }
        _ => false,
    })
}

fn cell_comment_is_modified(sheet: &str, address: &str, operations: &[EditOperation]) -> bool {
    operations.iter().any(|operation| match operation {
        EditOperation::SetComment {
            sheet: target_sheet,
            cell,
            ..
        } => target_sheet == sheet && normalize_address(cell) == address,
        EditOperation::CopyRange {
            target_sheet,
            target_start,
            source_range,
            copy_comments,
            ..
        }
        | EditOperation::ImportRange {
            target_sheet,
            target_start,
            source_range,
            copy_comments,
            ..
        } => {
            *copy_comments
                && target_sheet == sheet
                && address_in_copied_range(address, target_start, source_range)
        }
        _ => false,
    })
}

fn address_in_copied_range(address: &str, target_start: &str, source_range: &str) -> bool {
    let Ok((row, column)) = parse_cell(address) else {
        return false;
    };
    let Ok((start_row, start_column)) = parse_cell(target_start) else {
        return false;
    };
    let Ok((source_start_row, source_start_column, source_end_row, source_end_column)) =
        parse_range(source_range)
    else {
        return false;
    };
    row >= start_row
        && row <= start_row + source_end_row - source_start_row
        && column >= start_column
        && column <= start_column + source_end_column - source_start_column
}

fn json_value_equal(expected: &serde_json::Value, actual: Option<&serde_json::Value>) -> bool {
    let Some(actual) = actual else {
        return false;
    };
    let expected_type = expected.get("type").and_then(serde_json::Value::as_str);
    let actual_type = actual.get("type").and_then(serde_json::Value::as_str);
    if expected_type != actual_type {
        return false;
    }
    let (Some(left), Some(right)) = (expected.get("value"), actual.get("value")) else {
        return false;
    };
    match (left.as_f64(), right.as_f64()) {
        (Some(left), Some(right)) => {
            (left - right).abs() <= left.abs().max(right.abs()).max(1.0) * 1e-15
        }
        _ => left == right,
    }
}

fn normalize_address(address: &str) -> String {
    address.replace('$', "").to_ascii_uppercase()
}

fn parse_cell(value: &str) -> Result<(u32, u32)> {
    let mut column = 0u32;
    let mut row = 0u32;
    let mut digits = false;
    for character in value.trim().trim_start_matches('$').chars() {
        if character == '$' {
            continue;
        }
        if character.is_ascii_alphabetic() && !digits {
            column = column
                .checked_mul(26)
                .and_then(|v| v.checked_add(character.to_ascii_uppercase() as u32 - 'A' as u32 + 1))
                .context("column overflow")?;
        } else if character.is_ascii_digit() {
            digits = true;
            row = row
                .checked_mul(10)
                .and_then(|v| v.checked_add(character as u32 - '0' as u32))
                .context("row overflow")?;
        } else {
            bail!("invalid A1 coordinate {value:?}");
        }
    }
    anyhow::ensure!(row > 0 && column > 0, "invalid A1 coordinate {value:?}");
    Ok((row, column))
}

fn parse_range(value: &str) -> Result<(u32, u32, u32, u32)> {
    let (start, end) = value.split_once(':').unwrap_or((value, value));
    let (start_row, start_column) = parse_cell(start)?;
    let (end_row, end_column) = parse_cell(end)?;
    Ok((start_row, start_column, end_row, end_column))
}

#[derive(Deserialize)]
struct WorkerResponse {
    schema_version: u32,
    request_id: Option<String>,
    status: String,
    operation_count: Option<usize>,
    verified_operation_count: Option<usize>,
    input_sha256_before: Option<String>,
    input_sha256_after: Option<String>,
    output_sha256: Option<String>,
    error: Option<WorkerResponseError>,
}

#[derive(Deserialize)]
struct WorkerResponseError {
    kind: String,
    message: String,
}

struct ScratchFile(PathBuf);

impl ScratchFile {
    fn create(parent: &Path, purpose: &str) -> Result<Self> {
        for _ in 0..128 {
            let sequence = SCRATCH_SEQUENCE.fetch_add(1, Ordering::Relaxed);
            let path = parent.join(format!(
                ".openpyxl-{}-{sequence}-{purpose}.tmp",
                std::process::id()
            ));
            match OpenOptions::new().write(true).create_new(true).open(&path) {
                Ok(_) => return Ok(Self(path)),
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => return Err(error).with_context(|| format!("reserve {purpose}")),
            }
        }
        bail!("could not reserve a unique {purpose} path")
    }
    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for ScratchFile {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.0);
    }
}

fn read_diagnostic(path: &Path) -> String {
    fs::read_to_string(path)
        .unwrap_or_default()
        .chars()
        .take(4096)
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn defined_name(scope: Option<&str>, name: &str, text: &str) -> CanonicalDefinedName {
        CanonicalDefinedName {
            scope: scope.map(str::to_owned),
            element: CanonicalXmlElement {
                name: format!("{SPREADSHEET_MAIN_NS}:definedName"),
                attributes: vec![(String::new(), "name".to_owned(), name.to_owned())],
                text: Some(text.to_owned()),
                children: Vec::new(),
            },
        }
    }

    fn defined_names_from_xml(
        xml: &str,
        sheet_order: &[String],
    ) -> Result<Vec<CanonicalDefinedName>> {
        let document = roxmltree::Document::parse(xml)?;
        canonical_defined_names(document.root_element(), sheet_order)
    }

    fn defined_name_plan(
        operations: Vec<EditOperation>,
        drop_allowed: BTreeMap<String, BTreeSet<PreservationRequirement>>,
    ) -> ExcelEditPlan {
        ExcelEditPlan {
            request_id: "defined-name-validation".to_owned(),
            input: PathBuf::from("input.xlsx"),
            output: PathBuf::from("output.xlsx"),
            format: WorkbookFormat::Xlsx,
            workbook: crate::WorkbookProfile {
                format: WorkbookFormat::Xlsx,
                inventory_complete: true,
                features_present: BTreeSet::new(),
            },
            operations,
            preserve: [PreservationRequirement::DefinedNames]
                .into_iter()
                .collect(),
            drop_allowed,
        }
    }

    fn dimension_plan(
        operations: Vec<EditOperation>,
        drop_allowed: BTreeMap<String, BTreeSet<PreservationRequirement>>,
    ) -> ExcelEditPlan {
        let mut plan = defined_name_plan(operations, drop_allowed);
        plan.preserve = [PreservationRequirement::RowAndColumnDimensions]
            .into_iter()
            .collect();
        plan
    }

    fn column_dimension(min: u32, max: u32, attributes: &[(&str, &str)]) -> CanonicalXmlElement {
        let mut canonical_attributes = vec![
            (String::new(), "min".to_owned(), min.to_string()),
            (String::new(), "max".to_owned(), max.to_string()),
        ];
        canonical_attributes.extend(
            attributes
                .iter()
                .map(|(name, value)| (String::new(), (*name).to_owned(), (*value).to_owned())),
        );
        CanonicalXmlElement {
            name: format!("{SPREADSHEET_MAIN_NS}:col"),
            attributes: canonical_attributes,
            text: None,
            children: Vec::new(),
        }
    }

    fn row_dimension(row: u32, attributes: &[(&str, &str)]) -> (String, Vec<(String, String)>) {
        let mut canonical_attributes = vec![("r".to_owned(), row.to_string())];
        canonical_attributes.extend(
            attributes
                .iter()
                .map(|(name, value)| ((*name).to_owned(), (*value).to_owned())),
        );
        (row.to_string(), canonical_attributes)
    }

    fn dimension_snapshot(
        sheet: &str,
        columns: Vec<CanonicalXmlElement>,
        rows: Vec<(String, Vec<(String, String)>)>,
    ) -> WorkbookPreservationSnapshot {
        WorkbookPreservationSnapshot {
            sheet_order: vec![sheet.to_owned()],
            dimensions: vec![(sheet.to_owned(), columns, rows)],
            ..WorkbookPreservationSnapshot::default()
        }
    }

    fn auto_filter(range: &str) -> Option<CanonicalXmlElement> {
        Some(CanonicalXmlElement {
            name: format!("{SPREADSHEET_MAIN_NS}:autoFilter"),
            attributes: vec![(String::new(), "ref".to_owned(), range.to_owned())],
            text: None,
            children: Vec::new(),
        })
    }

    #[test]
    fn empty_default_page_setup_properties_are_normalized_but_explicit_settings_remain(
    ) -> Result<()> {
        const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let source_xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetPr><pageSetUpPr/></sheetPr><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#
        );
        let normalized_xml = format!(
            r#"<worksheet xmlns="{NS}"><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#
        );
        let explicit_xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#
        );
        let source = roxmltree::Document::parse(&source_xml)?;
        let normalized = roxmltree::Document::parse(&normalized_xml)?;
        let explicit = roxmltree::Document::parse(&explicit_xml)?;

        assert_eq!(
            canonical_page_settings(source.root_element()),
            canonical_page_settings(normalized.root_element())
        );
        assert_ne!(
            canonical_page_settings(source.root_element()),
            canonical_page_settings(explicit.root_element())
        );
        Ok(())
    }

    #[test]
    fn row_dimensions_allow_only_openpyxl_custom_height_serialization_normalization() -> Result<()>
    {
        let source = dimension_snapshot(
            "Report",
            Vec::new(),
            vec![row_dimension(12, &[("ht", "15.0")])],
        );
        let output = dimension_snapshot(
            "Report",
            Vec::new(),
            vec![row_dimension(12, &[("ht", "15.0"), ("customHeight", "1")])],
        );
        let plan = dimension_plan(Vec::new(), BTreeMap::new());

        validate_preserved_dimensions(&source, &output, &plan)
    }

    #[test]
    fn row_dimensions_still_reject_a_real_height_or_style_change() {
        let source = dimension_snapshot(
            "Report",
            Vec::new(),
            vec![row_dimension(12, &[("ht", "15.0"), ("s", "2")])],
        );
        let changed_height = dimension_snapshot(
            "Report",
            Vec::new(),
            vec![row_dimension(
                12,
                &[("ht", "16.0"), ("customHeight", "1"), ("s", "2")],
            )],
        );
        let changed_style = dimension_snapshot(
            "Report",
            Vec::new(),
            vec![row_dimension(
                12,
                &[("ht", "15.0"), ("customHeight", "1"), ("s", "3")],
            )],
        );
        let plan = dimension_plan(Vec::new(), BTreeMap::new());

        assert!(validate_preserved_dimensions(&source, &changed_height, &plan).is_err());
        assert!(validate_preserved_dimensions(&source, &changed_style, &plan).is_err());
    }

    #[test]
    fn explicit_row_height_only_exempts_the_requested_height_property() -> Result<()> {
        let source = dimension_snapshot(
            "Report",
            Vec::new(),
            vec![row_dimension(12, &[("ht", "15.0"), ("hidden", "1")])],
        );
        let output = dimension_snapshot(
            "Report",
            Vec::new(),
            vec![row_dimension(
                12,
                &[("ht", "20.0"), ("customHeight", "1"), ("hidden", "1")],
            )],
        );
        let changed_hidden = dimension_snapshot(
            "Report",
            Vec::new(),
            vec![row_dimension(
                12,
                &[("ht", "20.0"), ("customHeight", "1"), ("hidden", "0")],
            )],
        );
        let plan = dimension_plan(
            vec![EditOperation::SetRowHeight {
                sheet: "Report".to_owned(),
                row_index: 12,
                height: 20.0,
            }],
            BTreeMap::new(),
        );

        validate_preserved_dimensions(&source, &output, &plan)?;
        assert!(validate_preserved_dimensions(&source, &changed_hidden, &plan).is_err());
        Ok(())
    }

    #[test]
    fn column_ranges_compare_by_effective_column_properties() -> Result<()> {
        let source = dimension_snapshot(
            "Report",
            vec![column_dimension(
                1,
                2,
                &[("width", "22"), ("customWidth", "1")],
            )],
            Vec::new(),
        );
        let output = dimension_snapshot(
            "Report",
            vec![
                column_dimension(1, 1, &[("width", "22"), ("customWidth", "1")]),
                column_dimension(2, 2, &[("width", "22"), ("customWidth", "1")]),
            ],
            Vec::new(),
        );
        let plan = dimension_plan(Vec::new(), BTreeMap::new());

        validate_preserved_dimensions(&source, &output, &plan)
    }

    #[test]
    fn explicit_column_width_only_exempts_the_requested_width_property() -> Result<()> {
        let source = dimension_snapshot(
            "Report",
            vec![column_dimension(
                1,
                2,
                &[("width", "9"), ("customWidth", "1"), ("hidden", "1")],
            )],
            Vec::new(),
        );
        let output = dimension_snapshot(
            "Report",
            vec![
                column_dimension(
                    1,
                    1,
                    &[("width", "22"), ("customWidth", "1"), ("hidden", "1")],
                ),
                column_dimension(
                    2,
                    2,
                    &[("width", "9"), ("customWidth", "1"), ("hidden", "1")],
                ),
            ],
            Vec::new(),
        );
        let plan = dimension_plan(
            vec![EditOperation::SetColumnWidth {
                sheet: "Report".to_owned(),
                column_index: 1,
                width: 22.0,
            }],
            BTreeMap::new(),
        );

        validate_preserved_dimensions(&source, &output, &plan)
    }

    #[test]
    fn recreated_sheet_dimensions_need_a_scoped_drop_allowance() -> Result<()> {
        let source = dimension_snapshot(
            "Managed",
            vec![column_dimension(
                1,
                3,
                &[("width", "9"), ("customWidth", "1")],
            )],
            vec![row_dimension(4, &[("ht", "18")])],
        );
        let output = dimension_snapshot(
            "Managed",
            vec![column_dimension(
                1,
                1,
                &[("width", "22"), ("customWidth", "1")],
            )],
            Vec::new(),
        );
        let operations = vec![
            EditOperation::DeleteSheet {
                name: "Managed".to_owned(),
            },
            EditOperation::CreateSheetAt {
                name: "Managed".to_owned(),
                index: 0,
            },
            EditOperation::SetColumnWidth {
                sheet: "Managed".to_owned(),
                column_index: 1,
                width: 22.0,
            },
        ];
        let allowed = dimension_plan(
            operations.clone(),
            [(
                "Managed".to_owned(),
                [PreservationRequirement::RowAndColumnDimensions].into(),
            )]
            .into_iter()
            .collect(),
        );
        let disallowed = dimension_plan(operations, BTreeMap::new());

        validate_preserved_dimensions(&source, &output, &allowed)?;
        assert!(validate_preserved_dimensions(&source, &output, &disallowed).is_err());
        Ok(())
    }

    #[test]
    fn row_dimensions_follow_planned_row_insertions() -> Result<()> {
        let source = dimension_snapshot(
            "Report",
            Vec::new(),
            vec![row_dimension(10, &[("ht", "18")])],
        );
        let output = dimension_snapshot(
            "Report",
            Vec::new(),
            vec![row_dimension(12, &[("ht", "18"), ("customHeight", "1")])],
        );
        let plan = dimension_plan(
            vec![EditOperation::InsertRows {
                sheet: "Report".to_owned(),
                first_row: 5,
                count: 2,
            }],
            BTreeMap::new(),
        );

        validate_preserved_dimensions(&source, &output, &plan)
    }

    #[test]
    fn defined_names_follow_sheet_scope_when_sheet_order_changes() -> Result<()> {
        let source_order = vec!["Before".to_owned(), "Managed".to_owned(), "Tail".to_owned()];
        let output_order = vec!["Before".to_owned(), "Tail".to_owned(), "Managed".to_owned()];
        let source_names = defined_names_from_xml(
            &format!(
                r#"<definedNames xmlns="{SPREADSHEET_MAIN_NS}">
                    <definedName name="GlobalRange">'Tail'!$A$1</definedName>
                    <definedName name="ManagedPrintArea" localSheetId="1">'Managed'!$A$1:$B$2</definedName>
                    <definedName name="TailLocal" localSheetId="2">'Tail'!$C$3</definedName>
                </definedNames>"#
            ),
            &source_order,
        )?;
        let output_names = defined_names_from_xml(
            &format!(
                r#"<definedNames xmlns="{SPREADSHEET_MAIN_NS}">
                    <definedName name="GlobalRange">'Tail'!$A$1</definedName>
                    <definedName name="TailLocal" localSheetId="1">'Tail'!$C$3</definedName>
                </definedNames>"#
            ),
            &output_order,
        )?;
        let source = WorkbookPreservationSnapshot {
            sheet_order: source_order,
            defined_names: source_names,
            ..WorkbookPreservationSnapshot::default()
        };
        let output = WorkbookPreservationSnapshot {
            sheet_order: output_order,
            defined_names: output_names,
            ..WorkbookPreservationSnapshot::default()
        };
        let plan = defined_name_plan(
            vec![
                EditOperation::DeleteSheet {
                    name: "Managed".to_owned(),
                },
                EditOperation::CreateSheetAt {
                    name: "Managed".to_owned(),
                    index: 2,
                },
            ],
            [(
                "Managed".to_owned(),
                [PreservationRequirement::DefinedNames].into(),
            )]
            .into_iter()
            .collect(),
        );

        validate_preserved_defined_names(&source, &output, &plan)
    }

    #[test]
    fn defined_name_drop_allowance_does_not_waive_other_names() {
        let source = WorkbookPreservationSnapshot {
            sheet_order: vec!["Managed".to_owned(), "Other".to_owned()],
            defined_names: vec![
                defined_name(Some("Managed"), "ManagedLocal", "'Managed'!$A$1"),
                defined_name(Some("Other"), "OtherLocal", "'Other'!$B$2"),
                defined_name(None, "GlobalRange", "'Other'!$C$3"),
            ],
            ..WorkbookPreservationSnapshot::default()
        };
        let output = WorkbookPreservationSnapshot {
            sheet_order: vec!["Other".to_owned(), "Managed".to_owned()],
            defined_names: vec![
                defined_name(Some("Other"), "OtherLocal", "'Other'!$B$9"),
                defined_name(None, "GlobalRange", "'Other'!$C$3"),
            ],
            ..WorkbookPreservationSnapshot::default()
        };
        let plan = defined_name_plan(
            vec![
                EditOperation::DeleteSheet {
                    name: "Managed".to_owned(),
                },
                EditOperation::CreateSheetAt {
                    name: "Managed".to_owned(),
                    index: 1,
                },
            ],
            [(
                "Managed".to_owned(),
                [PreservationRequirement::DefinedNames].into(),
            )]
            .into_iter()
            .collect(),
        );

        assert!(validate_preserved_defined_names(&source, &output, &plan).is_err());
    }

    #[test]
    fn defined_names_on_deleted_sheet_require_an_explicit_scoped_allowance() {
        let source = WorkbookPreservationSnapshot {
            sheet_order: vec!["Managed".to_owned(), "Other".to_owned()],
            defined_names: vec![defined_name(
                Some("Managed"),
                "ManagedLocal",
                "'Managed'!$A$1",
            )],
            ..WorkbookPreservationSnapshot::default()
        };
        let output = WorkbookPreservationSnapshot {
            sheet_order: vec!["Other".to_owned(), "Managed".to_owned()],
            ..WorkbookPreservationSnapshot::default()
        };
        let plan = defined_name_plan(
            vec![
                EditOperation::DeleteSheet {
                    name: "Managed".to_owned(),
                },
                EditOperation::CreateSheetAt {
                    name: "Managed".to_owned(),
                    index: 1,
                },
            ],
            BTreeMap::new(),
        );

        assert!(validate_preserved_defined_names(&source, &output, &plan).is_err());
    }

    #[test]
    fn defined_name_reassigned_to_a_different_sheet_scope_is_rejected() {
        let source = WorkbookPreservationSnapshot {
            sheet_order: vec!["Before".to_owned(), "Managed".to_owned(), "Tail".to_owned()],
            defined_names: vec![defined_name(Some("Tail"), "TailLocal", "'Tail'!$A$1")],
            ..WorkbookPreservationSnapshot::default()
        };
        let output = WorkbookPreservationSnapshot {
            sheet_order: vec!["Before".to_owned(), "Managed".to_owned(), "Tail".to_owned()],
            defined_names: vec![defined_name(Some("Managed"), "TailLocal", "'Tail'!$A$1")],
            ..WorkbookPreservationSnapshot::default()
        };
        let plan = defined_name_plan(
            vec![
                EditOperation::DeleteSheet {
                    name: "Managed".to_owned(),
                },
                EditOperation::CreateSheetAt {
                    name: "Managed".to_owned(),
                    index: 1,
                },
            ],
            [(
                "Managed".to_owned(),
                [PreservationRequirement::DefinedNames].into(),
            )]
            .into_iter()
            .collect(),
        );

        assert!(validate_preserved_defined_names(&source, &output, &plan).is_err());
    }

    #[test]
    fn explicitly_requested_print_area_is_checked_by_its_operation_readback() -> Result<()> {
        let source = WorkbookPreservationSnapshot {
            sheet_order: vec!["Report".to_owned()],
            defined_names: vec![defined_name(
                Some("Report"),
                "_xlnm.Print_Area",
                "'Report'!$A$1:$B$2",
            )],
            ..WorkbookPreservationSnapshot::default()
        };
        let output = WorkbookPreservationSnapshot {
            sheet_order: vec!["Report".to_owned()],
            defined_names: vec![defined_name(
                Some("Report"),
                "_xlnm.Print_Area",
                "'Report'!$C$3:$D$4",
            )],
            ..WorkbookPreservationSnapshot::default()
        };
        let plan = defined_name_plan(
            vec![EditOperation::SetPageSettings {
                sheet: "Report".to_owned(),
                settings: crate::PageSettings {
                    print_area: Some("C3:D4".to_owned()),
                    ..crate::PageSettings::default()
                },
            }],
            BTreeMap::new(),
        );

        validate_preserved_defined_names(&source, &output, &plan)
    }

    #[test]
    fn filter_database_names_are_checked_against_validated_auto_filter_ranges() -> Result<()> {
        let source = WorkbookPreservationSnapshot {
            sheet_order: vec!["Other".to_owned(), "History".to_owned()],
            defined_names: vec![
                defined_name(
                    Some("Other"),
                    "_xlnm._FilterDatabase",
                    "'Other'!$A$1:$Q$2393",
                ),
                defined_name(
                    Some("History"),
                    "_xlnm._FilterDatabase",
                    "'History'!$A$1:$L$10",
                ),
            ],
            auto_filters: vec![
                ("Other".to_owned(), auto_filter("A1:Q2393")),
                ("History".to_owned(), auto_filter("A1:L10")),
            ],
            ..WorkbookPreservationSnapshot::default()
        };
        let output = WorkbookPreservationSnapshot {
            sheet_order: vec!["History".to_owned(), "Other".to_owned()],
            defined_names: vec![
                defined_name(
                    Some("History"),
                    "_xlnm._FilterDatabase",
                    "'History'!$A$1:$M$2",
                ),
                defined_name(Some("Other"), "_xlnm._FilterDatabase", "Other!A1:Q2393"),
            ],
            auto_filters: vec![
                ("History".to_owned(), auto_filter("A1:M2")),
                ("Other".to_owned(), auto_filter("A1:Q2393")),
            ],
            ..WorkbookPreservationSnapshot::default()
        };
        let mut plan = defined_name_plan(
            vec![
                EditOperation::DeleteSheet {
                    name: "History".to_owned(),
                },
                EditOperation::CreateSheetAt {
                    name: "History".to_owned(),
                    index: 0,
                },
                EditOperation::SetAutoFilter {
                    sheet: "History".to_owned(),
                    range: Some("A1:M2".to_owned()),
                },
            ],
            [(
                "History".to_owned(),
                [PreservationRequirement::DefinedNames].into(),
            )]
            .into_iter()
            .collect(),
        );
        plan.preserve.insert(PreservationRequirement::AutoFilter);

        validate_preserved_defined_names(&source, &output, &plan)?;

        let mut damaged = output.clone();
        damaged.defined_names[1].element.text = Some("Other!A1:Q2392".to_owned());
        assert!(validate_preserved_defined_names(&source, &damaged, &plan).is_err());
        Ok(())
    }

    #[test]
    fn implicit_xf_zero_properties_are_visible_to_umya_expected_style_reader() -> Result<()> {
        let source = ScratchFile::create(&std::env::temp_dir(), "implicit-xf0-source")?;
        let mut workbook = rust_xlsxwriter::Workbook::new();
        workbook.add_worksheet().write_string(4, 10, "value")?;
        workbook.save(source.path())?;

        let styled_source =
            ScratchFile::create(&std::env::temp_dir(), "implicit-xf0-styled-source")?;
        rewrite_test_zip_part(
            source.path(),
            styled_source.path(),
            "xl/styles.xml",
            |xml| add_vertical_center_to_xf_zero(xml),
        )?;

        let source_umya = crate::open_xlsx(styled_source.path())?;
        let source_style = style_at(&source_umya, "Sheet1", "K5")?;
        assert!(source_style.alignment().is_none());

        let explicit_style_zero =
            ScratchFile::create(&std::env::temp_dir(), "explicit-xf0-reader-copy")?;
        let probes =
            materialize_implicit_style_zero(styled_source.path(), explicit_style_zero.path())?;
        let normalized_umya = crate::open_xlsx(explicit_style_zero.path())?;
        let normalized_style = style_at(&normalized_umya, "Sheet1", "K5")?;
        assert_eq!(
            normalized_style.alignment().unwrap().vertical(),
            &umya_spreadsheet::VerticalAlignmentValues::Center
        );
        let probe = probes.get("Sheet1").unwrap();
        let absent_cell_style = style_at(&normalized_umya, "Sheet1", probe)?;
        assert_eq!(
            absent_cell_style.alignment().unwrap().vertical(),
            &umya_spreadsheet::VerticalAlignmentValues::Center
        );

        let inherited_source =
            ScratchFile::create(&std::env::temp_dir(), "implicit-xf0-dimension-source")?;
        rewrite_test_zip_part(
            styled_source.path(),
            inherited_source.path(),
            "xl/worksheets/sheet1.xml",
            add_test_row_and_column_styles,
        )?;
        let plan = ExcelEditPlan {
            request_id: "missing-source-cell-xf0".to_owned(),
            input: inherited_source.path().to_owned(),
            output: inherited_source.path().with_extension("out.xlsx"),
            format: WorkbookFormat::Xlsx,
            workbook: crate::WorkbookProfile {
                format: WorkbookFormat::Xlsx,
                inventory_complete: true,
                features_present: BTreeSet::new(),
            },
            operations: vec![EditOperation::ImportRange {
                source_workbook: inherited_source.path().to_owned(),
                source_sheet: "Sheet1".to_owned(),
                source_range: "AB4".to_owned(),
                target_sheet: "Target".to_owned(),
                target_start: "A1".to_owned(),
                copy_values: false,
                copy_formulas: false,
                style_mode: RangeStyleMode::Full,
                copy_comments: false,
                copy_merges: false,
                copy_row_heights: false,
                copy_column_widths: false,
                copy_row_hidden: false,
                copy_column_hidden: false,
                copy_outline_levels: false,
                copy_all_row_dimensions: false,
                copy_all_column_dimensions: false,
                materialize_blank_cells: false,
                hide_zero: false,
            }],
            preserve: BTreeSet::new(),
            drop_allowed: BTreeMap::new(),
        };
        let expected = expected_import_styles(&plan)?;
        assert_eq!(
            expected[&(String::from("Target"), String::from("A1"))]
                .alignment()
                .unwrap()
                .vertical(),
            &umya_spreadsheet::VerticalAlignmentValues::Center
        );
        Ok(())
    }

    #[test]
    fn detects_stored_content_in_merged_non_anchor_cells() -> Result<()> {
        let xml = format!(
            r#"<worksheet xmlns="{SPREADSHEET_MAIN_NS}"><sheetData>
                <row r="5"><c r="A5" t="s"><v>1</v></c></row>
                <row r="10"><c r="A10" t="s"><v>34</v></c><c r="B10" s="2"/></row>
            </sheetData><mergeCells count="1"><mergeCell ref="A5:A26"/></mergeCells></worksheet>"#
        );
        let cells = non_anchor_payload_cells("Sheet1", &xml)?;
        assert_eq!(cells, vec![("Sheet1".to_owned(), "A10".to_owned())]);
        assert_eq!(
            merged_payload_cells_in_range(cells.clone(), "Sheet1", "A10")?,
            cells
        );
        assert!(merged_payload_cells_in_range(cells.clone(), "Sheet1", "A5:A9")?.is_empty());
        assert!(merged_payload_cells_in_range(cells, "Other", "A5:A26")?.is_empty());
        Ok(())
    }

    #[test]
    fn per_sheet_snapshot_comparison_ignores_sheet_order_only() {
        let source = vec![
            ("使用说明".to_owned(), vec![1]),
            ("核查表校验结果".to_owned(), vec![2, 3]),
        ];
        // Reordered worksheets with identical per-sheet content match.
        let reordered = vec![
            ("核查表校验结果".to_owned(), vec![2, 3]),
            ("使用说明".to_owned(), vec![1]),
        ];
        assert!(per_sheet_entries_match(&source, &reordered));
        // Changed per-sheet content still fails.
        let changed = vec![
            ("核查表校验结果".to_owned(), vec![2, 4]),
            ("使用说明".to_owned(), vec![1]),
        ];
        assert!(!per_sheet_entries_match(&source, &changed));
        // A dropped worksheet still fails.
        assert!(!per_sheet_entries_match(&source, &reordered[0..1].to_vec()));
    }

    fn sheet_dump(name: &str, merges: &[&str]) -> SheetDump {
        SheetDump {
            name: name.to_owned(),
            values: BTreeMap::new(),
            formulas: BTreeMap::new(),
            formula_status: "ok".to_owned(),
            merges: merges.iter().map(|merge| merge.to_string()).collect(),
            merge_status: "ok".to_owned(),
            comments: BTreeMap::new(),
            comments_status: "ok".to_owned(),
            conditional_formatting: Vec::new(),
            conditional_formatting_status: "ok".to_owned(),
            data_validations: Vec::new(),
            data_validations_status: "ok".to_owned(),
        }
    }

    fn exemption_plan(
        directory: &Path,
        operations: Vec<EditOperation>,
        allowance_sheet: Option<&str>,
    ) -> ExcelEditPlan {
        let input = directory.join("input.xlsx");
        std::fs::write(&input, b"original workbook bytes").unwrap();
        let mut plan = ExcelEditPlan {
            request_id: "payload-allowance-test".to_owned(),
            input,
            output: directory.join("output.xlsx"),
            format: WorkbookFormat::Xlsx,
            workbook: crate::WorkbookProfile {
                format: WorkbookFormat::Xlsx,
                inventory_complete: true,
                features_present: BTreeSet::new(),
            },
            operations,
            preserve: [PreservationRequirement::UnmodifiedPackageParts]
                .into_iter()
                .collect(),
            drop_allowed: BTreeMap::new(),
        };
        if let Some(sheet) = allowance_sheet {
            plan.drop_allowed.insert(
                sheet.to_owned(),
                [PreservationRequirement::MergedNonAnchorPayloads]
                    .into_iter()
                    .collect(),
            );
        }
        plan
    }

    fn exemption_test_directory(label: &str) -> PathBuf {
        let directory = std::env::temp_dir().join(format!(
            "excel-editor-payload-allowance-{label}-{}",
            std::process::id()
        ));
        std::fs::create_dir_all(&directory).unwrap();
        directory
    }

    #[test]
    fn payload_allowance_exempts_only_outside_operation_rectangles() -> Result<()> {
        let directory = exemption_test_directory("preflight");
        let operations = vec![
            EditOperation::ImportRange {
                source_workbook: directory.join("source.xlsx"),
                source_sheet: "模板".to_owned(),
                source_range: "A1:M10".to_owned(),
                target_sheet: "汇总".to_owned(),
                target_start: "A1".to_owned(),
                copy_values: true,
                copy_formulas: true,
                style_mode: RangeStyleMode::Basic,
                copy_comments: false,
                copy_merges: false,
                copy_row_heights: false,
                copy_column_widths: false,
                copy_row_hidden: false,
                copy_column_hidden: false,
                copy_outline_levels: false,
                copy_all_row_dimensions: false,
                copy_all_column_dimensions: false,
                materialize_blank_cells: false,
                hide_zero: false,
            },
            EditOperation::SetCellValue {
                sheet: "其他".to_owned(),
                cell: "A1".to_owned(),
                value: crate::CellValue::Text("unrelated".to_owned()),
            },
        ];
        let plan = exemption_plan(&directory, operations, Some("汇总"));

        // Payload on an allowed sheet outside the A1:M10 import rectangle.
        assert_eq!(
            offending_target_merged_payload(&plan, vec![("汇总".to_owned(), "B20".to_owned())]),
            None
        );
        // The same payload one row inside the import rectangle stays protected.
        assert_eq!(
            offending_target_merged_payload(&plan, vec![("汇总".to_owned(), "M10".to_owned())]),
            Some(("汇总".to_owned(), "M10".to_owned()))
        );
        // A payload on a worksheet without the allowance stays protected.
        assert_eq!(
            offending_target_merged_payload(&plan, vec![("其他".to_owned(), "B2".to_owned())]),
            Some(("其他".to_owned(), "B2".to_owned()))
        );
        Ok(())
    }

    #[test]
    fn readback_exemption_requires_allowance_and_unchanged_merge() {
        let source = sheet_dump("汇总", &["A1:C1"]);
        let output = sheet_dump("汇总", &["A1:C1"]);
        let operations = vec![EditOperation::SetCellValue {
            sheet: "汇总".to_owned(),
            cell: "F5".to_owned(),
            value: crate::CellValue::Text("unrelated".to_owned()),
        }];
        let directory = exemption_test_directory("readback");
        let mut plan = exemption_plan(&directory, operations, None);

        // Without the allowance nothing is exempt, even for merged members.
        assert!(!merged_non_anchor_payload_loss_exempt(
            &plan, &source, &output, "B1"
        ));
        assert_eq!(
            non_anchor_merge_containing(&source, "B1"),
            Some("A1:C1".to_owned())
        );
        assert_eq!(non_anchor_merge_containing(&source, "A1"), None);
        assert_eq!(non_anchor_merge_containing(&source, "D4"), None);

        plan.drop_allowed.insert(
            "汇总".to_owned(),
            [PreservationRequirement::MergedNonAnchorPayloads]
                .into_iter()
                .collect(),
        );
        assert!(merged_non_anchor_payload_loss_exempt(
            &plan, &source, &output, "B1"
        ));

        // A changed merge set drops the exemption; the strict merge comparison
        // reports that change separately.
        let rewritten = sheet_dump("汇总", &["A1:D1"]);
        assert!(!merged_non_anchor_payload_loss_exempt(
            &plan, &source, &rewritten, "B1"
        ));
    }

    fn add_test_row_and_column_styles(xml: &str) -> Result<String> {
        let document = roxmltree::Document::parse(xml)?;
        let sheet_data = document
            .descendants()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && node.tag_name().name() == "sheetData"
            })
            .context("test worksheet has no sheetData")?;
        let first_row = sheet_data
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "row")
            .context("test worksheet has no row")?;
        let mut result = xml.to_owned();
        result.insert_str(
            first_row.range().start,
            "<row xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" r=\"4\" s=\"0\" customFormat=\"1\"/>",
        );
        let document = roxmltree::Document::parse(&result)?;
        let sheet_data = document
            .descendants()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && node.tag_name().name() == "sheetData"
            })
            .context("test worksheet has no sheetData after row insertion")?;
        result.insert_str(
            sheet_data.range().start,
            "<cols xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><col min=\"1\" max=\"28\" style=\"0\"/></cols>",
        );
        Ok(result)
    }

    fn add_vertical_center_to_xf_zero(styles_xml: &str) -> Result<String> {
        let document = roxmltree::Document::parse(styles_xml)?;
        let cell_xfs = document
            .descendants()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && node.tag_name().name() == "cellXfs"
            })
            .context("test workbook has no cellXfs")?;
        let xf = cell_xfs
            .children()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && node.tag_name().name() == "xf"
            })
            .context("test workbook has no default cellXf")?;
        let range = xf.range();
        let bytes = styles_xml.as_bytes();
        let start_tag_length = xml_start_tag_length(&bytes[range.start..])
            .context("find test cellXf start-tag boundary")?;
        let start_tag_end = range.start + start_tag_length;
        let start_tag = &styles_xml[range.start..start_tag_end];
        if start_tag.ends_with("/>") {
            let mut replacement = start_tag[..start_tag.len() - 2].to_owned();
            replacement.push('>');
            replacement.push_str("<alignment vertical=\"center\"/></xf>");
            let mut output = String::with_capacity(styles_xml.len() + 42);
            output.push_str(&styles_xml[..range.start]);
            output.push_str(&replacement);
            output.push_str(&styles_xml[range.end..]);
            Ok(output)
        } else {
            let mut output = String::with_capacity(styles_xml.len() + 32);
            output.push_str(&styles_xml[..start_tag_end]);
            output.push_str("<alignment vertical=\"center\"/>");
            output.push_str(&styles_xml[start_tag_end..]);
            Ok(output)
        }
    }

    fn rewrite_test_zip_part<F>(
        source: &Path,
        output: &Path,
        part_name: &str,
        rewrite: F,
    ) -> Result<()>
    where
        F: FnOnce(&str) -> Result<String>,
    {
        let mut archive = zip::ZipArchive::new(File::open(source)?)?;
        let mut writer = zip::ZipWriter::new(File::create(output)?);
        let mut rewrite = Some(rewrite);
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            let name = entry.name().to_owned();
            let is_directory = entry.is_dir();
            let options =
                zip::write::SimpleFileOptions::default().compression_method(entry.compression());
            if is_directory {
                writer.add_directory(name, options)?;
                continue;
            }
            if name == part_name {
                let mut bytes = Vec::new();
                entry.read_to_end(&mut bytes)?;
                let xml = std::str::from_utf8(&bytes)?;
                writer.start_file(name, options)?;
                writer.write_all(
                    rewrite
                        .take()
                        .context("test ZIP part rewrite was already consumed")?(
                        xml
                    )?
                    .as_bytes(),
                )?;
            } else {
                writer.start_file(name, options)?;
                std::io::copy(&mut entry, &mut writer)?;
            }
        }
        anyhow::ensure!(rewrite.is_none(), "test ZIP part {part_name} missing");
        writer.finish()?;
        Ok(())
    }

    #[test]
    fn merged_style_readback_treats_only_non_anchor_cells_as_normalized() {
        let mut source = umya_spreadsheet::new_file();
        let source_sheet = source.sheet_by_name_mut("Sheet1").unwrap();
        source_sheet.add_merge_cells("A1:I1");
        source_sheet
            .cell_mut("A1")
            .style_mut()
            .borders_mut()
            .top_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);
        source_sheet
            .cell_mut("A1")
            .style_mut()
            .borders_mut()
            .bottom_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);
        source_sheet
            .cell_mut("I1")
            .style_mut()
            .borders_mut()
            .right_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);
        source_sheet
            .cell_mut("I1")
            .style_mut()
            .borders_mut()
            .bottom_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Double);
        source_sheet
            .cell_mut("G1")
            .style_mut()
            .borders_mut()
            .bottom_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Double);

        let mut output = umya_spreadsheet::new_file();
        let output_sheet = output.sheet_by_name_mut("Sheet1").unwrap();
        output_sheet.add_merge_cells("A1:I1");
        output_sheet
            .cell_mut("A1")
            .style_mut()
            .borders_mut()
            .top_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);
        output_sheet
            .cell_mut("A1")
            .style_mut()
            .borders_mut()
            .bottom_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);
        output_sheet
            .cell_mut("I1")
            .style_mut()
            .borders_mut()
            .right_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);

        // G1 inherits top/bottom from the anchor. The source bottom-right
        // cell's conflicting raw bottom border is ignored because the anchor
        // already defines that side.
        assert_eq!(
            merged_member_effective_borders_equal(source_sheet, output_sheet, "G1"),
            Some(true)
        );
        // The anchor remains under the full style comparison.
        assert_eq!(
            merged_member_effective_borders_equal(source_sheet, output_sheet, "A1"),
            None
        );
        // The right edge must still match: the anchor leaves it empty, so
        // openpyxl fills it from the bottom-right cell.
        output_sheet
            .cell_mut("I1")
            .style_mut()
            .borders_mut()
            .right_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Double);
        assert_eq!(
            merged_member_effective_borders_equal(source_sheet, output_sheet, "I1"),
            Some(false)
        );
        assert_eq!(
            merged_member_effective_borders_equal(source_sheet, output_sheet, "J1"),
            None
        );
    }

    #[test]
    fn packaged_runtime_uses_an_absolute_internal_path_and_never_path_lookup() {
        let config =
            OpenpyxlWorkerConfig::for_product_binary("/usr/bin/pytools", "pytools").unwrap();
        assert_eq!(
            config.python_executable,
            PathBuf::from("/usr/libexec/pytools/openpyxl/runtime/bin/python3")
        );
        assert_eq!(
            config.worker_script,
            PathBuf::from("/usr/libexec/pytools/openpyxl/worker/openpyxl_worker.py")
        );
        assert!(
            OpenpyxlWorkerConfig::for_product_binary("/usr/bin/pytools", "../python").is_none()
        );
    }

    #[test]
    fn capability_is_narrowed_to_implemented_operations_and_semantic_parts() {
        let capability = BackendCapability {
            engine: crate::EngineId::OpenpyxlWorker,
            status: crate::CapabilityStatus::Unverified,
            format: WorkbookFormat::Xlsx,
            operations: [
                EditOperationKind::SetCellValue,
                EditOperationKind::DeleteRows,
            ]
            .into_iter()
            .collect(),
            supported_features: BTreeSet::new(),
            preserved: [
                PreservationRequirement::CellValuesAndTypes,
                PreservationRequirement::Styles,
                PreservationRequirement::PageSettings,
                PreservationRequirement::RowAndColumnDimensions,
                PreservationRequirement::Relationships,
                PreservationRequirement::DefinedNames,
                PreservationRequirement::HeaderFooterText,
                PreservationRequirement::ConditionalFormatting,
                PreservationRequirement::DataValidations,
                PreservationRequirement::UnmodifiedPackageParts,
            ]
            .into_iter()
            .collect(),
            platform: "test".into(),
            evidence: None,
        };
        let backend = OpenpyxlXlsxBackend::build(capability, None).unwrap();
        assert!(backend
            .capability()
            .operations
            .contains(&EditOperationKind::SetCellValue));
        assert!(!backend
            .capability()
            .operations
            .contains(&EditOperationKind::DeleteRows));
        assert!(!backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::UnmodifiedPackageParts));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::Styles));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::PageSettings));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::RowAndColumnDimensions));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::Relationships));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::DefinedNames));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::HeaderFooterText));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::ConditionalFormatting));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::DataValidations));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::CellValuesAndTypes));
        assert_eq!(
            backend.capability().status,
            crate::CapabilityStatus::Unsupported
        );
    }

    #[test]
    fn style_only_copy_does_not_exempt_values_formulas_or_comments_from_preservation() {
        let operation = EditOperation::CopyRange {
            source_sheet: "Data".into(),
            source_range: "A1:B2".into(),
            target_sheet: "Data".into(),
            target_start: "D4".into(),
            copy_values: false,
            copy_formulas: false,
            copy_styles: true,
            copy_comments: false,
        };
        let operations = [operation];
        assert!(!cell_value_is_modified("Data", "D4", &operations));
        assert!(!cell_formula_is_modified("Data", "D4", &operations));
        assert!(!cell_comment_is_modified("Data", "D4", &operations));
    }

    #[test]
    fn changing_a_value_does_not_exempt_existing_comment_from_preservation() {
        let operations = [EditOperation::SetCellValue {
            sheet: "Data".into(),
            cell: "A1".into(),
            value: crate::CellValue::Text("replacement".into()),
        }];
        assert!(cell_value_is_modified("Data", "A1", &operations));
        assert!(cell_formula_is_modified("Data", "A1", &operations));
        assert!(!cell_comment_is_modified("Data", "A1", &operations));
    }
}
