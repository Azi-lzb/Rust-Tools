//! Engine-independent edit plans and a conservative execution boundary.
//!
//! This module does not implement an XLSX backend. A backend is routable only
//! when its registered evidence exactly covers the plan's format, operations,
//! workbook features, preservation requirements, and runtime platform.

use anyhow::{anyhow, bail, Context, Result};
use serde::{Deserialize, Serialize};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{self, OpenOptions},
    path::{Path, PathBuf},
    sync::{
        atomic::{AtomicU64, Ordering},
        Arc,
    },
    time::Instant,
};

static TEMP_SEQUENCE: AtomicU64 = AtomicU64::new(0);

#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum WorkbookFormat {
    Xlsx,
    Xlsm,
    Xls,
    Xlsb,
}

impl WorkbookFormat {
    fn extension(self) -> &'static str {
        match self {
            Self::Xlsx => "xlsx",
            Self::Xlsm => "xlsm",
            Self::Xls => "xls",
            Self::Xlsb => "xlsb",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum WorkbookFeature {
    Comments,
    ConditionalFormatting,
    DataValidations,
    MergedCells,
    DefinedNames,
    Formulas,
    Macros,
    ExternalLinks,
    Charts,
    PivotTables,
    Slicers,
    Drawings,
    Images,
    SheetProtection,
    ActiveX,
    DigitalSignatures,
    Relationships,
    UnknownExtensions,
    HeadersFooters,
    FreezePanes,
    AutoFilter,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum PreservationRequirement {
    CellValuesAndTypes,
    FormulaText,
    CommentsAndPositions,
    Styles,
    MergedRanges,
    ConditionalFormatting,
    DataValidations,
    DefinedNames,
    RowAndColumnDimensions,
    PageSettings,
    HeaderFooterText,
    Relationships,
    UnmodifiedPackageParts,
    SheetIdentityAndOrder,
    FreezePanes,
    AutoFilter,
    /// Allowance (used in `drop_allowed`, never in `preserve`): the openpyxl
    /// rewrite drops payloads stored on non-anchor cells of merged ranges,
    /// which sample evidence has shown to match the legacy Python baseline
    /// for the declaring worksheet. Declaring it narrows the openpyxl
    /// preflight to payloads intersecting operation target rectangles and
    /// exempts the emptied non-anchor cells from unmodified-readback
    /// comparison while their merged ranges stay intact. It never requires
    /// a DeleteSheet operation and is not a production capability on its own.
    MergedNonAnchorPayloads,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum EditOperationKind {
    SetCellValue,
    SetFormulaText,
    SetComment,
    SetColumnWidth,
    SetRowHeight,
    DeleteRows,
    InsertRows,
    RenameSheet,
    CreateSheet,
    CreateSheetAt,
    MoveSheet,
    SetFreezePanes,
    SetAutoFilter,
    DeleteSheet,
    SetCellFillColor,
    CopyRange,
    ImportRange,
    SetWorkbookCalculationProperties,
    SetPageSettings,
    ReplaceHeaderText,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "snake_case")]
pub enum CellValue {
    Text(String),
    /// Decimal text is retained to avoid silently imposing an f64 precision policy.
    Decimal(String),
    /// Excel date/time serial plus the number format required to retain its
    /// displayed calendar value when writing into a newly-created cell.
    DateTime {
        serial: String,
        number_format: String,
    },
    Boolean(bool),
    Error(String),
    Blank,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

/// `None` means leave that setting untouched. `clear_scale` explicitly removes
/// a previously stored scale when switching to fit-to-page behavior.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct PageSettings {
    pub orientation: Option<PageOrientation>,
    pub fit_to_page: Option<bool>,
    pub fit_to_width: Option<u16>,
    pub fit_to_height: Option<u16>,
    pub clear_scale: bool,
    pub paper_size: Option<u16>,
    pub print_area: Option<String>,
    pub horizontal_centered: Option<bool>,
    pub vertical_centered: Option<bool>,
    pub footer: Option<String>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum RangeStyleMode {
    None,
    Basic,
    Full,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "operation", rename_all = "snake_case")]
pub enum EditOperation {
    SetCellValue {
        sheet: String,
        cell: String,
        value: CellValue,
    },
    SetFormulaText {
        sheet: String,
        cell: String,
        formula: String,
    },
    SetComment {
        sheet: String,
        cell: String,
        text: Option<String>,
        author: Option<String>,
    },
    SetColumnWidth {
        sheet: String,
        column_index: u32,
        width: f64,
    },
    SetRowHeight {
        sheet: String,
        row_index: u32,
        height: f64,
    },
    DeleteRows {
        sheet: String,
        first_row: u32,
        count: u32,
    },
    InsertRows {
        sheet: String,
        first_row: u32,
        count: u32,
    },
    RenameSheet {
        old_name: String,
        new_name: String,
    },
    CreateSheet {
        name: String,
    },
    /// Create a worksheet at a zero-based workbook position.
    CreateSheetAt {
        name: String,
        index: usize,
    },
    /// Move an existing worksheet to a zero-based position in the final sheet
    /// order. Worksheet content and page settings stay untouched; only
    /// workbook-level position metadata follows the sheet.
    MoveSheet {
        sheet: String,
        new_index: usize,
    },
    /// Freeze panes at the first unfrozen cell, such as `A2`; `None` clears panes.
    SetFreezePanes {
        sheet: String,
        cell: Option<String>,
    },
    /// Set the worksheet AutoFilter reference, such as `A1:M42`; `None` clears it.
    SetAutoFilter {
        sheet: String,
        range: Option<String>,
    },
    DeleteSheet {
        name: String,
    },
    SetCellFillColor {
        sheet: String,
        cell: String,
        argb: String,
    },
    CopyRange {
        source_sheet: String,
        source_range: String,
        target_sheet: String,
        target_start: String,
        copy_values: bool,
        copy_formulas: bool,
        copy_styles: bool,
        copy_comments: bool,
    },
    ImportRange {
        source_workbook: PathBuf,
        source_sheet: String,
        source_range: String,
        target_sheet: String,
        target_start: String,
        copy_values: bool,
        copy_formulas: bool,
        style_mode: RangeStyleMode,
        copy_comments: bool,
        copy_merges: bool,
        copy_row_heights: bool,
        copy_column_widths: bool,
        #[serde(default)]
        copy_row_hidden: bool,
        #[serde(default)]
        copy_column_hidden: bool,
        #[serde(default)]
        copy_outline_levels: bool,
        /// Copy explicitly stored row dimensions outside `source_range`.
        /// Intended for full-sheet import plans rooted at A1.
        #[serde(default)]
        copy_all_row_dimensions: bool,
        /// Copy explicitly stored column dimensions outside `source_range`.
        /// Intended for full-sheet import plans rooted at A1.
        #[serde(default)]
        copy_all_column_dimensions: bool,
        /// Create destination cells for empty coordinates within the imported
        /// rectangle, matching APIs that materialize every source coordinate.
        #[serde(default)]
        materialize_blank_cells: bool,
        hide_zero: bool,
    },
    /// Update workbook-level calculation flags. `None` preserves the existing
    /// attribute; this operation requests flags only and never recalculates formulas.
    SetWorkbookCalculationProperties {
        full_calc_on_load: Option<bool>,
        force_full_calc: Option<bool>,
    },
    SetPageSettings {
        sheet: String,
        settings: PageSettings,
    },
    ReplaceHeaderText {
        sheet: String,
        search: String,
        replacement: String,
    },
}

impl EditOperation {
    pub fn kind(&self) -> EditOperationKind {
        match self {
            Self::SetCellValue { .. } => EditOperationKind::SetCellValue,
            Self::SetFormulaText { .. } => EditOperationKind::SetFormulaText,
            Self::SetComment { .. } => EditOperationKind::SetComment,
            Self::SetColumnWidth { .. } => EditOperationKind::SetColumnWidth,
            Self::SetRowHeight { .. } => EditOperationKind::SetRowHeight,
            Self::DeleteRows { .. } => EditOperationKind::DeleteRows,
            Self::InsertRows { .. } => EditOperationKind::InsertRows,
            Self::RenameSheet { .. } => EditOperationKind::RenameSheet,
            Self::CreateSheet { .. } => EditOperationKind::CreateSheet,
            Self::CreateSheetAt { .. } => EditOperationKind::CreateSheetAt,
            Self::MoveSheet { .. } => EditOperationKind::MoveSheet,
            Self::SetFreezePanes { .. } => EditOperationKind::SetFreezePanes,
            Self::SetAutoFilter { .. } => EditOperationKind::SetAutoFilter,
            Self::DeleteSheet { .. } => EditOperationKind::DeleteSheet,
            Self::SetCellFillColor { .. } => EditOperationKind::SetCellFillColor,
            Self::CopyRange { .. } => EditOperationKind::CopyRange,
            Self::ImportRange { .. } => EditOperationKind::ImportRange,
            Self::SetWorkbookCalculationProperties { .. } => {
                EditOperationKind::SetWorkbookCalculationProperties
            }
            Self::SetPageSettings { .. } => EditOperationKind::SetPageSettings,
            Self::ReplaceHeaderText { .. } => EditOperationKind::ReplaceHeaderText,
        }
    }

    fn validate(&self) -> Result<()> {
        let require_sheet_cell = |sheet: &str, cell: &str| -> Result<()> {
            if sheet.trim().is_empty() || cell.trim().is_empty() {
                bail!("sheet name and cell address must not be empty");
            }
            Ok(())
        };
        match self {
            Self::SetCellValue { sheet, cell, value } => {
                require_sheet_cell(sheet, cell)?;
                match value {
                    CellValue::Decimal(decimal) if decimal.trim().is_empty() => {
                        bail!("decimal cell values must not be empty");
                    }
                    CellValue::DateTime {
                        serial,
                        number_format,
                    } => {
                        if serial.trim().is_empty() || number_format.trim().is_empty() {
                            bail!("date/time cell values require a serial and number format");
                        }
                        let number = serial.parse::<f64>().context("parse date/time serial")?;
                        if !number.is_finite() {
                            bail!("date/time serial must be finite");
                        }
                    }
                    _ => {}
                }
            }
            Self::SetFormulaText {
                sheet,
                cell,
                formula,
            } => {
                require_sheet_cell(sheet, cell)?;
                if formula.trim().is_empty() {
                    bail!("formula text must not be empty");
                }
            }
            Self::SetComment { sheet, cell, .. } => require_sheet_cell(sheet, cell)?,
            Self::SetColumnWidth {
                sheet,
                column_index,
                width,
            } => {
                if sheet.trim().is_empty() || *column_index == 0 {
                    bail!("column-width updates require a sheet and a 1-based column");
                }
                if !width.is_finite() || *width < 0.0 {
                    bail!("column width must be finite and non-negative");
                }
            }
            Self::SetRowHeight {
                sheet,
                row_index,
                height,
            } => {
                if sheet.trim().is_empty() || *row_index == 0 {
                    bail!("row-height updates require a sheet and a 1-based row");
                }
                if !height.is_finite() || *height < 0.0 {
                    bail!("row height must be finite and non-negative");
                }
            }
            Self::DeleteRows {
                sheet,
                first_row,
                count,
            } => {
                if sheet.trim().is_empty() || *first_row == 0 || *count == 0 {
                    bail!("row deletion requires a sheet and a non-empty 1-based range");
                }
                first_row
                    .checked_add(*count - 1)
                    .ok_or_else(|| anyhow!("row deletion range overflows"))?;
            }
            Self::InsertRows {
                sheet,
                first_row,
                count,
            } => {
                if sheet.trim().is_empty() || *first_row == 0 || *count == 0 {
                    bail!("row insertion requires a sheet and a non-empty 1-based range");
                }
                first_row
                    .checked_add(*count - 1)
                    .ok_or_else(|| anyhow!("row insertion range overflows"))?;
            }
            Self::RenameSheet { old_name, new_name } => {
                if old_name.trim().is_empty() || new_name.trim().is_empty() {
                    bail!("sheet rename requires non-empty old and new names");
                }
            }
            Self::CreateSheet { name } | Self::DeleteSheet { name } => {
                if name.trim().is_empty() {
                    bail!("worksheet name must not be empty");
                }
            }
            Self::CreateSheetAt { name, .. } => {
                if name.trim().is_empty() {
                    bail!("worksheet name must not be empty");
                }
            }
            Self::MoveSheet { sheet, .. } => {
                if sheet.trim().is_empty() {
                    bail!("sheet move requires a worksheet name");
                }
            }
            Self::SetFreezePanes { sheet, cell } => {
                if sheet.trim().is_empty() {
                    bail!("freeze-panes update requires a sheet name");
                }
                if let Some(cell) = cell {
                    if !valid_a1_cell(cell) {
                        bail!("freeze panes requires a valid A1 cell address");
                    }
                }
            }
            Self::SetAutoFilter { sheet, range } => {
                if sheet.trim().is_empty() {
                    bail!("AutoFilter update requires a sheet name");
                }
                if let Some(range) = range {
                    if !valid_a1_range(range) {
                        bail!("AutoFilter requires a valid A1 cell or range reference");
                    }
                }
            }
            Self::SetCellFillColor { sheet, cell, argb } => {
                require_sheet_cell(sheet, cell)?;
                if !matches!(argb.len(), 6 | 8)
                    || !argb.bytes().all(|byte| byte.is_ascii_hexdigit())
                {
                    bail!("cell fill color must be RGB or ARGB hexadecimal");
                }
            }
            Self::CopyRange {
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                copy_values,
                copy_formulas,
                copy_styles,
                copy_comments,
            } => {
                if source_sheet.trim().is_empty()
                    || target_sheet.trim().is_empty()
                    || source_range.trim().is_empty()
                    || target_start.trim().is_empty()
                {
                    bail!("range copy requires source/target sheets and ranges");
                }
                if !copy_values && !copy_formulas && !copy_styles && !copy_comments {
                    bail!("range copy must request at least one property");
                }
            }
            Self::ImportRange {
                source_workbook,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                copy_values,
                copy_formulas,
                style_mode,
                copy_comments,
                copy_merges,
                copy_row_heights,
                copy_column_widths,
                copy_row_hidden,
                copy_column_hidden,
                copy_outline_levels,
                copy_all_row_dimensions,
                copy_all_column_dimensions,
                materialize_blank_cells,
                hide_zero: _,
            } => {
                if source_workbook.as_os_str().is_empty()
                    || source_sheet.trim().is_empty()
                    || source_range.trim().is_empty()
                    || target_sheet.trim().is_empty()
                    || target_start.trim().is_empty()
                {
                    bail!("cross-workbook range import requires source and target paths, sheets, and ranges");
                }
                if *copy_all_row_dimensions && !copy_row_heights && !copy_row_hidden {
                    bail!("copy_all_row_dimensions requires copying row heights or hidden flags");
                }
                if *copy_all_column_dimensions && !copy_column_widths && !copy_column_hidden {
                    bail!(
                        "copy_all_column_dimensions requires copying column widths or hidden flags"
                    );
                }
                if (*copy_all_row_dimensions || *copy_all_column_dimensions)
                    && (!source_range_starts_at_a1(source_range)
                        || !target_start.eq_ignore_ascii_case("A1"))
                {
                    bail!("copying all worksheet dimensions requires source and target ranges rooted at A1");
                }
                if !copy_values
                    && !copy_formulas
                    && matches!(style_mode, RangeStyleMode::None)
                    && !copy_comments
                    && !copy_merges
                    && !copy_row_heights
                    && !copy_column_widths
                    && !copy_row_hidden
                    && !copy_column_hidden
                    && !copy_outline_levels
                    && !copy_all_row_dimensions
                    && !copy_all_column_dimensions
                    && !materialize_blank_cells
                {
                    bail!("cross-workbook range import must copy at least one property");
                }
            }
            Self::SetPageSettings { sheet, settings } => {
                if sheet.trim().is_empty() {
                    bail!("page-settings update requires a sheet name");
                }
                if settings.fit_to_width == Some(0) {
                    bail!("fit-to-page width must be positive when specified");
                }
                if settings
                    .print_area
                    .as_ref()
                    .is_some_and(|area| area.trim().is_empty())
                {
                    bail!("print area must not be empty when specified");
                }
            }
            Self::SetWorkbookCalculationProperties {
                full_calc_on_load,
                force_full_calc,
            } => {
                if full_calc_on_load.is_none() && force_full_calc.is_none() {
                    bail!("calculation-property update must request at least one flag");
                }
            }
            Self::ReplaceHeaderText {
                sheet,
                search,
                replacement,
            } => {
                if sheet.trim().is_empty()
                    || search.is_empty()
                    || replacement.is_empty()
                    || search == replacement
                {
                    bail!("header replacement requires a sheet and distinct non-empty text values");
                }
            }
        }
        Ok(())
    }
}

fn valid_a1_cell(value: &str) -> bool {
    let value = value.trim().replace('$', "");
    let split = value.find(|character: char| character.is_ascii_digit());
    let Some(split) = split else { return false };
    let (column, row) = value.split_at(split);
    if column.is_empty() || row.is_empty() || !column.bytes().all(|byte| byte.is_ascii_alphabetic())
    {
        return false;
    }
    let column = column.bytes().fold(Some(0u32), |value, byte| {
        value
            .and_then(|value| value.checked_mul(26))
            .and_then(|value| value.checked_add(byte.to_ascii_uppercase() as u32 - b'A' as u32 + 1))
    });
    let row = row.parse::<u32>().ok();
    matches!((column, row), (Some(column), Some(row)) if (1..=16_384).contains(&column) && (1..=1_048_576).contains(&row))
}

fn valid_a1_range(value: &str) -> bool {
    let mut parts = value.split(':');
    let Some(first) = parts.next() else {
        return false;
    };
    let Some(last) = parts.next() else {
        return valid_a1_cell(first);
    };
    parts.next().is_none() && valid_a1_cell(first) && valid_a1_cell(last)
}

fn source_range_starts_at_a1(source_range: &str) -> bool {
    source_range
        .split_once(':')
        .map(|(start, _)| start)
        .unwrap_or(source_range)
        .trim()
        .eq_ignore_ascii_case("A1")
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct WorkbookProfile {
    pub format: WorkbookFormat,
    /// True only when the adapter has completed the package/object inventory
    /// needed to match this request against backend evidence.
    pub inventory_complete: bool,
    pub features_present: BTreeSet<WorkbookFeature>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ExcelEditPlan {
    pub request_id: String,
    pub input: PathBuf,
    pub output: PathBuf,
    pub format: WorkbookFormat,
    pub workbook: WorkbookProfile,
    pub operations: Vec<EditOperation>,
    /// Properties required to survive outside any sheet-specific exceptions
    /// listed in `drop_allowed`.
    pub preserve: BTreeSet<PreservationRequirement>,
    /// Intentional property losses keyed by worksheet name. A drop on one
    /// worksheet does not authorize the same loss on untouched worksheets.
    #[serde(default)]
    pub drop_allowed: BTreeMap<String, BTreeSet<PreservationRequirement>>,
}

impl ExcelEditPlan {
    pub fn validate(&self) -> Result<()> {
        if self.request_id.trim().is_empty() {
            bail!("edit plan request_id must not be empty");
        }
        if self.input.as_os_str().is_empty() || self.output.as_os_str().is_empty() {
            bail!("edit plan input and output paths must not be empty");
        }
        if self.format != self.workbook.format {
            bail!("edit plan format does not match the inspected workbook format");
        }
        if self.operations.is_empty() {
            bail!("edit plan must contain at least one operation");
        }
        for operation in &self.operations {
            operation.validate()?;
        }
        let calc_property_operations = self
            .operations
            .iter()
            .filter(|operation| {
                matches!(
                    operation,
                    EditOperation::SetWorkbookCalculationProperties { .. }
                )
            })
            .count();
        if calc_property_operations > 1 {
            bail!("an edit plan may contain at most one workbook calculation-property operation");
        }
        if calc_property_operations > 0 && self.format != WorkbookFormat::Xlsx {
            bail!("workbook calculation-property patch currently supports XLSX only");
        }
        if self
            .drop_allowed
            .keys()
            .any(|sheet| sheet.trim().is_empty())
        {
            bail!("drop_allowed worksheet names must not be empty");
        }
        for (sheet, allowed) in &self.drop_allowed {
            // Legacy allowances describe property loss on worksheets the plan
            // deletes, so they keep requiring an explicit DeleteSheet. The
            // merged non-anchor payload allowance covers engine rewrite loss
            // on retained worksheets and stands alone.
            let requires_delete_sheet = allowed.iter().any(|requirement| {
                *requirement != PreservationRequirement::MergedNonAnchorPayloads
            });
            if requires_delete_sheet
                && !self.operations.iter().any(
                    |operation| matches!(operation, EditOperation::DeleteSheet { name } if name == sheet),
                )
            {
                bail!("drop_allowed for worksheet {sheet:?} requires an explicit DeleteSheet operation");
            }
        }
        Ok(())
    }

    fn operation_kinds(&self) -> BTreeSet<EditOperationKind> {
        self.operations.iter().map(EditOperation::kind).collect()
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum EngineId {
    OoxmlPatch,
    UmyaSpreadsheet,
    ExcelizeWorker,
    OpenpyxlWorker,
}

impl std::fmt::Display for EngineId {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::OoxmlPatch => "ooxml_patch",
            Self::UmyaSpreadsheet => "umya_spreadsheet",
            Self::ExcelizeWorker => "excelize_worker",
            Self::OpenpyxlWorker => "openpyxl_worker",
        })
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum CapabilityStatus {
    Passed,
    Unverified,
    Unsupported,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct CapabilityEvidence {
    pub sample_id: String,
    pub report_path: String,
    pub adapter_version: String,
    pub engine_version: String,
    pub platform: String,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct BackendCapability {
    pub engine: EngineId,
    pub status: CapabilityStatus,
    pub format: WorkbookFormat,
    pub operations: BTreeSet<EditOperationKind>,
    pub supported_features: BTreeSet<WorkbookFeature>,
    pub preserved: BTreeSet<PreservationRequirement>,
    pub platform: String,
    pub evidence: Option<CapabilityEvidence>,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct BackendRouteDecision {
    pub engine: EngineId,
    pub eligible: bool,
    pub rejected_reasons: Vec<String>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum BackendErrorKind {
    /// The backend completed preflight and rejected the file before writing.
    /// The modifier may try another already-qualified backend from the original.
    UnsupportedBeforeWrite,
    /// Includes any error after writes may have started. Never auto-fallback.
    Failed,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BackendFailure {
    pub kind: BackendErrorKind,
    pub message: String,
}

impl BackendFailure {
    pub fn unsupported_before_write(message: impl Into<String>) -> Self {
        Self {
            kind: BackendErrorKind::UnsupportedBeforeWrite,
            message: message.into(),
        }
    }

    pub fn failed(message: impl Into<String>) -> Self {
        Self {
            kind: BackendErrorKind::Failed,
            message: message.into(),
        }
    }
}

pub trait EditBackend: Send + Sync {
    fn capability(&self) -> &BackendCapability;
    /// The plan's input is a fresh temporary copy of the original file and its
    /// output is a fresh temporary path. Implementations must not edit either
    /// path in place. `UnsupportedBeforeWrite` must be returned before changing
    /// the output path.
    fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure>;

    /// Optional backend-specific validation. This runs after the worker writes
    /// its candidate and before the modifier's shared output validator. Any
    /// failure is terminal; it must never trigger another engine.
    fn validate_output(&self, _plan: &ExcelEditPlan, _candidate_output: &Path) -> Result<()> {
        Ok(())
    }
}

pub trait OutputValidator: Send + Sync {
    /// Validate the candidate output against the original request contract.
    fn validate(&self, original_plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()>;
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct EditExecutionReport {
    pub engine: EngineId,
    pub input: PathBuf,
    pub output: PathBuf,
    pub elapsed_ms: u128,
    pub route_decisions: Vec<BackendRouteDecision>,
    pub attempts: Vec<BackendAttempt>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum BackendAttemptOutcome {
    Succeeded,
    RejectedBeforeWrite,
    Failed,
    ValidationFailed,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct BackendAttempt {
    pub engine: EngineId,
    pub outcome: BackendAttemptOutcome,
    pub message: String,
    pub elapsed_ms: u128,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ExcelModifierErrorKind {
    InvalidPlan,
    InvalidPaths,
    NoEligibleBackend,
    PreparationFailed,
    BackendFailed,
    BackendProtocolViolation,
    OutputValidationFailed,
    OutputCommitFailed,
    AllBackendsRejectedBeforeWrite,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct ExcelModifierError {
    pub kind: ExcelModifierErrorKind,
    pub engine: Option<EngineId>,
    pub message: String,
    pub route_decisions: Vec<BackendRouteDecision>,
    pub attempts: Vec<BackendAttempt>,
}

impl std::fmt::Display for ExcelModifierError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "{:?}: {}", self.kind, self.message)
    }
}

impl std::error::Error for ExcelModifierError {}

fn execution_error(
    kind: ExcelModifierErrorKind,
    engine: Option<EngineId>,
    message: impl Into<String>,
    route_decisions: &[BackendRouteDecision],
    attempts: &[BackendAttempt],
) -> ExcelModifierError {
    ExcelModifierError {
        kind,
        engine,
        message: message.into(),
        route_decisions: route_decisions.to_vec(),
        attempts: attempts.to_vec(),
    }
}

/// Routes one complete operation batch to one backend and commits a validated
/// output without overwriting an existing destination.
pub struct ExcelModifier {
    platform: String,
    backends: Vec<Arc<dyn EditBackend>>,
    validator: Arc<dyn OutputValidator>,
}

impl ExcelModifier {
    /// Backend order is explicit configuration, not a language-based ranking.
    pub fn new(
        platform: impl Into<String>,
        backends: Vec<Arc<dyn EditBackend>>,
        validator: Arc<dyn OutputValidator>,
    ) -> Result<Self> {
        let platform = platform.into();
        if platform.trim().is_empty() {
            bail!("runtime platform identifier must not be empty");
        }
        Ok(Self {
            platform,
            backends,
            validator,
        })
    }

    pub fn route(&self, plan: &ExcelEditPlan) -> Vec<BackendRouteDecision> {
        self.backends
            .iter()
            .map(|backend| {
                let capability = backend.capability();
                let mut rejected_reasons = Vec::new();
                if capability.status != CapabilityStatus::Passed {
                    rejected_reasons.push(format!("capability status is {:?}", capability.status));
                }
                if capability.format != plan.format {
                    rejected_reasons.push("input format is not covered".to_owned());
                }
                if !plan.workbook.inventory_complete {
                    rejected_reasons.push("workbook feature inventory is incomplete".to_owned());
                }
                if !plan.operation_kinds().is_subset(&capability.operations) {
                    rejected_reasons
                        .push("backend does not cover the whole operation batch".to_owned());
                }
                if !plan
                    .workbook
                    .features_present
                    .is_subset(&capability.supported_features)
                {
                    rejected_reasons
                        .push("present workbook objects exceed tested scope".to_owned());
                }
                if !plan.preserve.is_subset(&capability.preserved) {
                    rejected_reasons
                        .push("required preservation contract is not covered".to_owned());
                }
                if capability.platform != self.platform {
                    rejected_reasons.push("runtime platform is not covered".to_owned());
                }
                match &capability.evidence {
                    Some(evidence)
                        if !evidence.sample_id.trim().is_empty()
                            && !evidence.report_path.trim().is_empty()
                            && !evidence.adapter_version.trim().is_empty()
                            && !evidence.engine_version.trim().is_empty()
                            && evidence.platform == self.platform => {}
                    _ => rejected_reasons
                        .push("passed capability has no matching evidence".to_owned()),
                }
                BackendRouteDecision {
                    engine: capability.engine,
                    eligible: rejected_reasons.is_empty(),
                    rejected_reasons,
                }
            })
            .collect()
    }

    pub fn execute(
        &self,
        plan: &ExcelEditPlan,
    ) -> std::result::Result<EditExecutionReport, ExcelModifierError> {
        plan.validate().map_err(|error| {
            execution_error(
                ExcelModifierErrorKind::InvalidPlan,
                None,
                error.to_string(),
                &[],
                &[],
            )
        })?;
        validate_paths(plan).map_err(|error| {
            execution_error(
                ExcelModifierErrorKind::InvalidPaths,
                None,
                error.to_string(),
                &[],
                &[],
            )
        })?;
        let started = Instant::now();
        let route = self.route(plan);
        let eligible: Vec<_> = self
            .backends
            .iter()
            .zip(route.iter())
            .filter(|(_, decision)| decision.eligible)
            .map(|(backend, decision)| (backend, decision.engine))
            .collect();
        if eligible.is_empty() {
            let details = route
                .iter()
                .map(|decision| {
                    format!(
                        "{}: {}",
                        decision.engine,
                        decision.rejected_reasons.join(", ")
                    )
                })
                .collect::<Vec<_>>()
                .join("; ");
            return Err(execution_error(
                ExcelModifierErrorKind::NoEligibleBackend,
                None,
                format!("no qualified backend; {details}"),
                &route,
                &[],
            ));
        }

        let output_parent = parent_or_current(&plan.output);
        let extension = plan.format.extension();
        let mut attempts = Vec::new();
        for (backend, engine) in eligible {
            let attempt_started = Instant::now();
            // Every attempt receives a new copy of the original, never the
            // preceding backend's input or output.
            let input_snapshot =
                TempArtifact::new(output_parent, "source", extension).map_err(|error| {
                    execution_error(
                        ExcelModifierErrorKind::PreparationFailed,
                        Some(engine),
                        format!("create source snapshot: {error}"),
                        &route,
                        &attempts,
                    )
                })?;
            fs::copy(&plan.input, input_snapshot.path()).map_err(|error| {
                execution_error(
                    ExcelModifierErrorKind::PreparationFailed,
                    Some(engine),
                    format!("copy original input for {engine} attempt: {error}"),
                    &route,
                    &attempts,
                )
            })?;
            let output_candidate = TempArtifact::new(output_parent, "candidate", extension)
                .map_err(|error| {
                    execution_error(
                        ExcelModifierErrorKind::PreparationFailed,
                        Some(engine),
                        format!("create candidate output: {error}"),
                        &route,
                        &attempts,
                    )
                })?;
            let mut attempt_plan = plan.clone();
            attempt_plan.input = input_snapshot.path().to_path_buf();
            attempt_plan.output = output_candidate.path().to_path_buf();

            match backend.apply(&attempt_plan) {
                Ok(()) => {}
                Err(failure) if failure.kind == BackendErrorKind::UnsupportedBeforeWrite => {
                    let metadata = fs::metadata(output_candidate.path()).map_err(|error| {
                        attempts.push(BackendAttempt {
                            engine,
                            outcome: BackendAttemptOutcome::Failed,
                            message: format!("could not inspect rejected candidate: {error}"),
                            elapsed_ms: attempt_started.elapsed().as_millis(),
                        });
                        execution_error(
                            ExcelModifierErrorKind::BackendProtocolViolation,
                            Some(engine),
                            "backend removed or replaced its candidate before rejecting",
                            &route,
                            &attempts,
                        )
                    })?;
                    if metadata.len() != 0 {
                        let message = format!(
                            "backend wrote {} bytes before reporting UnsupportedBeforeWrite",
                            metadata.len()
                        );
                        attempts.push(BackendAttempt {
                            engine,
                            outcome: BackendAttemptOutcome::Failed,
                            message: message.clone(),
                            elapsed_ms: attempt_started.elapsed().as_millis(),
                        });
                        return Err(execution_error(
                            ExcelModifierErrorKind::BackendProtocolViolation,
                            Some(engine),
                            message,
                            &route,
                            &attempts,
                        ));
                    }
                    attempts.push(BackendAttempt {
                        engine,
                        outcome: BackendAttemptOutcome::RejectedBeforeWrite,
                        message: failure.message,
                        elapsed_ms: attempt_started.elapsed().as_millis(),
                    });
                    continue;
                }
                Err(failure) => {
                    attempts.push(BackendAttempt {
                        engine,
                        outcome: BackendAttemptOutcome::Failed,
                        message: failure.message.clone(),
                        elapsed_ms: attempt_started.elapsed().as_millis(),
                    });
                    return Err(execution_error(
                        ExcelModifierErrorKind::BackendFailed,
                        Some(engine),
                        format!(
                            "backend failed; automatic fallback stopped: {}",
                            failure.message
                        ),
                        &route,
                        &attempts,
                    ));
                }
            }

            let metadata = fs::metadata(output_candidate.path()).map_err(|error| {
                attempts.push(BackendAttempt {
                    engine,
                    outcome: BackendAttemptOutcome::Failed,
                    message: format!("candidate output is missing: {error}"),
                    elapsed_ms: attempt_started.elapsed().as_millis(),
                });
                execution_error(
                    ExcelModifierErrorKind::BackendFailed,
                    Some(engine),
                    "backend did not create its candidate output",
                    &route,
                    &attempts,
                )
            })?;
            if !metadata.is_file() || metadata.len() == 0 {
                let message = "backend produced an empty or non-file output".to_owned();
                attempts.push(BackendAttempt {
                    engine,
                    outcome: BackendAttemptOutcome::Failed,
                    message: message.clone(),
                    elapsed_ms: attempt_started.elapsed().as_millis(),
                });
                return Err(execution_error(
                    ExcelModifierErrorKind::BackendFailed,
                    Some(engine),
                    message,
                    &route,
                    &attempts,
                ));
            }
            if let Err(error) = backend.validate_output(plan, output_candidate.path()) {
                let message = error.to_string();
                attempts.push(BackendAttempt {
                    engine,
                    outcome: BackendAttemptOutcome::ValidationFailed,
                    message: message.clone(),
                    elapsed_ms: attempt_started.elapsed().as_millis(),
                });
                return Err(execution_error(
                    ExcelModifierErrorKind::OutputValidationFailed,
                    Some(engine),
                    format!("backend output validation failed; fallback stopped: {message}"),
                    &route,
                    &attempts,
                ));
            }
            if let Err(error) = self.validator.validate(plan, output_candidate.path()) {
                let message = error.to_string();
                attempts.push(BackendAttempt {
                    engine,
                    outcome: BackendAttemptOutcome::ValidationFailed,
                    message: message.clone(),
                    elapsed_ms: attempt_started.elapsed().as_millis(),
                });
                return Err(execution_error(
                    ExcelModifierErrorKind::OutputValidationFailed,
                    Some(engine),
                    format!("output validation failed; fallback stopped: {message}"),
                    &route,
                    &attempts,
                ));
            }
            OpenOptions::new()
                .write(true)
                .open(output_candidate.path())
                .and_then(|file| file.sync_all())
                .map_err(|error| {
                    attempts.push(BackendAttempt {
                        engine,
                        outcome: BackendAttemptOutcome::Failed,
                        message: format!("sync validated output: {error}"),
                        elapsed_ms: attempt_started.elapsed().as_millis(),
                    });
                    execution_error(
                        ExcelModifierErrorKind::OutputCommitFailed,
                        Some(engine),
                        format!("sync validated output: {error}"),
                        &route,
                        &attempts,
                    )
                })?;

            // A same-directory hard link is an atomic no-clobber publication:
            // if another process creates the destination meanwhile, commit fails.
            fs::hard_link(output_candidate.path(), &plan.output).map_err(|error| {
                attempts.push(BackendAttempt {
                    engine,
                    outcome: BackendAttemptOutcome::Failed,
                    message: format!("publish output without overwriting: {error}"),
                    elapsed_ms: attempt_started.elapsed().as_millis(),
                });
                execution_error(
                    ExcelModifierErrorKind::OutputCommitFailed,
                    Some(engine),
                    format!("publish output without overwriting: {error}"),
                    &route,
                    &attempts,
                )
            })?;
            attempts.push(BackendAttempt {
                engine,
                outcome: BackendAttemptOutcome::Succeeded,
                message: "validated output published".to_owned(),
                elapsed_ms: attempt_started.elapsed().as_millis(),
            });
            return Ok(EditExecutionReport {
                engine,
                input: plan.input.clone(),
                output: plan.output.clone(),
                elapsed_ms: started.elapsed().as_millis(),
                route_decisions: route,
                attempts,
            });
        }

        let details = attempts
            .iter()
            .map(|attempt| format!("{}: {}", attempt.engine, attempt.message))
            .collect::<Vec<_>>()
            .join("; ");
        Err(execution_error(
            ExcelModifierErrorKind::AllBackendsRejectedBeforeWrite,
            None,
            format!("all qualified backends rejected the file before writing: {details}"),
            &route,
            &attempts,
        ))
    }
}

fn validate_paths(plan: &ExcelEditPlan) -> Result<()> {
    if !plan.input.is_file() {
        bail!(
            "input workbook does not exist or is not a file: {}",
            plan.input.display()
        );
    }
    let extension_matches = |path: &Path| {
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case(plan.format.extension()))
    };
    if !extension_matches(&plan.input) || !extension_matches(&plan.output) {
        bail!(
            "ExcelModifier requires matching input/output extensions for {:?}",
            plan.format
        );
    }
    let output_parent = parent_or_current(&plan.output);
    if !output_parent.is_dir() {
        bail!(
            "output directory does not exist: {}",
            output_parent.display()
        );
    }
    let canonical_input = fs::canonicalize(&plan.input).context("resolve input workbook path")?;
    let output_name = plan
        .output
        .file_name()
        .ok_or_else(|| anyhow!("output path must include a filename"))?;
    let canonical_parent = fs::canonicalize(output_parent).context("resolve output directory")?;
    let canonical_output = canonical_parent.join(output_name);
    if same_path(&canonical_input, &canonical_output) {
        bail!("ExcelModifier never edits the original input in place");
    }
    if plan.output.exists() {
        bail!(
            "output already exists; refusing to overwrite: {}",
            plan.output.display()
        );
    }
    Ok(())
}

fn same_path(left: &Path, right: &Path) -> bool {
    if cfg!(windows) {
        left.to_string_lossy()
            .eq_ignore_ascii_case(&right.to_string_lossy())
    } else {
        left == right
    }
}

fn parent_or_current(path: &Path) -> &Path {
    path.parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."))
}

struct TempArtifact(PathBuf);

impl TempArtifact {
    fn new(parent: &Path, purpose: &str, extension: &str) -> Result<Self> {
        for _ in 0..128 {
            let sequence = TEMP_SEQUENCE.fetch_add(1, Ordering::Relaxed);
            let name = format!(
                ".excel-edit-{}-{sequence}-{purpose}.tmp.{extension}",
                std::process::id()
            );
            let path = parent.join(name);
            match OpenOptions::new().write(true).create_new(true).open(&path) {
                Ok(file) => {
                    drop(file);
                    return Ok(Self(path));
                }
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => {
                    return Err(error).with_context(|| format!("create {} temporary file", purpose))
                }
            }
        }
        bail!("could not reserve a unique temporary {purpose} path")
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for TempArtifact {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.0);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::{fs::File, io::Write, sync::Mutex};

    static TEST_DIR_SEQUENCE: AtomicU64 = AtomicU64::new(0);

    struct TestDir(PathBuf);

    impl TestDir {
        fn new() -> Self {
            let sequence = TEST_DIR_SEQUENCE.fetch_add(1, Ordering::Relaxed);
            let path = std::env::temp_dir().join(format!(
                "excel-editor-plan-{}-{sequence}",
                std::process::id()
            ));
            fs::create_dir_all(&path).unwrap();
            Self(path)
        }
    }

    impl Drop for TestDir {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    #[derive(Clone, Copy)]
    enum Action {
        UnsupportedBeforeWrite,
        Write,
        Fail,
    }

    struct FakeBackend {
        capability: BackendCapability,
        action: Action,
        observed_inputs: Arc<Mutex<Vec<(PathBuf, Vec<u8>)>>>,
    }

    impl EditBackend for FakeBackend {
        fn capability(&self) -> &BackendCapability {
            &self.capability
        }

        fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure> {
            let input =
                fs::read(&plan.input).map_err(|error| BackendFailure::failed(error.to_string()))?;
            self.observed_inputs
                .lock()
                .unwrap()
                .push((plan.input.clone(), input.clone()));
            match self.action {
                Action::UnsupportedBeforeWrite => Err(BackendFailure::unsupported_before_write(
                    "synthetic unsupported workbook trait",
                )),
                Action::Write => {
                    let mut output = File::create(&plan.output)
                        .map_err(|error| BackendFailure::failed(error.to_string()))?;
                    output
                        .write_all(b"edited:")
                        .and_then(|_| output.write_all(&input))
                        .map_err(|error| BackendFailure::failed(error.to_string()))
                }
                Action::Fail => Err(BackendFailure::failed("synthetic write failure")),
            }
        }
    }

    struct PrefixValidator;

    impl OutputValidator for PrefixValidator {
        fn validate(&self, _plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
            let bytes = fs::read(candidate_output)?;
            if bytes.starts_with(b"edited:") {
                Ok(())
            } else {
                bail!("candidate does not contain expected marker")
            }
        }
    }

    struct RejectValidator;

    impl OutputValidator for RejectValidator {
        fn validate(&self, _plan: &ExcelEditPlan, _candidate_output: &Path) -> Result<()> {
            bail!("synthetic preservation difference")
        }
    }

    fn plan(directory: &Path, operations: Vec<EditOperation>) -> ExcelEditPlan {
        let input = directory.join("input.xlsx");
        fs::write(&input, b"original workbook bytes").unwrap();
        ExcelEditPlan {
            request_id: "test-request".to_owned(),
            input,
            output: directory.join("output.xlsx"),
            format: WorkbookFormat::Xlsx,
            workbook: WorkbookProfile {
                format: WorkbookFormat::Xlsx,
                inventory_complete: true,
                features_present: BTreeSet::new(),
            },
            operations,
            preserve: [PreservationRequirement::UnmodifiedPackageParts]
                .into_iter()
                .collect(),
            drop_allowed: BTreeMap::new(),
        }
    }

    fn set_value() -> EditOperation {
        EditOperation::SetCellValue {
            sheet: "Data".to_owned(),
            cell: "A1".to_owned(),
            value: CellValue::Text("new value".to_owned()),
        }
    }

    #[test]
    fn merged_non_anchor_payload_allowance_stands_without_delete_sheet() {
        let directory = TestDir::new();
        let directory = directory.0.as_path();
        let mut payload_plan = plan(directory, vec![set_value()]);
        payload_plan.drop_allowed.insert(
            "汇总".to_owned(),
            [PreservationRequirement::MergedNonAnchorPayloads]
                .into_iter()
                .collect(),
        );
        assert!(payload_plan.validate().is_ok());

        // Mixing a legacy allowance into the same key brings back the
        // DeleteSheet requirement; existing semantics stay unchanged.
        let mut mixed = plan(directory, vec![set_value()]);
        mixed.drop_allowed.insert(
            "汇总".to_owned(),
            [
                PreservationRequirement::MergedNonAnchorPayloads,
                PreservationRequirement::DefinedNames,
            ]
            .into_iter()
            .collect(),
        );
        assert!(mixed.validate().is_err());
    }

    #[test]
    fn sheet_layout_operations_validate_a1_references() {
        assert!(EditOperation::SetFreezePanes {
            sheet: "核查表校验结果".to_owned(),
            cell: Some("A2".to_owned()),
        }
        .validate()
        .is_ok());
        assert!(EditOperation::SetAutoFilter {
            sheet: "核查表校验结果".to_owned(),
            range: Some("A1:M42".to_owned()),
        }
        .validate()
        .is_ok());
        assert!(EditOperation::SetFreezePanes {
            sheet: "Data".to_owned(),
            cell: Some("A0".to_owned()),
        }
        .validate()
        .is_err());
        assert!(EditOperation::SetAutoFilter {
            sheet: "Data".to_owned(),
            range: Some("A1:M0".to_owned()),
        }
        .validate()
        .is_err());
    }

    fn capability(
        engine: EngineId,
        operations: impl IntoIterator<Item = EditOperationKind>,
    ) -> BackendCapability {
        BackendCapability {
            engine,
            status: CapabilityStatus::Passed,
            format: WorkbookFormat::Xlsx,
            operations: operations.into_iter().collect(),
            supported_features: BTreeSet::new(),
            preserved: [PreservationRequirement::UnmodifiedPackageParts]
                .into_iter()
                .collect(),
            platform: "test-linux-x86_64".to_owned(),
            evidence: Some(CapabilityEvidence {
                sample_id: "synthetic-workbook".to_owned(),
                report_path: "probe/reports/test.json".to_owned(),
                adapter_version: "adapter-test".to_owned(),
                engine_version: "engine-test".to_owned(),
                platform: "test-linux-x86_64".to_owned(),
            }),
        }
    }

    fn backend(
        capability: BackendCapability,
        action: Action,
        observed_inputs: Arc<Mutex<Vec<(PathBuf, Vec<u8>)>>>,
    ) -> Arc<dyn EditBackend> {
        Arc::new(FakeBackend {
            capability,
            action,
            observed_inputs,
        })
    }

    #[test]
    fn only_passed_exact_capabilities_are_routable() {
        let directory = TestDir::new();
        let plan = plan(&directory.0, vec![set_value()]);
        let observed = Arc::new(Mutex::new(Vec::new()));
        let mut unverified =
            capability(EngineId::UmyaSpreadsheet, [EditOperationKind::SetCellValue]);
        unverified.status = CapabilityStatus::Unverified;
        let mut unsupported =
            capability(EngineId::ExcelizeWorker, [EditOperationKind::SetCellValue]);
        unsupported.status = CapabilityStatus::Unsupported;
        let modifier = ExcelModifier::new(
            "test-linux-x86_64",
            vec![
                backend(unverified, Action::Write, Arc::clone(&observed)),
                backend(unsupported, Action::Write, observed),
            ],
            Arc::new(PrefixValidator),
        )
        .unwrap();

        assert!(modifier
            .route(&plan)
            .iter()
            .all(|decision| !decision.eligible));
    }

    #[test]
    fn one_backend_must_cover_the_entire_operation_batch() {
        let directory = TestDir::new();
        let plan = plan(
            &directory.0,
            vec![
                set_value(),
                EditOperation::SetComment {
                    sheet: "Data".to_owned(),
                    cell: "A1".to_owned(),
                    text: Some("note".to_owned()),
                    author: Some("test".to_owned()),
                },
            ],
        );
        let observed = Arc::new(Mutex::new(Vec::new()));
        let modifier = ExcelModifier::new(
            "test-linux-x86_64",
            vec![
                backend(
                    capability(EngineId::OoxmlPatch, [EditOperationKind::SetCellValue]),
                    Action::Write,
                    Arc::clone(&observed),
                ),
                backend(
                    capability(EngineId::UmyaSpreadsheet, [EditOperationKind::SetComment]),
                    Action::Write,
                    observed,
                ),
            ],
            Arc::new(PrefixValidator),
        )
        .unwrap();

        assert!(modifier
            .route(&plan)
            .iter()
            .all(|decision| !decision.eligible));
        assert_eq!(
            modifier.execute(&plan).unwrap_err().kind,
            ExcelModifierErrorKind::NoEligibleBackend
        );
        assert!(!plan.output.exists());
    }

    #[test]
    fn fallback_uses_a_fresh_copy_of_the_original_and_publishes_only_validated_output() {
        let directory = TestDir::new();
        let plan = plan(&directory.0, vec![set_value()]);
        let original_bytes = fs::read(&plan.input).unwrap();
        let observed = Arc::new(Mutex::new(Vec::new()));
        let modifier = ExcelModifier::new(
            "test-linux-x86_64",
            vec![
                backend(
                    capability(EngineId::OoxmlPatch, [EditOperationKind::SetCellValue]),
                    Action::UnsupportedBeforeWrite,
                    Arc::clone(&observed),
                ),
                backend(
                    capability(EngineId::UmyaSpreadsheet, [EditOperationKind::SetCellValue]),
                    Action::Write,
                    Arc::clone(&observed),
                ),
            ],
            Arc::new(PrefixValidator),
        )
        .unwrap();

        let report = modifier.execute(&plan).unwrap();
        let attempts = observed.lock().unwrap();
        assert_eq!(attempts.len(), 2);
        assert_ne!(attempts[0].0, plan.input);
        assert_ne!(attempts[1].0, plan.input);
        assert_ne!(attempts[0].0, attempts[1].0);
        assert_eq!(attempts[0].1, original_bytes);
        assert_eq!(attempts[1].1, original_bytes);
        assert_eq!(fs::read(&plan.input).unwrap(), original_bytes);
        assert_eq!(
            fs::read(&plan.output).unwrap(),
            b"edited:original workbook bytes"
        );
        assert_eq!(report.engine, EngineId::UmyaSpreadsheet);
        assert_eq!(report.attempts.len(), 2);
        assert_eq!(
            report.attempts[0].outcome,
            BackendAttemptOutcome::RejectedBeforeWrite
        );
        assert_eq!(report.attempts[1].outcome, BackendAttemptOutcome::Succeeded);
        assert_eq!(report.route_decisions.len(), 2);
        assert!(!attempts.iter().any(|(path, _)| path.exists()));
    }

    #[test]
    fn output_validation_failure_stops_without_trying_another_engine() {
        let directory = TestDir::new();
        let plan = plan(&directory.0, vec![set_value()]);
        let observed = Arc::new(Mutex::new(Vec::new()));
        let modifier = ExcelModifier::new(
            "test-linux-x86_64",
            vec![
                backend(
                    capability(EngineId::OoxmlPatch, [EditOperationKind::SetCellValue]),
                    Action::Write,
                    Arc::clone(&observed),
                ),
                backend(
                    capability(EngineId::UmyaSpreadsheet, [EditOperationKind::SetCellValue]),
                    Action::Write,
                    Arc::clone(&observed),
                ),
            ],
            Arc::new(RejectValidator),
        )
        .unwrap();

        let error = modifier.execute(&plan).unwrap_err();
        assert_eq!(error.kind, ExcelModifierErrorKind::OutputValidationFailed);
        assert_eq!(error.attempts.len(), 1);
        assert_eq!(
            error.attempts[0].outcome,
            BackendAttemptOutcome::ValidationFailed
        );
        assert_eq!(observed.lock().unwrap().len(), 1);
        assert!(!plan.output.exists());
        assert_eq!(fs::read(&plan.input).unwrap(), b"original workbook bytes");
    }

    #[test]
    fn failure_after_write_may_have_started_never_falls_back() {
        let directory = TestDir::new();
        let plan = plan(&directory.0, vec![set_value()]);
        let observed = Arc::new(Mutex::new(Vec::new()));
        let modifier = ExcelModifier::new(
            "test-linux-x86_64",
            vec![
                backend(
                    capability(EngineId::OoxmlPatch, [EditOperationKind::SetCellValue]),
                    Action::Fail,
                    Arc::clone(&observed),
                ),
                backend(
                    capability(EngineId::UmyaSpreadsheet, [EditOperationKind::SetCellValue]),
                    Action::Write,
                    Arc::clone(&observed),
                ),
            ],
            Arc::new(PrefixValidator),
        )
        .unwrap();

        let error = modifier.execute(&plan).unwrap_err();
        assert_eq!(error.kind, ExcelModifierErrorKind::BackendFailed);
        assert_eq!(error.attempts.len(), 1);
        assert_eq!(error.attempts[0].outcome, BackendAttemptOutcome::Failed);
        assert_eq!(observed.lock().unwrap().len(), 1);
        assert!(!plan.output.exists());
        assert_eq!(fs::read(&plan.input).unwrap(), b"original workbook bytes");
    }

    #[test]
    fn existing_destination_is_never_overwritten() {
        let directory = TestDir::new();
        let plan = plan(&directory.0, vec![set_value()]);
        fs::write(&plan.output, b"existing output").unwrap();
        let observed = Arc::new(Mutex::new(Vec::new()));
        let modifier = ExcelModifier::new(
            "test-linux-x86_64",
            vec![backend(
                capability(EngineId::UmyaSpreadsheet, [EditOperationKind::SetCellValue]),
                Action::Write,
                Arc::clone(&observed),
            )],
            Arc::new(PrefixValidator),
        )
        .unwrap();

        assert_eq!(
            modifier.execute(&plan).unwrap_err().kind,
            ExcelModifierErrorKind::InvalidPaths
        );
        assert!(observed.lock().unwrap().is_empty());
        assert_eq!(fs::read(&plan.output).unwrap(), b"existing output");
    }
}
