//! Shared primitives for editing existing XLSX workbooks.

use anyhow::{Context, Result};
use std::path::Path;

mod ooxml_header_backend;
mod openpyxl_backend;
mod plan;
mod umya_backend;
mod workbook_profile;

pub use ooxml_header_backend::{OoxmlHeaderTextBackend, OoxmlHeaderTextValidator};
pub use openpyxl_backend::{OpenpyxlWorkerConfig, OpenpyxlXlsxBackend, OpenpyxlXlsxValidator};
pub use plan::{
    BackendAttempt, BackendAttemptOutcome, BackendCapability, BackendErrorKind, BackendFailure,
    BackendRouteDecision, CapabilityEvidence, CapabilityStatus, CellValue, EditBackend,
    EditExecutionReport, EditOperation, EditOperationKind, EngineId, ExcelEditPlan, ExcelModifier,
    ExcelModifierError, ExcelModifierErrorKind, OutputValidator, PageOrientation, PageSettings,
    PreservationRequirement, RangeStyleMode, WorkbookFeature, WorkbookFormat, WorkbookProfile,
};
pub use umya_backend::{UmyaXlsxBackend, UmyaXlsxValidator};
pub use umya_spreadsheet::{BorderStyleValues, Cell, Workbook};
pub use workbook_profile::inspect_xlsx_profile;

/// Opens an existing XLSX workbook with umya-spreadsheet.
pub fn open_xlsx(path: impl AsRef<Path>) -> Result<Workbook> {
    let path = path.as_ref();
    umya_spreadsheet::reader::xlsx::read(path)
        .with_context(|| format!("open workbook {}", path.display()))
}

/// Saves an edited workbook to a new or existing path.
pub fn save_xlsx(workbook: &Workbook, path: impl AsRef<Path>) -> Result<()> {
    let path = path.as_ref();
    umya_spreadsheet::writer::xlsx::write(workbook, path)
        .with_context(|| format!("save workbook {}", path.display()))
}
