//! Narrow, package-preserving OOXML edits for simple values and print metadata.
//!
//! The adapter rewrites only the worksheet/workbook XML parts named by a plan.
//! ZIP members outside those parts are copied without decoding. Cell edits are
//! intentionally conservative: cells carrying OOXML attributes or children
//! outside the plain value/formula shape are rejected before the candidate is
//! written.

use crate::{
    BackendCapability, BackendFailure, EditBackend, EditOperation, EditOperationKind,
    ExcelEditPlan, OutputValidator, WorkbookFormat,
};
use anyhow::{bail, ensure, Context, Result};
use quick_xml::{
    events::{BytesEnd, BytesStart, BytesText, Event},
    Writer,
};
use roxmltree::{Document, Node};
use std::{
    collections::{BTreeMap, BTreeSet, HashMap},
    fs::{self, OpenOptions},
    io::{Cursor, Read, Seek, Write},
    ops::Range,
    path::Path,
};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

#[derive(Clone)]
struct WorksheetPart {
    name: String,
    part: String,
    workbook_index: usize,
}

#[derive(Clone)]
struct CellUpdate {
    row: u32,
    column: u32,
    address: String,
    value: crate::CellValue,
}

#[derive(Clone)]
struct DefinedNameUpdate {
    name: String,
    formula: Option<String>,
    local_sheet_id: Option<usize>,
}

#[derive(Default)]
struct PackageEdits {
    worksheet_cells: BTreeMap<String, Vec<CellUpdate>>,
    worksheet_page_settings: Vec<(String, crate::PageSettings)>,
    worksheet_margins: Vec<(String, crate::PageMargins)>,
    worksheet_headers: Vec<(String, String, String)>,
    defined_names: Vec<DefinedNameUpdate>,
    calculation_properties: Option<(Option<bool>, Option<bool>)>,
}

pub struct RustNativeOoxmlBackend {
    capability: BackendCapability,
}

impl RustNativeOoxmlBackend {
    /// The caller supplies operation and preservation evidence. The adapter
    /// never changes an unverified capability to Passed on its own.
    pub fn new(capability: BackendCapability) -> Result<Self> {
        ensure!(
            capability.engine == crate::EngineId::RustNativeOoxml,
            "RustNative OOXML adapter requires a rust_native_ooxml capability"
        );
        Ok(Self { capability })
    }

    pub fn apply_bytes(source: &[u8], plan: &ExcelEditPlan) -> Result<Vec<u8>> {
        let mut archive =
            ZipArchive::new(Cursor::new(source)).context("open source XLSX package")?;
        let (workbook_part, sheets) = resolve_worksheet_parts(&mut archive)?;
        let mut edits = edits_from_plan(plan, &sheets)?;
        let mut replacements = BTreeMap::<String, Vec<u8>>::new();

        for (sheet_name, cell_updates) in &mut edits.worksheet_cells {
            cell_updates.sort_by_key(|cell| (cell.row, cell.column));
            ensure_unique_cell_updates(sheet_name, cell_updates)?;
            let worksheet = sheets
                .iter()
                .find(|sheet| sheet.name == *sheet_name)
                .with_context(|| format!("worksheet {sheet_name:?} was not found"))?;
            let original = read_part(&mut archive, &worksheet.part)?;
            let mut xml = String::from_utf8(original)
                .with_context(|| format!("worksheet {} is not UTF-8", worksheet.part))?;
            xml = patch_cells_xml(&xml, cell_updates)?;
            replacements.insert(worksheet.part.clone(), xml.into_bytes());
        }

        for (sheet_name, settings) in &edits.worksheet_page_settings {
            let worksheet = sheets
                .iter()
                .find(|sheet| sheet.name == *sheet_name)
                .with_context(|| format!("worksheet {sheet_name:?} was not found"))?;
            let xml = replacements
                .entry(worksheet.part.clone())
                .or_insert(read_part(&mut archive, &worksheet.part)?);
            let mut text = String::from_utf8(xml.clone())
                .with_context(|| format!("worksheet {} is not UTF-8", worksheet.part))?;
            text = patch_page_settings_xml(&text, settings)?;
            *xml = text.into_bytes();
        }

        for (sheet_name, margins) in &edits.worksheet_margins {
            let worksheet = sheets
                .iter()
                .find(|sheet| sheet.name == *sheet_name)
                .with_context(|| format!("worksheet {sheet_name:?} was not found"))?;
            let xml = replacements
                .entry(worksheet.part.clone())
                .or_insert(read_part(&mut archive, &worksheet.part)?);
            let mut text = String::from_utf8(xml.clone())
                .with_context(|| format!("worksheet {} is not UTF-8", worksheet.part))?;
            text = patch_margins_xml(&text, margins)?;
            *xml = text.into_bytes();
        }

        for (sheet_name, search, replacement) in &edits.worksheet_headers {
            let worksheet = sheets
                .iter()
                .find(|sheet| sheet.name == *sheet_name)
                .with_context(|| format!("worksheet {sheet_name:?} was not found"))?;
            let xml = replacements
                .entry(worksheet.part.clone())
                .or_insert(read_part(&mut archive, &worksheet.part)?);
            let mut text = String::from_utf8(xml.clone())
                .with_context(|| format!("worksheet {} is not UTF-8", worksheet.part))?;
            text = patch_header_text_xml(&text, search, replacement)?;
            *xml = text.into_bytes();
        }

        if !edits.defined_names.is_empty() || edits.calculation_properties.is_some() {
            let original = read_part(&mut archive, &workbook_part)?;
            let mut xml = String::from_utf8(original).context("workbook.xml is not UTF-8")?;
            if !edits.defined_names.is_empty() {
                xml = patch_defined_names_xml(&xml, &edits.defined_names)?;
            }
            if let Some((full_calc_on_load, force_full_calc)) = edits.calculation_properties {
                xml = patch_calculation_properties_xml(&xml, full_calc_on_load, force_full_calc)?;
            }
            replacements.insert(workbook_part, xml.into_bytes());
        }

        rewrite_package(&mut archive, &replacements)
    }
}

impl EditBackend for RustNativeOoxmlBackend {
    fn capability(&self) -> &BackendCapability {
        &self.capability
    }

    fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure> {
        if plan.format != WorkbookFormat::Xlsx {
            return Err(BackendFailure::unsupported_before_write(
                "RustNative OOXML adapter currently accepts XLSX only",
            ));
        }
        if !plan
            .operations
            .iter()
            .all(|operation| supported_operation(operation.kind()))
        {
            return Err(BackendFailure::unsupported_before_write(
                "RustNative OOXML cannot execute the complete operation batch",
            ));
        }
        let source = fs::read(&plan.input).map_err(|error| {
            BackendFailure::failed(format!("read RustNative input package: {error}"))
        })?;
        let candidate = Self::apply_bytes(&source, plan).map_err(|error| {
            BackendFailure::unsupported_before_write(format!("RustNative preflight: {error:#}"))
        })?;
        let mut output = OpenOptions::new()
            .write(true)
            .truncate(true)
            .open(&plan.output)
            .map_err(|error| {
                BackendFailure::failed(format!("create RustNative candidate: {error}"))
            })?;
        output
            .write_all(&candidate)
            .and_then(|_| output.sync_all())
            .map_err(|error| BackendFailure::failed(format!("write RustNative candidate: {error}")))
    }

    fn validate_output(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        NativeOoxmlValidator.validate(plan, candidate_output)
    }
}

pub struct NativeOoxmlValidator;

impl OutputValidator for NativeOoxmlValidator {
    fn validate(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        crate::OpenpyxlXlsxValidator.validate(plan, candidate_output)?;
        crate::UmyaXlsxValidator.validate(plan, candidate_output)?;
        let mut archive = ZipArchive::new(fs::File::open(candidate_output)?)?;
        let (workbook_part, sheets) = resolve_worksheet_parts(&mut archive)?;
        let edited_sheets: BTreeSet<_> = plan
            .operations
            .iter()
            .filter_map(|operation| {
                if let EditOperation::SetCellValue { sheet, .. } = operation {
                    Some(sheet)
                } else {
                    None
                }
            })
            .collect();
        for sheet_name in edited_sheets {
            let sheet = sheets
                .iter()
                .find(|sheet| sheet.name == *sheet_name)
                .context("edited worksheet is missing from native candidate")?;
            let xml = String::from_utf8(read_part(&mut archive, &sheet.part)?)?;
            let expanded = expand_worksheet_dimension(&xml)?;
            let original_doc = Document::parse(&xml)?;
            let expanded_doc = Document::parse(&expanded)?;
            let dimension_ref = |doc: &Document<'_>| {
                doc.root_element()
                    .children()
                    .find(|n| is_element(*n, "dimension"))
                    .and_then(|n| n.attribute("ref"))
                    .map(str::to_owned)
            };
            ensure!(
                dimension_ref(&original_doc) == dimension_ref(&expanded_doc),
                "native candidate dimension does not cover its stored cells"
            );
        }
        for operation in &plan.operations {
            match operation {
                EditOperation::SetDefinedName {
                    name,
                    formula,
                    local_sheet,
                } => {
                    let local_sheet_id = local_sheet
                        .as_ref()
                        .map(|name| {
                            sheets
                                .iter()
                                .find(|sheet| &sheet.name == name)
                                .map(|sheet| sheet.workbook_index)
                                .with_context(|| format!("worksheet {name:?} missing after save"))
                        })
                        .transpose()?;
                    validate_defined_name(
                        &mut archive,
                        &workbook_part,
                        name,
                        formula.as_deref(),
                        local_sheet_id,
                    )?;
                }
                EditOperation::SetMargins { sheet, margins } => {
                    let worksheet = sheets
                        .iter()
                        .find(|item| &item.name == sheet)
                        .with_context(|| format!("worksheet {sheet:?} missing after save"))?;
                    validate_margins(&mut archive, worksheet, margins)?;
                }
                EditOperation::SetPageSettings { sheet, settings } => {
                    if settings.print_area.is_some() {
                        let worksheet = sheets
                            .iter()
                            .find(|item| &item.name == sheet)
                            .with_context(|| format!("worksheet {sheet:?} missing after save"))?;
                        let quoted_sheet = sheet.replace('\'', "''");
                        let expected = format!(
                            "'{quoted_sheet}'!{}",
                            settings.print_area.as_deref().unwrap_or_default()
                        );
                        validate_defined_name(
                            &mut archive,
                            &workbook_part,
                            "_xlnm.Print_Area",
                            Some(&expected),
                            Some(worksheet.workbook_index),
                        )?;
                    }
                }
                _ => {}
            }
        }
        // The strict package validator must see the entire header batch.
        // Calling it per operation incorrectly rejects other planned sheets.
        if plan
            .operations
            .iter()
            .all(|operation| matches!(operation, EditOperation::ReplaceHeaderText { .. }))
        {
            let updates = plan
                .operations
                .iter()
                .filter_map(|operation| match operation {
                    EditOperation::ReplaceHeaderText {
                        sheet,
                        search,
                        replacement,
                    } => Some(ooxml_patch::HeaderTextUpdate {
                        sheet_name: sheet.clone(),
                        search: search.clone(),
                        replacement: replacement.clone(),
                    }),
                    _ => None,
                })
                .collect::<Vec<_>>();
            ooxml_patch::validate_header_text_patch(&plan.input, candidate_output, &updates)?;
        }
        Ok(())
    }
}

fn supported_operation(operation: EditOperationKind) -> bool {
    matches!(
        operation,
        EditOperationKind::SetCellValue
            | EditOperationKind::SetPageSettings
            | EditOperationKind::SetMargins
            | EditOperationKind::SetDefinedName
            | EditOperationKind::SetWorkbookCalculationProperties
            | EditOperationKind::ReplaceHeaderText
    )
}

fn edits_from_plan(plan: &ExcelEditPlan, sheets: &[WorksheetPart]) -> Result<PackageEdits> {
    let mut edits = PackageEdits::default();
    let has_structural_sheet_change = plan.operations.iter().any(|operation| {
        matches!(
            operation,
            EditOperation::RenameSheet { .. }
                | EditOperation::CreateSheet { .. }
                | EditOperation::CreateSheetAt { .. }
                | EditOperation::MoveSheet { .. }
                | EditOperation::DeleteSheet { .. }
        )
    });
    let mut defined_name_keys = BTreeSet::new();

    for operation in &plan.operations {
        match operation {
            EditOperation::SetCellValue { sheet, cell, value } => {
                let target = sheets
                    .iter()
                    .find(|candidate| candidate.name == *sheet)
                    .with_context(|| format!("worksheet {sheet:?} was not found"))?;
                let (row, column, address) = parse_cell_address(cell)?;
                edits
                    .worksheet_cells
                    .entry(target.name.clone())
                    .or_default()
                    .push(CellUpdate {
                        row,
                        column,
                        address,
                        value: value.clone(),
                    });
            }
            EditOperation::SetPageSettings { sheet, settings } => {
                ensure_sheet_exists(sheets, sheet)?;
                if has_structural_sheet_change && settings.print_area.is_some() {
                    bail!("RustNative print-area updates cannot be combined with sheet create, delete, move, or rename operations");
                }
                edits
                    .worksheet_page_settings
                    .push((sheet.clone(), settings.clone()));
                if let Some(area) = settings.print_area.as_deref() {
                    let target = sheets
                        .iter()
                        .find(|candidate| candidate.name == *sheet)
                        .context("print-area worksheet disappeared")?;
                    let key = (
                        "_xlnm.Print_Area".to_ascii_lowercase(),
                        Some(target.workbook_index),
                    );
                    ensure!(
                        defined_name_keys.insert(key),
                        "plan contains duplicate print-area updates for worksheet {sheet:?}"
                    );
                    let quoted_sheet = sheet.replace('\'', "''");
                    edits.defined_names.push(DefinedNameUpdate {
                        name: "_xlnm.Print_Area".to_owned(),
                        formula: Some(format!("'{quoted_sheet}'!{area}")),
                        local_sheet_id: Some(target.workbook_index),
                    });
                }
            }
            EditOperation::SetMargins { sheet, margins } => {
                ensure_sheet_exists(sheets, sheet)?;
                edits
                    .worksheet_margins
                    .push((sheet.clone(), margins.clone()));
            }
            EditOperation::SetDefinedName {
                name,
                formula,
                local_sheet,
            } => {
                if has_structural_sheet_change {
                    bail!("RustNative defined-name updates cannot be combined with sheet create, delete, move, or rename operations");
                }
                let local_sheet_id = local_sheet
                    .as_ref()
                    .map(|name| {
                        sheets
                            .iter()
                            .find(|sheet| &sheet.name == name)
                            .map(|sheet| sheet.workbook_index)
                            .with_context(|| format!("worksheet scope {name:?} was not found"))
                    })
                    .transpose()?;
                let key = (name.to_ascii_lowercase(), local_sheet_id);
                ensure!(
                    defined_name_keys.insert(key),
                    "plan contains duplicate defined-name updates for {name:?}"
                );
                edits.defined_names.push(DefinedNameUpdate {
                    name: name.clone(),
                    formula: formula.clone(),
                    local_sheet_id,
                });
            }
            EditOperation::SetWorkbookCalculationProperties {
                full_calc_on_load,
                force_full_calc,
            } => {
                ensure!(
                    edits.calculation_properties.is_none(),
                    "plan contains multiple calculation-property updates"
                );
                edits.calculation_properties = Some((*full_calc_on_load, *force_full_calc));
            }
            EditOperation::ReplaceHeaderText {
                sheet,
                search,
                replacement,
            } => {
                ensure_sheet_exists(sheets, sheet)?;
                edits
                    .worksheet_headers
                    .push((sheet.clone(), search.clone(), replacement.clone()));
            }
            other => bail!("RustNative OOXML does not support {:?}", other.kind()),
        }
    }
    Ok(edits)
}

fn ensure_sheet_exists(sheets: &[WorksheetPart], name: &str) -> Result<()> {
    ensure!(
        sheets.iter().any(|sheet| sheet.name == name),
        "worksheet {name:?} was not found"
    );
    Ok(())
}

fn ensure_unique_cell_updates(sheet: &str, updates: &[CellUpdate]) -> Result<()> {
    let mut unique = BTreeSet::new();
    for update in updates {
        ensure!(
            unique.insert((update.row, update.column)),
            "plan contains duplicate cell updates for {sheet}!{}",
            update.address
        );
    }
    Ok(())
}

fn parse_cell_address(value: &str) -> Result<(u32, u32, String)> {
    let normalized = value.replace('$', "").to_ascii_uppercase();
    let split = normalized
        .find(|character: char| character.is_ascii_digit())
        .context("cell address must use A1 notation")?;
    let (column_text, row_text) = normalized.split_at(split);
    ensure!(
        !column_text.is_empty() && !row_text.is_empty(),
        "cell address must use A1 notation"
    );
    let mut column = 0u32;
    for byte in column_text.bytes() {
        ensure!(byte.is_ascii_uppercase(), "invalid cell column");
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add(u32::from(byte - b'A' + 1)))
            .context("cell column overflows")?;
    }
    let row = row_text.parse::<u32>().context("parse cell row")?;
    ensure!(
        (1..=1_048_576).contains(&row) && (1..=16_384).contains(&column),
        "cell address is outside XLSX limits"
    );
    Ok((row, column, normalized))
}

fn patch_cells_xml(xml: &str, updates: &[CellUpdate]) -> Result<String> {
    let document = Document::parse(xml).context("parse worksheet before native cell edits")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().name() == "worksheet" && root.tag_name().namespace() == Some(MAIN_NS),
        "worksheet XML has an unexpected root element"
    );
    let sheet_data = root
        .children()
        .find(|node| is_element(*node, "sheetData"))
        .context("worksheet has no sheetData element")?;
    let prefix = lexical_prefix(xml, root)?;
    let row_qname = qualified(prefix.as_deref(), "row");
    let cell_qname = qualified(prefix.as_deref(), "c");
    let mut edits = Vec::<(Range<usize>, Vec<u8>)>::new();
    let mut insertions = BTreeMap::<usize, Vec<(u32, u32, String)>>::new();
    let mut by_row = BTreeMap::<u32, Vec<&CellUpdate>>::new();
    for update in updates {
        by_row.entry(update.row).or_default().push(update);
    }

    let rows = sheet_data
        .children()
        .filter(|node| is_element(*node, "row"))
        .collect::<Vec<_>>();
    let mut missing_rows = Vec::new();
    for (row_number, row_updates) in by_row {
        let row_node = rows.iter().find(|node| {
            node.attribute("r")
                .and_then(|value| value.parse::<u32>().ok())
                == Some(row_number)
        });
        let Some(row_node) = row_node else {
            let cells = row_updates
                .iter()
                .map(|update| {
                    cell_xml(&cell_qname, &update.address, None, &update.value)
                        .map(|xml| (update.column, xml))
                })
                .collect::<Result<Vec<_>>>()?;
            let mut cells = cells;
            cells.sort_by_key(|(column, _)| *column);
            let row_xml = format!(
                "<{row_qname} r=\"{row_number}\">{}</{row_qname}>",
                cells.into_iter().map(|(_, xml)| xml).collect::<String>()
            );
            missing_rows.push((row_number, row_xml));
            continue;
        };

        let existing_cells = row_node
            .children()
            .filter(|node| is_element(*node, "c"))
            .collect::<Vec<_>>();
        let mut new_cells_by_offset = BTreeMap::<usize, Vec<(u32, String)>>::new();
        for update in row_updates {
            let matching = existing_cells
                .iter()
                .filter(|node| {
                    node.attribute("r")
                        .is_some_and(|value| value.eq_ignore_ascii_case(&update.address))
                })
                .copied()
                .collect::<Vec<_>>();
            ensure!(
                matching.len() <= 1,
                "worksheet contains duplicate cell {}",
                update.address
            );
            if let Some(cell) = matching.first() {
                validate_plain_cell(*cell)?;
                let style = cell.attribute("s");
                let replacement = cell_xml(&cell_qname, &update.address, style, &update.value)?;
                edits.push((cell.range(), replacement.into_bytes()));
            } else {
                let position = existing_cells
                    .iter()
                    .filter_map(|cell| {
                        let address = cell.attribute("r")?;
                        let (_, column, _) = parse_cell_address(address).ok()?;
                        (column > update.column).then_some(cell.range().start)
                    })
                    .min()
                    .unwrap_or(closing_tag_start(xml, *row_node)?);
                let rendered = cell_xml(&cell_qname, &update.address, None, &update.value)?;
                new_cells_by_offset
                    .entry(position)
                    .or_default()
                    .push((update.column, rendered));
            }
        }
        for (position, mut cells) in new_cells_by_offset {
            cells.sort_by_key(|(column, _)| *column);
            insertions.entry(position).or_default().extend(
                cells
                    .into_iter()
                    .map(|(column, xml)| (row_number, column, xml)),
            );
        }
    }

    if !missing_rows.is_empty() {
        let empty_self_closing = is_self_closing(xml, sheet_data.range())?;
        missing_rows.sort_by_key(|(row, _)| *row);
        if empty_self_closing {
            let rows_xml = missing_rows
                .into_iter()
                .map(|(_, value)| value)
                .collect::<String>();
            // The worksheet may have both a self-closing sheetData and edits
            // to existing rows only if the package was malformed; reject it.
            ensure!(by_row_len_without_missing_rows(updates, &rows) == 0, "cannot add missing rows to a self-closing sheetData that also contains existing target rows");
            let replacement = expand_self_closing_element(xml, sheet_data, &rows_xml)?;
            edits.push((sheet_data.range(), replacement.into_bytes()));
        } else {
            let insertion_offset = rows
                .iter()
                .filter_map(|row| {
                    let number = row.attribute("r")?.parse::<u32>().ok()?;
                    missing_rows
                        .iter()
                        .any(|(missing, _)| *missing < number)
                        .then_some(row.range().start)
                })
                .min()
                .unwrap_or(closing_tag_start(xml, sheet_data)?);
            insertions.entry(insertion_offset).or_default().extend(
                missing_rows
                    .into_iter()
                    .map(|(row, row_xml)| (row, 0, row_xml)),
            );
        }
    }

    for (position, mut values) in insertions {
        values.sort_by_key(|(row, column, _)| (*row, *column));
        let content = values
            .into_iter()
            .map(|(_, _, fragment)| fragment)
            .collect::<String>();
        edits.push((position..position, content.into_bytes()));
    }
    expand_worksheet_dimension(&apply_xml_edits(xml, edits)?)
}

/// A streaming reader trusts dimension/@ref and can otherwise miss appended
/// cells. Expand the existing declaration to cover every stored cell, keeping
/// any larger source extent (including formatting-only areas) intact.
fn expand_worksheet_dimension(xml: &str) -> Result<String> {
    let document = Document::parse(xml).context("parse native candidate dimension")?;
    let root = document.root_element();
    let dimensions = root
        .children()
        .filter(|n| is_element(*n, "dimension"))
        .collect::<Vec<_>>();
    ensure!(
        dimensions.len() <= 1,
        "worksheet contains duplicate dimension elements"
    );
    let mut bounds: Option<(u32, u32, u32, u32)> = None;
    let mut include = |address: &str| -> Result<()> {
        let (row, column, _) = parse_cell_address(address)?;
        bounds = Some(match bounds {
            Some((r1, c1, r2, c2)) => (r1.min(row), c1.min(column), r2.max(row), c2.max(column)),
            None => (row, column, row, column),
        });
        Ok(())
    };
    if let Some(dimension) = dimensions.first() {
        let range = dimension
            .attribute("ref")
            .context("worksheet dimension has no ref")?;
        let (first, last) = range.split_once(':').unwrap_or((range, range));
        let (r1, c1, _) = parse_cell_address(first)?;
        let (r2, c2, _) = parse_cell_address(last)?;
        ensure!(r1 <= r2 && c1 <= c2, "worksheet dimension is reversed");
        include(first)?;
        include(last)?;
    }
    for cell in root
        .children()
        .filter(|n| is_element(*n, "sheetData"))
        .flat_map(|n| n.descendants())
        .filter(|n| is_element(*n, "c"))
    {
        include(
            cell.attribute("r")
                .context("stored cell has no coordinate")?,
        )?;
    }
    let (r1, c1, r2, c2) = bounds.unwrap_or((1, 1, 1, 1));
    let address = |row, mut column: u32| {
        let mut letters = Vec::new();
        while column > 0 {
            letters.push((b'A' + ((column - 1) % 26) as u8) as char);
            column = (column - 1) / 26;
        }
        format!("{}{row}", letters.into_iter().rev().collect::<String>())
    };
    let first = address(r1, c1);
    let last = address(r2, c2);
    let range = if first == last {
        first
    } else {
        format!("{first}:{last}")
    };
    update_worksheet_element_attributes(xml, None, "dimension", &[("ref".into(), Some(range))])
}

// Used to avoid adding worksheet rows under an empty self-closing sheetData
// when the same plan also expects pre-existing rows.
fn by_row_len_without_missing_rows(updates: &[CellUpdate], rows: &[Node<'_, '_>]) -> usize {
    let existing = rows
        .iter()
        .filter_map(|node| {
            node.attribute("r")
                .and_then(|value| value.parse::<u32>().ok())
        })
        .collect::<BTreeSet<_>>();
    updates
        .iter()
        .filter(|update| existing.contains(&update.row))
        .count()
}

fn validate_plain_cell(cell: Node<'_, '_>) -> Result<()> {
    for attribute in cell.attributes() {
        ensure!(
            matches!(attribute.name(), "r" | "s" | "t"),
            "RustNative refuses to replace cell {} with non-value metadata attribute {:?}",
            cell.attribute("r").unwrap_or("?"),
            attribute.name()
        );
    }
    for child in cell.children().filter(|node| node.is_element()) {
        ensure!(
            matches!(child.tag_name().name(), "f" | "v" | "is"),
            "RustNative refuses to replace cell {} containing unsupported child {:?}",
            cell.attribute("r").unwrap_or("?"),
            child.tag_name().name()
        );
    }
    Ok(())
}

fn cell_xml(
    qname: &str,
    address: &str,
    style: Option<&str>,
    value: &crate::CellValue,
) -> Result<String> {
    let mut writer = Writer::new(Vec::new());
    let mut cell = BytesStart::new(qname);
    cell.push_attribute(("r", address));
    if let Some(style) = style {
        cell.push_attribute(("s", style));
    }
    let cell_type = match value {
        crate::CellValue::Text(_) => Some("inlineStr"),
        crate::CellValue::Boolean(_) => Some("b"),
        crate::CellValue::Error(_) => Some("e"),
        crate::CellValue::Decimal(_) | crate::CellValue::Blank => None,
        crate::CellValue::DateTime { .. } => {
            bail!("RustNative currently refuses DateTime cell writes because it cannot create or verify a number-format style")
        }
    };
    if let Some(cell_type) = cell_type {
        cell.push_attribute(("t", cell_type));
    }
    if matches!(value, crate::CellValue::Blank) {
        writer.write_event(Event::Empty(cell))?;
        return Ok(String::from_utf8(writer.into_inner())?);
    }
    writer.write_event(Event::Start(cell))?;
    match value {
        crate::CellValue::Text(text) => {
            let is_name = qname_prefix(qname, "is");
            let text_name = qname_prefix(qname, "t");
            writer.write_event(Event::Start(BytesStart::new(is_name.as_str())))?;
            let mut text_element = BytesStart::new(text_name.as_str());
            if text.chars().next().is_some_and(char::is_whitespace)
                || text.chars().last().is_some_and(char::is_whitespace)
            {
                text_element.push_attribute(("xml:space", "preserve"));
            }
            writer.write_event(Event::Start(text_element))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new(text_name.as_str())))?;
            writer.write_event(Event::End(BytesEnd::new(is_name.as_str())))?;
        }
        crate::CellValue::Decimal(decimal) => {
            let number = decimal
                .parse::<f64>()
                .context("parse numeric value for RustNative")?;
            ensure!(number.is_finite(), "numeric value must be finite");
            write_value_element(&mut writer, qname, decimal)?;
        }
        crate::CellValue::Boolean(value) => {
            write_value_element(&mut writer, qname, if *value { "1" } else { "0" })?;
        }
        crate::CellValue::Error(value) => {
            ensure!(
                matches!(
                    value.as_str(),
                    "#DIV/0!"
                        | "#N/A"
                        | "#NAME?"
                        | "#NULL!"
                        | "#NUM!"
                        | "#REF!"
                        | "#VALUE!"
                        | "#SPILL!"
                        | "#CALC!"
                ),
                "unsupported Excel error token {value:?}"
            );
            write_value_element(&mut writer, qname, value)?;
        }
        crate::CellValue::Blank | crate::CellValue::DateTime { .. } => unreachable!(),
    }
    let cell_name = qname.to_owned();
    writer.write_event(Event::End(BytesEnd::new(cell_name.as_str())))?;
    Ok(String::from_utf8(writer.into_inner())?)
}

fn write_value_element(writer: &mut Writer<Vec<u8>>, qname: &str, value: &str) -> Result<()> {
    let name = qname_prefix(qname, "v");
    writer.write_event(Event::Start(BytesStart::new(name.as_str())))?;
    writer.write_event(Event::Text(BytesText::new(value)))?;
    writer.write_event(Event::End(BytesEnd::new(name.as_str())))?;
    Ok(())
}

fn patch_page_settings_xml(xml: &str, settings: &crate::PageSettings) -> Result<String> {
    let mut output = xml.to_owned();
    let mut page_attributes = Vec::<(String, Option<String>)>::new();
    if let Some(orientation) = settings.orientation {
        page_attributes.push((
            "orientation".to_owned(),
            Some(match orientation {
                crate::PageOrientation::Portrait => "portrait".to_owned(),
                crate::PageOrientation::Landscape => "landscape".to_owned(),
            }),
        ));
    }
    if let Some(width) = settings.fit_to_width {
        page_attributes.push(("fitToWidth".to_owned(), Some(width.to_string())));
    }
    if let Some(height) = settings.fit_to_height {
        page_attributes.push(("fitToHeight".to_owned(), Some(height.to_string())));
    }
    if settings.clear_scale {
        page_attributes.push(("scale".to_owned(), None));
    }
    if let Some(size) = settings.paper_size {
        page_attributes.push(("paperSize".to_owned(), Some(size.to_string())));
    }
    if !page_attributes.is_empty() {
        output = update_worksheet_element_attributes(&output, None, "pageSetup", &page_attributes)?;
    }
    if let Some(fit_to_page) = settings.fit_to_page {
        output = update_worksheet_element_attributes(
            &output,
            Some("sheetPr"),
            "pageSetUpPr",
            &[(
                "fitToPage".to_owned(),
                Some(xml_bool(fit_to_page).to_owned()),
            )],
        )?;
    }
    let print_attributes = [
        settings.horizontal_centered.map(|value| {
            (
                "horizontalCentered".to_owned(),
                Some(xml_bool(value).to_owned()),
            )
        }),
        settings.vertical_centered.map(|value| {
            (
                "verticalCentered".to_owned(),
                Some(xml_bool(value).to_owned()),
            )
        }),
    ]
    .into_iter()
    .flatten()
    .collect::<Vec<_>>();
    if !print_attributes.is_empty() {
        output =
            update_worksheet_element_attributes(&output, None, "printOptions", &print_attributes)?;
    }
    if let Some(footer) = &settings.footer {
        let value = if footer.starts_with("&C") {
            footer.clone()
        } else {
            format!("&C{footer}")
        };
        output = set_nested_text(&output, None, "headerFooter", "oddFooter", &value)?;
    }
    Ok(output)
}

fn patch_margins_xml(xml: &str, margins: &crate::PageMargins) -> Result<String> {
    let attributes = [
        ("left", margins.left),
        ("right", margins.right),
        ("top", margins.top),
        ("bottom", margins.bottom),
        ("header", margins.header),
        ("footer", margins.footer),
    ]
    .into_iter()
    .filter_map(|(name, value)| value.map(|value| (name.to_owned(), Some(value.to_string()))))
    .collect::<Vec<_>>();
    update_worksheet_element_attributes(xml, None, "pageMargins", &attributes)
}

fn patch_header_text_xml(xml: &str, search: &str, replacement: &str) -> Result<String> {
    let document = Document::parse(xml).context("parse worksheet header/footer XML")?;
    let root = document.root_element();
    let headers = root
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && (node.tag_name().name().ends_with("Header")
                    || node.tag_name().name().ends_with("Footer"))
        })
        .filter(|node| node.text().is_some_and(|text| text.contains(search)))
        .collect::<Vec<_>>();
    ensure!(
        !headers.is_empty(),
        "header/footer text {search:?} was not found"
    );
    let mut edits = Vec::new();
    for node in headers {
        ensure!(
            node.children()
                .filter(|child| child.is_element())
                .next()
                .is_none(),
            "RustNative refuses nested markup inside header/footer text"
        );
        let text = node.text().unwrap_or_default().replace(search, replacement);
        let rendered = replace_element_text(xml, node, &text)?;
        edits.push((node.range(), rendered.into_bytes()));
    }
    apply_xml_edits(xml, edits)
}

fn patch_defined_names_xml(xml: &str, updates: &[DefinedNameUpdate]) -> Result<String> {
    let document = Document::parse(xml).context("parse workbook defined names")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().name() == "workbook" && root.tag_name().namespace() == Some(MAIN_NS),
        "workbook XML has an unexpected root element"
    );
    let prefix = lexical_prefix(xml, root)?;
    let name_qname = qualified(prefix.as_deref(), "definedName");
    let container = root
        .children()
        .find(|node| is_element(*node, "definedNames"));
    let mut edits = Vec::new();
    let mut additions = Vec::new();
    for update in updates {
        let matches = container
            .into_iter()
            .flat_map(|container| container.children())
            .filter(|node| {
                is_element(*node, "definedName")
                    && node
                        .attribute("name")
                        .is_some_and(|name| name.eq_ignore_ascii_case(&update.name))
                    && node
                        .attribute("localSheetId")
                        .and_then(|value| value.parse::<usize>().ok())
                        == update.local_sheet_id
            })
            .collect::<Vec<_>>();
        ensure!(
            matches.len() <= 1,
            "workbook has ambiguous duplicate defined name {:?}",
            update.name
        );
        if let Some(existing) = matches.first() {
            match &update.formula {
                Some(formula) => {
                    let rendered = defined_name_xml(
                        &name_qname,
                        &update.name,
                        formula,
                        update.local_sheet_id,
                        Some(*existing),
                    )?;
                    edits.push((existing.range(), rendered.into_bytes()));
                }
                None => edits.push((existing.range(), Vec::new())),
            }
        } else if let Some(formula) = &update.formula {
            additions.push(defined_name_xml(
                &name_qname,
                &update.name,
                formula,
                update.local_sheet_id,
                None,
            )?);
        }
    }
    if !additions.is_empty() {
        if let Some(container) = container {
            if is_self_closing(xml, container.range())? {
                let qname = qualified(lexical_prefix(xml, container)?.as_deref(), "definedNames");
                let rendered = format!("<{qname}>{}</{qname}>", additions.concat());
                edits.push((container.range(), rendered.into_bytes()));
            } else {
                let at = closing_tag_start(xml, container)?;
                edits.push((at..at, additions.concat().into_bytes()));
            }
        } else {
            let qname = qualified(prefix.as_deref(), "definedNames");
            let rendered = format!("<{qname}>{}</{qname}>", additions.concat());
            let at = ordered_child_insertion(xml, root, "definedNames", workbook_child_order)?;
            edits.push((at..at, rendered.into_bytes()));
        }
    }
    apply_xml_edits(xml, edits)
}

fn defined_name_xml(
    qname: &str,
    name: &str,
    formula: &str,
    local_sheet_id: Option<usize>,
    existing: Option<Node<'_, '_>>,
) -> Result<String> {
    let mut attrs = BTreeMap::<String, String>::new();
    if let Some(existing) = existing {
        for attribute in existing.attributes() {
            let key = attribute.name().to_owned();
            attrs.insert(key, attribute.value().to_owned());
        }
    }
    attrs.insert("name".to_owned(), name.to_owned());
    match local_sheet_id {
        Some(index) => {
            attrs.insert("localSheetId".to_owned(), index.to_string());
        }
        None => {
            attrs.remove("localSheetId");
        }
    }
    let mut writer = Writer::new(Vec::new());
    let mut start = BytesStart::new(qname);
    let attributes = attrs
        .iter()
        .map(|(key, value)| (key.as_str(), value.as_str()))
        .collect::<Vec<_>>();
    for attribute in attributes {
        start.push_attribute(attribute);
    }
    writer.write_event(Event::Start(start))?;
    let formula = formula.trim_start_matches('=');
    writer.write_event(Event::Text(BytesText::new(formula)))?;
    writer.write_event(Event::End(BytesEnd::new(qname)))?;
    Ok(String::from_utf8(writer.into_inner())?)
}

fn patch_calculation_properties_xml(
    xml: &str,
    full_calc_on_load: Option<bool>,
    force_full_calc: Option<bool>,
) -> Result<String> {
    let mut updates = Vec::new();
    if let Some(value) = full_calc_on_load {
        updates.push((
            "fullCalcOnLoad".to_owned(),
            Some(xml_bool(value).to_owned()),
        ));
    }
    if let Some(value) = force_full_calc {
        updates.push(("forceFullCalc".to_owned(), Some(xml_bool(value).to_owned())));
    }
    update_workbook_element_attributes(xml, "calcPr", &updates)
}

fn update_worksheet_element_attributes(
    xml: &str,
    parent_name: Option<&str>,
    element_name: &str,
    attributes: &[(String, Option<String>)],
) -> Result<String> {
    let document = Document::parse(xml).context("parse worksheet XML for attribute update")?;
    let root = document.root_element();
    let prefix = lexical_prefix(xml, root)?;
    let parent = parent_name
        .map(|name| root.children().find(|node| is_element(*node, name)))
        .unwrap_or(Some(root));
    if let Some(parent) = parent {
        if let Some(element) = parent
            .children()
            .find(|node| is_element(*node, element_name))
        {
            return patch_start_tag(xml, element, attributes);
        }
        let child_prefix = lexical_prefix(xml, parent)?.or(prefix);
        let qname = qualified(child_prefix.as_deref(), element_name);
        let rendered = render_empty_element(&qname, attributes)?;
        let insertion = ordered_child_insertion(xml, parent, element_name, worksheet_child_order)?;
        return insert_at(xml, insertion, &rendered);
    }

    // A missing sheetPr parent is valid and must precede the other worksheet
    // children. Create it with the requested pageSetUpPr in one edit.
    let parent_name = parent_name.context("missing worksheet parent element")?;
    ensure!(
        parent_name == "sheetPr",
        "cannot synthesize worksheet parent {parent_name}"
    );
    let parent_qname = qualified(prefix.as_deref(), parent_name);
    let child_qname = qualified(prefix.as_deref(), element_name);
    let child = render_empty_element(&child_qname, attributes)?;
    let parent_xml = format!("<{parent_qname}>{child}</{parent_qname}>");
    let at = ordered_child_insertion(xml, root, parent_name, worksheet_child_order)?;
    insert_at(xml, at, &parent_xml)
}

fn update_workbook_element_attributes(
    xml: &str,
    element_name: &str,
    attributes: &[(String, Option<String>)],
) -> Result<String> {
    let document = Document::parse(xml).context("parse workbook XML for attribute update")?;
    let root = document.root_element();
    let prefix = lexical_prefix(xml, root)?;
    if let Some(element) = root.children().find(|node| is_element(*node, element_name)) {
        return patch_start_tag(xml, element, attributes);
    }
    let qname = qualified(prefix.as_deref(), element_name);
    let rendered = render_empty_element(&qname, attributes)?;
    let insertion = ordered_child_insertion(xml, root, element_name, workbook_child_order)?;
    insert_at(xml, insertion, &rendered)
}

fn set_nested_text(
    xml: &str,
    parent_name: Option<&str>,
    element_name: &str,
    child_name: &str,
    value: &str,
) -> Result<String> {
    let mut output = xml.to_owned();
    let document = Document::parse(&output).context("parse worksheet header/footer XML")?;
    let root = document.root_element();
    let parent = parent_name
        .map(|name| root.children().find(|node| is_element(*node, name)))
        .unwrap_or(Some(root));
    if parent.is_none() {
        let prefix = lexical_prefix(&output, root)?;
        let element_qname = qualified(prefix.as_deref(), element_name);
        let child_qname = qualified(prefix.as_deref(), child_name);
        let child = render_text_element(&child_qname, value)?;
        let container = format!("<{element_qname}>{child}</{element_qname}>");
        let at = ordered_child_insertion(&output, root, element_name, worksheet_child_order)?;
        return insert_at(&output, at, &container);
    }
    let parent = parent.context("header/footer parent missing")?;
    let element = parent
        .children()
        .find(|node| is_element(*node, element_name));
    if element.is_none() {
        let prefix = lexical_prefix(&output, parent)?;
        let element_qname = qualified(prefix.as_deref(), element_name);
        let child_qname = qualified(prefix.as_deref(), child_name);
        let child = render_text_element(&child_qname, value)?;
        let container = format!("<{element_qname}>{child}</{element_qname}>");
        let at = ordered_child_insertion(&output, parent, element_name, worksheet_child_order)?;
        return insert_at(&output, at, &container);
    }
    let element = element.context("header/footer parent missing")?;
    let child = element
        .children()
        .find(|node| is_element(*node, child_name));
    if let Some(child) = child {
        if is_self_closing(&output, child.range())? {
            let escaped = quick_xml::escape::escape(value);
            let replacement = expand_self_closing_element(&output, child, &escaped)?;
            return Ok(replace_range(
                &output,
                child.range(),
                replacement.as_bytes(),
            ));
        }
        let replacement = replace_element_text(&output, child, value)?;
        output = replace_range(&output, child.range(), replacement.as_bytes());
        return Ok(output);
    }
    if is_self_closing(&output, element.range())? {
        let prefix = lexical_prefix(&output, element)?;
        let child_qname = qualified(prefix.as_deref(), child_name);
        let child_xml = render_text_element(&child_qname, value)?;
        let replacement = expand_self_closing_element(&output, element, &child_xml)?;
        return Ok(replace_range(
            &output,
            element.range(),
            replacement.as_bytes(),
        ));
    }
    let prefix = lexical_prefix(&output, element)?;
    let child_qname = qualified(prefix.as_deref(), child_name);
    let child_xml = render_text_element(&child_qname, value)?;
    let at = closing_tag_start(&output, element)?;
    insert_at(&output, at, &child_xml)
}

fn patch_start_tag(
    xml: &str,
    node: Node<'_, '_>,
    updates: &[(String, Option<String>)],
) -> Result<String> {
    let range = node.range();
    let close = xml[range.clone()]
        .find('>')
        .map(|offset| range.start + offset + 1)
        .context("XML element start tag is unterminated")?;
    let raw_start = &xml[range.start..close];
    let self_closing = raw_start.trim_end().ends_with("/>");
    let mut values = BTreeMap::<String, String>::new();
    for attribute in node.attributes() {
        values.insert(attribute.name().to_owned(), attribute.value().to_owned());
    }
    for (name, value) in updates {
        match value {
            Some(value) => {
                values.insert(name.clone(), value.clone());
            }
            None => {
                values.remove(name);
            }
        }
    }
    let qname = lexical_qname(xml, node)?;
    let mut replacement = format!("<{qname}");
    for (name, value) in values {
        replacement.push(' ');
        replacement.push_str(&name);
        replacement.push_str("=\"");
        replacement.push_str(&escape_attribute(&value));
        replacement.push('"');
    }
    replacement.push_str(if self_closing { "/>" } else { ">" });
    Ok(replace_range(
        xml,
        range.start..close,
        replacement.as_bytes(),
    ))
}

fn render_empty_element(qname: &str, attributes: &[(String, Option<String>)]) -> Result<String> {
    let mut writer = Writer::new(Vec::new());
    let mut start = BytesStart::new(qname);
    for (name, value) in attributes {
        if let Some(value) = value {
            start.push_attribute((name.as_str(), value.as_str()));
        }
    }
    writer.write_event(Event::Empty(start))?;
    Ok(String::from_utf8(writer.into_inner())?)
}

fn render_text_element(qname: &str, text: &str) -> Result<String> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Start(BytesStart::new(qname)))?;
    writer.write_event(Event::Text(BytesText::new(text)))?;
    writer.write_event(Event::End(BytesEnd::new(qname)))?;
    Ok(String::from_utf8(writer.into_inner())?)
}

fn replace_element_text(xml: &str, node: Node<'_, '_>, text: &str) -> Result<String> {
    let range = node.range();
    let start_close = xml[range.clone()]
        .find('>')
        .map(|offset| range.start + offset + 1)
        .context("XML text element start tag is unterminated")?;
    let closing = xml[range.clone()]
        .rfind("</")
        .map(|offset| range.start + offset)
        .context("XML text element has no closing tag")?;
    let mut replacement = String::with_capacity(range.end - range.start + text.len());
    replacement.push_str(&xml[range.start..start_close]);
    replacement.push_str(&quick_xml::escape::escape(text));
    replacement.push_str(&xml[closing..range.end]);
    Ok(replacement)
}

fn expand_self_closing_element(xml: &str, node: Node<'_, '_>, inner: &str) -> Result<String> {
    let range = node.range();
    let close = xml[range.clone()]
        .find('>')
        .map(|offset| range.start + offset)
        .context("XML element start tag is unterminated")?;
    let raw_start = &xml[range.start..=close];
    ensure!(
        raw_start.trim_end().ends_with("/>"),
        "element is not self-closing"
    );
    let prefix_end = raw_start
        .rfind("/>")
        .context("self-closing marker missing")?;
    let mut output = raw_start[..prefix_end].to_owned();
    output.push('>');
    output.push_str(inner);
    let qname = lexical_qname(xml, node)?;
    output.push_str("</");
    output.push_str(&qname);
    output.push('>');
    Ok(output)
}

fn ordered_child_insertion(
    xml: &str,
    parent: Node<'_, '_>,
    new_child: &str,
    order: fn(&str) -> usize,
) -> Result<usize> {
    let new_order = order(new_child);
    if let Some(node) = parent
        .children()
        .filter(|node| node.is_element())
        .filter(|node| order(node.tag_name().name()) > new_order)
        .min_by_key(|node| node.range().start)
    {
        return Ok(node.range().start);
    }
    closing_tag_start(xml, parent)
}

fn worksheet_child_order(name: &str) -> usize {
    match name {
        "sheetPr" => 0,
        "dimension" => 1,
        "sheetViews" => 2,
        "sheetFormatPr" => 3,
        "cols" => 4,
        "sheetData" => 5,
        "sheetCalcPr" => 6,
        "sheetProtection" => 7,
        "protectedRanges" => 8,
        "scenarios" => 9,
        "autoFilter" => 10,
        "sortState" => 11,
        "dataConsolidate" => 12,
        "customSheetViews" => 13,
        "mergeCells" => 14,
        "phoneticPr" => 15,
        "conditionalFormatting" => 16,
        "dataValidations" => 17,
        "hyperlinks" => 18,
        "printOptions" => 19,
        "pageMargins" => 20,
        "pageSetup" => 21,
        "headerFooter" => 22,
        "rowBreaks" => 23,
        "colBreaks" => 24,
        _ => 100,
    }
}

fn workbook_child_order(name: &str) -> usize {
    match name {
        "fileVersion" => 0,
        "fileSharing" => 1,
        "workbookPr" => 2,
        "workbookProtection" => 3,
        "bookViews" => 4,
        "sheets" => 5,
        "functionGroups" => 6,
        "externalReferences" => 7,
        "definedNames" => 8,
        "calcPr" => 9,
        "oleSize" => 10,
        "customWorkbookViews" => 11,
        "pivotCaches" => 12,
        "smartTagPr" => 13,
        "smartTagTypes" => 14,
        "webPublishing" => 15,
        "fileRecoveryPr" => 16,
        "webPublishObjects" => 17,
        "extLst" => 18,
        _ => 100,
    }
}

fn validate_defined_name<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    workbook_part: &str,
    name: &str,
    expected_formula: Option<&str>,
    local_sheet_id: Option<usize>,
) -> Result<()> {
    let xml = String::from_utf8(read_part(archive, workbook_part)?)?;
    let document = Document::parse(&xml)?;
    let matches = document
        .descendants()
        .filter(|node| {
            is_element(*node, "definedName")
                && node
                    .attribute("name")
                    .is_some_and(|value| value.eq_ignore_ascii_case(name))
                && node
                    .attribute("localSheetId")
                    .and_then(|value| value.parse::<usize>().ok())
                    == local_sheet_id
        })
        .collect::<Vec<_>>();
    if let Some(expected) = expected_formula {
        ensure!(
            matches.len() == 1,
            "defined name {name:?} missing or duplicated after RustNative save"
        );
        ensure!(
            matches[0].text().unwrap_or_default() == expected.trim_start_matches('='),
            "defined name {name:?} formula readback mismatch"
        );
    } else {
        ensure!(matches.is_empty(), "defined name {name:?} was not removed");
    }
    Ok(())
}

fn validate_margins<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    worksheet: &WorksheetPart,
    expected: &crate::PageMargins,
) -> Result<()> {
    let xml = String::from_utf8(read_part(archive, &worksheet.part)?)?;
    let document = Document::parse(&xml)?;
    let node = document
        .descendants()
        .find(|node| is_element(*node, "pageMargins"))
        .context("pageMargins was missing after RustNative save")?;
    for (name, expected) in [
        ("left", expected.left),
        ("right", expected.right),
        ("top", expected.top),
        ("bottom", expected.bottom),
        ("header", expected.header),
        ("footer", expected.footer),
    ] {
        if let Some(expected) = expected {
            let actual = node
                .attribute(name)
                .with_context(|| format!("page margin {name} missing after save"))?
                .parse::<f64>()?;
            ensure!(
                (actual - expected).abs() <= 1e-10,
                "page margin {name} readback mismatch"
            );
        }
    }
    Ok(())
}

fn resolve_worksheet_parts<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<(String, Vec<WorksheetPart>)> {
    let root_rels_xml = String::from_utf8(read_part(archive, "_rels/.rels")?)?;
    let root_rels = Document::parse(&root_rels_xml)?;
    let workbook_target = root_rels
        .descendants()
        .find(|node| {
            is_element_ns(*node, "Relationship", PACKAGE_REL_NS)
                && node
                    .attribute("Type")
                    .is_some_and(|kind| kind.ends_with("/officeDocument"))
                && !node
                    .attribute("TargetMode")
                    .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
        })
        .and_then(|node| node.attribute("Target"))
        .context("package root relationships do not name a workbook part")?;
    let workbook_part = normalize_part("", workbook_target)?;
    let workbook_xml = String::from_utf8(read_part(archive, &workbook_part)?)?;
    let workbook = Document::parse(&workbook_xml)?;
    let rels_part = relationships_part(&workbook_part);
    let rels_xml = String::from_utf8(read_part(archive, &rels_part)?)?;
    let rels = Document::parse(&rels_xml)?;
    let mut relationships = HashMap::new();
    for node in rels
        .descendants()
        .filter(|node| is_element_ns(*node, "Relationship", PACKAGE_REL_NS))
    {
        if node
            .attribute("TargetMode")
            .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
        {
            continue;
        }
        if let (Some(id), Some(kind), Some(target)) = (
            node.attribute("Id"),
            node.attribute("Type"),
            node.attribute("Target"),
        ) {
            relationships.insert(id.to_owned(), (kind.to_owned(), target.to_owned()));
        }
    }
    let workbook_parent = workbook_part
        .rsplit_once('/')
        .map(|(parent, _)| parent)
        .unwrap_or("");
    let sheets_node = workbook
        .root_element()
        .children()
        .find(|node| is_element(*node, "sheets"))
        .context("workbook has no sheets catalog")?;
    let mut sheets = Vec::new();
    for (workbook_index, sheet) in sheets_node
        .children()
        .filter(|node| is_element(*node, "sheet"))
        .enumerate()
    {
        let name = sheet
            .attribute("name")
            .context("workbook sheet has no name")?;
        let relation_id = sheet
            .attribute((REL_NS, "id"))
            .context("workbook sheet has no relationship id")?;
        let (kind, target) = relationships
            .get(relation_id)
            .with_context(|| format!("relationship {relation_id} is unresolved"))?;
        if !kind.ends_with("/worksheet") {
            continue;
        }
        sheets.push(WorksheetPart {
            name: name.to_owned(),
            part: normalize_part(workbook_parent, target)?,
            workbook_index,
        });
    }
    ensure!(!sheets.is_empty(), "workbook has no worksheet parts");
    Ok((workbook_part, sheets))
}

fn normalize_part(parent: &str, target: &str) -> Result<String> {
    let mut parts = Vec::<&str>::new();
    if !target.starts_with('/') {
        parts.extend(parent.split('/').filter(|part| !part.is_empty()));
    }
    for part in target.trim_start_matches('/').split('/') {
        match part {
            "" | "." => {}
            ".." => {
                ensure!(
                    parts.pop().is_some(),
                    "relationship target escapes package root"
                );
            }
            part => parts.push(part),
        }
    }
    ensure!(!parts.is_empty(), "empty OOXML relationship target");
    Ok(parts.join("/"))
}

fn relationships_part(part: &str) -> String {
    let (parent, filename) = part.rsplit_once('/').unwrap_or(("", part));
    if parent.is_empty() {
        format!("_rels/{filename}.rels")
    } else {
        format!("{parent}/_rels/{filename}.rels")
    }
}

fn read_part<R: Read + Seek>(archive: &mut ZipArchive<R>, name: &str) -> Result<Vec<u8>> {
    let mut entry = archive
        .by_name(name)
        .with_context(|| format!("ZIP member {name} is missing"))?;
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes)?;
    Ok(bytes)
}

fn rewrite_package<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    replacements: &BTreeMap<String, Vec<u8>>,
) -> Result<Vec<u8>> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index)?;
        let name = entry.name().to_owned();
        let mut options = SimpleFileOptions::default().compression_method(entry.compression());
        if let Some(modified) = entry.last_modified() {
            options = options.last_modified_time(modified);
        }
        if entry.is_dir() {
            writer.add_directory(name, options)?;
        } else {
            writer.start_file(&name, options)?;
            if let Some(replacement) = replacements.get(&name) {
                writer.write_all(replacement)?;
            } else {
                std::io::copy(&mut entry, &mut writer)?;
            }
        }
    }
    Ok(writer.finish()?.into_inner())
}

fn is_element(node: Node<'_, '_>, name: &str) -> bool {
    node.is_element()
        && node.tag_name().namespace() == Some(MAIN_NS)
        && node.tag_name().name() == name
}

fn is_element_ns(node: Node<'_, '_>, name: &str, namespace: &str) -> bool {
    node.is_element()
        && node.tag_name().namespace() == Some(namespace)
        && node.tag_name().name() == name
}

fn qualified(prefix: Option<&str>, local: &str) -> String {
    prefix.map_or_else(|| local.to_owned(), |prefix| format!("{prefix}:{local}"))
}

fn lexical_qname(xml: &str, node: Node<'_, '_>) -> Result<String> {
    let range = node.range();
    let fragment = xml
        .get(range.start..range.end)
        .context("XML node range is outside its source document")?;
    let start = fragment.find('<').context("XML element has no start tag")?;
    let bytes = fragment[start + 1..].as_bytes();
    let end = bytes
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'/' || *byte == b'>')
        .context("XML element name is missing")?;
    Ok(std::str::from_utf8(&bytes[..end])?.to_owned())
}

fn lexical_prefix(xml: &str, node: Node<'_, '_>) -> Result<Option<String>> {
    let name = lexical_qname(xml, node)?;
    Ok(name.split_once(':').map(|(prefix, _)| prefix.to_owned()))
}

fn qname_prefix(qname: &str, local: &str) -> String {
    qname.split_once(':').map_or_else(
        || local.to_owned(),
        |(prefix, _)| format!("{prefix}:{local}"),
    )
}

fn closing_tag_start(xml: &str, node: Node<'_, '_>) -> Result<usize> {
    let range = node.range();
    xml[..range.end]
        .rfind("</")
        .filter(|offset| *offset >= range.start)
        .context("XML element has no closing tag")
}

fn is_self_closing(xml: &str, range: Range<usize>) -> Result<bool> {
    let close = xml[range.clone()]
        .find('>')
        .map(|offset| range.start + offset)
        .context("XML element has no closing angle bracket")?;
    Ok(xml[range.start..=close].trim_end().ends_with("/>"))
}

fn apply_xml_edits(xml: &str, mut edits: Vec<(Range<usize>, Vec<u8>)>) -> Result<String> {
    edits.sort_by_key(|(range, _)| (std::cmp::Reverse(range.start), std::cmp::Reverse(range.end)));
    let mut output = xml.as_bytes().to_vec();
    let mut previous_start = usize::MAX;
    for (range, replacement) in edits {
        ensure!(
            range.end <= previous_start,
            "overlapping native OOXML edits"
        );
        output.splice(range.clone(), replacement);
        previous_start = range.start;
    }
    String::from_utf8(output).context("native OOXML output is not UTF-8")
}

fn insert_at(xml: &str, offset: usize, fragment: &str) -> Result<String> {
    ensure!(
        offset <= xml.len(),
        "XML insertion offset is outside the document"
    );
    Ok(replace_range(xml, offset..offset, fragment.as_bytes()))
}

fn replace_range(xml: &str, range: Range<usize>, replacement: &[u8]) -> String {
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml.as_bytes()[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml.as_bytes()[range.end..]);
    String::from_utf8(output).expect("replacement preserves UTF-8")
}

fn escape_attribute(value: &str) -> String {
    quick_xml::escape::escape(value).replace('"', "&quot;")
}

fn xml_bool(value: bool) -> &'static str {
    if value {
        "1"
    } else {
        "0"
    }
}

#[cfg(test)]
mod dimension_tests {
    use super::*;

    #[test]
    fn native_validator_rejects_stale_dimension_without_changing_source() -> Result<()> {
        let root = std::env::temp_dir().join(format!(
            "native-dimension-validation-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        fs::create_dir(&root)?;
        let input = root.join("source.xlsx");
        let output = root.join("candidate.xlsx");
        let mut book = rust_xlsxwriter::Workbook::new();
        let sheet = book.add_worksheet();
        sheet.set_name("Data")?.write_string(0, 0, "header")?;
        book.save(&input)?;
        let original = fs::read(&input)?;
        let plan = ExcelEditPlan {
            request_id: "native-dimension-validation".into(),
            input: input.clone(),
            output: output.clone(),
            format: WorkbookFormat::Xlsx,
            workbook: crate::inspect_xlsx_profile(&input)?,
            operations: vec![EditOperation::SetCellValue {
                sheet: "Data".into(),
                cell: "AA100".into(),
                value: crate::CellValue::Text("tail".into()),
            }],
            preserve: [crate::PreservationRequirement::CellValuesAndTypes].into(),
            drop_allowed: Default::default(),
        };
        let edited = RustNativeOoxmlBackend::apply_bytes(&original, &plan)?;
        fs::write(&output, &edited)?;
        NativeOoxmlValidator.validate(&plan, &output)?;
        let mut archive = ZipArchive::new(Cursor::new(edited))?;
        let part = "xl/worksheets/sheet1.xml";
        let xml = String::from_utf8(read_part(&mut archive, part)?)?;
        ensure!(xml.contains("A1:AA100"), "fixture dimension missing");
        let replacements =
            BTreeMap::from([(part.to_owned(), xml.replace("A1:AA100", "A1").into_bytes())]);
        fs::write(&output, rewrite_package(&mut archive, &replacements)?)?;
        let error = NativeOoxmlValidator.validate(&plan, &output).unwrap_err();
        assert!(
            error.to_string().contains("dimension does not cover"),
            "{error:#}"
        );
        assert_eq!(fs::read(&input)?, original);
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn native_cell_append_expands_dimension_at_row_and_column_boundaries() -> Result<()> {
        let xml = format!(
            r#"<worksheet xmlns="{MAIN_NS}"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>header</t></is></c></row></sheetData></worksheet>"#
        );
        let updates = ["Z9", "AA10", "Z99", "AA100"]
            .into_iter()
            .map(|address| {
                let (row, column, address) = parse_cell_address(address).unwrap();
                CellUpdate {
                    row,
                    column,
                    address,
                    value: crate::CellValue::Decimal("0".into()),
                }
            })
            .collect::<Vec<_>>();
        let result = patch_cells_xml(&xml, &updates)?;
        let doc = Document::parse(&result)?;
        assert_eq!(
            doc.descendants()
                .find(|n| is_element(*n, "dimension"))
                .and_then(|n| n.attribute("ref")),
            Some("A1:AA100")
        );
        for update in updates {
            assert!(doc
                .descendants()
                .any(|n| is_element(n, "c") && n.attribute("r") == Some(update.address.as_str())));
        }
        assert!(result.contains("header"));
        Ok(())
    }

    #[test]
    fn dimensions_handle_missing_prefixed_and_oversized_declarations_without_shrinking(
    ) -> Result<()> {
        for (xml, expected) in [
            (
                format!(r#"<worksheet xmlns="{MAIN_NS}"><sheetData/></worksheet>"#),
                "AA100",
            ),
            (
                format!(r#"<x:worksheet xmlns:x="{MAIN_NS}"><x:sheetData /></x:worksheet>"#),
                "AA100",
            ),
            (
                format!(
                    r#"<worksheet xmlns="{MAIN_NS}"><dimension ref="A1:XFD1048576"/><sheetData/></worksheet>"#
                ),
                "A1:XFD1048576",
            ),
        ] {
            let (row, column, address) = parse_cell_address("AA100")?;
            let result = patch_cells_xml(
                &xml,
                &[CellUpdate {
                    row,
                    column,
                    address,
                    value: crate::CellValue::Text("tail".into()),
                }],
            )?;
            let doc = Document::parse(&result)?;
            assert_eq!(
                doc.root_element()
                    .children()
                    .find(|n| is_element(*n, "dimension"))
                    .and_then(|n| n.attribute("ref")),
                Some(expected)
            );
            let children: Vec<_> = doc
                .root_element()
                .children()
                .filter(|n| n.is_element())
                .map(|n| n.tag_name().name())
                .collect();
            assert_eq!(children, ["dimension", "sheetData"]);
        }
        for invalid in ["A0:B2", "XFE1", "B2:A1", "A1:B2:C3", ""] {
            let xml = format!(
                r#"<worksheet xmlns="{MAIN_NS}"><dimension ref="{invalid}"/><sheetData/></worksheet>"#
            );
            assert!(
                expand_worksheet_dimension(&xml).is_err(),
                "invalid ref: {invalid}"
            );
        }
        Ok(())
    }
}
