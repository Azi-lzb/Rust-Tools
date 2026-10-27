use anyhow::{bail, ensure, Context, Result};
use calamine::{open_workbook_auto, Data, Reader, Sheets};
use roxmltree::Document;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::{
    collections::BTreeMap,
    fs,
    io::{Read, Write},
    ops::Range,
    path::Path,
};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

const SPREADSHEET_MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// Engine-independent representation of a cell value returned by the reader.
/// The tagged JSON form is part of the existing WorkbookDump probe contract.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value")]
pub enum CellValue {
    #[serde(rename = "empty")]
    Empty,
    #[serde(rename = "string")]
    Text(String),
    #[serde(rename = "number")]
    Number(f64),
    #[serde(rename = "integer")]
    Integer(i64),
    #[serde(rename = "boolean")]
    Boolean(bool),
    #[serde(rename = "datetime_serial")]
    DateTimeSerial(f64),
    #[serde(rename = "datetime_iso")]
    DateTimeIso(String),
    #[serde(rename = "duration_iso")]
    DurationIso(String),
    #[serde(rename = "error")]
    Error(String),
}

impl From<&Data> for CellValue {
    fn from(data: &Data) -> Self {
        match data {
            Data::Empty => Self::Empty,
            Data::String(value) => Self::Text(value.clone()),
            Data::Float(value) => Self::Number(*value),
            Data::Int(value) => Self::Integer(*value),
            Data::Bool(value) => Self::Boolean(*value),
            Data::DateTime(value) => Self::DateTimeSerial(value.as_f64()),
            Data::DateTimeIso(value) => Self::DateTimeIso(value.clone()),
            Data::DurationIso(value) => Self::DurationIso(value.clone()),
            Data::Error(value) => Self::Error(format!("{value:?}")),
        }
    }
}

impl CellValue {
    /// Decode the tagged JSON representation emitted by `WorkbookDump`.
    /// Non-finite numbers are emitted by `json!` with a null value; decode those
    /// as NaN so callers can consistently classify them as non-finite.
    pub fn from_probe_json(value: &Value) -> Result<Self> {
        if value.get("type").and_then(Value::as_str) == Some("number")
            && value.get("value").is_some_and(Value::is_null)
        {
            return Ok(Self::Number(f64::NAN));
        }
        serde_json::from_value(value.clone()).context("decode workbook cell JSON")
    }

    fn to_probe_json(self) -> Option<Value> {
        Some(match self {
            Self::Empty => return None,
            Self::Text(value) => json!({"type":"string", "value":value}),
            Self::Number(value) => json!({"type":"number", "value":value}),
            Self::Integer(value) => json!({"type":"integer", "value":value}),
            Self::Boolean(value) => json!({"type":"boolean", "value":value}),
            Self::DateTimeSerial(value) => json!({"type":"datetime_serial", "value":value}),
            Self::DateTimeIso(value) => json!({"type":"datetime_iso", "value":value}),
            Self::DurationIso(value) => json!({"type":"duration_iso", "value":value}),
            Self::Error(value) => json!({"type":"error", "value":value}),
        })
    }
}

#[derive(Debug, Serialize)]
pub struct SheetDump {
    pub name: String,
    pub values: BTreeMap<String, Value>,
    pub formulas: BTreeMap<String, String>,
    pub formula_status: String,
    pub merges: Vec<String>,
    pub merge_status: String,
    pub comments: BTreeMap<String, String>,
    pub comments_status: String,
    pub conditional_formatting: Vec<Value>,
    pub conditional_formatting_status: String,
    pub data_validations: Vec<Value>,
    pub data_validations_status: String,
}

/// A traditional OOXML cell comment with its author resolved from the
/// worksheet's comments part. Calamine does not expose this metadata.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct WorkbookComment {
    pub text: String,
    pub author: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct WorkbookDump {
    pub schema_version: u32,
    pub reader: &'static str,
    pub elapsed_ms: u128,
    pub sheets: Vec<SheetDump>,
}

/// Diagnostics returned by the opt-in unsupported error-token compatibility
/// reader. The count is safe for structured logs and does not expose cell data.
#[derive(Debug)]
pub struct UnknownErrorTokenRead {
    pub workbook: WorkbookDump,
    pub compatibility_normalization_used: bool,
    pub normalized_cell_count: usize,
}

/// A lightweight projection of stored values from a sheet.
///
/// Values use absolute A1 coordinates and preserve the same typed `CellValue`
/// interpretation as `WorkbookDump`. The calamine `Range` is still materialized
/// for the used sheet, but callers can avoid serializing unrelated columns and
/// reading OOXML metadata they do not need.
#[derive(Clone, Debug, PartialEq)]
pub struct SheetValues {
    pub name: String,
    pub values: BTreeMap<String, CellValue>,
    /// Absolute worksheet extent reported by Calamine, in one-based rows and columns.
    /// Empty worksheets report `(0, 0)`; callers that need application-style defaults
    /// should apply those at their own contract boundary.
    pub max_row: usize,
    pub max_column: usize,
}

/// Cell and merged-range bounds stored in one XLSX worksheet part. This is a
/// small OOXML metadata projection used when Calamine's value range does not
/// include styled blank cells or merge-only coordinates. Zero min/max values
/// mean that the worksheet contains no stored cells and no merged ranges.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct WorksheetExtent {
    pub name: String,
    pub min_row: usize,
    pub min_column: usize,
    pub max_row: usize,
    pub max_column: usize,
}

/// One worksheet cell whose serialized number format is requested from OOXML.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CellNumberFormatRequest {
    pub sheet_name: String,
    /// A1 cell reference such as `B7` or `$B$7` (without a sheet prefix).
    pub cell_reference: String,
}

/// Number-format metadata returned by [`read_xlsx_cell_number_formats`].
///
/// `format_code` contains the exact custom code from `styles.xml`, or a
/// canonical code for a recognized built-in number format. It is `None` for
/// reserved/unknown built-in IDs; `number_format_id` remains available so the
/// caller can handle those explicitly. This legacy API reports an explicit
/// cell-level `s` style reference. If no cell-level style exists but a matching
/// row or column default style is present, it fails closed instead of assuming
/// the default style; use [`EffectiveCellNumberFormatInfo`] to query resolved
/// row/column inheritance.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CellNumberFormatInfo {
    pub sheet_name: String,
    /// Normalized A1 reference (uppercase column letters, no `$` markers).
    pub cell_reference: String,
    pub style_index: u32,
    pub number_format_id: u32,
    pub format_code: Option<String>,
}

/// The OOXML style location used to resolve a cell's number format.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CellNumberFormatSource {
    /// An explicit `s` attribute on the cell.
    Cell,
    /// A row-level `s` style with `customFormat="1"`.
    Row,
    /// A column-level `style` attribute covering the cell's column.
    Column,
    /// The workbook's default cell style (cellXfs index zero).
    Default,
}

/// Number-format metadata resolved from the cell, row, column, or default style.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EffectiveCellNumberFormatInfo {
    pub sheet_name: String,
    /// Normalized A1 reference (uppercase column letters, no `$` markers).
    pub cell_reference: String,
    pub style_index: u32,
    pub number_format_id: u32,
    pub format_code: Option<String>,
    pub source: CellNumberFormatSource,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct InheritedNumberFormatStyle {
    style_index: u32,
    source: CellNumberFormatSource,
    description: String,
}

#[derive(Default)]
struct OoxmlSheetMetadata {
    comments: BTreeMap<String, String>,
    comment_records: BTreeMap<String, WorkbookComment>,
    conditional_formatting: Vec<Value>,
    data_validations: Vec<Value>,
}

/// The three SpreadsheetML worksheet header variants.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd)]
pub enum HeaderVariant {
    Odd,
    Even,
    First,
}

/// Plain text from one worksheet header variant, including Excel control codes.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct WorksheetHeaderText {
    pub sheet_name: String,
    pub variant: HeaderVariant,
    pub text: String,
}

/// Header text inventory for one worksheet, in workbook sheet order.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct WorksheetHeaderInventory {
    pub sheet_name: String,
    pub headers: Vec<WorksheetHeaderText>,
}

/// Scope of a workbook defined name. `local_sheet_id` follows the zero-based
/// index in the workbook's `<sheets>` collection, including non-worksheet tabs.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum DefinedNameScope {
    Workbook,
    Worksheet {
        local_sheet_id: usize,
        sheet_name: Option<String>,
        sheet_kind: Option<String>,
    },
    UnresolvedWorksheet {
        local_sheet_id: usize,
    },
    InvalidWorksheetScope {
        raw_local_sheet_id: String,
    },
}

/// A resolved A1 area in a defined-name formula.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct DefinedNameArea {
    pub sheet_name: String,
    pub sheet_kind: String,
    pub reference: String,
}

/// How much of the defined-name formula this reader understands.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum DefinedNameParseStatus {
    Areas,
    NonAreaExpression,
    UnsupportedAreaExpression,
    InvalidAreaExpression,
}

/// Namespace-aware inventory of workbook defined names. `raw_reference` is
/// preserved exactly as stored (apart from XML entity decoding); `areas` is
/// populated only when every union member is a supported A1 area reference.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct DefinedNameRecord {
    pub name: String,
    pub scope: DefinedNameScope,
    pub raw_reference: String,
    pub parse_status: DefinedNameParseStatus,
    pub areas: Vec<DefinedNameArea>,
}

#[derive(Clone, Debug)]
struct WorkbookSheetIdentity {
    name: String,
    kind: String,
}

/// Read workbook and worksheet-scoped defined names from XLSX/XLSM OOXML.
/// This is a metadata adapter and does not read or infer cell values. Supported
/// area expressions are returned with sheet identity; formulas, constants,
/// external references and unsupported reference forms remain available in
/// `raw_reference` with an explicit parse status.
pub fn read_defined_names(path: impl AsRef<Path>) -> Result<Vec<DefinedNameRecord>> {
    let path = path.as_ref();
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "defined-name reader only supports .xlsx and .xlsm"
    );
    let mut zip = ZipArchive::new(fs::File::open(path).context("open workbook package")?)
        .context("open workbook ZIP package")?;
    let (workbook_part, sheets) = workbook_sheet_identities(&mut zip)?;
    let workbook_xml = zip_read(&mut zip, &workbook_part).context("read workbook part")?;
    let document = Document::parse(&workbook_xml).context("parse workbook part")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && root.tag_name().name() == "workbook",
        "workbook part has an unexpected root element"
    );

    let Some(container) = root.children().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "definedNames"
    }) else {
        return Ok(Vec::new());
    };
    let mut records = Vec::new();
    for node in container.children().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "definedName"
    }) {
        let name = node.attribute("name").context("definedName has no name")?;
        let raw_reference = node.text().unwrap_or_default().to_owned();
        let scope = match node.attribute("localSheetId") {
            None => DefinedNameScope::Workbook,
            Some(raw) => match raw.parse::<usize>() {
                Ok(local_sheet_id) => match sheets.get(local_sheet_id) {
                    Some(sheet) => DefinedNameScope::Worksheet {
                        local_sheet_id,
                        sheet_name: Some(sheet.name.clone()),
                        sheet_kind: Some(sheet.kind.clone()),
                    },
                    None => DefinedNameScope::UnresolvedWorksheet { local_sheet_id },
                },
                Err(_) => DefinedNameScope::InvalidWorksheetScope {
                    raw_local_sheet_id: raw.to_owned(),
                },
            },
        };
        let (parse_status, areas) = parse_defined_name_areas(&raw_reference, &sheets);
        records.push(DefinedNameRecord {
            name: name.to_owned(),
            scope,
            raw_reference,
            parse_status,
            areas,
        });
    }
    Ok(records)
}

fn workbook_sheet_identities(
    zip: &mut ZipArchive<fs::File>,
) -> Result<(String, Vec<WorkbookSheetIdentity>)> {
    let root_rels_xml = zip_read(zip, "_rels/.rels")?;
    let root_rels = Document::parse(&root_rels_xml)?;
    let workbook_target = root_rels
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
        })
        .find_map(|relationship| {
            let workbook = relationship
                .attribute("Type")
                .is_some_and(|kind| kind.ends_with("/officeDocument"));
            let external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            (workbook && !external)
                .then(|| relationship.attribute("Target"))
                .flatten()
        })
        .context("package relationships do not identify an internal officeDocument part")?;
    let workbook_part = norm_part("", workbook_target);
    let workbook_parent = workbook_part
        .rsplit_once('/')
        .map(|(parent, _)| parent)
        .unwrap_or("");
    let workbook_xml = zip_read(zip, &workbook_part)?;
    let workbook_doc = Document::parse(&workbook_xml)?;
    let workbook_rels_xml = zip_read(zip, &relationships_part(&workbook_part))?;
    let rels_doc = Document::parse(&workbook_rels_xml)?;
    let mut rels = BTreeMap::new();
    for relationship in rels_doc.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
            && node.tag_name().name() == "Relationship"
    }) {
        if let (Some(id), Some(kind), Some(target)) = (
            relationship.attribute("Id"),
            relationship.attribute("Type"),
            relationship.attribute("Target"),
        ) {
            let external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            rels.insert(
                id.to_owned(),
                (kind.to_owned(), target.to_owned(), external),
            );
        }
    }
    let mut sheets = Vec::new();
    for sheet in workbook_doc.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "sheet"
    }) {
        let name = sheet
            .attribute("name")
            .context("workbook sheet has no name")?;
        let rel_id = sheet
            .attribute((OFFICE_REL_NS, "id"))
            .context("workbook sheet has no relationship id")?;
        let (kind, target, external) = rels
            .get(rel_id)
            .with_context(|| format!("unresolved relationship for sheet {name:?}"))?;
        ensure!(
            !*external,
            "workbook sheet {name:?} has an external relationship"
        );
        let part = norm_part(workbook_parent, target);
        ensure!(
            !part.is_empty(),
            "workbook sheet {name:?} has an empty part path"
        );
        sheets.push(WorkbookSheetIdentity {
            name: name.to_owned(),
            kind: kind.rsplit('/').next().unwrap_or(kind).to_owned(),
        });
    }
    Ok((workbook_part, sheets))
}

fn parse_defined_name_areas(
    expression: &str,
    sheets: &[WorkbookSheetIdentity],
) -> (DefinedNameParseStatus, Vec<DefinedNameArea>) {
    if expression.contains('(') || expression.contains(')') {
        return (DefinedNameParseStatus::NonAreaExpression, Vec::new());
    }
    let parts = split_formula_union(expression);
    if parts.is_empty() {
        return (DefinedNameParseStatus::NonAreaExpression, Vec::new());
    }
    let mut areas = Vec::new();
    for part in parts {
        let Some((sheet_token, reference)) = part.rsplit_once('!') else {
            let status = if is_a1_area(part) {
                DefinedNameParseStatus::UnsupportedAreaExpression
            } else {
                DefinedNameParseStatus::NonAreaExpression
            };
            return (status, Vec::new());
        };
        let Some(sheet_name) = decode_sheet_token(sheet_token) else {
            return (
                DefinedNameParseStatus::UnsupportedAreaExpression,
                Vec::new(),
            );
        };
        if !is_a1_area(reference) {
            return (
                DefinedNameParseStatus::UnsupportedAreaExpression,
                Vec::new(),
            );
        }
        let Some(sheet) = sheets.iter().find(|sheet| sheet.name == sheet_name) else {
            return (DefinedNameParseStatus::InvalidAreaExpression, Vec::new());
        };
        areas.push(DefinedNameArea {
            sheet_name: sheet.name.clone(),
            sheet_kind: sheet.kind.clone(),
            reference: reference.to_owned(),
        });
    }
    (DefinedNameParseStatus::Areas, areas)
}

fn split_formula_union(expression: &str) -> Vec<&str> {
    split_formula_union_bytes(expression)
}

fn split_formula_union_bytes(expression: &str) -> Vec<&str> {
    let (mut start, mut quoted) = (0, false);
    let mut parts = Vec::new();
    let chars: Vec<(usize, char)> = expression.char_indices().collect();
    let mut i = 0;
    while i < chars.len() {
        let (byte, ch) = chars[i];
        if ch == '\'' {
            if quoted && chars.get(i + 1).is_some_and(|(_, next)| *next == '\'') {
                i += 2;
                continue;
            }
            quoted = !quoted;
        } else if ch == ',' && !quoted {
            parts.push(expression[start..byte].trim());
            start = byte + 1;
        }
        i += 1;
    }
    if quoted {
        return Vec::new();
    }
    parts.push(expression[start..].trim());
    parts
}

fn decode_sheet_token(token: &str) -> Option<String> {
    let token = token.strip_prefix('=').unwrap_or(token);
    let decoded = if token.starts_with('\'') && token.ends_with('\'') && token.len() >= 2 {
        token[1..token.len() - 1].replace("''", "'")
    } else if token
        .chars()
        .all(|ch| ch.is_alphanumeric() || ch == '_' || ch == '.')
    {
        token.to_owned()
    } else {
        return None;
    };
    // External workbook qualifiers and 3-D references are deliberately unsupported.
    (!decoded.contains('[') && !decoded.contains(']') && !decoded.contains(':')).then_some(decoded)
}

fn is_a1_area(reference: &str) -> bool {
    fn cell(value: &str) -> bool {
        let value = value.strip_prefix('$').unwrap_or(value);
        let letters = value
            .chars()
            .take_while(|ch| ch.is_ascii_alphabetic())
            .count();
        if letters == 0 || letters > 3 {
            return false;
        }
        let column = value[..letters].chars().fold(0_usize, |column, ch| {
            column * 26 + (ch.to_ascii_uppercase() as usize - 'A' as usize + 1)
        });
        if column > 16_384 {
            return false;
        }
        let rest = &value[letters..];
        let rest = rest.strip_prefix('$').unwrap_or(rest);
        !rest.is_empty()
            && rest.chars().all(|ch| ch.is_ascii_digit())
            && rest
                .parse::<u32>()
                .is_ok_and(|row| (1..=1_048_576).contains(&row))
    }
    let mut pair = reference.split(':');
    let first = pair.next().unwrap_or_default();
    let second = pair.next();
    pair.next().is_none() && cell(first) && second.map_or(true, cell)
}

/// Read odd/even/first-page header text from an OOXML workbook.
///
/// This is an OOXML metadata read, separate from Calamine cell-value reading.
/// Rich/nested header markup is rejected so callers cannot mistake a partial
/// text projection for the original header contract.
pub fn read_header_inventory(path: impl AsRef<Path>) -> Result<Vec<WorksheetHeaderInventory>> {
    let path = path.as_ref();
    let is_ooxml = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            matches!(
                extension.to_ascii_lowercase().as_str(),
                "xlsx" | "xlsm" | "xlam"
            )
        });
    ensure!(
        is_ooxml,
        "header metadata reader only supports XLSX/XLSM/XLAM"
    );

    let mut zip = ZipArchive::new(fs::File::open(path)?)?;
    let worksheets = ooxml_worksheet_parts(&mut zip)?;
    let mut inventory = Vec::with_capacity(worksheets.len());
    for (sheet_name, sheet_part) in worksheets {
        let xml = zip_read(&mut zip, &sheet_part)?;
        let document =
            Document::parse(&xml).with_context(|| format!("parse worksheet part {sheet_part}"))?;
        let mut headers = Vec::new();
        if let Some(header_footer) = document.root_element().descendants().find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "headerFooter"
        }) {
            let mut seen = std::collections::BTreeSet::new();
            for header in header_footer.children().filter(|node| node.is_element()) {
                if header.tag_name().namespace() != Some(SPREADSHEET_MAIN_NS) {
                    continue;
                }
                let variant = match header.tag_name().name() {
                    "oddHeader" => HeaderVariant::Odd,
                    "evenHeader" => HeaderVariant::Even,
                    "firstHeader" => HeaderVariant::First,
                    _ => continue,
                };
                ensure!(
                    seen.insert(variant),
                    "worksheet {sheet_name:?} contains duplicate {variant:?} header elements"
                );
                if header.children().any(|child| child.is_element()) {
                    bail!(
                        "worksheet {sheet_name:?} contains unsupported nested markup in {variant:?} header"
                    );
                }
                let text = header
                    .children()
                    .filter(|child| child.is_text())
                    .filter_map(|child| child.text())
                    .collect();
                headers.push(WorksheetHeaderText {
                    sheet_name: sheet_name.clone(),
                    variant,
                    text,
                });
            }
        }
        inventory.push(WorksheetHeaderInventory {
            sheet_name,
            headers,
        });
    }
    Ok(inventory)
}

fn value_json(data: &Data) -> Option<Value> {
    CellValue::from(data).to_probe_json()
}

fn column_letters(mut col: usize) -> String {
    let mut out = String::new();
    loop {
        let rem = col % 26;
        out.insert(0, (b'A' + rem as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    out
}

fn coord(row: usize, col: usize) -> String {
    format!("{}{}", column_letters(col), row + 1)
}

fn dimensions_to_range(start: (u32, u32), end: (u32, u32)) -> String {
    format!(
        "{}{}:{}{}",
        column_letters(start.1 as usize),
        start.0 + 1,
        column_letters(end.1 as usize),
        end.0 + 1
    )
}

fn merges_for(
    workbook: &mut Sheets<std::io::BufReader<fs::File>>,
    name: &str,
) -> Result<(Vec<String>, &'static str)> {
    match workbook {
        Sheets::Xlsx(reader) => Ok((
            reader
                .merge_cells_by_sheet_name(name)?
                .iter()
                .map(|d| dimensions_to_range(d.start, d.end))
                .collect(),
            "read_by_calamine",
        )),
        Sheets::Xls(reader) => Ok((
            reader
                .merge_cells_by_sheet_name(name)?
                .iter()
                .map(|d| dimensions_to_range(d.start, d.end))
                .collect(),
            "read_by_calamine",
        )),
        Sheets::Xlsb(_) => Ok((
            Vec::new(),
            "unsupported_api: Calamine 0.37.0 exposes no xlsb merge reader",
        )),
        Sheets::Ods(_) => Ok((Vec::new(), "out_of_scope_format")),
    }
}

pub fn read_workbook(path: impl AsRef<Path>) -> Result<WorkbookDump> {
    let path = path.as_ref();
    validate_legacy_xls_signature(path)?;
    let started = std::time::Instant::now();
    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    let is_ooxml = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| matches!(e.to_ascii_lowercase().as_str(), "xlsx" | "xlsm" | "xlam"))
        .unwrap_or(false);
    let metadata_by_sheet = if is_ooxml {
        extract_ooxml_metadata(path)?
    } else {
        BTreeMap::new()
    };
    let mut sheets = Vec::new();
    for name in names {
        let range = workbook
            .worksheet_range(&name)
            .with_context(|| format!("read values for {name}"))?;
        let (start_row, start_col) = range.start().unwrap_or((0, 0));
        let mut values = BTreeMap::new();
        for (local_row, local_col, value) in range.cells() {
            if let Some(v) = value_json(value) {
                values.insert(
                    coord(
                        start_row as usize + local_row,
                        start_col as usize + local_col,
                    ),
                    v,
                );
            }
        }
        let (formulas, formula_status) = match workbook.worksheet_formula(&name) {
            Ok(range) => {
                let (fr, fc) = range.start().unwrap_or((0, 0));
                (
                    range
                        .cells()
                        .filter_map(|(r, c, v)| {
                            (!v.trim().is_empty())
                                .then(|| (coord(fr as usize + r, fc as usize + c), v.to_owned()))
                        })
                        .collect(),
                    "read_by_calamine".to_owned(),
                )
            }
            Err(error) => (BTreeMap::new(), format!("read_error:{error}")),
        };
        let (merges, merge_status) = merges_for(&mut workbook, &name)
            .with_context(|| format!("read merge regions for {name}"))?;
        let (
            comments,
            comments_status,
            conditional_formatting,
            conditional_formatting_status,
            data_validations,
            data_validations_status,
        ) = if is_ooxml {
            let metadata = metadata_by_sheet
                .get(&name)
                .with_context(|| format!("OOXML metadata missing for sheet {name}"))?;
            (
                metadata.comments.clone(),
                "read_by_namespace_aware_ooxml_supplement".to_owned(),
                metadata.conditional_formatting.clone(),
                "read_by_namespace_aware_ooxml_supplement".to_owned(),
                metadata.data_validations.clone(),
                "read_by_namespace_aware_ooxml_supplement".to_owned(),
            )
        } else {
            (
                BTreeMap::new(),
                "unsupported_api: comment adapter only supports OOXML xlsx/xlsm".to_owned(),
                Vec::new(),
                "unsupported_api: conditional-formatting adapter only supports OOXML xlsx/xlsm"
                    .to_owned(),
                Vec::new(),
                "unsupported_api: data-validation adapter only supports OOXML xlsx/xlsm".to_owned(),
            )
        };
        sheets.push(SheetDump {
            name,
            values,
            formulas,
            formula_status,
            merges,
            merge_status: merge_status.to_owned(),
            comments,
            comments_status,
            conditional_formatting,
            conditional_formatting_status,
            data_validations,
            data_validations_status,
        });
    }
    Ok(WorkbookDump {
        schema_version: 3,
        reader: "calamine",
        elapsed_ms: started.elapsed().as_millis(),
        sheets,
    })
}

/// Read traditional cell comments and author names from an OOXML workbook.
///
/// The map is keyed by worksheet name and then absolute A1 coordinate. This is
/// a metadata-only supplement; cell values continue to be read through
/// [`read_workbook`] and Calamine.
pub fn read_ooxml_comments(
    path: impl AsRef<Path>,
) -> Result<BTreeMap<String, BTreeMap<String, WorkbookComment>>> {
    let path = path.as_ref();
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    ensure!(
        matches!(extension.as_str(), "xlsx" | "xlsm" | "xlam"),
        "OOXML comment reader only supports .xlsx, .xlsm, and .xlam files"
    );
    let metadata = extract_ooxml_metadata(path)?;
    Ok(metadata
        .into_iter()
        .map(|(sheet, metadata)| (sheet, metadata.comment_records))
        .collect())
}

/// Read a workbook while treating OOXML error tokens unsupported by Calamine
/// as ordinary text values. The default `read_workbook` contract is unchanged:
/// this compatibility path is opt-in, limited to XLSX/XLSM, and normalizes
/// only literal `<c t="e"><v>…</v></c>` cells in a temporary package copy.
/// Formula cells and malformed/unknown error-cell structures are rejected.
pub fn read_workbook_with_unknown_error_tokens_as_text(
    path: impl AsRef<Path>,
) -> Result<UnknownErrorTokenRead> {
    let path = path.as_ref();
    match read_workbook(path) {
        Ok(workbook) => {
            return Ok(UnknownErrorTokenRead {
                workbook,
                compatibility_normalization_used: false,
                normalized_cell_count: 0,
            });
        }
        Err(error) if !format!("{error:#}").contains("Unsupported cell error value") => {
            return Err(error);
        }
        Err(_) => {}
    }
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")
                || extension.eq_ignore_ascii_case("xlsm")),
        "unknown error-token compatibility reader supports XLSX/XLSM only"
    );
    let (replacements, normalized_cell_count) = prepare_unknown_error_token_replacements(path)?;
    ensure!(!replacements.is_empty(), "Calamine reported an unsupported cell error but no literal unsupported OOXML error token was found");
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)?
        .as_nanos();
    let temp = std::env::temp_dir().join(format!(
        "excel-reader-error-text-{}-{nonce}.xlsx",
        std::process::id()
    ));
    let write_result = write_package_with_replacements(path, &temp, &replacements);
    if let Err(error) = write_result {
        let _ = fs::remove_file(&temp);
        return Err(error)
            .context("write temporary workbook with unsupported error tokens normalized");
    }
    let result = read_workbook(&temp).context("read normalized temporary workbook with Calamine");
    let _ = fs::remove_file(&temp);
    Ok(UnknownErrorTokenRead {
        workbook: result?,
        compatibility_normalization_used: true,
        normalized_cell_count,
    })
}

fn prepare_unknown_error_token_replacements(
    path: &Path,
) -> Result<(BTreeMap<String, Vec<u8>>, usize)> {
    let mut zip = ZipArchive::new(fs::File::open(path).context("open workbook package")?)?;
    let mut replacements = BTreeMap::new();
    let mut normalized_cell_count = 0;
    for (_, part) in ooxml_worksheet_parts(&mut zip)? {
        let xml = zip_read(&mut zip, &part)?.into_bytes();
        let (patched, cell_count) = normalize_unknown_error_cells(&xml)?;
        if patched != xml {
            replacements.insert(part, patched);
            normalized_cell_count += cell_count;
        }
    }
    Ok((replacements, normalized_cell_count))
}

fn normalize_unknown_error_cells(xml: &[u8]) -> Result<(Vec<u8>, usize)> {
    let document = Document::parse(std::str::from_utf8(xml)?)?;
    let mut edits = Vec::<(Range<usize>, Vec<u8>)>::new();
    let mut normalized_cell_count = 0;
    for cell in document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "c"
            && node.attribute("t") == Some("e")
    }) {
        ensure!(
            cell.attribute("r").is_some(),
            "error cell is missing a coordinate"
        );
        let value_nodes = cell
            .children()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && node.tag_name().name() == "v"
            })
            .collect::<Vec<_>>();
        ensure!(
            value_nodes.len() == 1,
            "error cell {:?} must contain exactly one cached value",
            cell.attribute("r")
        );
        ensure!(
            cell.children().all(|child| !child.is_element()
                || (child.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && matches!(child.tag_name().name(), "v" | "f"))),
            "unsupported child element in literal error cell {:?}",
            cell.attribute("r")
        );
        let value = value_nodes[0].text().unwrap_or_default();
        if is_calamine_error_token(value) {
            continue;
        }
        ensure!(
            cell.children().all(|child| {
                !child.is_element()
                    || child.tag_name().namespace() != Some(SPREADSHEET_MAIN_NS)
                    || child.tag_name().name() != "f"
            }),
            "formula error cell {:?} cannot be normalized as a literal string",
            cell.attribute("r")
        );
        let type_attribute = cell
            .attribute_node("t")
            .context("error cell is missing type attribute")?;
        edits.push((type_attribute.range_value(), b"inlineStr".to_vec()));
        let value_node = value_nodes[0];
        let prefix = xml_element_qname(xml, cell)?;
        let is_name = child_qname(&prefix, "is");
        let text_name = child_qname(&prefix, "t");
        let text = escape_xml_text(value);
        let replacement = format!("<{is_name}><{text_name}>{text}</{text_name}></{is_name}>");
        edits.push((value_node.range(), replacement.into_bytes()));
        normalized_cell_count += 1;
    }
    Ok((apply_xml_edits(xml, edits)?, normalized_cell_count))
}

fn is_calamine_error_token(value: &str) -> bool {
    matches!(
        value,
        "#DIV/0!" | "#N/A" | "#NAME?" | "#NULL!" | "#NUM!" | "#REF!" | "#VALUE!"
    )
}

fn xml_element_qname(xml: &[u8], node: roxmltree::Node<'_, '_>) -> Result<String> {
    let fragment = std::str::from_utf8(&xml[node.range()])?;
    let start = fragment
        .find('<')
        .context("XML element has no opening tag")?
        + 1;
    let end = fragment[start..]
        .find(|ch: char| ch.is_whitespace() || matches!(ch, '>' | '/'))
        .map(|offset| start + offset)
        .context("XML element has no name")?;
    Ok(fragment[start..end].to_owned())
}

fn child_qname(parent_qname: &str, child: &str) -> String {
    parent_qname
        .rsplit_once(':')
        .map(|(prefix, _)| format!("{prefix}:{child}"))
        .unwrap_or_else(|| child.to_owned())
}

fn escape_xml_text(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn apply_xml_edits(xml: &[u8], mut edits: Vec<(Range<usize>, Vec<u8>)>) -> Result<Vec<u8>> {
    edits.sort_by_key(|(range, _)| std::cmp::Reverse((range.start, range.end)));
    let mut output = xml.to_vec();
    for (range, replacement) in edits {
        ensure!(
            range.start <= range.end && range.end <= output.len(),
            "OOXML edit range is out of bounds"
        );
        output.splice(range, replacement);
    }
    Ok(output)
}

fn write_package_with_replacements(
    source: &Path,
    output: &Path,
    replacements: &BTreeMap<String, Vec<u8>>,
) -> Result<()> {
    let mut input = ZipArchive::new(fs::File::open(source)?)?;
    let file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(output)?;
    let mut writer = ZipWriter::new(file);
    for index in 0..input.len() {
        let mut entry = input.by_index(index)?;
        let name = entry.name().to_owned();
        let mut options = SimpleFileOptions::default().compression_method(entry.compression());
        if let Some(modified) = entry.last_modified() {
            options = options.last_modified_time(modified);
        }
        if entry.is_dir() {
            writer.add_directory(name, options)?;
            continue;
        }
        writer.start_file(&name, options)?;
        if let Some(bytes) = replacements.get(&name) {
            writer.write_all(bytes)?;
        } else {
            std::io::copy(&mut entry, &mut writer)?;
        }
    }
    writer.finish()?;
    Ok(())
}

/// Read worksheet names without reading worksheet cell values.
///
/// This is intended for workflows that identify workbook/sheet-level items
/// from file metadata alone. The reader and supported file formats follow the
/// same Calamine dispatch and legacy XLS signature checks as the value APIs.
pub fn read_sheet_names(path: impl AsRef<Path>) -> Result<Vec<String>> {
    let path = path.as_ref();
    validate_legacy_xls_signature(path)?;
    let workbook = open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    ensure!(
        !names.is_empty(),
        "workbook has no worksheets: {}",
        path.display()
    );
    Ok(names)
}

/// Return whether a workbook uses the 1904 date epoch.
///
/// XLSX-family workbooks use the namespace-aware `workbookPr` metadata. BIFF8
/// XLS and XLSB use Calamine's format-specific epoch metadata. Other formats
/// return `false` under this adapter contract.
pub fn uses_1904_date_system(path: impl AsRef<Path>) -> Result<bool> {
    let path = path.as_ref();
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .map(str::to_ascii_lowercase)
        .unwrap_or_default();
    match extension.as_str() {
        "xls" | "xlsb" => {
            validate_legacy_xls_signature(path)?;
            let workbook = open_workbook_auto(path)
                .with_context(|| format!("open {} to read date epoch", path.display()))?;
            match (extension.as_str(), workbook) {
                ("xls", Sheets::Xls(reader)) => Ok(reader.has_1904_epoch()),
                ("xlsb", Sheets::Xlsb(reader)) => Ok(reader.has_1904_epoch()),
                (format, _) => anyhow::bail!(
                    "Calamine did not recognize the .{format} file as its expected workbook format"
                ),
            }
        }
        "xlsx" | "xlsm" | "xlam" => read_ooxml_1904_date_system(path),
        _ => Ok(false),
    }
}

fn read_ooxml_1904_date_system(path: &Path) -> Result<bool> {
    let mut zip = ZipArchive::new(fs::File::open(path).context("open OOXML workbook")?)
        .context("open OOXML ZIP package")?;
    let xml =
        zip_read(&mut zip, "xl/workbook.xml").context("read workbook date-system metadata")?;
    let document = Document::parse(&xml).context("parse workbook date-system metadata")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && root.tag_name().name() == "workbook",
        "xl/workbook.xml has an unexpected root element"
    );
    let Some(properties) = root.descendants().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "workbookPr"
    }) else {
        return Ok(false);
    };
    Ok(matches!(
        properties
            .attribute("date1904")
            .unwrap_or_default()
            .to_ascii_lowercase()
            .as_str(),
        "1" | "true" | "on"
    ))
}

/// Read the worksheet directory from the standard XLSX `xl/workbook.xml`
/// part using a namespace-aware OOXML parser. This metadata-only adapter does
/// not read cell values, and it deliberately does not fall back to Calamine
/// when the workbook part is missing or malformed.
pub fn read_xlsx_sheet_names_ooxml(path: impl AsRef<Path>) -> Result<Vec<String>> {
    let path = path.as_ref();
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "OOXML sheet-name adapter only supports .xlsx and .xlsm"
    );

    let mut zip = ZipArchive::new(fs::File::open(path).context("open XLSX package")?)
        .context("open XLSX ZIP package")?;
    let xml =
        zip_read(&mut zip, "xl/workbook.xml").context("read xl/workbook.xml for sheet names")?;
    let document = Document::parse(&xml).context("parse xl/workbook.xml")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && root.tag_name().name() == "workbook",
        "xl/workbook.xml has an unexpected root element"
    );

    let names = root
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "sheet"
        })
        .filter_map(|node| node.attribute("name"))
        .map(str::to_owned)
        .collect::<Vec<_>>();
    ensure!(!names.is_empty(), "workbook has no worksheets");
    Ok(names)
}

/// Read actual cell and merged-range bounds from an OOXML workbook. The
/// `<dimension>` hint is deliberately ignored because some writers leave it
/// stale or broader than the stored cell set. XLSX and XLSM are accepted for
/// this read-only metadata projection; VBA parts are not opened or changed.
pub fn read_xlsx_worksheet_extents(path: impl AsRef<Path>) -> Result<Vec<WorksheetExtent>> {
    let path = path.as_ref();
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "worksheet extent reader only supports .xlsx and .xlsm"
    );
    let mut zip = ZipArchive::new(fs::File::open(path).context("open XLSX package")?)
        .context("open XLSX ZIP package")?;
    let worksheet_parts = ooxml_worksheet_parts(&mut zip)?;
    let mut seen = std::collections::BTreeSet::new();
    let mut extents = Vec::with_capacity(worksheet_parts.len());
    for (name, part) in worksheet_parts {
        ensure!(
            seen.insert(name.clone()),
            "duplicate worksheet name {name:?}"
        );
        let xml =
            zip_read(&mut zip, &part).with_context(|| format!("read worksheet part {part}"))?;
        let document =
            Document::parse(&xml).with_context(|| format!("parse worksheet part {part}"))?;
        let root = document.root_element();
        ensure!(
            root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && root.tag_name().name() == "worksheet",
            "worksheet part {part} has an unexpected root element"
        );
        let mut extent = WorksheetExtent {
            name,
            ..WorksheetExtent::default()
        };
        for cell in root.descendants().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "c"
        }) {
            let reference = cell
                .attribute("r")
                .with_context(|| format!("cell in {part} has no A1 reference"))?;
            let (row, column) = parse_ooxml_cell_reference(reference)
                .with_context(|| format!("invalid cell reference {reference:?} in {part}"))?;
            include_extent_coordinate(&mut extent, row, column);
        }
        for merged in root.descendants().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "mergeCell"
        }) {
            let reference = merged
                .attribute("ref")
                .with_context(|| format!("merged range in {part} has no reference"))?;
            let (row1, column1, row2, column2) = parse_ooxml_range_reference(reference)
                .with_context(|| format!("invalid merged range {reference:?} in {part}"))?;
            include_extent_coordinate(&mut extent, row1, column1);
            include_extent_coordinate(&mut extent, row2, column2);
        }
        extents.push(extent);
    }
    Ok(extents)
}

/// Read number-format metadata for cells with explicit cell-level styles or
/// the default style from XLSX/XLSM.
///
/// This API intentionally fails closed when a requested cell inherits a row
/// or column style. Call [`read_xlsx_effective_cell_number_formats`] when that
/// inheritance should be resolved. Keeping this behavior allows existing
/// callers to distinguish explicit cell formatting from inherited formatting.
pub fn read_xlsx_cell_number_formats(
    path: impl AsRef<Path>,
    requests: &[CellNumberFormatRequest],
) -> Result<Vec<CellNumberFormatInfo>> {
    read_xlsx_cell_number_format_records(path, requests, false).map(|records| {
        records
            .into_iter()
            .map(|record| CellNumberFormatInfo {
                sheet_name: record.sheet_name,
                cell_reference: record.cell_reference,
                style_index: record.style_index,
                number_format_id: record.number_format_id,
                format_code: record.format_code,
            })
            .collect()
    })
}

/// Read the effective number format for a batch of worksheet cells from
/// XLSX/XLSM, including a uniquely identifiable row- or column-level style.
///
/// An explicit cell `s` style takes precedence. If no cell style exists, a
/// single applicable row or column style is used; if neither exists, the
/// workbook default style (cellXfs index zero) is used. When row and column
/// styles both apply, the source is ambiguous to this metadata adapter and the
/// request is rejected rather than guessing. Invalid style indexes are also
/// rejected. The workbook is opened once for the whole batch; cell values and
/// formulas are not read or changed.
pub fn read_xlsx_effective_cell_number_formats(
    path: impl AsRef<Path>,
    requests: &[CellNumberFormatRequest],
) -> Result<Vec<EffectiveCellNumberFormatInfo>> {
    read_xlsx_cell_number_format_records(path, requests, true)
}

fn read_xlsx_cell_number_format_records(
    path: impl AsRef<Path>,
    requests: &[CellNumberFormatRequest],
    resolve_inherited: bool,
) -> Result<Vec<EffectiveCellNumberFormatInfo>> {
    let path = path.as_ref();
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "cell number-format reader only supports .xlsx and .xlsm"
    );

    let mut parsed_requests = Vec::with_capacity(requests.len());
    let mut requested_by_sheet: BTreeMap<String, std::collections::BTreeSet<(usize, usize)>> =
        BTreeMap::new();
    for request in requests {
        ensure!(
            !request.cell_reference.contains('!'),
            "cell number-format request must use a separate sheet name and A1 reference"
        );
        let coordinate = parse_ooxml_cell_reference(&request.cell_reference)
            .with_context(|| format!("invalid cell reference {:?}", request.cell_reference))?;
        requested_by_sheet
            .entry(request.sheet_name.clone())
            .or_default()
            .insert(coordinate);
        parsed_requests.push((request, coordinate));
    }

    let mut zip = ZipArchive::new(fs::File::open(path).context("open workbook package")?)
        .context("open workbook ZIP package")?;
    let worksheet_parts = ooxml_worksheet_parts(&mut zip)?;
    let worksheet_parts: BTreeMap<_, _> = worksheet_parts.into_iter().collect();
    for sheet_name in requested_by_sheet.keys() {
        ensure!(
            worksheet_parts.contains_key(sheet_name),
            "worksheet {sheet_name:?} was not found"
        );
    }
    let (style_number_formats, custom_number_formats) = read_ooxml_number_format_styles(&mut zip)?;
    let mut styles_by_cell = BTreeMap::new();
    for (sheet_name, requested_coordinates) in requested_by_sheet {
        let part = worksheet_parts
            .get(&sheet_name)
            .expect("requested worksheet was validated above");
        let xml =
            zip_read(&mut zip, part).with_context(|| format!("read worksheet part {part}"))?;
        let document =
            Document::parse(&xml).with_context(|| format!("parse worksheet part {part}"))?;
        let root = document.root_element();
        ensure!(
            root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && root.tag_name().name() == "worksheet",
            "worksheet part {part} has an unexpected root element"
        );
        let mut cell_styles = BTreeMap::new();
        for cell in root.descendants().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "c"
        }) {
            let Some(reference) = cell.attribute("r") else {
                continue;
            };
            let coordinate = parse_ooxml_cell_reference(reference)
                .with_context(|| format!("invalid cell reference {reference:?} in {part}"))?;
            if !requested_coordinates.contains(&coordinate) {
                continue;
            }
            let Some(style_index) = cell
                .attribute("s")
                .map(str::parse::<u32>)
                .transpose()
                .with_context(|| format!("invalid style index for {reference} in {part}"))?
            else {
                continue;
            };
            ensure!(
                (style_index as usize) < style_number_formats.len(),
                "cell {reference} in {part} references missing cell style {style_index}"
            );
            cell_styles.insert(coordinate, (style_index, CellNumberFormatSource::Cell));
        }
        for &(row, column) in &requested_coordinates {
            if cell_styles.contains_key(&(row, column)) {
                continue;
            }
            let sources = inherited_style_sources(root, row, column)?;
            let cell_reference = format_ooxml_cell_reference(row, column);
            let (style_index, source) = match sources.as_slice() {
                [] => (0, CellNumberFormatSource::Default),
                [inherited] if resolve_inherited => {
                    (inherited.style_index, inherited.source)
                }
                _ if resolve_inherited => bail!(
                    "cell {cell_reference} in worksheet {sheet_name:?} has ambiguous inherited row/column styles: {}; refusing to guess the effective number format",
                    sources
                        .iter()
                        .map(|source| source.description.as_str())
                        .collect::<Vec<_>>()
                        .join(", ")
                ),
                _ => bail!(
                    "cell {cell_reference} in worksheet {sheet_name:?} has unresolved inherited style(s): {}; refusing to assume the default number format",
                    sources
                        .iter()
                        .map(|source| source.description.as_str())
                        .collect::<Vec<_>>()
                        .join(", ")
                ),
            };
            ensure!(
                (style_index as usize) < style_number_formats.len(),
                "cell {cell_reference} in worksheet {sheet_name:?} inherits missing cell style {style_index} from {source:?}"
            );
            cell_styles.insert((row, column), (style_index, source));
        }
        styles_by_cell.insert(sheet_name, cell_styles);
    }

    parsed_requests
        .into_iter()
        .map(|(request, coordinate)| {
            let (style_index, source) = styles_by_cell
                .get(&request.sheet_name)
                .and_then(|styles| styles.get(&coordinate))
                .copied()
                .unwrap_or((0, CellNumberFormatSource::Default));
            let number_format_id = style_number_formats[style_index as usize];
            let format_code = custom_number_formats
                .get(&number_format_id)
                .cloned()
                .or_else(|| builtin_number_format_code(number_format_id).map(str::to_owned));
            Ok(EffectiveCellNumberFormatInfo {
                sheet_name: request.sheet_name.clone(),
                cell_reference: format_ooxml_cell_reference(coordinate.0, coordinate.1),
                style_index,
                number_format_id,
                format_code,
                source,
            })
        })
        .collect()
}

fn inherited_style_sources(
    root: roxmltree::Node<'_, '_>,
    target_row: usize,
    target_column: usize,
) -> Result<Vec<InheritedNumberFormatStyle>> {
    let mut sources = Vec::new();

    for row in root.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "row"
    }) {
        if row
            .attribute("r")
            .and_then(|value| value.parse::<usize>().ok())
            != Some(target_row)
        {
            continue;
        }
        let custom_format = match row.attribute("customFormat") {
            Some("1" | "true") => true,
            Some("0" | "false") | None => false,
            Some(value) => {
                bail!("invalid customFormat value {value:?} for worksheet row {target_row}")
            }
        };
        if custom_format {
            if let Some(style) = row.attribute("s") {
                let style_index = style.parse::<u32>().with_context(|| {
                    format!("invalid row style index {style:?} for worksheet row {target_row}")
                })?;
                sources.push(InheritedNumberFormatStyle {
                    style_index,
                    source: CellNumberFormatSource::Row,
                    description: format!("row default style {style_index}"),
                });
            }
        }
    }

    for column in root.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "col"
    }) {
        let Some(style) = column.attribute("style") else {
            continue;
        };
        let min = column
            .attribute("min")
            .context("column style has no min column")?
            .parse::<usize>()
            .context("parse column style min column")?;
        let max = column
            .attribute("max")
            .context("column style has no max column")?
            .parse::<usize>()
            .context("parse column style max column")?;
        ensure!(
            min > 0 && max >= min,
            "invalid column style range {min}:{max}"
        );
        if !(min..=max).contains(&target_column) {
            continue;
        }
        let style_index = style
            .parse::<u32>()
            .with_context(|| format!("invalid column style index {style:?} for {min}:{max}"))?;
        sources.push(InheritedNumberFormatStyle {
            style_index,
            source: CellNumberFormatSource::Column,
            description: format!("column default style {style_index} for {min}:{max}"),
        });
    }

    Ok(sources)
}

fn read_ooxml_number_format_styles(
    zip: &mut ZipArchive<fs::File>,
) -> Result<(Vec<u32>, BTreeMap<u32, String>)> {
    let Some(styles_part) = ooxml_styles_part(zip)? else {
        return Ok((vec![0], BTreeMap::new()));
    };
    let xml = zip_read(zip, &styles_part).context("read workbook styles part")?;
    let document = Document::parse(&xml).context("parse workbook styles part")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && root.tag_name().name() == "styleSheet",
        "styles part {styles_part} has an unexpected root element"
    );

    let mut custom_number_formats = BTreeMap::new();
    for node in root.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "numFmt"
    }) {
        let id = node
            .attribute("numFmtId")
            .context("numFmt has no numFmtId")?
            .parse::<u32>()
            .context("parse numFmtId")?;
        let code = node
            .attribute("formatCode")
            .context("numFmt has no formatCode")?
            .to_owned();
        ensure!(
            custom_number_formats.insert(id, code).is_none(),
            "styles part {styles_part} contains duplicate custom numFmtId {id}"
        );
    }

    let cell_xfs = root.children().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "cellXfs"
    });
    let mut style_number_formats = cell_xfs
        .into_iter()
        .flat_map(|container| container.children())
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "xf"
        })
        .map(|node| {
            node.attribute("numFmtId")
                .unwrap_or("0")
                .parse::<u32>()
                .context("parse cell style numFmtId")
        })
        .collect::<Result<Vec<_>>>()?;
    if style_number_formats.is_empty() {
        style_number_formats.push(0);
    }
    Ok((style_number_formats, custom_number_formats))
}

fn ooxml_styles_part(zip: &mut ZipArchive<fs::File>) -> Result<Option<String>> {
    let root_rels_xml = zip_read(zip, "_rels/.rels")?;
    let root_rels_doc = Document::parse(&root_rels_xml)?;
    let workbook_target = root_rels_doc
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
        })
        .find_map(|relationship| {
            let is_workbook = relationship
                .attribute("Type")
                .is_some_and(|kind| kind.ends_with("/officeDocument"));
            let is_external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            (is_workbook && !is_external)
                .then(|| relationship.attribute("Target"))
                .flatten()
        })
        .context("package relationships do not identify an internal officeDocument part")?;
    let workbook_part = norm_part("", workbook_target);
    let workbook_rels_xml = zip_read(zip, &relationships_part(&workbook_part))?;
    let workbook_rels_doc = Document::parse(&workbook_rels_xml)?;
    let workbook_parent = workbook_part
        .rsplit_once('/')
        .map(|(parent, _)| parent)
        .unwrap_or("");
    let styles = workbook_rels_doc
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
                && node
                    .attribute("Type")
                    .is_some_and(|kind| kind.ends_with("/styles"))
        })
        .map(|relationship| {
            ensure!(
                !relationship
                    .attribute("TargetMode")
                    .is_some_and(|mode| mode.eq_ignore_ascii_case("External")),
                "workbook styles relationship is external"
            );
            let target = relationship
                .attribute("Target")
                .context("workbook styles relationship has no target")?;
            Ok(norm_part(workbook_parent, target))
        })
        .collect::<Result<Vec<_>>>()?;
    ensure!(
        styles.len() <= 1,
        "workbook has multiple styles relationships"
    );
    Ok(styles.into_iter().next())
}

fn builtin_number_format_code(id: u32) -> Option<&'static str> {
    Some(match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "\"$\"#,##0_);(\"$\"#,##0)",
        6 => "\"$\"#,##0_);[Red](\"$\"#,##0)",
        7 => "\"$\"#,##0.00_);(\"$\"#,##0.00)",
        8 => "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        41 => "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
        42 => "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",
        43 => "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
        44 => "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    })
}

fn format_ooxml_cell_reference(row: usize, column: usize) -> String {
    let mut column = column;
    let mut letters = Vec::new();
    while column > 0 {
        column -= 1;
        letters.push((b'A' + (column % 26) as u8) as char);
        column /= 26;
    }
    letters.iter().rev().copied().collect::<String>() + &row.to_string()
}

fn include_extent_coordinate(extent: &mut WorksheetExtent, row: usize, column: usize) {
    extent.min_row = if extent.min_row == 0 {
        row
    } else {
        extent.min_row.min(row)
    };
    extent.min_column = if extent.min_column == 0 {
        column
    } else {
        extent.min_column.min(column)
    };
    extent.max_row = extent.max_row.max(row);
    extent.max_column = extent.max_column.max(column);
}

fn parse_ooxml_cell_reference(reference: &str) -> Result<(usize, usize)> {
    let reference = reference
        .rsplit_once('!')
        .map(|(_, cell)| cell)
        .unwrap_or(reference)
        .trim()
        .replace('$', "");
    let split = reference
        .find(|character: char| character.is_ascii_digit())
        .context("A1 reference has no row number")?;
    let (letters, digits) = reference.split_at(split);
    ensure!(
        !letters.is_empty() && letters.bytes().all(|byte| byte.is_ascii_alphabetic()),
        "A1 reference has an invalid column"
    );
    ensure!(
        !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()),
        "A1 reference has an invalid row"
    );
    let mut column = 0usize;
    for letter in letters.bytes() {
        column = column
            .checked_mul(26)
            .and_then(|value| {
                value.checked_add(usize::from(letter.to_ascii_uppercase() - b'A' + 1))
            })
            .context("A1 column index overflows")?;
    }
    let row = digits.parse::<usize>().context("parse A1 row number")?;
    ensure!(row > 0 && column > 0, "A1 coordinates are 1-based");
    Ok((row, column))
}

fn parse_ooxml_range_reference(reference: &str) -> Result<(usize, usize, usize, usize)> {
    let (first, last) = reference.split_once(':').unwrap_or((reference, reference));
    let (row1, column1) = parse_ooxml_cell_reference(first)?;
    let (row2, column2) = parse_ooxml_cell_reference(last)?;
    Ok((row1, column1, row2, column2))
}

/// Read stored values from the first `max_columns` columns of each worksheet.
///
/// This is a value-only API: it does not read formula text or OOXML metadata.
/// Formula cells are represented by the value available from Calamine's
/// worksheet range (normally the cached result); callers that require formula
/// text must use `read_workbook` and inspect `SheetDump::formulas` separately.
pub fn read_value_columns(path: impl AsRef<Path>, max_columns: usize) -> Result<Vec<SheetValues>> {
    read_value_columns_impl(path.as_ref(), max_columns, false)
}

/// Read stored values while representing Calamine date-formatted cells as ISO
/// datetime strings. This preserves the workbook's 1900/1904 date epoch for
/// callers that need calendar dates rather than raw serial values.
pub fn read_value_columns_with_iso_dates(
    path: impl AsRef<Path>,
    max_columns: usize,
) -> Result<Vec<SheetValues>> {
    read_value_columns_impl(path.as_ref(), max_columns, true)
}

fn read_value_columns_impl(
    path: &Path,
    max_columns: usize,
    iso_dates: bool,
) -> Result<Vec<SheetValues>> {
    anyhow::ensure!(max_columns > 0, "max_columns must be greater than zero");
    validate_legacy_xls_signature(path)?;
    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    let mut sheets = Vec::with_capacity(names.len());

    for name in names {
        let range = workbook
            .worksheet_range(&name)
            .with_context(|| format!("read values for {name}"))?;
        let (start_row, start_col) = range.start().unwrap_or((0, 0));
        let (max_row, max_column) = range
            .end()
            .map(|(row, column)| (row as usize + 1, column as usize + 1))
            .unwrap_or_default();
        let mut values = BTreeMap::new();
        for (local_row, local_col, value) in range.cells() {
            let absolute_col = start_col as usize + local_col;
            if absolute_col >= max_columns {
                continue;
            }
            if !matches!(value, Data::Empty) {
                let cell_value = match value {
                    Data::DateTime(datetime) if iso_dates && datetime.is_datetime() => {
                        let (year, month, day, hour, minute, second, millisecond) =
                            datetime.to_ymd_hms_milli();
                        CellValue::DateTimeIso(format!(
                            "{year:04}-{month:02}-{day:02}T{hour:02}:{minute:02}:{second:02}.{millisecond:03}"
                        ))
                    }
                    other => CellValue::from(other),
                };
                values.insert(
                    coord(start_row as usize + local_row, absolute_col),
                    cell_value,
                );
            }
        }
        sheets.push(SheetValues {
            name,
            values,
            max_row,
            max_column,
        });
    }
    Ok(sheets)
}

fn validate_legacy_xls_signature(path: &Path) -> Result<()> {
    let is_xls = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"));
    if !is_xls {
        return Ok(());
    }

    const CFB_OLE_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let mut file = fs::File::open(path).with_context(|| format!("open {}", path.display()))?;
    let mut signature = [0_u8; 8];
    if file.read_exact(&mut signature).is_err() || signature != CFB_OLE_SIGNATURE {
        anyhow::bail!(
            "格式不支持：.xls 文件的 CFB/OLE 签名无效（{}）",
            path.display()
        );
    }
    Ok(())
}

fn zip_read(zip: &mut ZipArchive<fs::File>, name: &str) -> Result<String> {
    let mut file = zip
        .by_name(name)
        .with_context(|| format!("missing ZIP part {name}"))?;
    let mut text = String::new();
    file.read_to_string(&mut text)?;
    Ok(text)
}

fn norm_part(base_dir: &str, target: &str) -> String {
    let combined = if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else if base_dir.is_empty() {
        target.to_owned()
    } else {
        format!("{base_dir}/{target}")
    };
    let mut parts: Vec<&str> = Vec::new();
    for p in combined.split('/') {
        match p {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(p),
        }
    }
    parts.join("/")
}

fn relationships_part(part: &str) -> String {
    let (parent, file_name) = part.rsplit_once('/').unwrap_or(("", part));
    if parent.is_empty() {
        format!("_rels/{file_name}.rels")
    } else {
        format!("{parent}/_rels/{file_name}.rels")
    }
}

fn ooxml_worksheet_parts(zip: &mut ZipArchive<fs::File>) -> Result<Vec<(String, String)>> {
    let root_rels_xml = zip_read(zip, "_rels/.rels")?;
    let root_rels_doc = Document::parse(&root_rels_xml)?;
    let workbook_target = root_rels_doc
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
        })
        .find_map(|relationship| {
            let is_workbook = relationship
                .attribute("Type")
                .is_some_and(|kind| kind.ends_with("/officeDocument"));
            let is_external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            if is_workbook && !is_external {
                relationship.attribute("Target")
            } else {
                None
            }
        })
        .context("package relationships do not identify an internal officeDocument part")?;
    let workbook_part = norm_part("", workbook_target);
    let workbook_xml = zip_read(zip, &workbook_part)?;
    let workbook_rels_xml = zip_read(zip, &relationships_part(&workbook_part))?;
    let workbook_doc = Document::parse(&workbook_xml)?;
    let rels_doc = Document::parse(&workbook_rels_xml)?;
    let workbook_parent = workbook_part
        .rsplit_once('/')
        .map(|(parent, _)| parent)
        .unwrap_or("");

    let mut rel_targets = BTreeMap::new();
    for relationship in rels_doc.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
            && node.tag_name().name() == "Relationship"
    }) {
        if let (Some(id), Some(kind), Some(target)) = (
            relationship.attribute("Id"),
            relationship.attribute("Type"),
            relationship.attribute("Target"),
        ) {
            let is_external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            rel_targets.insert(
                id.to_owned(),
                (kind.to_owned(), target.to_owned(), is_external),
            );
        }
    }

    let mut worksheets = Vec::new();
    let mut names = std::collections::BTreeSet::new();
    let mut parts = std::collections::BTreeSet::new();
    for sheet in workbook_doc.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "sheet"
    }) {
        let name = sheet
            .attribute("name")
            .context("workbook sheet has no name")?;
        let id = sheet
            .attribute((OFFICE_REL_NS, "id"))
            .with_context(|| format!("workbook sheet {name:?} has no relationship id"))?;
        let (kind, target, is_external) = rel_targets
            .get(id)
            .with_context(|| format!("workbook sheet {name:?} has an unresolved relationship"))?;
        if !kind.ends_with("/worksheet") {
            continue;
        }
        ensure!(
            !*is_external,
            "workbook worksheet {name:?} has an external relationship"
        );
        let part = norm_part(workbook_parent, target);
        ensure!(
            !part.is_empty(),
            "workbook worksheet {name:?} has an empty part path"
        );
        ensure!(
            names.insert(name.to_owned()),
            "duplicate worksheet name {name:?}"
        );
        ensure!(
            parts.insert(part.clone()),
            "multiple worksheet names resolve to part {part}"
        );
        worksheets.push((name.to_owned(), part));
    }
    ensure!(
        !worksheets.is_empty(),
        "workbook contains no resolvable worksheet relationships"
    );
    Ok(worksheets)
}

fn xml_element_value(node: roxmltree::Node<'_, '_>) -> Value {
    let mut attributes = BTreeMap::new();
    for attribute in node.attributes() {
        let key = match attribute.namespace() {
            Some(namespace) => format!("{{{namespace}}}{}", attribute.name()),
            None => attribute.name().to_owned(),
        };
        attributes.insert(key, attribute.value().to_owned());
    }

    let content = node
        .children()
        .filter_map(|child| {
            if child.is_element() {
                Some(xml_element_value(child))
            } else {
                child
                    .text()
                    .filter(|text| !text.trim().is_empty())
                    .map(|text| json!({"text": text}))
            }
        })
        .collect::<Vec<_>>();

    json!({
        "namespace": node.tag_name().namespace().unwrap_or(""),
        "name": node.tag_name().name(),
        "attributes": attributes,
        "content": content,
    })
}

// Comments, conditional formatting, and data validations are not exposed by Calamine. This
// namespace-aware supplement follows package relationships and records normalized XML structure.
fn extract_ooxml_metadata(path: &Path) -> Result<BTreeMap<String, OoxmlSheetMetadata>> {
    let file = fs::File::open(path)?;
    let mut zip = ZipArchive::new(file)?;
    let worksheets = ooxml_worksheet_parts(&mut zip)?;
    let mut result = BTreeMap::new();
    for (sheet_name, sheet_part) in worksheets {
        let sheet_xml = zip_read(&mut zip, &sheet_part)?;
        let sheet_doc = Document::parse(&sheet_xml)
            .with_context(|| format!("parse worksheet part {sheet_part}"))?;
        let conditional_formatting = sheet_doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "conditionalFormatting")
            .map(xml_element_value)
            .collect();
        let data_validations = sheet_doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "dataValidations")
            .map(xml_element_value)
            .collect();

        let parent = sheet_part
            .rsplit_once('/')
            .map(|(parent, _)| parent)
            .unwrap_or("");
        let rels_part = relationships_part(&sheet_part);
        let mut comment_parts = Vec::new();
        if let Ok(sheet_rels_xml) = zip_read(&mut zip, &rels_part) {
            let sheet_rels_doc = Document::parse(&sheet_rels_xml)?;
            for rel in sheet_rels_doc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
            {
                let ty = rel.attribute("Type").unwrap_or_default();
                if ty.ends_with("/comments") {
                    if let Some(target) = rel.attribute("Target") {
                        comment_parts.push(norm_part(parent, target));
                    }
                }
            }
        }
        let mut comments = BTreeMap::new();
        let mut comment_records = BTreeMap::new();
        for part in comment_parts {
            let xml = zip_read(&mut zip, &part)?;
            let doc = Document::parse(&xml)?;
            let authors = doc
                .descendants()
                .filter(|node| {
                    node.is_element()
                        && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                        && node.tag_name().name() == "author"
                        && node.parent().is_some_and(|parent| {
                            parent.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                                && parent.tag_name().name() == "authors"
                        })
                })
                .map(|author| author.text().unwrap_or_default().to_owned())
                .collect::<Vec<_>>();
            for comment in doc.descendants().filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && node.tag_name().name() == "comment"
            }) {
                let Some(reference) = comment.attribute("ref") else {
                    continue;
                };
                let mut text = String::new();
                for t in comment.descendants().filter(|node| {
                    node.is_element()
                        && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                        && node.tag_name().name() == "t"
                }) {
                    if let Some(s) = t.text() {
                        text.push_str(s);
                    }
                }
                let author = comment
                    .attribute("authorId")
                    .and_then(|index| index.parse::<usize>().ok())
                    .and_then(|index| authors.get(index).cloned());
                comments.insert(reference.to_owned(), text.clone());
                comment_records.insert(reference.to_owned(), WorkbookComment { text, author });
            }
        }
        result.insert(
            sheet_name.to_owned(),
            OoxmlSheetMetadata {
                comments,
                comment_records,
                conditional_formatting,
                data_validations,
            },
        );
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use std::path::PathBuf;
    use zip::{write::SimpleFileOptions, ZipWriter};

    fn cell_number_format_fixture(extension: &str) -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-number-formats-{}-{}.{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos(),
            extension
        ));
        let package_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="Data" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>"#
        );
        let workbook_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/data.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="custom/styles.xml"/></Relationships>"#
        );
        let styles = format!(
            r#"<s:styleSheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:numFmts count="2"><s:numFmt numFmtId="164" formatCode="h:mm:ss"/><s:numFmt numFmtId="165" formatCode="000000"/></s:numFmts><s:cellXfs count="4"><s:xf numFmtId="0"/><s:xf numFmtId="14"/><s:xf numFmtId="164"/><s:xf numFmtId="165"/></s:cellXfs></s:styleSheet>"#
        );
        let worksheet = format!(
            r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData><s:row r="1"><s:c r="A1" s="1"><s:v>45000</s:v></s:c><s:c r="B1" s="2"><s:v>0.5</s:v></s:c><s:c r="C1"><s:v>42</s:v></s:c><s:c r="D1" s="3"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
        );
        let entries = [
            ("_rels/.rels", package_relationships),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_relationships),
            ("xl/custom/styles.xml", styles),
            ("xl/worksheets/data.xml", worksheet),
        ];
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, data) in entries {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(data.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    fn inherited_number_format_fixture(extension: &str, worksheet: &str) -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-inherited-number-format-{}-{}.{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos(),
            extension
        ));
        let package_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="Data" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>"#
        );
        let workbook_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/data.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#
        );
        let styles = format!(
            r#"<s:styleSheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:numFmts count="1"><s:numFmt numFmtId="165" formatCode="000000"/></s:numFmts><s:cellXfs count="2"><s:xf numFmtId="0"/><s:xf numFmtId="165"/></s:cellXfs></s:styleSheet>"#
        );
        let entries = [
            ("_rels/.rels", package_relationships),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_relationships),
            ("xl/styles.xml", styles),
            ("xl/worksheets/data.xml", worksheet.to_owned()),
        ];
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, data) in entries {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(data.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    #[test]
    fn cell_number_formats_read_custom_builtin_and_default_styles_for_xlsx_and_xlsm() -> Result<()>
    {
        for extension in ["xlsx", "xlsm"] {
            let fixture = cell_number_format_fixture(extension)?;
            let formats = read_xlsx_cell_number_formats(
                &fixture,
                &[
                    CellNumberFormatRequest {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "B1".to_owned(),
                    },
                    CellNumberFormatRequest {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "$A$1".to_owned(),
                    },
                    CellNumberFormatRequest {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "C1".to_owned(),
                    },
                    CellNumberFormatRequest {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "D1".to_owned(),
                    },
                    CellNumberFormatRequest {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "E1".to_owned(),
                    },
                ],
            )?;
            fs::remove_file(fixture)?;

            assert_eq!(
                formats,
                vec![
                    CellNumberFormatInfo {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "B1".to_owned(),
                        style_index: 2,
                        number_format_id: 164,
                        format_code: Some("h:mm:ss".to_owned()),
                    },
                    CellNumberFormatInfo {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "A1".to_owned(),
                        style_index: 1,
                        number_format_id: 14,
                        format_code: Some("mm-dd-yy".to_owned()),
                    },
                    CellNumberFormatInfo {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "C1".to_owned(),
                        style_index: 0,
                        number_format_id: 0,
                        format_code: Some("General".to_owned()),
                    },
                    CellNumberFormatInfo {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "D1".to_owned(),
                        style_index: 3,
                        number_format_id: 165,
                        format_code: Some("000000".to_owned()),
                    },
                    CellNumberFormatInfo {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "E1".to_owned(),
                        style_index: 0,
                        number_format_id: 0,
                        format_code: Some("General".to_owned()),
                    },
                ]
            );
        }
        Ok(())
    }

    #[test]
    fn cell_number_format_reader_fails_closed_for_inherited_row_and_column_styles() -> Result<()> {
        let request = [CellNumberFormatRequest {
            sheet_name: "Data".to_owned(),
            cell_reference: "A2".to_owned(),
        }];
        for extension in ["xlsx", "xlsm"] {
            for (source, worksheet) in [
                (
                    "row default style 1",
                    format!(
                        r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData><s:row r="2" s="1" customFormat="1"><s:c r="A2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                    ),
                ),
                (
                    "column default style 1",
                    format!(
                        r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:cols><s:col min="1" max="1" style="1"/></s:cols><s:sheetData><s:row r="2"><s:c r="A2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                    ),
                ),
            ] {
                let fixture = inherited_number_format_fixture(extension, &worksheet)?;
                let error = read_xlsx_cell_number_formats(&fixture, &request)
                    .expect_err("inherited styles must not be reported as General");
                fs::remove_file(fixture)?;
                assert!(
                    error.to_string().contains("unresolved inherited style"),
                    "unexpected reader error: {error:#}"
                );
                assert!(
                    error.to_string().contains(source),
                    "reader error: {error:#}"
                );
            }

            let row_style_disabled = inherited_number_format_fixture(
                extension,
                &format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData><s:row r="2" s="1" customFormat="0"><s:c r="A2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
            )?;
            let info = read_xlsx_cell_number_formats(&row_style_disabled, &request)?;
            fs::remove_file(row_style_disabled)?;
            assert_eq!(info[0].format_code.as_deref(), Some("General"));

            let explicit_cell_override = inherited_number_format_fixture(
                extension,
                &format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:cols><s:col min="1" max="1" style="1"/></s:cols><s:sheetData><s:row r="2" s="1" customFormat="1"><s:c r="A2" s="0"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
            )?;
            let info = read_xlsx_cell_number_formats(&explicit_cell_override, &request)?;
            fs::remove_file(explicit_cell_override)?;
            assert_eq!(info[0].format_code.as_deref(), Some("General"));
        }
        Ok(())
    }

    #[test]
    fn effective_cell_number_formats_resolve_cell_row_column_and_default_sources() -> Result<()> {
        let request = [CellNumberFormatRequest {
            sheet_name: "Data".to_owned(),
            cell_reference: "A2".to_owned(),
        }];
        let cases = [
            (
                "row style",
                format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData><s:row r="2" s="1" customFormat="1"><s:c r="A2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
                CellNumberFormatSource::Row,
                1,
                165,
                Some("000000"),
            ),
            (
                "column style",
                format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:cols><s:col min="1" max="1" style="1"/></s:cols><s:sheetData><s:row r="2"><s:c r="A2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
                CellNumberFormatSource::Column,
                1,
                165,
                Some("000000"),
            ),
            (
                "default style",
                format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData><s:row r="2"><s:c r="A2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
                CellNumberFormatSource::Default,
                0,
                0,
                Some("General"),
            ),
            (
                "disabled row style",
                format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData><s:row r="2" s="1" customFormat="0"><s:c r="A2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
                CellNumberFormatSource::Default,
                0,
                0,
                Some("General"),
            ),
            (
                "explicit cell override",
                format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:cols><s:col min="1" max="1" style="1"/></s:cols><s:sheetData><s:row r="2" s="1" customFormat="1"><s:c r="A2" s="0"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
                CellNumberFormatSource::Cell,
                0,
                0,
                Some("General"),
            ),
        ];

        for extension in ["xlsx", "xlsm"] {
            for (name, worksheet, source, style_index, number_format_id, format_code) in &cases {
                let fixture = inherited_number_format_fixture(extension, worksheet)?;
                let formats = read_xlsx_effective_cell_number_formats(&fixture, &request)
                    .with_context(|| format!("resolve {name} for {extension}"))?;
                fs::remove_file(fixture)?;
                assert_eq!(
                    formats[0],
                    EffectiveCellNumberFormatInfo {
                        sheet_name: "Data".to_owned(),
                        cell_reference: "A2".to_owned(),
                        style_index: *style_index,
                        number_format_id: *number_format_id,
                        format_code: format_code.map(str::to_owned),
                        source: *source,
                    },
                    "unexpected {name} result for {extension}"
                );
            }
        }
        Ok(())
    }

    #[test]
    fn effective_cell_number_formats_reject_ambiguous_and_invalid_styles() -> Result<()> {
        let request = [CellNumberFormatRequest {
            sheet_name: "Data".to_owned(),
            cell_reference: "A2".to_owned(),
        }];
        let cases = [
            (
                "ambiguous row and column styles",
                format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:cols><s:col min="1" max="1" style="0"/></s:cols><s:sheetData><s:row r="2" s="1" customFormat="1"><s:c r="A2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
                "ambiguous inherited row/column styles",
            ),
            (
                "row style index out of range",
                format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData><s:row r="2" s="2" customFormat="1"><s:c r="A2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
                "inherits missing cell style 2",
            ),
            (
                "column style index out of range",
                format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:cols><s:col min="1" max="1" style="2"/></s:cols><s:sheetData><s:row r="2"><s:c r="A2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
                "inherits missing cell style 2",
            ),
            (
                "cell style index out of range",
                format!(
                    r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData><s:row r="2"><s:c r="A2" s="2"><s:v>123</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
                ),
                "references missing cell style 2",
            ),
        ];

        for (name, worksheet, expected_error) in cases {
            let fixture = inherited_number_format_fixture("xlsx", &worksheet)?;
            let error =
                read_xlsx_effective_cell_number_formats(&fixture, &request).expect_err(name);
            fs::remove_file(fixture)?;
            assert!(
                error.to_string().contains(expected_error),
                "{name}: unexpected error: {error:#}"
            );
        }
        Ok(())
    }

    #[test]
    fn cell_number_format_reader_rejects_unknown_sheets_and_non_ooxml_inputs() -> Result<()> {
        let fixture = cell_number_format_fixture("xlsx")?;
        let error = read_xlsx_cell_number_formats(
            &fixture,
            &[CellNumberFormatRequest {
                sheet_name: "Missing".to_owned(),
                cell_reference: "A1".to_owned(),
            }],
        )
        .expect_err("unknown worksheets must be rejected");
        assert!(error
            .to_string()
            .contains("worksheet \"Missing\" was not found"));
        fs::remove_file(&fixture)?;

        let non_ooxml = std::env::temp_dir().join(format!(
            "excel-reader-number-formats-{}.xls",
            std::process::id()
        ));
        fs::write(&non_ooxml, b"not an OOXML package")?;
        let error = read_xlsx_cell_number_formats(&non_ooxml, &[])
            .expect_err("BIFF input must be rejected by the OOXML metadata adapter");
        fs::remove_file(non_ooxml)?;
        assert!(error
            .to_string()
            .contains("cell number-format reader only supports .xlsx and .xlsm"));
        Ok(())
    }

    fn nonstandard_header_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-header-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let package_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="workbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/books/main.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="外汇报表" sheetId="1" r:id="worksheet"/><s:sheet name="图表页" sheetId="2" r:id="chartsheet"/></s:sheets></s:workbook>"#
        );
        let workbook_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="worksheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/fx.xml"/><Relationship Id="chartsheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="../charts/chart1.xml"/></Relationships>"#
        );
        let worksheet = format!(
            r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData/><s:headerFooter differentOddEven="1" differentFirst="1"><s:oddHeader>&amp;L单位：万元</s:oddHeader><s:evenHeader>&amp;C第二页</s:evenHeader><s:firstHeader/></s:headerFooter></s:worksheet>"#
        );
        let entries = [
            ("_rels/.rels", package_relationships),
            ("custom/books/main.xml", workbook),
            ("custom/books/_rels/main.xml.rels", workbook_relationships),
            ("custom/worksheets/fx.xml", worksheet),
            ("custom/charts/chart1.xml", "<chart/>".to_owned()),
        ];
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, data) in entries {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(data.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    fn comment_records_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-comment-records-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let package_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="Data" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>"#
        );
        let workbook_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let worksheet =
            format!(r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData/></s:worksheet>"#);
        let worksheet_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comments1.xml"/></Relationships>"#
        );
        let comments = format!(
            r#"<x:comments xmlns:x="{SPREADSHEET_MAIN_NS}"><x:authors><x:author>Data Team</x:author><x:author></x:author></x:authors><x:commentList><x:comment ref="A1" authorId="0"><x:text><x:r><x:t>First </x:t></x:r><x:r><x:t>note</x:t></x:r></x:text></x:comment><x:comment ref="B2" authorId="1"><x:text><x:t>Empty author</x:t></x:text></x:comment><x:comment ref="C3" authorId="9"><x:text><x:t>Unknown author</x:t></x:text></x:comment></x:commentList></x:comments>"#
        );
        let entries = [
            ("_rels/.rels", package_relationships),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_relationships),
            ("xl/worksheets/sheet1.xml", worksheet),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                worksheet_relationships,
            ),
            ("xl/comments/comments1.xml", comments),
        ];
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, data) in entries {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(data.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    fn sheet_name_ooxml_fixture(extension: &str) -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-sheet-names-{}-{}.{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos(),
            extension
        ));
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:x="urn:unrelated"><s:sheets><s:sheet name="单位贷款" sheetId="1"/><s:sheet name="图表页" sheetId="2"/><x:sheet name="must-be-ignored" sheetId="3"/></s:sheets></s:workbook>"#
        );
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        archive.start_file("xl/workbook.xml", SimpleFileOptions::default())?;
        archive.write_all(workbook.as_bytes())?;
        archive.finish()?;
        Ok(path)
    }

    fn date_epoch_ooxml_fixture(extension: &str, workbook_properties: &str) -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-date-epoch-{}-{}.{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos(),
            extension
        ));
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}">{workbook_properties}<s:sheets/></s:workbook>"#
        );
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        archive.start_file("xl/workbook.xml", SimpleFileOptions::default())?;
        archive.write_all(workbook.as_bytes())?;
        archive.finish()?;
        Ok(path)
    }

    fn defined_names_ooxml_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-defined-names-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let root_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/book.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="Data Set" sheetId="1" r:id="rId1"/><s:sheet name="Chart" sheetId="2" r:id="rId2"/><s:sheet name="存量单位贷款信息" sheetId="3" r:id="rId3"/></s:sheets><s:definedNames><s:definedName name="GlobalRange">'Data Set'!$A$1:$B$8</s:definedName><s:definedName name="LocalUnion" localSheetId="0">'Data Set'!A1:B2,'Data Set'!D4</s:definedName><s:definedName name="Formula">SUM('Data Set'!A:A)</s:definedName><s:definedName name="UnicodeRange">存量单位贷款信息!$A$4:$B$90</s:definedName><s:definedName name="UnicodeLocal" localSheetId="2">存量单位贷款信息!A1:B2</s:definedName><s:definedName name="Broken" localSheetId="9">'Missing'!A1</s:definedName></s:definedNames></s:workbook>"#
        );
        let workbook_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/s1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/c1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/s3.xml"/></Relationships>"#
        );
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, xml) in [
            ("_rels/.rels", root_rels),
            ("custom/book.xml", workbook),
            ("custom/_rels/book.xml.rels", workbook_rels),
        ] {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(xml.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    fn worksheet_extent_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-worksheet-extents-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let package_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="Layout" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>"#
        );
        let workbook_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let worksheet = format!(
            r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:dimension ref="A1:XFD1048576"/><s:sheetData><s:row r="3"><s:c r="B3"><s:v>4</s:v></s:c></s:row><s:row r="8"><s:c r="G8" s="1"/></s:row></s:sheetData><s:mergeCells count="1"><s:mergeCell ref="C10:D12"/></s:mergeCells></s:worksheet>"#
        );
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, xml) in [
            ("_rels/.rels", package_rels),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", worksheet),
        ] {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(xml.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    #[test]
    fn cell_value_json_shape_remains_compatible_and_typed() {
        let cases = [
            (
                Data::String("001.20".to_owned()),
                json!({"type":"string", "value":"001.20"}),
            ),
            (Data::Float(1.0), json!({"type":"number", "value":1.0})),
            (Data::Int(42), json!({"type":"integer", "value":42})),
            (Data::Bool(true), json!({"type":"boolean", "value":true})),
        ];

        for (data, expected) in cases {
            assert_eq!(value_json(&data), Some(expected));
        }
        assert_eq!(value_json(&Data::Empty), None);
        assert_eq!(
            value_json(&Data::Float(f64::NAN)),
            Some(json!({"type":"number", "value":null}))
        );

        let decoded = CellValue::from_probe_json(&json!({
            "type": "number",
            "value": 12.5
        }))
        .unwrap();
        assert_eq!(decoded, CellValue::Number(12.5));
        let non_finite = CellValue::from_probe_json(&json!({
            "type": "number",
            "value": null
        }))
        .unwrap();
        assert!(matches!(non_finite, CellValue::Number(value) if value.is_nan()));
    }

    #[test]
    fn ooxml_sheet_name_adapter_matches_namespace_and_workbook_order() -> Result<()> {
        let fixture = sheet_name_ooxml_fixture("xlsx")?;
        let names = read_xlsx_sheet_names_ooxml(&fixture)?;
        fs::remove_file(fixture)?;
        assert_eq!(names, vec!["单位贷款".to_owned(), "图表页".to_owned()]);
        Ok(())
    }

    #[test]
    fn ooxml_sheet_name_adapter_accepts_xlsm_metadata() -> Result<()> {
        let fixture = sheet_name_ooxml_fixture("xlsm")?;
        let names = read_xlsx_sheet_names_ooxml(&fixture)?;
        fs::remove_file(fixture)?;
        assert_eq!(names, vec!["单位贷款".to_owned(), "图表页".to_owned()]);
        Ok(())
    }

    #[test]
    fn date_epoch_reader_detects_ooxml_1904_setting_for_xlsx_and_xlsm() -> Result<()> {
        for extension in ["xlsx", "xlsm"] {
            let fixture =
                date_epoch_ooxml_fixture(extension, &format!(r#"<s:workbookPr date1904="1"/>"#))?;
            assert!(uses_1904_date_system(&fixture)?, "{extension} epoch");
            fs::remove_file(fixture)?;
        }

        let fixture =
            date_epoch_ooxml_fixture("xlsx", &format!(r#"<s:workbookPr date1904="false"/>"#))?;
        assert!(!uses_1904_date_system(&fixture)?);
        fs::remove_file(fixture)?;
        Ok(())
    }

    #[test]
    fn date_epoch_reader_rejects_invalid_legacy_xls_signature() -> Result<()> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-invalid-epoch-{}-{}.xls",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        fs::write(&path, b"<html>not an xls</html>")?;
        let error = uses_1904_date_system(&path).expect_err("invalid .xls must be rejected");
        assert!(error.to_string().contains("CFB/OLE 签名无效"));
        fs::remove_file(path)?;
        Ok(())
    }

    #[test]
    fn defined_name_reader_preserves_scope_raw_text_and_resolves_area_sheet_identity() -> Result<()>
    {
        let fixture = defined_names_ooxml_fixture()?;
        let names = read_defined_names(&fixture)?;
        fs::remove_file(fixture)?;

        assert_eq!(names.len(), 6);
        assert_eq!(names[0].name, "GlobalRange");
        assert_eq!(names[0].scope, DefinedNameScope::Workbook);
        assert_eq!(names[0].raw_reference, "'Data Set'!$A$1:$B$8");
        assert_eq!(names[0].parse_status, DefinedNameParseStatus::Areas);
        assert_eq!(
            names[0].areas,
            vec![DefinedNameArea {
                sheet_name: "Data Set".to_owned(),
                sheet_kind: "worksheet".to_owned(),
                reference: "$A$1:$B$8".to_owned(),
            }]
        );
        assert_eq!(
            names[1].scope,
            DefinedNameScope::Worksheet {
                local_sheet_id: 0,
                sheet_name: Some("Data Set".to_owned()),
                sheet_kind: Some("worksheet".to_owned()),
            }
        );
        assert_eq!(names[1].parse_status, DefinedNameParseStatus::Areas);
        assert_eq!(names[1].areas.len(), 2);
        assert_eq!(
            names[2].parse_status,
            DefinedNameParseStatus::NonAreaExpression
        );
        assert_eq!(names[2].raw_reference, "SUM('Data Set'!A:A)");
        assert_eq!(names[3].name, "UnicodeRange");
        assert_eq!(names[3].parse_status, DefinedNameParseStatus::Areas);
        assert_eq!(
            names[3].areas,
            vec![DefinedNameArea {
                sheet_name: "存量单位贷款信息".to_owned(),
                sheet_kind: "worksheet".to_owned(),
                reference: "$A$4:$B$90".to_owned(),
            }]
        );
        assert_eq!(
            names[4].scope,
            DefinedNameScope::Worksheet {
                local_sheet_id: 2,
                sheet_name: Some("存量单位贷款信息".to_owned()),
                sheet_kind: Some("worksheet".to_owned()),
            }
        );
        assert_eq!(
            names[5].scope,
            DefinedNameScope::UnresolvedWorksheet { local_sheet_id: 9 }
        );
        assert_eq!(
            names[5].parse_status,
            DefinedNameParseStatus::InvalidAreaExpression
        );
        assert!(decode_sheet_token("存量单位贷款信息").is_some());
        assert!(decode_sheet_token("[Other.xlsx]Data").is_none());
        assert!(decode_sheet_token("Data:Sheet").is_none());
        Ok(())
    }

    #[test]
    fn worksheet_extents_include_styled_blanks_and_merge_coordinates_not_dimension_hint(
    ) -> Result<()> {
        let fixture = worksheet_extent_fixture()?;
        let extents = read_xlsx_worksheet_extents(&fixture)?;
        fs::remove_file(fixture)?;
        assert_eq!(
            extents,
            vec![WorksheetExtent {
                name: "Layout".to_owned(),
                min_row: 3,
                min_column: 2,
                max_row: 12,
                max_column: 7,
            }]
        );
        Ok(())
    }

    #[test]
    fn ooxml_sheet_name_adapter_rejects_biff_formats_without_fallback() -> Result<()> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-sheet-names-{}.xls",
            std::process::id()
        ));
        fs::write(&path, b"not an OOXML package")?;
        let error = read_xlsx_sheet_names_ooxml(&path).expect_err("reject unsupported format");
        fs::remove_file(path)?;
        assert!(error.to_string().contains("only supports .xlsx and .xlsm"));
        Ok(())
    }

    #[test]
    fn value_column_projection_matches_dump_and_keeps_absolute_coordinates() {
        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../audit-core/tests/fixtures/synthetic_probe.xlsx");
        let dump = read_workbook(&fixture).expect("read full probe fixture");
        let projected = read_value_columns(&fixture, 3).expect("read first three columns");

        assert_eq!(projected.len(), dump.sheets.len());
        for (projection, sheet) in projected.iter().zip(&dump.sheets) {
            assert_eq!(projection.name, sheet.name);
            let expected: BTreeMap<_, _> = sheet
                .values
                .iter()
                .filter(|(address, _)| {
                    address
                        .chars()
                        .take_while(|character| character.is_ascii_alphabetic())
                        .fold(0_usize, |column, character| {
                            column * 26
                                + (character.to_ascii_uppercase() as usize - 'A' as usize + 1)
                        })
                        <= 3
                })
                .map(|(address, value)| {
                    (
                        address.clone(),
                        CellValue::from_probe_json(value).expect("decode dumped cell value"),
                    )
                })
                .collect();
            assert_eq!(projection.values, expected);
        }
    }

    #[test]
    fn invalid_legacy_xls_signature_is_rejected_as_unsupported_format() {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-invalid-xls-{}.xls",
            std::process::id()
        ));
        fs::write(&path, b"<html>not an xls</html>").expect("write temporary pseudo XLS");
        let error = read_value_columns(&path, 3).expect_err("reject invalid CFB/OLE signature");
        let _ = fs::remove_file(&path);
        assert!(error.to_string().contains("格式不支持"));
        assert!(error.to_string().contains("CFB/OLE"));
    }

    #[test]
    fn normalized_ooxml_keeps_namespaces_attributes_and_text() {
        let xml = r#"<x:conditionalFormatting xmlns:x="urn:example:sheet" sqref="A1">
            <x:cfRule type="expression"><x:formula>A1&gt;0</x:formula></x:cfRule>
        </x:conditionalFormatting>"#;
        let document = Document::parse(xml).unwrap();
        let value = xml_element_value(document.root_element());

        assert_eq!(value["namespace"], "urn:example:sheet");
        assert_eq!(value["attributes"]["sqref"], "A1");
        assert_eq!(value["content"][0]["name"], "cfRule");
        assert_eq!(
            value["content"][0]["content"][0]["content"][0]["text"],
            "A1>0"
        );
    }

    #[test]
    fn ooxml_comment_reader_returns_text_and_resolved_authors() -> Result<()> {
        let fixture = comment_records_fixture()?;
        let comments = read_ooxml_comments(&fixture)?;
        fs::remove_file(fixture)?;

        let sheet = comments.get("Data").context("worksheet comments missing")?;
        assert_eq!(
            sheet.get("A1"),
            Some(&WorkbookComment {
                text: "First note".to_owned(),
                author: Some("Data Team".to_owned()),
            })
        );
        assert_eq!(
            sheet.get("B2"),
            Some(&WorkbookComment {
                text: "Empty author".to_owned(),
                author: Some(String::new()),
            })
        );
        assert_eq!(
            sheet.get("C3"),
            Some(&WorkbookComment {
                text: "Unknown author".to_owned(),
                author: None,
            })
        );
        Ok(())
    }

    #[test]
    fn ooxml_comment_reader_rejects_non_ooxml_formats() {
        let error = read_ooxml_comments("not-a-workbook.xls")
            .expect_err("the comment supplement does not read BIFF8 workbooks");
        assert!(error.to_string().contains("only supports"));
    }

    #[test]
    fn relationship_targets_are_resolved_from_their_source_part() {
        assert_eq!(
            norm_part("custom/worksheets", "../comments/comments1.xml"),
            "custom/comments/comments1.xml"
        );
        assert_eq!(norm_part("", "/custom/workbook.xml"), "custom/workbook.xml");
        assert_eq!(
            relationships_part("custom/worksheets/sheet1.xml"),
            "custom/worksheets/_rels/sheet1.xml.rels"
        );
    }

    #[test]
    fn header_inventory_resolves_custom_workbook_part_and_excludes_chartsheets() -> Result<()> {
        let fixture = nonstandard_header_fixture()?;
        let inventory = read_header_inventory(&fixture)?;

        assert_eq!(inventory.len(), 1);
        assert_eq!(inventory[0].sheet_name, "外汇报表");
        assert_eq!(
            inventory[0].headers,
            vec![
                WorksheetHeaderText {
                    sheet_name: "外汇报表".to_owned(),
                    variant: HeaderVariant::Odd,
                    text: "&L单位：万元".to_owned(),
                },
                WorksheetHeaderText {
                    sheet_name: "外汇报表".to_owned(),
                    variant: HeaderVariant::Even,
                    text: "&C第二页".to_owned(),
                },
                WorksheetHeaderText {
                    sheet_name: "外汇报表".to_owned(),
                    variant: HeaderVariant::First,
                    text: String::new(),
                },
            ]
        );

        fs::remove_file(fixture)?;
        Ok(())
    }
}
