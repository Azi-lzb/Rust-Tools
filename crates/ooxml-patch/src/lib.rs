//! Namespace-aware, minimal OOXML package patching primitives.

pub mod excel_open_repair;
pub use excel_open_repair::{repair_excel_com_open_issues, ExcelOpenRepairReport};

use anyhow::{anyhow, bail, ensure, Context, Result};
use roxmltree::Document;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    io::{Read, Write},
    ops::Range,
    path::Path,
};
use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// A requested width change in one worksheet. Column indexes are 1-based.
#[derive(Clone, Debug)]
pub struct ColumnWidthUpdate {
    pub sheet_name: String,
    pub column_index: u32,
    pub width: f64,
}

impl ColumnWidthUpdate {
    pub fn new(sheet_name: impl Into<String>, column_index: u32, width: f64) -> Result<Self> {
        let sheet_name = sheet_name.into();
        if sheet_name.is_empty() {
            bail!("sheet name must not be empty");
        }
        if column_index == 0 {
            bail!("column indexes are 1-based");
        }
        if !width.is_finite() || width < 0.0 {
            bail!("column width must be finite and non-negative");
        }
        Ok(Self {
            sheet_name,
            column_index,
            width,
        })
    }
}

#[derive(Debug)]
pub struct PatchReport {
    pub updated_columns: usize,
    pub changed_worksheet_parts: Vec<String>,
    pub package_parts: usize,
}

#[derive(Debug)]
pub struct FormulaCachePatchReport {
    pub worksheet_part: String,
    pub cell_reference: String,
    pub package_parts: usize,
}

#[derive(Debug, Default, PartialEq, Eq)]
pub struct FormulaCachesClearedReport {
    pub worksheet_parts: Vec<String>,
    pub formula_cells: usize,
    pub cache_values_removed: usize,
}

/// Calculation flags stored in the workbook's `calcPr` element. `None`
/// means the caller does not request a change to that attribute.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct WorkbookCalculationProperties {
    pub full_calc_on_load: Option<bool>,
    pub force_full_calc: Option<bool>,
}

/// Read the explicit workbook calculation flags. `None` means the attribute
/// is absent (which is distinct from an explicit `false`).
pub fn read_workbook_calculation_properties(
    path: impl AsRef<Path>,
) -> Result<WorkbookCalculationProperties> {
    let path = path.as_ref();
    let mut zip = ZipArchive::new(fs::File::open(path)?)?;
    let workbook_part = workbook_part_from_zip(&mut zip)?;
    let xml = read_zip_part(&mut zip, &workbook_part)?;
    let document = Document::parse(std::str::from_utf8(&xml)?)?;
    let workbook = document.root_element();
    let calc_prs = workbook
        .children()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "calcPr"
        })
        .collect::<Vec<_>>();
    ensure!(
        calc_prs.len() <= 1,
        "workbook contains duplicate calcPr elements"
    );
    let Some(calc_pr) = calc_prs.first() else {
        return Ok(WorkbookCalculationProperties::default());
    };
    Ok(WorkbookCalculationProperties {
        full_calc_on_load: calc_pr
            .attribute("fullCalcOnLoad")
            .map(parse_xml_bool)
            .transpose()?,
        force_full_calc: calc_pr
            .attribute("forceFullCalc")
            .map(parse_xml_bool)
            .transpose()?,
    })
}

/// Write an XLSX candidate with targeted updates to workbook calculation flags.
/// The source remains untouched and an existing output is rejected. All
/// package parts except the workbook part are copied without XML modification.
pub fn patch_workbook_calculation_properties(
    source: impl AsRef<Path>,
    output: impl AsRef<Path>,
    requested: WorkbookCalculationProperties,
) -> Result<()> {
    let source = source.as_ref();
    let output = output.as_ref();
    ensure!(
        requested.full_calc_on_load.is_some() || requested.force_full_calc.is_some(),
        "calculation-property patch must request at least one attribute"
    );
    let canonical_source = fs::canonicalize(source).context("resolve source workbook")?;
    ensure!(
        source != output && (!output.exists() || fs::canonicalize(output)? != canonical_source),
        "calculation-property patch requires a distinct output path"
    );
    ensure!(
        !output.exists(),
        "calculation-property patch refuses to overwrite an existing output"
    );

    let mut input_zip = ZipArchive::new(fs::File::open(source)?)?;
    let workbook_part = workbook_part_from_zip(&mut input_zip)?;
    let original = read_zip_part(&mut input_zip, &workbook_part)?;
    let patched = patch_calc_pr_xml(&original, requested)?;
    let parent = output.parent().filter(|p| !p.as_os_str().is_empty());
    if let Some(parent) = parent {
        fs::create_dir_all(parent)?;
    }
    let mut created = false;
    let result = (|| -> Result<()> {
        let file = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(output)?;
        created = true;
        let mut writer = ZipWriter::new(file);
        for index in 0..input_zip.len() {
            let mut entry = input_zip.by_index(index)?;
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
            if name == workbook_part {
                writer.write_all(&patched)?;
            } else {
                std::io::copy(&mut entry, &mut writer)?;
            }
        }
        writer.finish()?;
        let readback = read_workbook_calculation_properties(output)?;
        ensure!(
            requested
                .full_calc_on_load
                .is_none_or(|v| readback.full_calc_on_load == Some(v)),
            "fullCalcOnLoad readback mismatch"
        );
        ensure!(
            requested
                .force_full_calc
                .is_none_or(|v| readback.force_full_calc == Some(v)),
            "forceFullCalc readback mismatch"
        );
        Ok(())
    })();
    if result.is_err() && created {
        let _ = fs::remove_file(output);
    }
    result
}

fn parse_xml_bool(value: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => bail!("invalid OOXML boolean {value:?}"),
    }
}

fn workbook_part_from_zip(zip: &mut ZipArchive<fs::File>) -> Result<String> {
    let root_rels_xml = read_zip_part(zip, "_rels/.rels")?;
    let root_rels = Document::parse(std::str::from_utf8(&root_rels_xml)?)?;
    let target = root_rels
        .descendants()
        .find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
                && node
                    .attribute("Type")
                    .is_some_and(|kind| kind.ends_with("/officeDocument"))
                && !node
                    .attribute("TargetMode")
                    .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
        })
        .and_then(|node| node.attribute("Target"))
        .context("package relationships do not identify an internal workbook part")?;
    Ok(normalize_part("", target))
}

fn patch_calc_pr_xml(xml: &[u8], requested: WorkbookCalculationProperties) -> Result<Vec<u8>> {
    let document = Document::parse(std::str::from_utf8(xml)?)?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(MAIN_NS) && root.tag_name().name() == "workbook",
        "workbook part has an unexpected root element"
    );
    let calc_prs = root
        .children()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "calcPr"
        })
        .collect::<Vec<_>>();
    ensure!(
        calc_prs.len() <= 1,
        "workbook contains duplicate calcPr elements"
    );
    if let Some(calc_pr) = calc_prs.first() {
        let mut edits = Vec::new();
        if let Some(value) = requested.full_calc_on_load {
            edits.push(attribute_update(
                xml,
                *calc_pr,
                "fullCalcOnLoad",
                if value { "1" } else { "0" },
            )?);
        }
        if let Some(value) = requested.force_full_calc {
            edits.push(attribute_update(
                xml,
                *calc_pr,
                "forceFullCalc",
                if value { "1" } else { "0" },
            )?);
        }
        return apply_xml_edits(xml, edits);
    }
    let qname = child_qname(xml, root, "calcPr");
    let mut attributes = String::new();
    if let Some(value) = requested.full_calc_on_load {
        attributes.push_str(&format!(
            " fullCalcOnLoad=\"{}\"",
            if value { "1" } else { "0" }
        ));
    }
    if let Some(value) = requested.force_full_calc {
        attributes.push_str(&format!(
            " forceFullCalc=\"{}\"",
            if value { "1" } else { "0" }
        ));
    }
    let element = format!("<{qname}{attributes}/>");
    let insertion = root
        .children()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && matches!(
                    node.tag_name().name(),
                    "oleSize"
                        | "customWorkbookViews"
                        | "pivotCaches"
                        | "smartTagPr"
                        | "smartTagTypes"
                        | "webPublishing"
                        | "fileRecoveryPr"
                        | "webPublishObjects"
                        | "extLst"
                )
        })
        .map(|node| node.range().start)
        .min()
        .unwrap_or(closing_tag_start(xml, root.range())?);
    apply_xml_edits(xml, vec![(insertion..insertion, element.into_bytes())])
}

/// Remove cached values and cached result types from ordinary formula cells in
/// an XLSX copy. This lets an external calculation engine treat formulas as
/// needing evaluation instead of reusing writer-provided placeholder caches.
/// Formula text and every non-worksheet package part are preserved. Formula
/// structures with attributes (such as shared/array formulas) are rejected.
pub fn clear_formula_caches_for_recalculation(
    source: impl AsRef<Path>,
    output: impl AsRef<Path>,
) -> Result<FormulaCachesClearedReport> {
    let source = source.as_ref();
    let output = output.as_ref();
    let canonical_source = fs::canonicalize(source).context("resolve source workbook")?;
    ensure!(
        source != output && (!output.exists() || fs::canonicalize(output)? != canonical_source),
        "formula-cache clearing requires a distinct output path"
    );
    ensure!(
        !output.exists(),
        "formula-cache clearing refuses to overwrite an existing output"
    );
    validate_header_package(source)
        .context("validate formula recalculation input package safety")?;

    let worksheet_parts = sheet_parts(source)?;
    let mut input_zip = ZipArchive::new(fs::File::open(source)?)?;
    let mut replacements = BTreeMap::new();
    let mut report = FormulaCachesClearedReport::default();
    for (_, worksheet_part) in worksheet_parts {
        let xml = read_zip_part(&mut input_zip, &worksheet_part)?;
        let (patched, formula_cells, cache_values_removed) =
            clear_formula_caches_for_recalculation_xml(&xml)?;
        report.formula_cells += formula_cells;
        report.cache_values_removed += cache_values_removed;
        if patched != xml {
            report.worksheet_parts.push(worksheet_part.clone());
            replacements.insert(worksheet_part, patched);
        }
    }
    ensure!(report.formula_cells > 0, "workbook contains no formulas");

    if let Some(parent) = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)?;
    }
    let mut created = false;
    let result = (|| -> Result<()> {
        let file = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(output)
            .with_context(|| format!("create formula-cache-free copy {}", output.display()))?;
        created = true;
        let mut writer = ZipWriter::new(file);
        for index in 0..input_zip.len() {
            let mut entry = input_zip.by_index(index)?;
            let name = entry.name().to_owned();
            let mut options = SimpleFileOptions::default().compression_method(entry.compression());
            if let Some(modified) = entry.last_modified() {
                options = options.last_modified_time(modified);
            }
            if entry.is_dir() {
                writer.add_directory(name, options)?;
            } else {
                writer.start_file(&name, options)?;
                if let Some(bytes) = replacements.get(&name) {
                    writer.write_all(bytes)?;
                } else {
                    std::io::copy(&mut entry, &mut writer)?;
                }
            }
        }
        writer.finish()?;

        let mut output_zip = ZipArchive::new(fs::File::open(output)?)?;
        for worksheet_part in &report.worksheet_parts {
            let xml = read_zip_part(&mut output_zip, worksheet_part)?;
            let (_, remaining_formula_cells, remaining_cache_values) =
                clear_formula_caches_for_recalculation_xml(&xml)?;
            ensure!(
                remaining_formula_cells > 0 && remaining_cache_values == 0,
                "formula cache removal readback failed for {worksheet_part}"
            );
        }
        Ok(())
    })();
    if result.is_err() && created {
        let _ = fs::remove_file(output);
    }
    result?;
    Ok(report)
}

fn clear_formula_caches_for_recalculation_xml(
    source_xml: &[u8],
) -> Result<(Vec<u8>, usize, usize)> {
    let xml = std::str::from_utf8(source_xml).context("worksheet XML must be UTF-8")?;
    let document = Document::parse(xml).context("parse worksheet XML")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(MAIN_NS) && root.tag_name().name() == "worksheet",
        "expected a SpreadsheetML worksheet root"
    );

    let mut formula_cells = 0usize;
    let mut cache_values_removed = 0usize;
    let mut ranges = Vec::new();
    for cell in root.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "c"
    }) {
        let formulas = cell
            .children()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "f"
            })
            .collect::<Vec<_>>();
        if formulas.is_empty() {
            continue;
        }
        ensure!(
            formulas.len() == 1 && formulas[0].attributes().next().is_none(),
            "formula cell {:?} uses an unsupported shared/array or annotated formula",
            cell.attribute("r")
        );
        formula_cells += 1;
        let values = cell
            .children()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "v"
            })
            .collect::<Vec<_>>();
        ensure!(
            values.len() <= 1,
            "formula cell {:?} contains multiple cached value nodes",
            cell.attribute("r")
        );
        if let Some(value) = values.first() {
            ensure!(
                !value.children().any(|node| node.is_element()),
                "formula cell {:?} has nested cache XML",
                cell.attribute("r")
            );
            ranges.push(value.range());
            cache_values_removed += 1;
        }
        let result_type = cell
            .attributes()
            .find(|attribute| attribute.namespace().is_none() && attribute.name() == "t");
        ensure!(
            !cell
                .attributes()
                .any(|attribute| attribute.name() == "t" && attribute.namespace().is_some()),
            "formula cell {:?} has a namespaced result type",
            cell.attribute("r")
        );
        if let Some(attribute) = result_type {
            ranges.push(attribute.range());
        }
    }

    ranges.sort_by_key(|range| std::cmp::Reverse(range.start));
    let mut patched = source_xml.to_vec();
    for range in ranges {
        patched = splice_bytes(&patched, range, b"");
    }
    Ok((patched, formula_cells, cache_values_removed))
}

/// The three worksheet header variants stored in a SpreadsheetML worksheet.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd)]
pub enum HeaderVariant {
    Odd,
    Even,
    First,
}

impl HeaderVariant {
    fn element_name(self) -> &'static str {
        match self {
            Self::Odd => "oddHeader",
            Self::Even => "evenHeader",
            Self::First => "firstHeader",
        }
    }
}

/// A text value found in one worksheet header variant. Left, center, and right
/// sections are represented by SpreadsheetML header control codes in text.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct WorksheetHeaderText {
    pub sheet_name: String,
    pub variant: HeaderVariant,
    pub text: String,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct WorksheetHeaderInventory {
    pub sheet_name: String,
    pub headers: Vec<WorksheetHeaderText>,
}

/// Replace every occurrence in all odd/even/first header variants of a sheet.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct HeaderTextUpdate {
    pub sheet_name: String,
    pub search: String,
    pub replacement: String,
}

#[derive(Debug, Default, Eq, PartialEq)]
pub struct HeaderTextPatchReport {
    pub modified_sheets: usize,
    pub replacements: usize,
    pub changed_worksheet_parts: Vec<String>,
    pub package_parts: usize,
}

/// Request to repair one worksheet hyperlink to an absolute `file://` URI.
/// Optional location/display values replace those attributes when supplied;
/// when omitted, any existing values are preserved.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ExternalHyperlinkPatch {
    pub sheet_name: String,
    pub cell_reference: String,
    pub target_uri: String,
    pub location: Option<String>,
    pub display: Option<String>,
}

impl ExternalHyperlinkPatch {
    pub fn new(
        sheet_name: impl Into<String>,
        cell_reference: impl Into<String>,
        target_uri: impl Into<String>,
    ) -> Result<Self> {
        let sheet_name = sheet_name.into();
        let cell_reference = cell_reference.into();
        let target_uri = target_uri.into();
        ensure!(!sheet_name.is_empty(), "sheet name must not be empty");
        ensure!(
            !cell_reference.is_empty(),
            "hyperlink cell reference must not be empty"
        );
        validate_absolute_file_uri(&target_uri)?;
        Ok(Self {
            sheet_name,
            cell_reference,
            target_uri,
            location: None,
            display: None,
        })
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ExternalHyperlinkPatchReport {
    pub sheet_name: String,
    pub cell_reference: String,
    pub worksheet_part: String,
    pub relationships_part: String,
    pub relationship_id: String,
    pub package_parts: usize,
}

/// Returns workbook parts outside the preservation scope of the current adapters.
pub fn unverified_parts(path: impl AsRef<Path>) -> Result<Vec<String>> {
    let file = fs::File::open(path.as_ref())
        .with_context(|| format!("open workbook {}", path.as_ref().display()))?;
    let mut zip = ZipArchive::new(file)?;
    let mut found = Vec::new();
    for index in 0..zip.len() {
        let entry = zip.by_index(index)?;
        if entry.is_dir() {
            continue;
        }
        let name = entry.name().to_owned();
        let lower = name.to_ascii_lowercase();
        let blocked = lower.starts_with("xl/charts/")
            || lower.starts_with("xl/pivot")
            || lower.starts_with("xl/activex/")
            || lower.starts_with("xl/embeddings/")
            || lower.starts_with("xl/externallinks/")
            || lower == "xl/vbaproject.bin"
            || (lower.starts_with("xl/drawings/")
                && !lower.contains("vml")
                && !lower.ends_with(".rels"));
        if blocked {
            found.push(name);
        }
    }
    Ok(found)
}

/// Refuses to process a workbook containing parts unsupported by this adapter.
pub fn reject_unverified_parts(path: impl AsRef<Path>) -> Result<()> {
    let found = unverified_parts(path)?;
    if found.is_empty() {
        Ok(())
    } else {
        bail!("refusing unverified workbook parts: {}", found.join(", "))
    }
}

#[cfg(test)]
mod unverified_directory_entry_tests {
    use super::unverified_parts;
    use anyhow::Result;
    use std::{fs, io::Write};
    use zip::{write::SimpleFileOptions, ZipWriter};

    fn fixture(include_drawing_part: bool) -> std::path::PathBuf {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "ooxml_unverified_directory_{}_{}.xlsx",
            std::process::id(),
            nonce
        ));
        let mut zip = ZipWriter::new(fs::File::create(&path).unwrap());
        zip.add_directory("xl/drawings/", SimpleFileOptions::default())
            .unwrap();
        zip.start_file("xl/worksheets/sheet1.xml", SimpleFileOptions::default())
            .unwrap();
        zip.write_all(b"<worksheet/>").unwrap();
        if include_drawing_part {
            zip.start_file("xl/drawings/drawing1.xml", SimpleFileOptions::default())
                .unwrap();
            zip.write_all(b"<drawing/>").unwrap();
        }
        zip.finish().unwrap();
        path
    }

    #[test]
    fn directory_markers_are_ignored_but_unverified_file_parts_stay_blocked() -> Result<()> {
        let directory_only = fixture(false);
        assert!(unverified_parts(&directory_only)?.is_empty());
        fs::remove_file(directory_only)?;

        let with_drawing = fixture(true);
        assert_eq!(
            unverified_parts(&with_drawing)?,
            ["xl/drawings/drawing1.xml"]
        );
        fs::remove_file(with_drawing)?;
        Ok(())
    }
}

fn read_zip_part(zip: &mut ZipArchive<fs::File>, name: &str) -> Result<Vec<u8>> {
    let mut file = zip
        .by_name(name)
        .with_context(|| format!("missing ZIP part {name}"))?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    Ok(bytes)
}

fn normalize_part(base: &str, target: &str) -> String {
    let combined = if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else if base.is_empty() {
        target.to_owned()
    } else {
        format!("{base}/{target}")
    };
    let mut parts: Vec<&str> = Vec::new();
    for part in combined.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
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

fn sheet_parts(path: &Path) -> Result<BTreeMap<String, String>> {
    Ok(strict_sheet_parts(path)?.into_iter().collect())
}

fn strict_sheet_parts(path: &Path) -> Result<Vec<(String, String)>> {
    let mut zip = ZipArchive::new(fs::File::open(path)?)?;
    strict_sheet_parts_from_zip(&mut zip, path)
}

fn strict_sheet_parts_from_zip(
    zip: &mut ZipArchive<fs::File>,
    path: &Path,
) -> Result<Vec<(String, String)>> {
    let root_rels_xml = read_zip_part(zip, "_rels/.rels")?;
    let root_rels = Document::parse(std::str::from_utf8(&root_rels_xml)?)?;
    let workbook_target = root_rels
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
    let workbook_part = normalize_part("", workbook_target);
    let workbook_xml = read_zip_part(zip, &workbook_part)?;
    let rels_xml = read_zip_part(zip, &relationships_part(&workbook_part))?;
    let workbook = Document::parse(std::str::from_utf8(&workbook_xml)?)?;
    let rels = Document::parse(std::str::from_utf8(&rels_xml)?)?;
    let workbook_parent = workbook_part
        .rsplit_once('/')
        .map(|(parent, _)| parent)
        .unwrap_or("");

    let mut targets = BTreeMap::new();
    for rel in rels.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
            && node.tag_name().name() == "Relationship"
    }) {
        if rel
            .attribute("TargetMode")
            .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
        {
            continue;
        }
        if let (Some(id), Some(kind), Some(target)) = (
            rel.attribute("Id"),
            rel.attribute("Type"),
            rel.attribute("Target"),
        ) {
            targets.insert(id.to_owned(), (kind.to_owned(), target.to_owned()));
        }
    }

    let mut result = Vec::new();
    let mut names = BTreeSet::new();
    let mut parts = BTreeSet::new();
    for sheet in workbook.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().name() == "sheet"
            && node.tag_name().namespace() == Some(MAIN_NS)
    }) {
        let name = sheet
            .attribute("name")
            .ok_or_else(|| anyhow!("workbook sheet has no name"))?;
        let id = sheet
            .attribute((REL_NS, "id"))
            .ok_or_else(|| anyhow!("workbook sheet {name:?} has no relationship id"))?;
        let (kind, target) = targets
            .get(id)
            .ok_or_else(|| anyhow!("workbook sheet {name:?} has an unresolved relationship"))?;
        if !kind.ends_with("/worksheet") {
            continue;
        }
        let part = normalize_part(workbook_parent, target);
        if !names.insert(name.to_owned()) {
            bail!("duplicate worksheet name {name:?}");
        }
        if !parts.insert(part.clone()) {
            bail!("multiple worksheet names resolve to part {part}");
        }
        result.push((name.to_owned(), part));
    }
    if result.is_empty() {
        bail!(
            "could not resolve workbook worksheet relationships in {}",
            path.display()
        );
    }
    Ok(result)
}

/// Patch one existing worksheet hyperlink and its external relationship.
/// Only the selected worksheet part and its `.rels` part are rewritten; every
/// other package part is copied unchanged. The source is never modified and an
/// existing output path is rejected. This is a convenience wrapper over the
/// batch API.
pub fn patch_external_hyperlink(
    source: impl AsRef<Path>,
    output: impl AsRef<Path>,
    patch: &ExternalHyperlinkPatch,
) -> Result<ExternalHyperlinkPatchReport> {
    patch_external_hyperlinks(source, output, std::slice::from_ref(patch))?
        .into_iter()
        .next()
        .context("single hyperlink patch returned no report")
}

/// Patch multiple worksheet hyperlinks with one package output pass and one
/// package-level read-back validation. All sheet/cell and relationship
/// mappings are prepared before the output file is created, so any invalid or
/// ambiguous request rejects the whole batch without leaving a candidate.
pub fn patch_external_hyperlinks(
    source: impl AsRef<Path>,
    output: impl AsRef<Path>,
    patches: &[ExternalHyperlinkPatch],
) -> Result<Vec<ExternalHyperlinkPatchReport>> {
    let source = source.as_ref();
    let output = output.as_ref();
    ensure!(
        !patches.is_empty(),
        "hyperlink patch batch must not be empty"
    );
    let canonical_source = fs::canonicalize(source).context("resolve source workbook")?;
    ensure!(
        source != output
            && (!output.exists()
                || fs::canonicalize(output).context("resolve output workbook")?
                    != canonical_source),
        "hyperlink patch requires a distinct output path"
    );
    ensure!(
        !output.exists(),
        "hyperlink patch refuses to overwrite an existing output"
    );
    validate_header_package(source)?;

    let mut input_zip = ZipArchive::new(fs::File::open(source)?)?;
    let package_parts = input_zip.len();
    let worksheet_parts = strict_sheet_parts_from_zip(&mut input_zip, source)?
        .into_iter()
        .collect::<BTreeMap<_, _>>();
    let mut grouped = BTreeMap::<String, Vec<(usize, &ExternalHyperlinkPatch)>>::new();
    let mut requested_cells = BTreeSet::new();
    let mut reports = vec![None; patches.len()];
    for (index, patch) in patches.iter().enumerate() {
        validate_absolute_file_uri(&patch.target_uri)?;
        let cell_reference = normalize_a1_cell(&patch.cell_reference)?;
        let worksheet_part = worksheet_parts
            .get(&patch.sheet_name)
            .with_context(|| format!("worksheet {:?} does not exist", patch.sheet_name))?;
        ensure!(
            requested_cells.insert((worksheet_part.clone(), cell_reference.clone())),
            "hyperlink patch batch contains duplicate target {}!{}",
            patch.sheet_name,
            cell_reference
        );
        grouped
            .entry(worksheet_part.clone())
            .or_default()
            .push((index, patch));
    }

    let mut replacements = BTreeMap::new();
    for (worksheet_part, grouped_patches) in grouped {
        let relationships_part = relationships_part(&worksheet_part);
        let worksheet_xml = read_zip_part(&mut input_zip, &worksheet_part)?;
        let rels_xml = read_zip_part(&mut input_zip, &relationships_part)?;
        let patch_refs = grouped_patches
            .iter()
            .map(|(_, patch)| *patch)
            .collect::<Vec<_>>();
        let (patched_worksheet, patched_rels, relationship_ids) =
            prepare_external_hyperlinks_parts(&worksheet_xml, &rels_xml, &patch_refs)?;
        for ((index, patch), relationship_id) in grouped_patches.iter().zip(relationship_ids) {
            reports[*index] = Some(ExternalHyperlinkPatchReport {
                sheet_name: patch.sheet_name.clone(),
                cell_reference: normalize_a1_cell(&patch.cell_reference)?,
                worksheet_part: worksheet_part.clone(),
                relationships_part: relationships_part.clone(),
                relationship_id,
                package_parts,
            });
        }
        replacements.insert(worksheet_part, patched_worksheet);
        replacements.insert(relationships_part, patched_rels);
    }

    if let Some(parent) = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)?;
    }
    let mut output_created = false;
    let create_result = (|| -> Result<()> {
        let file = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(output)
            .with_context(|| format!("create hyperlink patch output {}", output.display()))?;
        output_created = true;
        let mut writer = ZipWriter::new(file);
        for index in 0..input_zip.len() {
            let mut entry = input_zip.by_index(index)?;
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
    })();
    if let Err(error) = create_result {
        if output_created {
            let _ = fs::remove_file(output);
        }
        return Err(error);
    }
    if let Err(error) = validate_package_replacements(source, output, &replacements) {
        let _ = fs::remove_file(output);
        return Err(error).context("validate external hyperlink patch output");
    }
    reports
        .into_iter()
        .map(|report| report.context("hyperlink patch completed without a report"))
        .collect()
}

fn prepare_external_hyperlinks_parts(
    worksheet_xml: &[u8],
    rels_xml: &[u8],
    patches: &[&ExternalHyperlinkPatch],
) -> Result<(Vec<u8>, Vec<u8>, Vec<String>)> {
    ensure!(
        !patches.is_empty(),
        "worksheet hyperlink patch list must not be empty"
    );
    let worksheet_text =
        std::str::from_utf8(worksheet_xml).context("worksheet XML must be UTF-8")?;
    let worksheet = Document::parse(worksheet_text).context("parse worksheet XML")?;
    let root = worksheet.root_element();
    ensure!(
        root.tag_name().namespace() == Some(MAIN_NS) && root.tag_name().name() == "worksheet",
        "expected a SpreadsheetML worksheet root"
    );
    let all_links = root
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "hyperlink"
        })
        .collect::<Vec<_>>();
    let mut worksheet_edits = Vec::new();
    let mut requested_relationships = BTreeSet::new();
    let mut relationship_updates = Vec::new();
    let mut normalized_cells = Vec::new();
    for patch in patches {
        let normalized_cell = normalize_a1_cell(&patch.cell_reference)?;
        let matching_links = all_links
            .iter()
            .filter(|link| {
                link.attribute("ref")
                    .and_then(|reference| normalize_a1_cell(reference).ok())
                    .as_deref()
                    == Some(normalized_cell.as_str())
            })
            .collect::<Vec<_>>();
        ensure!(
            matching_links.len() == 1,
            "expected exactly one hyperlink for {}!{}, found {}",
            patch.sheet_name,
            normalized_cell,
            matching_links.len()
        );
        let hyperlink = *matching_links[0];
        let relationship_id = hyperlink
            .attribute((REL_NS, "id"))
            .with_context(|| {
                format!(
                    "hyperlink {}!{} has no relationship id",
                    patch.sheet_name, normalized_cell
                )
            })?
            .to_owned();
        ensure!(
            requested_relationships.insert(relationship_id.clone()),
            "multiple hyperlink patches resolve to relationship {relationship_id:?}"
        );
        let references = all_links
            .iter()
            .filter(|link| link.attribute((REL_NS, "id")) == Some(relationship_id.as_str()))
            .count();
        ensure!(
            references == 1,
            "hyperlink relationship {relationship_id:?} is shared by {references} worksheet links"
        );
        if let Some(display) = patch.display.as_deref() {
            worksheet_edits.push(attribute_update_escaped(
                worksheet_xml,
                hyperlink,
                "display",
                display,
            )?);
        }
        if let Some(location) = patch.location.as_deref() {
            worksheet_edits.push(attribute_update_escaped(
                worksheet_xml,
                hyperlink,
                "location",
                location,
            )?);
        }
        normalized_cells.push(normalized_cell);
        relationship_updates.push((relationship_id, patch.target_uri.as_str()));
    }

    let rels_text =
        std::str::from_utf8(rels_xml).context("worksheet relationships XML must be UTF-8")?;
    let rels = Document::parse(rels_text).context("parse worksheet relationships XML")?;
    ensure!(
        rels.root_element().tag_name().namespace() == Some(PACKAGE_REL_NS)
            && rels.root_element().tag_name().name() == "Relationships",
        "worksheet relationships part has an unexpected root element"
    );
    let mut rel_edits = Vec::new();
    for (relationship_id, target_uri) in &relationship_updates {
        let matches = rels
            .descendants()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                    && node.tag_name().name() == "Relationship"
                    && node.attribute("Id") == Some(relationship_id.as_str())
            })
            .collect::<Vec<_>>();
        ensure!(
            matches.len() == 1,
            "hyperlink relationship {relationship_id:?} is missing or ambiguous"
        );
        let relationship = matches[0];
        ensure!(
            relationship
                .attribute("Type")
                .is_some_and(|kind| kind.ends_with("/hyperlink")),
            "relationship {relationship_id:?} is not a hyperlink"
        );
        ensure!(
            relationship.attribute("Target").is_some(),
            "hyperlink relationship {relationship_id:?} has no Target attribute"
        );
        rel_edits.push(attribute_update_escaped(
            rels_xml,
            relationship,
            "Target",
            target_uri,
        )?);
        if relationship.attribute("TargetMode") != Some("External") {
            rel_edits.push(attribute_update_escaped(
                rels_xml,
                relationship,
                "TargetMode",
                "External",
            )?);
        }
    }
    let patched_worksheet = if worksheet_edits.is_empty() {
        worksheet_xml.to_vec()
    } else {
        let patched = apply_xml_edits(worksheet_xml, worksheet_edits)?;
        Document::parse(std::str::from_utf8(&patched)?)
            .context("patched worksheet XML is invalid")?;
        patched
    };
    let patched_rels = apply_xml_edits(rels_xml, rel_edits)?;
    Document::parse(std::str::from_utf8(&patched_rels)?)
        .context("patched worksheet relationships XML is invalid")?;
    Ok((
        patched_worksheet,
        patched_rels,
        relationship_updates
            .into_iter()
            .map(|(relationship_id, _)| relationship_id)
            .collect(),
    ))
}

fn validate_absolute_file_uri(target: &str) -> Result<()> {
    ensure!(
        target
            .get(..7)
            .is_some_and(|scheme| scheme.eq_ignore_ascii_case("file://")),
        "hyperlink target must be an absolute file:// URI"
    );
    let path_or_authority = &target[7..];
    ensure!(
        !path_or_authority.is_empty()
            && (path_or_authority.starts_with('/')
                || path_or_authority
                    .split_once('/')
                    .is_some_and(|(host, _)| !host.is_empty()))
            && !target
                .chars()
                .any(|ch| ch.is_control() || ch.is_whitespace() || ch == '\\'),
        "hyperlink target must contain an absolute filesystem path"
    );
    Ok(())
}

fn normalize_a1_cell(reference: &str) -> Result<String> {
    let (row, column) = parse_a1_cell(reference)?;
    ensure!(
        row <= 1_048_576 && column <= 16_384,
        "hyperlink cell reference is outside worksheet bounds"
    );
    let mut value = column;
    let mut letters = Vec::new();
    while value > 0 {
        let rem = (value - 1) % 26;
        letters.push((b'A' + rem as u8) as char);
        value = (value - 1) / 26;
    }
    letters.reverse();
    Ok(format!(
        "{}{}",
        letters.into_iter().collect::<String>(),
        row
    ))
}

fn attribute_update_escaped(
    xml: &[u8],
    node: roxmltree::Node<'_, '_>,
    name: &str,
    value: &str,
) -> Result<(Range<usize>, Vec<u8>)> {
    attribute_update(xml, node, name, &escape_xml_attribute(value))
}

fn escape_xml_attribute(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
        .replace('\t', "&#x9;")
        .replace('\n', "&#xA;")
        .replace('\r', "&#xD;")
}

fn validate_package_replacements(
    source: &Path,
    candidate: &Path,
    replacements: &BTreeMap<String, Vec<u8>>,
) -> Result<()> {
    let mut source_zip = ZipArchive::new(fs::File::open(source)?)?;
    let mut candidate_zip = ZipArchive::new(fs::File::open(candidate)?)?;
    ensure!(
        source_zip.len() == candidate_zip.len(),
        "candidate package part count differs from source"
    );
    for index in 0..source_zip.len() {
        let mut source_entry = source_zip.by_index(index)?;
        let mut candidate_entry = candidate_zip.by_index(index)?;
        let name = source_entry.name().to_owned();
        ensure!(
            candidate_entry.name() == name,
            "candidate package part order/name differs at index {index}"
        );
        ensure!(
            source_entry.compression() == candidate_entry.compression(),
            "candidate package compression differs for part {name}"
        );
        let mut source_bytes = Vec::new();
        let mut candidate_bytes = Vec::new();
        source_entry.read_to_end(&mut source_bytes)?;
        candidate_entry.read_to_end(&mut candidate_bytes)?;
        let expected = replacements.get(&name).unwrap_or(&source_bytes);
        ensure!(
            candidate_bytes == *expected,
            "candidate package part differs from expected content: {name}"
        );
    }
    Ok(())
}

fn validate_header_package(path: &Path) -> Result<()> {
    let mut zip = ZipArchive::new(fs::File::open(path)?)?;
    for index in 0..zip.len() {
        let mut entry = zip.by_index(index)?;
        let name = entry.name().to_ascii_lowercase();
        if name.starts_with("_xmlsignatures/")
            || name.starts_with("xl/_xmlsignatures/")
            || name == "xl/vbaproject.bin"
        {
            bail!(
                "refusing package with macros or digital signatures: {}",
                entry.name()
            );
        }
        if name.ends_with(".rels") {
            let mut xml = Vec::new();
            entry.read_to_end(&mut xml)?;
            let xml = std::str::from_utf8(&xml).context("relationships XML must be UTF-8")?;
            if xml.contains("digital-signature") || xml.contains("digitalSignature") {
                bail!("refusing package with digital-signature relationships");
            }
        }
    }
    Ok(())
}

fn worksheet_headers(sheet_name: &str, part: &str, xml: &[u8]) -> Result<WorksheetHeaderInventory> {
    let document =
        Document::parse(std::str::from_utf8(xml).context("worksheet XML must be UTF-8")?)
            .with_context(|| format!("parse worksheet part {part}"))?;
    let root = document.root_element();
    if root.tag_name().name() != "worksheet" || root.tag_name().namespace() != Some(MAIN_NS) {
        bail!("expected a SpreadsheetML worksheet root in {part}");
    }
    let header_footer = root.children().find(|node| {
        node.is_element()
            && node.tag_name().name() == "headerFooter"
            && node.tag_name().namespace() == Some(MAIN_NS)
    });
    let mut headers = Vec::new();
    if let Some(header_footer) = header_footer {
        for variant in [
            HeaderVariant::Odd,
            HeaderVariant::Even,
            HeaderVariant::First,
        ] {
            let mut matches = header_footer.children().filter(|node| {
                node.is_element()
                    && node.tag_name().name() == variant.element_name()
                    && node.tag_name().namespace() == Some(MAIN_NS)
            });
            let Some(header) = matches.next() else {
                continue;
            };
            if matches.next().is_some() {
                bail!(
                    "duplicate {} element in worksheet part {part}",
                    variant.element_name()
                );
            }
            let mut text_nodes = header.children().filter(|node| node.is_text());
            let text = text_nodes
                .next()
                .and_then(|node| node.text())
                .unwrap_or_default();
            if text_nodes.next().is_some() || header.children().any(|node| !node.is_text()) {
                bail!(
                    "unsupported nested markup in {} on sheet {sheet_name:?}",
                    variant.element_name()
                );
            }
            headers.push(WorksheetHeaderText {
                sheet_name: sheet_name.to_owned(),
                variant,
                text: text.to_owned(),
            });
        }
    }
    Ok(WorksheetHeaderInventory {
        sheet_name: sheet_name.to_owned(),
        headers,
    })
}

/// Reads odd/even/first header text for each worksheet using workbook
/// relationships. It does not modify the package.
pub fn read_header_inventory(path: impl AsRef<Path>) -> Result<Vec<WorksheetHeaderInventory>> {
    let path = path.as_ref();
    validate_header_package(path)?;
    let parts = strict_sheet_parts(path)?;
    let mut zip = ZipArchive::new(fs::File::open(path)?)?;
    parts
        .into_iter()
        .map(|(sheet_name, part)| {
            let xml = read_zip_part(&mut zip, &part)?;
            worksheet_headers(&sheet_name, &part, &xml)
        })
        .collect()
}

fn escape_xml_text(text: &str) -> String {
    let mut escaped = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

/// Element-content escaping for formula text: only `&`, `<` and `>` are
/// required there, so quoted sheet names keep their literal apostrophes.
fn escape_formula_text(text: &str) -> String {
    let mut escaped = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

fn replace_header_xml(
    sheet_name: &str,
    part: &str,
    source_xml: &[u8],
    update: &HeaderTextUpdate,
) -> Result<(Vec<u8>, usize)> {
    let xml = std::str::from_utf8(source_xml).context("worksheet XML must be UTF-8")?;
    let document = Document::parse(xml).with_context(|| format!("parse worksheet part {part}"))?;
    let root = document.root_element();
    if root.tag_name().name() != "worksheet" || root.tag_name().namespace() != Some(MAIN_NS) {
        bail!("expected a SpreadsheetML worksheet root in {part}");
    }
    let header_footer = root.children().find(|node| {
        node.is_element()
            && node.tag_name().name() == "headerFooter"
            && node.tag_name().namespace() == Some(MAIN_NS)
    });
    let Some(header_footer) = header_footer else {
        return Ok((source_xml.to_vec(), 0));
    };

    let mut edits = Vec::new();
    let mut replacement_count = 0;
    for variant in [
        HeaderVariant::Odd,
        HeaderVariant::Even,
        HeaderVariant::First,
    ] {
        let mut matches = header_footer.children().filter(|node| {
            node.is_element()
                && node.tag_name().name() == variant.element_name()
                && node.tag_name().namespace() == Some(MAIN_NS)
        });
        let Some(header) = matches.next() else {
            continue;
        };
        if matches.next().is_some() {
            bail!(
                "duplicate {} element in worksheet part {part}",
                variant.element_name()
            );
        }
        let children = header.children().collect::<Vec<_>>();
        if children.iter().any(|node| !node.is_text()) || children.len() > 1 {
            bail!(
                "unsupported nested markup in {} on sheet {sheet_name:?}",
                variant.element_name()
            );
        }
        let Some(text_node) = children.first().copied() else {
            continue;
        };
        let text = text_node.text().unwrap_or_default();
        let count = text.matches(&update.search).count();
        if count == 0 {
            continue;
        }
        let replaced = text.replace(&update.search, &update.replacement);
        edits.push((text_node.range(), escape_xml_text(&replaced).into_bytes()));
        replacement_count += count;
    }

    edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
    let mut output = source_xml.to_vec();
    for (range, replacement) in edits {
        output = splice_bytes(&output, range, &replacement);
    }
    Ok((output, replacement_count))
}

fn prepare_header_text_patch(
    source: &Path,
    updates: &[HeaderTextUpdate],
) -> Result<(BTreeMap<String, Vec<u8>>, HeaderTextPatchReport)> {
    if updates.is_empty() {
        bail!("header text patch requires at least one update");
    }
    validate_header_package(source)?;
    let source_parts = strict_sheet_parts(source)?;
    let mut seen_sheets = BTreeSet::new();
    for update in updates {
        if update.sheet_name.trim().is_empty()
            || update.search.is_empty()
            || update.replacement.is_empty()
            || update.search == update.replacement
        {
            bail!("header text update requires a sheet and distinct non-empty text values");
        }
        if !seen_sheets.insert(update.sheet_name.as_str()) {
            bail!(
                "duplicate header text update for sheet {:?}",
                update.sheet_name
            );
        }
    }

    let part_by_sheet: BTreeMap<_, _> = source_parts.iter().cloned().collect();
    let mut zip = ZipArchive::new(fs::File::open(source)?)?;
    let mut replacements = BTreeMap::new();
    let mut changed_parts = Vec::new();
    let mut replacement_count = 0;
    for update in updates {
        let part = part_by_sheet.get(&update.sheet_name).ok_or_else(|| {
            anyhow!(
                "worksheet {:?} is missing from source workbook",
                update.sheet_name
            )
        })?;
        let xml = read_zip_part(&mut zip, part)?;
        let (patched, count) = replace_header_xml(&update.sheet_name, part, &xml, update)?;
        if count == 0 {
            bail!(
                "header text {:?} was not found on worksheet {:?}",
                update.search,
                update.sheet_name
            );
        }
        replacement_count += count;
        replacements.insert(part.clone(), patched);
        changed_parts.push(part.clone());
    }
    changed_parts.sort();
    Ok((
        replacements,
        HeaderTextPatchReport {
            modified_sheets: updates.len(),
            replacements: replacement_count,
            changed_worksheet_parts: changed_parts,
            package_parts: zip.len(),
        },
    ))
}

/// Performs a read-only preflight and reports whether every requested header
/// replacement is present before any output is written.
pub fn preflight_header_text_patch(
    source: impl AsRef<Path>,
    updates: &[HeaderTextUpdate],
) -> Result<HeaderTextPatchReport> {
    let (_, report) = prepare_header_text_patch(source.as_ref(), updates)?;
    Ok(report)
}

/// Applies targeted header text replacements. All validation and XML edits are
/// prepared before the ZIP output is opened; only worksheet parts containing
/// requested header text are changed.
pub fn replace_header_text(
    source: impl AsRef<Path>,
    output: impl AsRef<Path>,
    updates: &[HeaderTextUpdate],
) -> Result<HeaderTextPatchReport> {
    let source = source.as_ref();
    let output = output.as_ref();
    let canonical_source = fs::canonicalize(source).context("resolve source workbook")?;
    if output.exists()
        && fs::canonicalize(output).context("resolve output workbook")? == canonical_source
    {
        bail!("header text patch never edits the source workbook in place");
    }
    let (replacements, mut report) = prepare_header_text_patch(source, updates)?;
    if let Some(parent) = output.parent() {
        fs::create_dir_all(parent)?;
    }

    let mut input_zip = ZipArchive::new(fs::File::open(source)?)?;
    let package_parts = input_zip.len();
    let mut writer = ZipWriter::new(fs::File::create(output)?);
    for index in 0..package_parts {
        let mut entry = input_zip.by_index(index)?;
        let name = entry.name().to_owned();
        let method = entry.compression();
        let mut options = SimpleFileOptions::default().compression_method(method);
        if let Some(modified) = entry.last_modified() {
            options = options.last_modified_time(modified);
        }
        if entry.is_dir() {
            writer.add_directory(name, options)?;
            continue;
        }
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes)?;
        writer.start_file(&name, options)?;
        writer.write_all(replacements.get(&name).unwrap_or(&bytes))?;
    }
    writer.finish()?;
    report.package_parts = package_parts;
    Ok(report)
}

/// Validates that a candidate is exactly the source package with only the
/// requested header text nodes changed. ZIP compression bytes may differ.
pub fn validate_header_text_patch(
    source: impl AsRef<Path>,
    candidate: impl AsRef<Path>,
    updates: &[HeaderTextUpdate],
) -> Result<()> {
    let source = source.as_ref();
    let candidate = candidate.as_ref();
    let (expected_replacements, _) = prepare_header_text_patch(source, updates)?;
    let mut source_zip = ZipArchive::new(fs::File::open(source)?)?;
    let mut candidate_zip = ZipArchive::new(fs::File::open(candidate)?)?;
    if source_zip.len() != candidate_zip.len() {
        bail!("candidate ZIP part count differs from source");
    }
    for index in 0..source_zip.len() {
        let mut source_entry = source_zip.by_index(index)?;
        let mut candidate_entry = candidate_zip.by_index(index)?;
        let name = source_entry.name().to_owned();
        if candidate_entry.name() != name {
            bail!("candidate ZIP part order/name differs at index {index}");
        }
        if source_entry.compression() != candidate_entry.compression() {
            bail!("candidate ZIP compression changed for part {name}");
        }
        let mut source_bytes = Vec::new();
        let mut candidate_bytes = Vec::new();
        source_entry.read_to_end(&mut source_bytes)?;
        candidate_entry.read_to_end(&mut candidate_bytes)?;
        let expected = expected_replacements.get(&name).unwrap_or(&source_bytes);
        if candidate_bytes != *expected {
            bail!("candidate contains an unexpected change in ZIP part {name}");
        }
    }
    Ok(())
}

fn splice_bytes(source: &[u8], range: Range<usize>, replacement: &[u8]) -> Vec<u8> {
    let mut output =
        Vec::with_capacity(source.len() - (range.end - range.start) + replacement.len());
    output.extend_from_slice(&source[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[range.end..]);
    output
}

fn patched_source_cols(
    source_xml: &[u8],
    cols: roxmltree::Node<'_, '_>,
    updates: &[&ColumnWidthUpdate],
) -> Result<Vec<u8>> {
    let cols_range = cols.range();
    let mut fragment = source_xml[cols_range.clone()].to_vec();
    let mut edits = Vec::new();
    for update in updates {
        let col = cols.children().find(|node| {
            node.is_element()
                && node.tag_name().name() == "col"
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.attribute("min").and_then(|v| v.parse::<u32>().ok()) == Some(update.column_index)
                && node.attribute("max").and_then(|v| v.parse::<u32>().ok()) == Some(update.column_index)
        }).ok_or_else(|| anyhow!("column {} on sheet {:?} is not represented by an exact source <col> range; refusing to widen the patch", update.column_index, update.sheet_name))?;
        let width = col
            .attributes()
            .find(|attr| attr.name() == "width")
            .ok_or_else(|| {
                anyhow!(
                    "column {} has no source width attribute",
                    update.column_index
                )
            })?;
        let range = width.range_value();
        let relative = (range.start - cols_range.start)..(range.end - cols_range.start);
        edits.push((relative, update.width.to_string().into_bytes()));
    }
    edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
    for (range, replacement) in edits {
        fragment = splice_bytes(&fragment, range, &replacement);
    }
    Ok(fragment)
}

fn replace_cols(
    source_xml: &[u8],
    output_xml: &[u8],
    updates: &[&ColumnWidthUpdate],
) -> Result<Vec<u8>> {
    let source_doc = Document::parse(std::str::from_utf8(source_xml)?)?;
    let output_doc = Document::parse(std::str::from_utf8(output_xml)?)?;
    let source_root = source_doc.root_element();
    let output_root = output_doc.root_element();
    if source_root.tag_name().name() != "worksheet"
        || output_root.tag_name().name() != "worksheet"
        || source_root.tag_name().namespace() != Some(MAIN_NS)
        || output_root.tag_name().namespace() != Some(MAIN_NS)
    {
        bail!("expected SpreadsheetML worksheet XML roots");
    }
    let source_cols = source_root.children().find(|node| {
        node.is_element()
            && node.tag_name().name() == "cols"
            && node.tag_name().namespace() == Some(MAIN_NS)
    });
    let output_cols = output_root.children().find(|node| {
        node.is_element()
            && node.tag_name().name() == "cols"
            && node.tag_name().namespace() == Some(MAIN_NS)
    });
    let replacement = match source_cols {
        Some(cols) => Some(patched_source_cols(source_xml, cols, updates)?),
        None if updates.is_empty() => None,
        None => {
            bail!("source worksheet has no <cols> node; refusing to synthesize column definitions")
        }
    };
    if let Some(output_cols) = output_cols {
        let range = output_cols.range();
        Ok(match replacement {
            Some(bytes) => splice_bytes(output_xml, range, &bytes),
            None => splice_bytes(output_xml, range, b""),
        })
    } else if let Some(bytes) = replacement {
        let sheet_data = output_root
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "sheetData")
            .ok_or_else(|| anyhow!("output worksheet has no sheetData node"))?;
        Ok(splice_bytes(
            output_xml,
            sheet_data.range().start..sheet_data.range().start,
            &bytes,
        ))
    } else {
        Ok(output_xml.to_vec())
    }
}

fn clear_empty_formula_cache_xml(source_xml: &[u8], cell_reference: &str) -> Result<Vec<u8>> {
    if cell_reference.is_empty() || !cell_reference.bytes().all(|b| b.is_ascii_alphanumeric()) {
        bail!("invalid cell reference {cell_reference:?}");
    }
    let xml = std::str::from_utf8(source_xml).context("worksheet XML must be UTF-8")?;
    let document = Document::parse(xml).context("parse worksheet XML")?;
    let root = document.root_element();
    if root.tag_name().name() != "worksheet" || root.tag_name().namespace() != Some(MAIN_NS) {
        bail!("expected a SpreadsheetML worksheet root");
    }

    let cells = root
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().name() == "c"
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.attribute("r") == Some(cell_reference)
        })
        .collect::<Vec<_>>();
    if cells.len() != 1 {
        bail!(
            "expected exactly one cell {cell_reference} in worksheet, found {}",
            cells.len()
        );
    }
    let cell = cells[0];
    let formula_nodes = cell
        .children()
        .filter(|node| {
            node.is_element()
                && node.tag_name().name() == "f"
                && node.tag_name().namespace() == Some(MAIN_NS)
        })
        .collect::<Vec<_>>();
    if formula_nodes.len() != 1 || formula_nodes[0].attributes().next().is_some() {
        bail!("cell {cell_reference} must contain one ordinary, unannotated formula");
    }
    let value_nodes = cell
        .children()
        .filter(|node| {
            node.is_element()
                && node.tag_name().name() == "v"
                && node.tag_name().namespace() == Some(MAIN_NS)
        })
        .collect::<Vec<_>>();
    if value_nodes.len() > 1 {
        bail!("cell {cell_reference} contains multiple formula cache values");
    }
    if cell.children().any(|node| {
        node.is_element()
            && (node.tag_name().namespace() != Some(MAIN_NS)
                || !matches!(node.tag_name().name(), "f" | "v"))
    }) {
        bail!("cell {cell_reference} contains unverified formula content");
    }
    if let Some(value) = value_nodes.first() {
        if value.children().any(|node| node.is_element())
            || value.text().is_some_and(|text| !text.trim().is_empty())
        {
            bail!("cell {cell_reference} has a non-empty formula cache; refusing to clear it");
        }
    }

    let type_attribute = cell
        .attributes()
        .find(|attribute| attribute.namespace().is_none() && attribute.name() == "t");
    if cell
        .attributes()
        .any(|attribute| attribute.name() == "t" && attribute.namespace().is_some())
    {
        bail!("cell {cell_reference} has an unverified namespaced result type");
    }
    if type_attribute.is_some_and(|attribute| attribute.value() != "str") {
        bail!("cell {cell_reference} has an unexpected formula result type");
    }

    let mut ranges = Vec::new();
    if let Some(value) = value_nodes.first() {
        ranges.push(value.range());
    }
    if let Some(attribute) = type_attribute {
        ranges.push(attribute.range());
    }
    ranges.sort_by_key(|range| std::cmp::Reverse(range.start));

    let mut patched = source_xml.to_vec();
    for range in ranges {
        patched = splice_bytes(&patched, range, b"");
    }
    Ok(patched)
}

/// One "copy this template sheet region into that copy sheet region" request.
#[derive(Clone, Debug)]
pub struct CellRegionCopyRequest {
    pub source_sheet: String,
    pub target_sheet: String,
    /// A1 range in both workbooks; the copy keeps the same coordinates.
    pub address: String,
}

#[derive(Debug, Default, PartialEq, Eq)]
pub struct CellRegionCopyReport {
    pub cells_with_formula: usize,
    pub cells_cleared: usize,
    pub worksheet_parts: Vec<String>,
    /// Template sheets copied into the target because copied formulas
    /// reference them and the target did not have them.
    pub sheets_copied: Vec<String>,
}

/// Return the missing template sheets referenced by formula regions, including
/// recursive formula dependencies. The result is intended for a narrowly
/// scoped Openpyxl fallback when direct OOXML copying encounters an unsupported
/// auxiliary-sheet relationship.
pub fn missing_formula_dependency_sheets(
    source: impl AsRef<Path>,
    target: impl AsRef<Path>,
    requests: &[CellRegionCopyRequest],
) -> Result<Vec<String>> {
    ensure!(!requests.is_empty(), "cell region copy requires requests");
    let source = source.as_ref();
    let target = target.as_ref();
    let source_sheets = sheet_parts(source)?;
    let target_sheets = sheet_parts(target)?;
    let mut source_zip = ZipArchive::new(fs::File::open(source)?)?;
    let mut formulas_by_sheet = BTreeMap::new();
    for request in requests {
        let part = source_sheets
            .get(&request.source_sheet)
            .with_context(|| format!("template sheet {:?} is missing", request.source_sheet))?;
        if !formulas_by_sheet.contains_key(&request.source_sheet) {
            let xml = read_zip_part(&mut source_zip, part)?;
            formulas_by_sheet.insert(
                request.source_sheet.clone(),
                region_formulas(&xml, &request.address)?,
            );
        }
    }
    let referenced = collect_referenced_sheets(&formulas_by_sheet);
    let missing = discover_missing_template_sheets(
        &mut source_zip,
        &source_sheets,
        &target_sheets,
        &referenced,
    )?;
    Ok(missing.into_keys().collect())
}

/// Copy formula cells from one workbook's sheet region into another workbook's
/// same-coordinate region, then write the target back in place.
///
/// This exists so an external calculator (Excel/WPS via COM) only ever has to
/// open ONE workbook: the caller copies the template's validation formulas into
/// its own working copy here, in Rust, and then only recalculates that copy.
/// Opening a second workbook doubles the ways an automation host can fail, and
/// the template usually lives in a user directory that may be open in Excel.
///
/// Cached values are dropped, never copied: the template's cache describes the
/// template's data, and the whole point of the following recalculation is to
/// compute this workbook's values. Cells the template does not define are
/// blanked so the region matches the template exactly.
pub fn copy_cell_region(
    source: impl AsRef<Path>,
    output: impl AsRef<Path>,
    requests: &[CellRegionCopyRequest],
) -> Result<CellRegionCopyReport> {
    let source = source.as_ref();
    let output = output.as_ref();
    ensure!(!requests.is_empty(), "cell region copy requires requests");
    let canonical_source = fs::canonicalize(source).context("resolve copy source workbook")?;
    ensure!(
        !output.exists() || fs::canonicalize(output)? != canonical_source,
        "cell region copy never edits the template workbook in place"
    );
    validate_header_package(source).context("validate cell region copy source")?;
    validate_header_package(output).context("validate cell region copy target")?;

    let source_sheets = sheet_parts(source)?;
    let target_sheets = sheet_parts(output)?;
    let mut source_zip = ZipArchive::new(fs::File::open(source)?)?;

    // template sheet -> address -> formula text
    let mut formulas_by_sheet: BTreeMap<String, BTreeMap<String, String>> = BTreeMap::new();
    for request in requests {
        let part = source_sheets
            .get(&request.source_sheet)
            .with_context(|| format!("template sheet {:?} is missing", request.source_sheet))?;
        if formulas_by_sheet.contains_key(&request.source_sheet) {
            continue;
        }
        let xml = read_zip_part(&mut source_zip, part)?;
        formulas_by_sheet.insert(
            request.source_sheet.clone(),
            region_formulas(&xml, &request.address)?,
        );
    }

    let mut replacements: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    let mut report = CellRegionCopyReport::default();

    // Sheets referenced by the formulas that will be copied must exist BEFORE
    // the formulas land, in this same plan. Excel resolves references when it
    // opens the package, so a plan that writes formulas first and sheets later
    // would briefly describe a workbook whose formulas point at nothing; doing
    // sheets first also means a failure here never leaves formulas behind that
    // reference a sheet nobody copied.
    let referenced = collect_referenced_sheets(&formulas_by_sheet);
    let (sheet_changes, sheets_copied) =
        plan_missing_template_sheets(source, output, &source_sheets, &referenced)?;
    report.sheets_copied = sheets_copied;
    replacements.extend(sheet_changes);

    // With every referenced sheet present, copy the region formulas in.
    for request in requests {
        let part = target_sheets
            .get(&request.target_sheet)
            .with_context(|| format!("target sheet {:?} is missing", request.target_sheet))?;
        let formulas = formulas_by_sheet
            .get(&request.source_sheet)
            .expect("template formulas loaded above");
        let current = replacements.get(part).cloned().map(Ok).unwrap_or_else(|| {
            read_zip_part(&mut ZipArchive::new(fs::File::open(output)?)?, part)
        })?;
        let (patched, with_formula, cleared) =
            apply_region_formulas(&current, &request.address, formulas)?;
        report.cells_with_formula += with_formula;
        report.cells_cleared += cleared;
        if !report
            .worksheet_parts
            .iter()
            .any(|existing| existing == part)
        {
            report.worksheet_parts.push(part.clone());
        }
        replacements.insert(part.clone(), patched);
    }
    ensure!(
        !replacements.is_empty(),
        "cell region copy produced no worksheet changes"
    );
    // Invariant: every sheet the copied formulas reference now exists either in
    // the target already or among the sheets just copied. Without this a silent
    // gap would make Excel compute against a missing sheet.
    for name in &referenced {
        ensure!(
            target_sheets
                .keys()
                .any(|existing| existing.eq_ignore_ascii_case(name))
                || report
                    .sheets_copied
                    .iter()
                    .any(|copied| copied.eq_ignore_ascii_case(name)),
            "公式引用了工作表 {name:?}，但它既不在目标工作簿中，也未能从模板补装"
        );
    }

    // Rewrite the package to a sibling temp file, then replace the target:
    // `output` is the caller's working copy and must survive a mid-write
    // failure intact.
    let staged = output.with_file_name(format!(
        ".{}.cell-region-{}.tmp",
        output
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| "workbook".to_owned()),
        std::process::id()
    ));
    let _ = fs::remove_file(&staged);
    let write_result = (|| -> Result<()> {
        let mut input_zip = ZipArchive::new(fs::File::open(output)?)?;
        let file = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&staged)
            .with_context(|| format!("create staged copy {}", staged.display()))?;
        let mut writer = ZipWriter::new(file);
        let mut written = BTreeSet::new();
        for index in 0..input_zip.len() {
            let mut entry = input_zip.by_index(index)?;
            let name = entry.name().to_owned();
            let mut options = SimpleFileOptions::default().compression_method(entry.compression());
            if let Some(modified) = entry.last_modified() {
                options = options.last_modified_time(modified);
            }
            if entry.is_dir() {
                writer.add_directory(name, options)?;
            } else {
                writer.start_file(&name, options)?;
                written.insert(name.clone());
                if is_worksheet_part(&name) {
                    // The working copy is disposable: normalize every
                    // worksheet's formulas so Excel can parse sheet references
                    // written unquoted (copied template sheets included). The
                    // caller's own file is never touched.
                    let bytes = match replacements.get(&name) {
                        Some(bytes) => bytes.clone(),
                        None => {
                            let mut bytes = Vec::new();
                            entry.read_to_end(&mut bytes)?;
                            bytes
                        }
                    };
                    writer.write_all(&normalize_worksheet_formula_references(&bytes)?)?;
                } else if let Some(bytes) = replacements.get(&name) {
                    writer.write_all(bytes)?;
                } else {
                    std::io::copy(&mut entry, &mut writer)?;
                }
            }
        }
        // Parts that exist only in the replacements (copied template sheets).
        for (name, bytes) in &replacements {
            if written.contains(name) {
                continue;
            }
            writer.start_file(name, SimpleFileOptions::default())?;
            if is_worksheet_part(name) {
                writer.write_all(&normalize_worksheet_formula_references(bytes)?)?;
            } else {
                writer.write_all(bytes)?;
            }
        }
        writer.finish()?;
        Ok(())
    })()
    .context("write target workbook with copied cell regions");
    if let Err(error) = write_result {
        let _ = fs::remove_file(&staged);
        return Err(error);
    }
    fs::rename(&staged, output)
        .with_context(|| format!("replace {} with copied cell regions", output.display()))?;
    Ok(report)
}

/// Sheet names referenced by the copied formulas (`Sheet!A1`, `'My Sheet'!A1`).
fn collect_referenced_sheets(
    formulas_by_sheet: &BTreeMap<String, BTreeMap<String, String>>,
) -> BTreeSet<String> {
    let mut referenced = BTreeSet::new();
    for formulas in formulas_by_sheet.values() {
        for text in formulas.values() {
            for name in sheet_references_in_formula(text)
                .into_iter()
                .filter(|name| !is_external_sheet_reference(name))
            {
                referenced.insert(name);
            }
        }
    }
    referenced
}

fn is_external_sheet_reference(name: &str) -> bool {
    name.contains('[') || name.contains(']')
}

fn is_worksheet_part(name: &str) -> bool {
    name.starts_with("xl/worksheets/") && name.ends_with(".xml") && !name.contains("/_rels/")
}

/// Extract `Sheet!` prefixes from a formula, honouring quoted names.
fn sheet_references_in_formula(text: &str) -> Vec<String> {
    let characters = text.chars().collect::<Vec<_>>();
    let mut names = Vec::new();
    let mut index = 0usize;
    let mut in_string = false;
    while index < characters.len() {
        let current = characters[index];
        if in_string {
            if current == '"' {
                if characters.get(index + 1) == Some(&'"') {
                    index += 2;
                    continue;
                }
                in_string = false;
            }
            index += 1;
            continue;
        }
        if current == '"' {
            in_string = true;
            index += 1;
            continue;
        }
        if current == '\'' {
            // Quoted sheet name: 'My Sheet'!A1
            let start = index + 1;
            let mut end = start;
            while end < characters.len() && characters[end] != '\'' {
                end += 1;
            }
            if end < characters.len() && characters.get(end + 1) == Some(&'!') {
                names.push(characters[start..end].iter().collect());
            }
            index = end + 1;
            continue;
        }
        if current.is_alphanumeric() || current == '_' {
            let start = index;
            while index < characters.len()
                && (characters[index].is_alphanumeric() || characters[index] == '_')
            {
                index += 1;
            }
            if characters.get(index) == Some(&'!') {
                names.push(characters[start..index].iter().collect());
            }
            continue;
        }
        index += 1;
    }
    names
}

/// Lowest `sheetN.xml` index not yet used by either package.
fn next_free_sheet_index(used_parts: &BTreeSet<String>) -> u32 {
    let mut index = 1u32;
    loop {
        if !used_parts.contains(&format!("xl/worksheets/sheet{index}.xml")) {
            return index;
        }
        index += 1;
    }
}

/// Copy a worksheet part together with the parts its relationships point at.
///
/// A worksheet commonly owns `commentsN.xml` and `vmlDrawingN.vml`; the sheet
/// XML refers to them by relationship id, so copying the sheet alone leaves
/// dangling references and Excel refuses to open the package. Relationship
/// targets are rewritten to freshly allocated part names.
fn copy_worksheet_with_dependencies(
    source_zip: &mut ZipArchive<fs::File>,
    source_part: &str,
    new_part: &str,
    sheet_xml: &[u8],
    used_parts: &mut BTreeSet<String>,
    sheet_index: u32,
) -> Result<(Vec<u8>, Vec<(String, Vec<u8>)>)> {
    let source_rels_part = relationships_part(source_part);
    let Ok(rels_xml) = read_zip_part(source_zip, &source_rels_part) else {
        // No relationships: nothing to carry over.
        return Ok((sheet_xml.to_vec(), Vec::new()));
    };
    let rels_text = std::str::from_utf8(&rels_xml).context("worksheet rels must be UTF-8")?;
    let document = Document::parse(rels_text).context("parse worksheet relationships")?;

    let source_dir = source_part
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    let mut new_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    let mut extra_parts = Vec::new();
    // comments and vml drawings are numbered per sheet in practice; reuse the
    // sheet's own index so a copied sheet keeps a readable, unique name.
    let mut comments_index: Option<u32> = None;
    for relationship in document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "Relationship")
    {
        let Some(id) = relationship.attribute("Id") else {
            continue;
        };
        let Some(kind) = relationship.attribute("Type") else {
            continue;
        };
        let Some(target) = relationship.attribute("Target") else {
            continue;
        };
        if relationship.attribute("TargetMode") == Some("External") {
            bail!(
                "OPENPYXL_FALLBACK_REQUIRED: auxiliary worksheet {source_part:?} has external relationship {kind:?}"
            );
        }
        let resolved = normalize_part(source_dir, target);
        let bytes = read_zip_part(source_zip, &resolved).with_context(|| {
            format!(
                "OPENPYXL_FALLBACK_REQUIRED: auxiliary worksheet relationship target {resolved:?} is unavailable"
            )
        })?;
        let file_name = resolved.rsplit('/').next().unwrap_or(&resolved);
        let (prefix, suffix) = match (
            file_name
                .strip_suffix(".xml")
                .and_then(|stem| stem.strip_prefix("comments")),
            file_name
                .strip_suffix(".vml")
                .and_then(|stem| stem.strip_prefix("vmlDrawing")),
        ) {
            (Some(_), _) => ("comments", ".xml"),
            (_, Some(_)) => ("vmlDrawing", ".vml"),
            _ => {
                bail!(
                    "OPENPYXL_FALLBACK_REQUIRED: auxiliary worksheet {source_part:?} uses relationship {kind:?} to {resolved:?}"
                );
            }
        };
        let new_index = if prefix == "comments" && comments_index.is_none() {
            comments_index = Some(sheet_index);
            sheet_index
        } else {
            allocate_indexed_part(used_parts, &resolved, prefix, suffix)
        };
        // Keep the source directory (e.g. `xl/drawings`), otherwise the
        // relationship target points at a part that does not exist.
        let source_dir_of_part = resolved
            .rsplit_once('/')
            .map(|(dir, _)| dir)
            .unwrap_or("xl");
        let new_resolved = format!("{source_dir_of_part}/{prefix}{new_index}{suffix}");
        used_parts.insert(new_resolved.clone());
        new_rels.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{}" Target="{}"/>"#,
            escape_xml_text(kind),
            escape_xml_text(&relative_target(new_part, &new_resolved))
        ));
        extra_parts.push((new_resolved, bytes));
    }
    new_rels.push_str("</Relationships>");
    if extra_parts.is_empty() {
        return Ok((sheet_xml.to_vec(), Vec::new()));
    }

    let new_rels_part = relationships_part(new_part);
    used_parts.insert(new_rels_part.clone());
    extra_parts.push((new_rels_part, new_rels.into_bytes()));
    Ok((sheet_xml.to_vec(), extra_parts))
}

/// Allocate a copy name for `resolved` that no package part uses yet.
fn allocate_indexed_part(
    used_parts: &BTreeSet<String>,
    resolved: &str,
    prefix: &str,
    suffix: &str,
) -> u32 {
    let dir = resolved
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("xl");
    let mut index = 1u32;
    loop {
        let candidate = format!("{dir}/{prefix}{index}{suffix}");
        if !used_parts.contains(&candidate) {
            return index;
        }
        index += 1;
    }
}

/// Relationship target from a part in `from_part`'s directory to `to_part`.
fn relative_target(from_part: &str, to_part: &str) -> String {
    let from_dir = from_part.rsplit_once('/').map(|(dir, _)| dir).unwrap_or("");
    let to_dir = to_part.rsplit_once('/').map(|(dir, _)| dir).unwrap_or("");
    let file = to_part
        .rsplit_once('/')
        .map(|(_, file)| file)
        .unwrap_or(to_part);
    if from_dir == to_dir {
        return file.to_owned();
    }
    // Targets are relative to the part owning the rels file. Both live under
    // `xl/`, so reaching `xl/` costs one `../`; any extra depth on either side
    // adds to that.
    let from_depth = from_dir.split('/').filter(|s| !s.is_empty()).count();
    let to_depth = to_dir.split('/').filter(|s| !s.is_empty()).count();
    let climb = from_depth.saturating_sub(1).max(to_depth.saturating_sub(1));
    let target_dir = to_dir
        .split('/')
        .filter(|s| !s.is_empty())
        .skip(1)
        .collect::<Vec<_>>()
        .join("/");
    let mut target = "../".repeat(climb);
    if !target_dir.is_empty() {
        target.push_str(&target_dir);
        target.push('/');
    }
    target.push_str(file);
    target
}

/// Rewrite `<c t="s"><v>N</v></c>` cells to inline strings using the source
/// package's shared-string table.
///
/// Shared-string indexes are package-local. A worksheet lifted from another
/// workbook must therefore stop referring to them; inlining is the smallest
/// change that keeps the copied sheet self-contained.
fn inline_shared_strings(
    source_zip: &mut ZipArchive<fs::File>,
    sheet_xml: &[u8],
) -> Result<Vec<u8>> {
    let strings = read_shared_strings(source_zip)?;
    if strings.is_empty() {
        return Ok(sheet_xml.to_vec());
    }
    let xml = std::str::from_utf8(sheet_xml).context("worksheet XML must be UTF-8")?;
    let document = Document::parse(xml).context("parse worksheet for shared-string inlining")?;
    let mut edits = Vec::<(Range<usize>, Vec<u8>)>::new();
    for cell in document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "c"
            && node.attribute("t") == Some("s")
    }) {
        let value = cell
            .children()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "v"
            })
            .and_then(|node| node.text())
            .unwrap_or_default()
            .trim();
        let Ok(index) = value.parse::<usize>() else {
            continue;
        };
        let Some(text) = strings.get(index) else {
            bail!(
                "工作表单元格 {:?} 的共享字符串索引 {index} 超出模板字符串表范围（{} 项）",
                cell.attribute("r"),
                strings.len()
            );
        };
        let range = cell.range();
        let fragment = &xml[range.clone()];
        let open_end = fragment
            .find('>')
            .context("cell element has no closing angle bracket")?;
        let open_tag = fragment[..open_end].trim_end_matches('/');
        // Swap `t="s"` for `t="inlineStr"` so the value is read from `<is>`.
        let swapped = open_tag.replace("t=\"s\"", "t=\"inlineStr\"");
        if swapped == open_tag {
            bail!("单元格 {:?} 缺少 t=\"s\" 属性", cell.attribute("r"));
        }
        let open_tag = swapped;
        let replacement = format!(
            "{open_tag}><is><t xml:space=\"preserve\">{}</t></is></c>",
            escape_xml_text(text)
        );
        edits.push((range, replacement.into_bytes()));
    }
    if edits.is_empty() {
        return Ok(sheet_xml.to_vec());
    }
    let mut patched = sheet_xml.to_vec();
    edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
    for (range, replacement) in edits {
        patched = splice_bytes(&patched, range, &replacement);
    }
    Ok(patched)
}

/// Plain text of every `<si>` in the package's shared-string table.
fn read_shared_strings(source_zip: &mut ZipArchive<fs::File>) -> Result<Vec<String>> {
    let Ok(xml) = read_zip_part(source_zip, "xl/sharedStrings.xml") else {
        return Ok(Vec::new());
    };
    let xml = std::str::from_utf8(&xml).context("sharedStrings.xml must be UTF-8")?;
    let document = Document::parse(xml).context("parse sharedStrings.xml")?;
    Ok(document
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "si"
        })
        .map(|node| {
            // Concatenate every <t> in the item; <r> runs and phonetic <rPh>
            // children both contribute their text nodes.
            node.descendants()
                .filter(|child| {
                    child.is_element()
                        && child.tag_name().namespace() == Some(MAIN_NS)
                        && child.tag_name().name() == "t"
                })
                .filter_map(|child| child.text())
                .collect::<String>()
        })
        .collect())
}

/// Copy template sheets into the target when copied formulas reference them
/// and the target lacks them. Formula-bearing helper sheets are copied too;
/// their own references are walked recursively so the calculation workbook
/// receives the full missing dependency chain before it is recalculated.
fn plan_missing_template_sheets(
    source: &Path,
    output: &Path,
    source_sheets: &BTreeMap<String, String>,
    referenced: &BTreeSet<String>,
) -> Result<(BTreeMap<String, Vec<u8>>, Vec<String>)> {
    let target_sheets = sheet_parts(output)?;
    let mut changes = BTreeMap::new();
    let mut source_zip = ZipArchive::new(fs::File::open(source)?)?;
    let mut target_zip = ZipArchive::new(fs::File::open(output)?)?;
    let missing = discover_missing_template_sheets(
        &mut source_zip,
        source_sheets,
        &target_sheets,
        referenced,
    )?;
    if missing.is_empty() {
        return Ok((changes, Vec::new()));
    }
    let mut workbook = read_zip_part(&mut target_zip, "xl/workbook.xml")?;
    let mut workbook_rels = read_zip_part(&mut target_zip, "xl/_rels/workbook.xml.rels")?;
    let mut content_types = read_zip_part(&mut target_zip, "[Content_Types].xml")?;

    let mut copied = Vec::new();
    // Parts already taken by either package, so copied sheets never overwrite
    // an existing comments/vml part.
    let mut used_parts = source_zip
        .file_names()
        .chain(target_zip.file_names())
        .map(str::to_owned)
        .collect::<BTreeSet<_>>();
    for (name, part) in missing {
        let sheet_xml = read_zip_part(&mut source_zip, &part)
            .with_context(|| format!("read template sheet {name:?}"))?;
        // The sheet's `<c t="s"><v>N</v></c>` indexes point into the TEMPLATE's
        // sharedStrings table. The target has its own table with different
        // entries, so leaving them alone makes every text cell resolve to the
        // wrong string (or out of range, which Excel rejects as a corrupt
        // package). Inline the text instead: the copied sheet then carries its
        // own strings and needs no shared-table surgery.
        let sheet_xml = inline_shared_strings(&mut source_zip, &sheet_xml)
            .with_context(|| format!("inline shared strings for template sheet {name:?}"))?;
        // A fresh part name that cannot collide with the target's.
        let next_index = next_free_sheet_index(&used_parts);
        let new_part = format!("xl/worksheets/sheet{next_index}.xml");
        used_parts.insert(new_part.clone());

        // The worksheet's own relationships (comments, VML drawing for comment
        // boxes) must come along, otherwise the copied sheet keeps `r:id`
        // references that resolve to nothing and Excel reports the package as
        // corrupt. Those targets are remapped to fresh part names as well.
        let (sheet_xml, extra_parts) = copy_worksheet_with_dependencies(
            &mut source_zip,
            &part,
            &new_part,
            &sheet_xml,
            &mut used_parts,
            next_index,
        )
        .with_context(|| format!("copy template sheet {name:?} with its dependencies"))?;
        changes.insert(new_part.clone(), sheet_xml);
        for (extra_name, bytes) in extra_parts {
            // Every copied part must be declared in [Content_Types].xml. A
            // source without any comments never declared the comments/vml
            // content types, so silently adding those parts produced a
            // package Excel rejects while WPS opens it anyway.
            content_types = append_copied_part_content_type(&content_types, &extra_name)?;
            changes.insert(extra_name, bytes);
        }

        let relationship_id = format!("rId{}", next_relationship_id(&workbook_rels)?);
        let sheet_id = next_sheet_id(&workbook)?;
        workbook = append_sheet_entry(&workbook, &name, sheet_id, &relationship_id)?;
        workbook_rels = append_worksheet_relationship(&workbook_rels, &relationship_id, &new_part)?;
        content_types = append_content_type_override(&content_types, &new_part)?;
        copied.push(name);
    }
    changes.insert("xl/workbook.xml".to_owned(), workbook);
    changes.insert("xl/_rels/workbook.xml.rels".to_owned(), workbook_rels);
    changes.insert("[Content_Types].xml".to_owned(), content_types);
    Ok((changes, copied))
}

fn discover_missing_template_sheets(
    source_zip: &mut ZipArchive<fs::File>,
    source_sheets: &BTreeMap<String, String>,
    target_sheets: &BTreeMap<String, String>,
    referenced: &BTreeSet<String>,
) -> Result<BTreeMap<String, String>> {
    let mut pending = referenced.iter().cloned().collect::<BTreeSet<_>>();
    let mut missing = BTreeMap::<String, String>::new();
    // Resolve the dependency graph first. A formula-bearing auxiliary sheet is
    // not a terminal dependency: its formulas may reference more template
    // sheets that must also be installed.
    while let Some(name) = pending.pop_first() {
        if target_sheets
            .keys()
            .any(|existing| existing.eq_ignore_ascii_case(&name))
            || missing
                .keys()
                .any(|existing| existing.eq_ignore_ascii_case(&name))
        {
            continue;
        }
        let Some(part) = source_sheets
            .iter()
            .find(|(sheet, _)| sheet.eq_ignore_ascii_case(&name))
            .map(|(_, part)| part.clone())
        else {
            bail!("公式引用了工作表 {name:?}，但模板中也没有这个表，无法补装");
        };
        let sheet_xml = read_zip_part(source_zip, &part)
            .with_context(|| format!("read template sheet {name:?}"))?;
        for formula in all_sheet_formula_texts(&sheet_xml)
            .with_context(|| format!("读取辅助工作表 {name:?} 的公式依赖"))?
            .values()
        {
            pending.extend(
                sheet_references_in_formula(formula)
                    .into_iter()
                    .filter(|reference| !is_external_sheet_reference(reference)),
            );
        }
        missing.insert(name, part);
    }
    Ok(missing)
}

/// Formula text for every formula cell on a sheet, expanding shared formulas
/// so dependency discovery sees the references used by each cell.
fn all_sheet_formula_texts(sheet_xml: &[u8]) -> Result<BTreeMap<String, String>> {
    let xml = std::str::from_utf8(sheet_xml).context("worksheet XML must be UTF-8")?;
    let document = Document::parse(xml).context("parse worksheet formula dependencies")?;
    let shared = shared_formula_groups(&document)?;
    let mut formulas = BTreeMap::new();
    for cell in formula_cells(&document) {
        let Some(reference) = cell.attribute("r") else {
            continue;
        };
        let Some(formula) = formula_node(cell) else {
            continue;
        };
        let position = parse_a1_cell(reference)
            .with_context(|| format!("parse auxiliary formula cell {reference}"))?;
        let text = if formula.attribute("t") == Some("shared") {
            effective_formula_text(formula, reference, position, &shared)?
        } else {
            // Array/data-table formulas are not interpreted here. Their text
            // still participates in dependency discovery and their XML is
            // copied verbatim; the external calculator remains responsible
            // for evaluating them.
            formula.text().unwrap_or_default().trim().to_owned()
        };
        ensure!(
            !text.is_empty(),
            "辅助工作表公式 {reference} 缺少公式文本，无法发现依赖"
        );
        formulas.insert(reference.to_owned(), text);
    }
    Ok(formulas)
}

fn next_relationship_id(workbook_rels: &[u8]) -> Result<u32> {
    let document = Document::parse(std::str::from_utf8(workbook_rels)?)
        .context("parse workbook relationships")?;
    Ok(document
        .descendants()
        .filter_map(|node| node.attribute("Id"))
        .filter_map(|id| id.strip_prefix("rId"))
        .filter_map(|index| index.parse::<u32>().ok())
        .max()
        .unwrap_or(0)
        + 1)
}

fn next_sheet_id(workbook: &[u8]) -> Result<u32> {
    let document =
        Document::parse(std::str::from_utf8(workbook)?).context("parse target workbook.xml")?;
    Ok(document
        .descendants()
        .filter(|node| node.tag_name().name() == "sheet")
        .filter_map(|node| node.attribute("sheetId"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .unwrap_or(0)
        + 1)
}

fn append_sheet_entry(
    workbook: &[u8],
    name: &str,
    sheet_id: u32,
    relationship_id: &str,
) -> Result<Vec<u8>> {
    let xml = std::str::from_utf8(workbook)?;
    let anchor = xml
        .rfind("</sheets>")
        .context("workbook.xml has no </sheets>")?;
    let entry = format!(
        r#"<sheet name="{}" sheetId="{sheet_id}" r:id="{relationship_id}"/>"#,
        escape_xml_text(name)
    );
    let mut patched = xml.to_owned();
    patched.insert_str(anchor, &entry);
    Ok(patched.into_bytes())
}

fn append_worksheet_relationship(
    workbook_rels: &[u8],
    relationship_id: &str,
    part: &str,
) -> Result<Vec<u8>> {
    let xml = std::str::from_utf8(workbook_rels)?;
    let anchor = xml
        .rfind("</Relationships>")
        .context("workbook relationships have no closing tag")?;
    let target = part.strip_prefix("xl/").unwrap_or(part);
    let entry = format!(
        r#"<Relationship Id="{relationship_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="{target}"/>"#
    );
    let mut patched = xml.to_owned();
    patched.insert_str(anchor, &entry);
    Ok(patched.into_bytes())
}

fn append_content_type_override(content_types: &[u8], part: &str) -> Result<Vec<u8>> {
    let xml = std::str::from_utf8(content_types)?;
    let anchor = xml
        .rfind("</Types>")
        .context("[Content_Types].xml has no closing tag")?;
    let entry = format!(
        r#"<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
    );
    let mut patched = xml.to_owned();
    patched.insert_str(anchor, &entry);
    Ok(patched.into_bytes())
}

/// Content-type override for the non-worksheet parts that travel with a
/// copied template sheet (comments, VML comment drawings). Their relationship
/// parts (`_rels/*.rels`) are covered by the package-wide `rels` Default.
fn append_copied_part_content_type(content_types: &[u8], part: &str) -> Result<Vec<u8>> {
    let xml = std::str::from_utf8(content_types)?;
    let content_type = match part {
        name if name.starts_with("xl/comments") && name.ends_with(".xml") => {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
        }
        name if name.starts_with("xl/drawings/") && name.ends_with(".vml") => {
            "application/vnd.openxmlformats-officedocument.vmlDrawing"
        }
        name if name.contains("/_rels/") && name.ends_with(".rels") => {
            return Ok(content_types.to_owned());
        }
        _ => anyhow::bail!("copied template part {part:?} has no known content type"),
    };
    let anchor = xml
        .rfind("</Types>")
        .context("[Content_Types].xml has no closing tag")?;
    let entry = format!(r#"<Override PartName="/{part}" ContentType="{content_type}"/>"#);
    let mut patched = xml.to_owned();
    patched.insert_str(anchor, &entry);
    Ok(patched.into_bytes())
}

/// Formula text for every cell the template defines inside `address`.
fn region_formulas(sheet_xml: &[u8], address: &str) -> Result<BTreeMap<String, String>> {
    let (row1, col1, row2, col2) = parse_a1_range(address)?;
    let xml = std::str::from_utf8(sheet_xml).context("worksheet XML must be UTF-8")?;
    let document = Document::parse(xml).context("parse template worksheet XML")?;
    let shared = shared_formula_groups(&document)?;
    let mut formulas = BTreeMap::new();
    for cell in formula_cells(&document) {
        let Some(reference) = cell.attribute("r") else {
            continue;
        };
        let (row, column) = match parse_a1_cell(reference) {
            Ok(position) => position,
            Err(_) => continue,
        };
        if row < row1 || row > row2 || column < col1 || column > col2 {
            continue;
        }
        let Some(formula) = formula_node(cell) else {
            continue;
        };
        let text = effective_formula_text(formula, reference, (row, column), &shared)?;
        formulas.insert(reference.to_owned(), quote_unquoted_sheet_references(&text));
    }
    Ok(formulas)
}

/// Wrap bare sheet-name references whose names contain characters Microsoft
/// Excel refuses to parse unquoted (here: full-width parentheses) in single
/// quotes. Quoting is semantically neutral and always accepted; WPS tolerates
/// the unquoted form, Excel rejects the whole workbook at `Workbooks.Open`.
/// String literals, already-quoted names and external `[n]Sheet!` prefixes are
/// left untouched. The transformation is idempotent.
fn quote_unquoted_sheet_references(text: &str) -> String {
    let characters: Vec<char> = text.chars().collect();
    let is_name_char = |c: char| c.is_alphanumeric() || c == '_' || c == '.' || c as u32 > 0x7F;
    let needs_quotes = |name: &str| {
        name.chars()
            .any(|c| !(c.is_ascii_alphanumeric() || c == '_' || c == '.'))
    };
    let mut out = String::with_capacity(text.len());
    let mut index = 0usize;
    let mut in_string = false;
    while index < characters.len() {
        let current = characters[index];
        if in_string {
            out.push(current);
            if current == '"' {
                in_string = false;
            }
            index += 1;
            continue;
        }
        if current == '"' {
            in_string = true;
            out.push(current);
            index += 1;
            continue;
        }
        if current == '\'' {
            // Already-quoted name: copy through the closing quote, plus the
            // trailing `!` when this is a sheet reference.
            out.push(current);
            index += 1;
            while index < characters.len() {
                out.push(characters[index]);
                if characters[index] == '\'' {
                    index += 1;
                    break;
                }
                index += 1;
            }
            if index < characters.len() && characters[index] == '!' {
                out.push('!');
                index += 1;
            }
            continue;
        }
        if !is_name_char(current) {
            out.push(current);
            index += 1;
            continue;
        }
        let start = index;
        while index < characters.len() && is_name_char(characters[index]) {
            index += 1;
        }
        let name: String = characters[start..index].iter().collect();
        if characters.get(index) == Some(&'!') {
            // External workbook references ([1]Sheet!) are never renamed here;
            // the `[` belongs to a preceding token and the name itself stays.
            let preceded_by_bracket = out.ends_with(']');
            if !preceded_by_bracket && needs_quotes(&name) {
                out.push('\'');
                out.push_str(&name);
                out.push('\'');
                out.push('!');
                index += 1;
                continue;
            }
        }
        out.push_str(&name);
    }
    out
}

/// Normalize every formula's sheet references in a worksheet part so that
/// Microsoft Excel accepts the workbook. Used on the disposable working copy
/// the validation flow hands to COM; the caller's own file is never edited.
fn normalize_worksheet_formula_references(sheet_xml: &[u8]) -> Result<Vec<u8>> {
    let xml = std::str::from_utf8(sheet_xml).context("worksheet XML must be UTF-8")?;
    let Some(document) = Document::parse(xml).ok() else {
        return Ok(sheet_xml.to_vec());
    };
    let mut edits: Vec<(std::ops::Range<usize>, Vec<u8>)> = Vec::new();
    for node in document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "f"
    }) {
        let Some(original) = node.text() else {
            continue;
        };
        let rewritten = quote_unquoted_sheet_references(original);
        if rewritten == original {
            continue;
        }
        let range = node.range();
        let open_end = xml[range.clone()]
            .find('>')
            .context("formula element has no closing angle bracket")?;
        let open_tag = xml[range.start..range.start + open_end + 1].to_owned();
        let replacement = format!("{open_tag}{}</f>", escape_formula_text(&rewritten));
        edits.push((range, replacement.into_bytes()));
    }
    if edits.is_empty() {
        return Ok(sheet_xml.to_vec());
    }
    let mut result = xml.as_bytes().to_vec();
    for (range, replacement) in edits.into_iter().rev() {
        result.splice(range, replacement);
    }
    Ok(result)
}

fn formula_cells<'a>(
    document: &'a Document<'a>,
) -> impl Iterator<Item = roxmltree::Node<'a, 'a>> + 'a {
    document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "c"
    })
}

fn formula_node<'a>(cell: roxmltree::Node<'a, 'a>) -> Option<roxmltree::Node<'a, 'a>> {
    cell.children().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "f"
    })
}

/// Anchor cell and formula text of every shared-formula master, keyed by `si`.
///
/// Excel stores one master with the formula and `t="shared" ref=… si=N`, and
/// the remaining cells of the group carry only `t="shared" si=N`. Their formula
/// is the master's with relative references shifted by the cell offset.
fn shared_formula_groups(document: &Document<'_>) -> Result<BTreeMap<u32, ((u32, u32), String)>> {
    let mut groups = BTreeMap::new();
    for cell in formula_cells(document) {
        let Some(reference) = cell.attribute("r") else {
            continue;
        };
        let Some(formula) = formula_node(cell) else {
            continue;
        };
        if formula.attribute("t") != Some("shared") {
            continue;
        }
        let Some(si) = formula
            .attribute("si")
            .and_then(|value| value.parse::<u32>().ok())
        else {
            continue;
        };
        // A dependent carries no text; only the master defines the group.
        let text = formula.text().unwrap_or_default().trim();
        if text.is_empty() {
            continue;
        }
        let (row, column) = parse_a1_cell(reference)?;
        groups.insert(si, ((row, column), text.to_owned()));
    }
    Ok(groups)
}

/// Resolve one cell's formula, expanding shared-formula dependents.
fn effective_formula_text(
    formula: roxmltree::Node<'_, '_>,
    reference: &str,
    position: (u32, u32),
    shared: &BTreeMap<u32, ((u32, u32), String)>,
) -> Result<String> {
    match formula.attribute("t") {
        None => {
            let text = formula.text().unwrap_or_default().trim();
            ensure!(!text.is_empty(), "校验区域公式 {reference} 缺少公式文本");
            Ok(text.to_owned())
        }
        Some("shared") => {
            let si = formula
                .attribute("si")
                .and_then(|value| value.parse::<u32>().ok())
                .with_context(|| format!("校验区域共享公式 {reference} 缺少 si"))?;
            let own = formula.text().unwrap_or_default().trim();
            if !own.is_empty() {
                // Master cell: its own text is already relative to itself.
                return Ok(own.to_owned());
            }
            let (anchor, master_text) = shared.get(&si).with_context(|| {
                format!(
                    "校验区域公式 {reference} 引用了共享组 si={si}，\
                     但模板中找不到该组的主公式"
                )
            })?;
            translate_shared_formula(master_text, *anchor, position)
                .with_context(|| format!("展开校验区域共享公式 {reference}（si={si}）失败"))
        }
        Some(other) => bail!(
            "校验区域公式 {reference} 的类型为 {other:?}（数组公式等），暂不支持复制；\
             请在模板中把该公式改为普通公式后重试"
        ),
    }
}

/// Shift every relative A1 reference in `text` by the offset between two cells.
fn translate_shared_formula(text: &str, from: (u32, u32), to: (u32, u32)) -> Result<String> {
    let delta_row = i64::from(to.0) - i64::from(from.0);
    let delta_col = i64::from(to.1) - i64::from(from.1);
    if delta_row == 0 && delta_col == 0 {
        return Ok(text.to_owned());
    }
    let characters = text.chars().collect::<Vec<_>>();
    let mut output = String::with_capacity(text.len());
    let mut index = 0usize;
    let mut in_string = false;
    while index < characters.len() {
        let current = characters[index];
        if in_string {
            output.push(current);
            if current == '"' {
                // `""` is an escaped quote inside a string literal.
                if characters.get(index + 1) == Some(&'"') {
                    output.push('"');
                    index += 2;
                    continue;
                }
                in_string = false;
            }
            index += 1;
            continue;
        }
        if current == '"' {
            in_string = true;
            output.push(current);
            index += 1;
            continue;
        }
        if current == '$' || current.is_ascii_alphabetic() {
            if let Some((consumed, replacement)) =
                shift_a1_reference(&characters, index, delta_row, delta_col)?
            {
                output.push_str(&replacement);
                index += consumed;
                continue;
            }
        }
        output.push(current);
        index += 1;
    }
    Ok(output)
}

/// Match an A1 reference at `start` and return it shifted, with its length.
fn shift_a1_reference(
    characters: &[char],
    start: usize,
    delta_row: i64,
    delta_col: i64,
) -> Result<Option<(usize, String)>> {
    // A reference must start at a token boundary so names such as `MyName1`
    // or the digits in `LOG10` are never rewritten.
    if start > 0 {
        let previous = characters[start - 1];
        if previous.is_ascii_alphanumeric() || previous == '_' || previous == '.' {
            return Ok(None);
        }
    }
    let mut index = start;
    let column_absolute = characters.get(index) == Some(&'$');
    if column_absolute {
        index += 1;
    }
    let column_start = index;
    while characters
        .get(index)
        .is_some_and(|character| character.is_ascii_alphabetic())
        && index - column_start < 4
    {
        index += 1;
    }
    let column_end = index;
    let column_letters = column_end - column_start;
    if column_letters == 0 || column_letters > 3 {
        return Ok(None);
    }
    let row_absolute = characters.get(index) == Some(&'$');
    if row_absolute {
        index += 1;
    }
    let row_start = index;
    while characters
        .get(index)
        .is_some_and(|character| character.is_ascii_digit())
        && index - row_start < 8
    {
        index += 1;
    }
    if index == row_start {
        return Ok(None);
    }
    // A following `(` means the letters are a function name, not a cell.
    if characters.get(index) == Some(&'(') {
        return Ok(None);
    }
    if characters
        .get(index)
        .is_some_and(|character| character.is_ascii_alphanumeric() || *character == '_')
    {
        return Ok(None);
    }

    let letters = characters[column_start..column_end]
        .iter()
        .collect::<String>();
    let mut column = 0i64;
    for byte in letters.bytes() {
        column = column * 26 + i64::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    let row = characters[row_start..index]
        .iter()
        .collect::<String>()
        .parse::<i64>()
        .context("shared formula row number overflows")?;
    let new_column = if column_absolute {
        column
    } else {
        column + delta_col
    };
    let new_row = if row_absolute { row } else { row + delta_row };
    ensure!(
        new_column >= 1 && new_row >= 1,
        "共享公式的引用平移后越界（列 {new_column}，行 {new_row}）"
    );
    let mut replacement = String::new();
    if column_absolute {
        replacement.push('$');
    }
    replacement.push_str(&column_letters_of(new_column as u32)?);
    if row_absolute {
        replacement.push('$');
    }
    replacement.push_str(&new_row.to_string());
    Ok(Some((index - start, replacement)))
}

fn column_letters_of(mut column: u32) -> Result<String> {
    ensure!(column > 0, "column index is 1-based");
    let mut letters = String::new();
    while column > 0 {
        let remainder = ((column - 1) % 26) as u8;
        letters.insert(0, (b'A' + remainder) as char);
        column = (column - 1) / 26;
    }
    Ok(letters)
}

/// Replace the formulas inside `address` on one sheet.
fn apply_region_formulas(
    sheet_xml: &[u8],
    address: &str,
    formulas: &BTreeMap<String, String>,
) -> Result<(Vec<u8>, usize, usize)> {
    let (row1, col1, row2, col2) = parse_a1_range(address)?;
    let xml = std::str::from_utf8(sheet_xml).context("worksheet XML must be UTF-8")?;
    let document = Document::parse(xml).context("parse target worksheet XML")?;

    let mut edits: Vec<(Range<usize>, Vec<u8>)> = Vec::new();
    let mut seen = BTreeSet::new();
    for cell in document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "c"
    }) {
        let Some(reference) = cell.attribute("r") else {
            continue;
        };
        let (row, column) = match parse_a1_cell(reference) {
            Ok(position) => position,
            Err(_) => continue,
        };
        if row < row1 || row > row2 || column < col1 || column > col2 {
            continue;
        }
        seen.insert(reference.to_owned());
        let formula = formulas.get(reference);
        // The template's style indexes belong to the template package, so keep
        // the target cell's own style and swap only its content.
        let style = cell
            .attribute_node("s")
            .map(|attribute| attribute.value().to_owned());
        let cell_range = cell.range();
        let fragment = &xml[cell_range.clone()];
        let open_end = fragment
            .find('>')
            .context("cell element has no closing angle bracket")?;
        let self_closing = fragment.as_bytes()[open_end - 1] == b'/';
        let mut open_tag = fragment[..open_end].trim_end_matches('/').to_owned();
        // Drop the cached type so the recalculation decides the real type.
        if let Some(position) = open_tag.find(" t=") {
            let end = open_tag[position + 1..]
                .find(|character: char| character == ' ' || character == '>')
                .map(|offset| position + 1 + offset)
                .unwrap_or(open_tag.len());
            open_tag.replace_range(position..end, "");
        }
        let replacement = match formula {
            Some(text) => format!(
                "{open_tag}><f>{text}</f></c>",
                text = escape_formula_text(text)
            ),
            None => format!("{open_tag}/>"),
        };
        let _ = (style, self_closing);
        edits.push((cell_range, replacement.into_bytes()));
    }

    // Cells the template defines but the target does not have. Excel's
    // Range.Copy (the previous implementation) materializes these, so do the
    // same: insert a `<c>` into the matching row, creating the row if needed.
    // Source workbooks legitimately differ here - one bank's sheet may leave
    // the validation column empty.
    let mut missing = formulas
        .keys()
        .filter(|reference| !seen.contains(*reference))
        .cloned()
        .collect::<Vec<_>>();
    missing.sort_by_key(|reference| {
        parse_a1_cell(reference)
            .map(|(row, column)| (row, column))
            .unwrap_or((u32::MAX, u32::MAX))
    });
    // Group by row: several new cells in one row share a single insertion.
    let mut new_by_row: BTreeMap<u32, Vec<&String>> = BTreeMap::new();
    for reference in &missing {
        let (row, _column) = parse_a1_cell(reference)?;
        new_by_row.entry(row).or_default().push(reference);
    }
    // Rows that already exist get their cells inserted at the right column
    // position. Rows that do not exist are created together at the end of
    // `sheetData`, in row order - inserting them one by one at the same offset
    // would interleave and produce nested rows.
    let mut appended_rows = String::new();
    let mut cell_insertions: BTreeMap<usize, String> = BTreeMap::new();
    for (row, references) in &new_by_row {
        // Cells are inserted one at a time, each resolved against the row as it
        // existed originally. Several inserts into one row therefore all get
        // the same anchor offset only when they are adjacent; when the row has
        // larger existing columns the anchor differs per cell, and edits are
        // spliced from the end so earlier offsets stay valid.
        let mut inserted_here = false;
        let mut row_cells = String::new();
        for reference in references {
            let cell = format!(
                "<c r=\"{reference}\"><f>{}</f></c>",
                escape_formula_text(
                    formulas
                        .get(*reference)
                        .expect("missing cell has a formula")
                )
            );
            let (_, column) = parse_a1_cell(reference)?;
            match row_insert_offset(&document, *row, &[column])? {
                Some(offset) => {
                    // Cells sharing an anchor must be inserted together: separate
                    // splices at the same offset reverse their column order.
                    cell_insertions.entry(offset).or_default().push_str(&cell);
                    inserted_here = true;
                }
                None => row_cells.push_str(&cell),
            }
        }
        if !row_cells.is_empty() && !inserted_here {
            appended_rows.push_str(&format!("<row r=\"{row}\">{row_cells}</row>"));
        }
    }
    for (offset, cells) in cell_insertions {
        edits.push((offset..offset, cells.into_bytes()));
    }
    if !appended_rows.is_empty() {
        // A source sheet can be almost empty (a bank that fills only the
        // summary columns). Create the rows rather than refuse: Excel's
        // Range.Copy does the same.
        match sheet_data_insert(&document)? {
            SheetDataInsert::BeforeClose(offset) => {
                edits.push((offset..offset, appended_rows.into_bytes()))
            }
            SheetDataInsert::ReplaceSelfClosing(offset) => edits.push((
                offset..offset + 1,
                format!(">{appended_rows}</sheetData>").into_bytes(),
            )),
        }
    }

    let mut patched = sheet_xml.to_vec();
    let with_formula = seen
        .iter()
        .filter(|reference| formulas.contains_key(*reference))
        .count()
        + missing.len();
    let cleared = seen.len() - (with_formula - missing.len());
    edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
    for (range, replacement) in edits {
        patched = splice_bytes(&patched, range, &replacement);
    }
    Ok((patched, with_formula, cleared))
}

/// Where new rows can be written inside a worksheet's `sheetData`.
enum SheetDataInsert {
    /// Insert just before `</sheetData>`.
    BeforeClose(usize),
    /// `<sheetData/>` is self-closing: replace the trailing `/` at this offset
    /// with `>` plus the new rows plus `</sheetData>`.
    ReplaceSelfClosing(usize),
}

fn sheet_data_insert(document: &Document<'_>) -> Result<SheetDataInsert> {
    let sheet_data = document
        .descendants()
        .find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "sheetData"
        })
        .context("worksheet has no sheetData element")?;
    let range = sheet_data.range();
    let fragment = &document.input_text()[range.clone()];
    if let Some(close) = fragment.rfind("</sheetData>") {
        return Ok(SheetDataInsert::BeforeClose(range.start + close));
    }
    let slash = fragment
        .rfind('/')
        .context("sheetData element is neither open nor self-closing")?;
    Ok(SheetDataInsert::ReplaceSelfClosing(range.start + slash))
}

/// Byte offset where cells for `columns` belong inside `row`, or `None` when
/// the row element does not exist.
///
/// Excel requires `<c>` elements in a row to be sorted by column; appending new
/// cells at `</row>` produced packages Excel refuses to open. New cells are
/// placed before the first existing cell whose column is greater.
fn row_insert_offset(document: &Document<'_>, row: u32, columns: &[u32]) -> Result<Option<usize>> {
    let sheet_data = document.descendants().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "sheetData"
    });
    let Some(sheet_data) = sheet_data else {
        return Ok(None);
    };
    for row_node in sheet_data.children().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "row"
    }) {
        if row_node
            .attribute("r")
            .and_then(|value| value.parse::<u32>().ok())
            == Some(row)
        {
            let range = row_node.range();
            let fragment = &document.input_text()[range.clone()];
            let close = fragment
                .rfind("</row>")
                .context("row element has no closing tag")?;
            let smallest = columns.iter().copied().min().unwrap_or(u32::MAX);
            let offset = row_node
                .children()
                .filter(|node| {
                    node.is_element()
                        && node.tag_name().namespace() == Some(MAIN_NS)
                        && node.tag_name().name() == "c"
                })
                .filter_map(|cell| {
                    let (_, column) = parse_a1_cell(cell.attribute("r")?).ok()?;
                    (column > smallest).then(|| cell.range().start - range.start)
                })
                .min()
                .unwrap_or(close);
            return Ok(Some(range.start + offset));
        }
    }
    Ok(None)
}

/// Removes an already-empty cached result from a simple formula cell.
///
/// This narrow patch is intended for writer output that must retain formula
/// text while matching a source workbook with no cached formula value. It
/// rejects shared/array formulas, non-empty caches, and unverified cell content.
/// The input package is copied to `patched_output`; only the resolved worksheet
/// part is changed. `patched_output` may be an existing reserved temporary file.
pub fn clear_empty_formula_cache(
    input_workbook: impl AsRef<Path>,
    patched_output: impl AsRef<Path>,
    sheet_name: &str,
    cell_reference: &str,
) -> Result<FormulaCachePatchReport> {
    let input_workbook = input_workbook.as_ref();
    let patched_output = patched_output.as_ref();
    let same_path = input_workbook == patched_output
        || (patched_output.exists()
            && fs::canonicalize(input_workbook)? == fs::canonicalize(patched_output)?);
    if same_path {
        bail!("input and output workbooks must be different paths");
    }
    let parts = sheet_parts(input_workbook)?;
    let worksheet_part = parts
        .get(sheet_name)
        .with_context(|| format!("worksheet {sheet_name:?} not found"))?
        .clone();
    let mut input_zip = ZipArchive::new(fs::File::open(input_workbook)?)?;
    let worksheet_xml = read_zip_part(&mut input_zip, &worksheet_part)?;
    let patched_worksheet = clear_empty_formula_cache_xml(&worksheet_xml, cell_reference)?;

    if let Some(parent) = patched_output.parent() {
        fs::create_dir_all(parent)?;
    }
    let write_result = (|| -> Result<usize> {
        let target = fs::File::create(patched_output)?;
        let mut writer = ZipWriter::new(target);
        let package_parts = input_zip.len();
        for index in 0..package_parts {
            let mut entry = input_zip.by_index(index)?;
            let name = entry.name().to_owned();
            let method = entry.compression();
            let mut options = SimpleFileOptions::default().compression_method(method);
            if let Some(modified) = entry.last_modified() {
                options = options.last_modified_time(modified);
            }
            if entry.is_dir() {
                writer.add_directory(name, options)?;
                continue;
            }
            let mut data = Vec::new();
            entry.read_to_end(&mut data)?;
            writer.start_file(&name, options)?;
            if name == worksheet_part {
                writer.write_all(&patched_worksheet)?;
            } else {
                writer.write_all(&data)?;
            }
        }
        writer.finish()?;
        Ok(package_parts)
    })();
    let package_parts = match write_result {
        Ok(package_parts) => package_parts,
        Err(error) => {
            let _ = fs::remove_file(patched_output);
            return Err(error);
        }
    };

    Ok(FormulaCachePatchReport {
        worksheet_part,
        cell_reference: cell_reference.to_owned(),
        package_parts,
    })
}

/// Restores each source worksheet's column definitions after an editor save and
/// applies explicitly requested width changes. Only worksheet `<cols>` nodes are
/// altered; other package parts are copied unchanged.
pub fn patch_column_dimensions(
    source: impl AsRef<Path>,
    editor_output: impl AsRef<Path>,
    patched_output: impl AsRef<Path>,
    updates: &[ColumnWidthUpdate],
) -> Result<PatchReport> {
    let source = source.as_ref();
    let editor_output = editor_output.as_ref();
    let patched_output = patched_output.as_ref();
    let mut seen_updates = BTreeSet::new();
    for update in updates {
        if !seen_updates.insert((update.sheet_name.as_str(), update.column_index)) {
            bail!(
                "duplicate width update for sheet {:?}, column {}",
                update.sheet_name,
                update.column_index
            );
        }
    }
    let source_parts = sheet_parts(source)?;
    let output_parts = sheet_parts(editor_output)?;
    if source_parts.len() != output_parts.len()
        || source_parts
            .keys()
            .any(|name| !output_parts.contains_key(name))
    {
        bail!("source and editor output have different worksheet names or counts");
    }
    let missing: Vec<_> = updates
        .iter()
        .filter(|u| !source_parts.contains_key(&u.sheet_name))
        .collect();
    if !missing.is_empty() {
        bail!("requested worksheet for a column update is missing from source workbook");
    }
    let mut source_zip = ZipArchive::new(fs::File::open(source)?)?;
    let mut output_zip = ZipArchive::new(fs::File::open(editor_output)?)?;
    let mut replacements = BTreeMap::new();
    let mut changed_parts = BTreeSet::new();
    for (sheet_name, output_part) in &output_parts {
        let source_part = source_parts
            .get(sheet_name)
            .ok_or_else(|| anyhow!("source worksheet {sheet_name:?} is missing"))?;
        let source_xml = read_zip_part(&mut source_zip, source_part)?;
        let output_xml = read_zip_part(&mut output_zip, output_part)?;
        let sheet_updates: Vec<_> = updates
            .iter()
            .filter(|u| u.sheet_name == *sheet_name)
            .collect();
        let replacement = replace_cols(&source_xml, &output_xml, &sheet_updates)?;
        if replacement != output_xml {
            changed_parts.insert(output_part.clone());
        }
        replacements.insert(output_part.clone(), replacement);
    }
    let mut seen_parts = BTreeSet::new();
    for output_part in output_parts.values() {
        seen_parts.insert(output_part.as_str());
    }
    if seen_parts.len() != output_parts.len() {
        bail!("workbook relationships resolve multiple sheets to the same worksheet part");
    }
    if let Some(parent) = patched_output.parent() {
        fs::create_dir_all(parent)?;
    }
    let target = fs::File::create(patched_output)?;
    let mut writer = ZipWriter::new(target);
    let package_parts = output_zip.len();
    for index in 0..package_parts {
        let mut entry = output_zip.by_index(index)?;
        let name = entry.name().to_owned();
        let method: CompressionMethod = entry.compression();
        let mut options = SimpleFileOptions::default().compression_method(method);
        if let Some(modified) = entry.last_modified() {
            options = options.last_modified_time(modified);
        }
        if entry.is_dir() {
            writer.add_directory(name, options)?;
            continue;
        }
        let mut data = Vec::new();
        entry.read_to_end(&mut data)?;
        writer.start_file(&name, options)?;
        if let Some(replacement) = replacements.get(&name) {
            writer.write_all(replacement)?;
        } else {
            writer.write_all(&data)?;
        }
    }
    writer.finish()?;
    let changed_worksheet_parts = changed_parts.into_iter().collect();
    Ok(PatchReport {
        updated_columns: updates.len(),
        changed_worksheet_parts,
        package_parts,
    })
}

/// Copy only row and column outline levels between aligned worksheet ranges.
/// Values, styles, widths, heights, and hidden flags are handled by the edit
/// adapter. This namespace-aware patch is limited to the corresponding
/// `outlineLevel` attributes and leaves every other package part untouched.
pub fn copy_outline_levels(
    source: impl AsRef<Path>,
    editor_output: impl AsRef<Path>,
    source_sheet: &str,
    source_range: &str,
    target_sheet: &str,
    target_start: &str,
) -> Result<usize> {
    let source = source.as_ref();
    let editor_output = editor_output.as_ref();
    if source == editor_output
        || (editor_output.exists() && fs::canonicalize(source)? == fs::canonicalize(editor_output)?)
    {
        bail!("outline-level copy requires distinct source and output workbooks");
    }

    let source_parts = sheet_parts(source)?;
    let output_parts = sheet_parts(editor_output)?;
    let source_part = source_parts
        .get(source_sheet)
        .with_context(|| format!("source worksheet {source_sheet:?} not found"))?;
    let output_part = output_parts
        .get(target_sheet)
        .with_context(|| format!("target worksheet {target_sheet:?} not found"))?;
    let (source_min_row, source_min_col, source_max_row, source_max_col) =
        parse_a1_range(source_range)?;
    let (target_start_row, target_start_col) = parse_a1_cell(target_start)?;

    let mut source_zip = ZipArchive::new(fs::File::open(source)?)?;
    let source_xml = read_zip_part(&mut source_zip, source_part)?;
    drop(source_zip);
    let mut output_zip = ZipArchive::new(fs::File::open(editor_output)?)?;
    let output_xml = read_zip_part(&mut output_zip, output_part)?;
    drop(output_zip);

    let source_doc = Document::parse(std::str::from_utf8(&source_xml)?)?;
    let output_doc = Document::parse(std::str::from_utf8(&output_xml)?)?;
    let source_root = source_doc.root_element();
    let output_root = output_doc.root_element();
    validate_worksheet_root(source_root)?;
    validate_worksheet_root(output_root)?;

    let mut edits = Vec::<(Range<usize>, Vec<u8>)>::new();
    let mut row_insertions = BTreeMap::<usize, Vec<(u32, String)>>::new();
    let mut column_insertions = Vec::<(u32, u8)>::new();
    let mut updated = 0usize;

    for source_row in source_root.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "row"
    }) {
        let Some(row) = source_row
            .attribute("r")
            .and_then(|value| value.parse::<u32>().ok())
        else {
            bail!("source worksheet row has a missing or invalid r attribute");
        };
        if row < source_min_row || row > source_max_row {
            continue;
        }
        let Some(level) = source_row
            .attribute("outlineLevel")
            .and_then(|value| value.parse::<u8>().ok())
            .filter(|level| *level > 0)
        else {
            continue;
        };
        let target_row = target_start_row + row - source_min_row;
        let matches = output_root
            .descendants()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "row"
                    && node
                        .attribute("r")
                        .and_then(|value| value.parse::<u32>().ok())
                        == Some(target_row)
            })
            .collect::<Vec<_>>();
        if matches.len() > 1 {
            bail!("target worksheet has duplicate row {target_row}");
        }
        if let Some(target_row_node) = matches.first().copied() {
            if target_row_node.attribute("outlineLevel") == Some(&level.to_string()) {
                continue;
            }
            edits.push(attribute_update(
                &output_xml,
                target_row_node,
                "outlineLevel",
                &level.to_string(),
            )?);
        } else {
            let sheet_data = output_root
                .children()
                .find(|node| {
                    node.is_element()
                        && node.tag_name().namespace() == Some(MAIN_NS)
                        && node.tag_name().name() == "sheetData"
                })
                .context("target worksheet has no sheetData element")?;
            let position = row_insertion_position(&output_xml, sheet_data, target_row)?;
            let qname = child_qname(&output_xml, sheet_data, "row");
            row_insertions.entry(position).or_default().push((
                target_row,
                format!("<{qname} r=\"{target_row}\" outlineLevel=\"{level}\"/>"),
            ));
        }
        updated += 1;
    }

    let source_cols = source_root.children().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "cols"
    });
    let mut column_level_updates = BTreeMap::<u32, u8>::new();
    if let Some(source_cols) = source_cols {
        for source_col in source_cols.children().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "col"
        }) {
            let min = source_col
                .attribute("min")
                .and_then(|value| value.parse::<u32>().ok())
                .context("source column definition has invalid min")?;
            let max = source_col
                .attribute("max")
                .and_then(|value| value.parse::<u32>().ok())
                .context("source column definition has invalid max")?;
            let Some(level) = source_col
                .attribute("outlineLevel")
                .and_then(|value| value.parse::<u8>().ok())
                .filter(|level| *level > 0)
            else {
                continue;
            };
            let first = min.max(source_min_col);
            let last = max.min(source_max_col);
            if first > last {
                continue;
            }
            for source_column in first..=last {
                let target_column = target_start_col + source_column - source_min_col;
                if let Some(previous) = column_level_updates.insert(target_column, level) {
                    if previous != level {
                        bail!("source worksheet has conflicting outline levels for column {source_column}");
                    }
                }
                updated += 1;
            }
        }
    }

    let output_cols = output_root.children().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "cols"
    });
    let mut column_node_updates =
        BTreeMap::<(usize, usize), (roxmltree::Node<'_, '_>, BTreeMap<u32, u8>)>::new();
    for (target_column, level) in &column_level_updates {
        let target_nodes = output_cols
            .into_iter()
            .flat_map(|cols| cols.children())
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "col"
                    && node
                        .attribute("min")
                        .and_then(|value| value.parse::<u32>().ok())
                        .is_some_and(|value| value <= *target_column)
                    && node
                        .attribute("max")
                        .and_then(|value| value.parse::<u32>().ok())
                        .is_some_and(|value| value >= *target_column)
            })
            .collect::<Vec<_>>();
        if target_nodes.len() > 1 {
            bail!("target worksheet has overlapping column definitions at {target_column}");
        }
        if let Some(target_col_node) = target_nodes.first().copied() {
            if target_col_node.attribute("outlineLevel") == Some(&level.to_string()) {
                continue;
            }
            let range = target_col_node.range();
            column_node_updates
                .entry((range.start, range.end))
                .or_insert_with(|| (target_col_node, BTreeMap::new()))
                .1
                .insert(*target_column, *level);
        } else {
            column_insertions.push((*target_column, *level));
        }
    }
    for (_, (column_node, updates)) in column_node_updates {
        let replacement = split_column_definition(&output_xml, column_node, &updates)?;
        edits.push((column_node.range(), replacement));
    }

    if !column_insertions.is_empty() {
        let output_cols = output_root.children().find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "cols"
        });
        if let Some(cols) = output_cols {
            let qname = child_qname(&output_xml, cols, "col");
            let mut insertions = BTreeMap::<usize, Vec<(u32, u8)>>::new();
            column_insertions.sort_by_key(|(column, _)| *column);
            for (column, level) in column_insertions {
                let position = column_insertion_position(&output_xml, cols, column)?;
                insertions
                    .entry(position)
                    .or_default()
                    .push((column, level));
            }
            for (position, columns) in insertions {
                let mut fragment = String::new();
                for (column, level) in columns {
                    fragment.push_str(&format!(
                        "<{qname} min=\"{column}\" max=\"{column}\" outlineLevel=\"{level}\"/>"
                    ));
                }
                edits.push((position..position, fragment.into_bytes()));
            }
        } else {
            let sheet_data = output_root
                .children()
                .find(|node| {
                    node.is_element()
                        && node.tag_name().namespace() == Some(MAIN_NS)
                        && node.tag_name().name() == "sheetData"
                })
                .context("target worksheet has no sheetData element")?;
            let position = sheet_data.range().start;
            let qname = child_qname(&output_xml, output_root, "cols");
            let col_qname = child_qname(&output_xml, output_root, "col");
            let mut fragment = format!("<{qname}>");
            column_insertions.sort_by_key(|(column, _)| *column);
            for (column, level) in column_insertions {
                fragment.push_str(&format!(
                    "<{col_qname} min=\"{column}\" max=\"{column}\" outlineLevel=\"{level}\"/>"
                ));
            }
            fragment.push_str(&format!("</{qname}>"));
            edits.push((position..position, fragment.into_bytes()));
        }
    }

    for (position, mut rows) in row_insertions {
        rows.sort_by_key(|(row, _)| *row);
        edits.push((
            position..position,
            rows.into_iter()
                .map(|(_, xml)| xml)
                .collect::<String>()
                .into_bytes(),
        ));
    }
    if edits.is_empty() {
        return Ok(0);
    }
    let patched_xml = apply_xml_edits(&output_xml, edits)?;
    Document::parse(std::str::from_utf8(&patched_xml)?)
        .context("patched worksheet XML is invalid")?;
    replace_worksheet_part(editor_output, output_part, &patched_xml)?;
    Ok(updated)
}

fn validate_worksheet_root(root: roxmltree::Node<'_, '_>) -> Result<()> {
    if root.tag_name().name() != "worksheet" || root.tag_name().namespace() != Some(MAIN_NS) {
        bail!("expected a SpreadsheetML worksheet root");
    }
    Ok(())
}

fn parse_a1_cell(address: &str) -> Result<(u32, u32)> {
    let address = address.trim();
    let address = address.strip_prefix('$').unwrap_or(address);
    let letters_end = address
        .bytes()
        .take_while(|byte| byte.is_ascii_alphabetic())
        .count();
    let (letters, row) = address.split_at(letters_end);
    if letters.is_empty() || !letters.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        bail!("invalid A1 cell address {address:?}");
    }
    let mut column = 0u32;
    for byte in letters.bytes() {
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .context("A1 column index overflows")?;
    }
    let row = row.strip_prefix('$').unwrap_or(row);
    if row.is_empty() || !row.bytes().all(|byte| byte.is_ascii_digit()) {
        bail!("invalid A1 cell address {address:?}");
    }
    let row = row.parse::<u32>()?;
    if row == 0 || column == 0 {
        bail!("A1 cell indexes are 1-based");
    }
    Ok((row, column))
}

fn parse_a1_range(address: &str) -> Result<(u32, u32, u32, u32)> {
    let (start, end) = address.split_once(':').unwrap_or((address, address));
    let (row1, col1) = parse_a1_cell(start)?;
    let (row2, col2) = parse_a1_cell(end)?;
    Ok((
        row1.min(row2),
        col1.min(col2),
        row1.max(row2),
        col1.max(col2),
    ))
}

fn attribute_update(
    xml: &[u8],
    node: roxmltree::Node<'_, '_>,
    name: &str,
    value: &str,
) -> Result<(Range<usize>, Vec<u8>)> {
    if let Some(attribute) = node.attribute_node(name) {
        return Ok((attribute.range_value(), value.as_bytes().to_vec()));
    }
    let node_range = node.range();
    let fragment = std::str::from_utf8(&xml[node_range.clone()])?;
    let tag_end = fragment
        .find('>')
        .context("XML element has no closing angle bracket")?;
    let insert_at = if tag_end > 0 && fragment.as_bytes()[tag_end - 1] == b'/' {
        node_range.start + tag_end - 1
    } else {
        node_range.start + tag_end
    };
    Ok((
        insert_at..insert_at,
        format!(" {name}=\"{value}\"").into_bytes(),
    ))
}

fn split_column_definition(
    xml: &[u8],
    column: roxmltree::Node<'_, '_>,
    outline_updates: &BTreeMap<u32, u8>,
) -> Result<Vec<u8>> {
    let min = column
        .attribute("min")
        .and_then(|value| value.parse::<u32>().ok())
        .context("target column definition has invalid min")?;
    let max = column
        .attribute("max")
        .and_then(|value| value.parse::<u32>().ok())
        .context("target column definition has invalid max")?;
    ensure!(
        min > 0 && max >= min,
        "target column definition has an invalid range"
    );
    ensure!(
        column.children().all(|child| child.is_text()),
        "target <col> definition has unsupported child nodes"
    );

    let mut attributes = Vec::<(String, String)>::new();
    for attribute in column.attributes() {
        ensure!(
            attribute.namespace().is_none(),
            "namespaced attributes on <col> definitions are not supported by this patch"
        );
        attributes.push((attribute.name().to_owned(), attribute.value().to_owned()));
    }
    ensure!(
        attributes.iter().any(|(name, _)| name == "min")
            && attributes.iter().any(|(name, _)| name == "max"),
        "target column definition is missing min/max attributes"
    );
    let qname = child_qname(
        xml,
        column.parent().context("target <col> has no parent")?,
        "col",
    );

    let mut segments = Vec::<(u32, u32, Vec<(String, String)>)>::new();
    for index in min..=max {
        let mut current = attributes.clone();
        if let Some((_, value)) = current.iter_mut().find(|(name, _)| name == "min") {
            *value = index.to_string();
        }
        if let Some((_, value)) = current.iter_mut().find(|(name, _)| name == "max") {
            *value = index.to_string();
        }
        if let Some(level) = outline_updates.get(&index) {
            if let Some((_, value)) = current.iter_mut().find(|(name, _)| name == "outlineLevel") {
                *value = level.to_string();
            } else {
                current.push(("outlineLevel".to_owned(), level.to_string()));
            }
        }

        if let Some((_, previous_max, previous_attrs)) = segments.last_mut() {
            let mut previous_normalized = previous_attrs.clone();
            let mut current_normalized = current.clone();
            for attrs in [&mut previous_normalized, &mut current_normalized] {
                if let Some((_, value)) = attrs.iter_mut().find(|(name, _)| name == "min") {
                    *value = String::new();
                }
                if let Some((_, value)) = attrs.iter_mut().find(|(name, _)| name == "max") {
                    *value = String::new();
                }
            }
            if *previous_max + 1 == index && previous_normalized == current_normalized {
                *previous_max = index;
                *previous_attrs = current;
                continue;
            }
        }
        segments.push((index, index, current));
    }

    let mut output = String::new();
    for (first, last, mut segment_attributes) in segments {
        if let Some((_, value)) = segment_attributes
            .iter_mut()
            .find(|(name, _)| name == "min")
        {
            *value = first.to_string();
        }
        if let Some((_, value)) = segment_attributes
            .iter_mut()
            .find(|(name, _)| name == "max")
        {
            *value = last.to_string();
        }
        output.push('<');
        output.push_str(&qname);
        for (name, value) in segment_attributes {
            output.push(' ');
            output.push_str(&name);
            output.push_str("=\"");
            output.push_str(&escape_xml_attribute(&value));
            output.push('"');
        }
        output.push_str("/>");
    }
    Ok(output.into_bytes())
}

fn row_insertion_position(
    xml: &[u8],
    sheet_data: roxmltree::Node<'_, '_>,
    target_row: u32,
) -> Result<usize> {
    for row in sheet_data.children().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "row"
    }) {
        let number = row
            .attribute("r")
            .and_then(|value| value.parse::<u32>().ok())
            .context("target row has a missing or invalid r attribute")?;
        if number > target_row {
            return Ok(row.range().start);
        }
    }
    closing_tag_start(xml, sheet_data.range())
}

fn column_insertion_position(
    xml: &[u8],
    cols: roxmltree::Node<'_, '_>,
    target_column: u32,
) -> Result<usize> {
    for column in cols.children().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "col"
    }) {
        let min = column
            .attribute("min")
            .and_then(|value| value.parse::<u32>().ok())
            .context("target column definition has invalid min")?;
        if min > target_column {
            return Ok(column.range().start);
        }
    }
    closing_tag_start(xml, cols.range())
}

fn child_qname(xml: &[u8], parent: roxmltree::Node<'_, '_>, child: &str) -> String {
    let parent_range = parent.range();
    let fragment = std::str::from_utf8(&xml[parent_range]).unwrap_or_default();
    let Some(start) = fragment.find('<') else {
        return child.to_owned();
    };
    let name_start = start + 1;
    let name_end = fragment[name_start..]
        .find(|character: char| character.is_whitespace() || matches!(character, '>' | '/'))
        .map(|offset| name_start + offset)
        .unwrap_or(fragment.len());
    let parent_name = &fragment[name_start..name_end];
    parent_name
        .rsplit_once(':')
        .map(|(prefix, _)| format!("{prefix}:{child}"))
        .unwrap_or_else(|| child.to_owned())
}

fn closing_tag_start(xml: &[u8], range: Range<usize>) -> Result<usize> {
    let fragment = std::str::from_utf8(&xml[range.clone()])?;
    let relative = fragment
        .rfind("</")
        .context("XML element has no closing tag")?;
    Ok(range.start + relative)
}

fn apply_xml_edits(xml: &[u8], mut edits: Vec<(Range<usize>, Vec<u8>)>) -> Result<Vec<u8>> {
    edits.sort_by_key(|(range, _)| std::cmp::Reverse((range.start, range.end)));
    let mut output = xml.to_vec();
    let mut previous_start = usize::MAX;
    for (range, replacement) in edits {
        if range.end > output.len() || range.start > range.end {
            bail!("XML patch range is out of bounds");
        }
        if range.start == previous_start && range.start != range.end {
            bail!("XML patch contains overlapping edits");
        }
        output = splice_bytes(&output, range.clone(), &replacement);
        previous_start = range.start;
    }
    Ok(output)
}

fn replace_worksheet_part(workbook: &Path, part: &str, replacement: &[u8]) -> Result<()> {
    let file_name = workbook
        .file_name()
        .and_then(|name| name.to_str())
        .context("workbook output path is not valid UTF-8")?;
    let sequence = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)?
        .as_nanos();
    let patched_path = workbook.with_file_name(format!(
        ".{file_name}.{}.{}.outline.xlsx",
        std::process::id(),
        sequence
    ));
    let write_result = (|| -> Result<()> {
        let mut input_zip = ZipArchive::new(fs::File::open(workbook)?)?;
        let mut writer = ZipWriter::new(fs::File::create(&patched_path)?);
        for index in 0..input_zip.len() {
            let mut entry = input_zip.by_index(index)?;
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
            if name == part {
                writer.write_all(replacement)?;
            } else {
                std::io::copy(&mut entry, &mut writer)?;
            }
        }
        writer.finish()?;
        Ok(())
    })();
    if let Err(error) = write_result {
        let _ = fs::remove_file(&patched_path);
        return Err(error);
    }
    fs::remove_file(workbook)
        .with_context(|| format!("replace disposable editor output {}", workbook.display()))?;
    if let Err(error) = fs::rename(&patched_path, workbook) {
        let _ = fs::remove_file(&patched_path);
        return Err(error).with_context(|| {
            format!(
                "publish outline-level patched output {}",
                workbook.display()
            )
        });
    }
    Ok(())
}

#[cfg(test)]
mod formula_cache_tests {
    use super::{clear_empty_formula_cache, MAIN_NS};
    use anyhow::Result;
    use roxmltree::Document;
    use std::{
        collections::BTreeMap,
        fs::{self, File},
        io::{Read, Write},
        path::PathBuf,
    };
    use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

    fn temp_path(label: &str) -> PathBuf {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("clock after epoch")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "ooxml_formula_cache_{}_{}_{}",
            std::process::id(),
            nonce,
            label
        ))
    }

    fn write_fixture(path: &PathBuf, cached_value: &str) -> Result<BTreeMap<String, Vec<u8>>> {
        const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
        let worksheet = format!(
            r#"<s:worksheet xmlns:s="{MAIN_NS}"><s:sheetData><s:row r="1"><s:c r="A1" t="str"><s:f>1+1</s:f><s:v>{cached_value}</s:v></s:c><s:c r="B1" t="inlineStr"><s:is><s:t>keep</s:t></s:is></s:c></s:row></s:sheetData></s:worksheet>"#
        );
        let other_worksheet = format!(
            r#"<s:worksheet xmlns:s="{MAIN_NS}"><s:sheetData><s:row r="1"><s:c r="A1"><s:v>7</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{MAIN_NS}" xmlns:r="{REL_NS}"><s:sheets><s:sheet name="Summary" sheetId="1" r:id="rId1"/><s:sheet name="Other" sheetId="2" r:id="rId2"/></s:sheets></s:workbook>"#
        );
        let relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheetA.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheetB.xml"/></Relationships>"#
        );
        let entries = BTreeMap::from([
            (
                "_rels/.rels".to_owned(),
                format!(
                    r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
                )
                .into_bytes(),
            ),
            ("xl/workbook.xml".to_owned(), workbook.into_bytes()),
            (
                "xl/_rels/workbook.xml.rels".to_owned(),
                relationships.into_bytes(),
            ),
            (
                "xl/worksheets/sheetA.xml".to_owned(),
                worksheet.into_bytes(),
            ),
            (
                "xl/worksheets/sheetB.xml".to_owned(),
                other_worksheet.into_bytes(),
            ),
            ("custom/opaque.bin".to_owned(), vec![0, 1, 2, 255]),
        ]);
        let mut archive = ZipWriter::new(File::create(path)?);
        for (name, data) in &entries {
            archive.start_file(
                name,
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
            )?;
            archive.write_all(data)?;
        }
        archive.finish()?;
        Ok(entries)
    }

    fn read_entries(path: &PathBuf) -> Result<BTreeMap<String, Vec<u8>>> {
        let mut archive = ZipArchive::new(File::open(path)?)?;
        let mut entries = BTreeMap::new();
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            let mut data = Vec::new();
            entry.read_to_end(&mut data)?;
            entries.insert(entry.name().to_owned(), data);
        }
        Ok(entries)
    }

    #[test]
    fn clears_only_an_empty_formula_cache_and_preserves_other_package_parts() -> Result<()> {
        let input = temp_path("input.xlsx");
        let output = temp_path("output.xlsx");
        let original_entries = write_fixture(&input, "")?;

        let report = clear_empty_formula_cache(&input, &output, "Summary", "A1")?;
        assert_eq!(report.worksheet_part, "xl/worksheets/sheetA.xml");
        assert_eq!(report.cell_reference, "A1");
        assert_eq!(report.package_parts, original_entries.len());

        let patched_entries = read_entries(&output)?;
        assert_eq!(
            original_entries.keys().collect::<Vec<_>>(),
            patched_entries.keys().collect::<Vec<_>>()
        );
        for (name, original) in &original_entries {
            if name == "xl/worksheets/sheetA.xml" {
                continue;
            }
            assert_eq!(
                patched_entries[name], *original,
                "changed package part {name}"
            );
        }
        let patched_xml = std::str::from_utf8(&patched_entries["xl/worksheets/sheetA.xml"])?;
        let document = Document::parse(patched_xml)?;
        let cell = document
            .root_element()
            .descendants()
            .find(|node| {
                node.is_element()
                    && node.tag_name().name() == "c"
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.attribute("r") == Some("A1")
            })
            .expect("target cell remains");
        assert!(cell.attribute("t").is_none());
        assert_eq!(
            cell.children()
                .find(|node| node.is_element() && node.tag_name().name() == "f")
                .and_then(|node| node.text()),
            Some("1+1")
        );
        assert!(!cell
            .children()
            .any(|node| node.is_element() && node.tag_name().name() == "v"));
        fs::remove_file(input)?;
        fs::remove_file(output)?;
        Ok(())
    }

    #[test]
    fn refuses_to_clear_a_non_empty_formula_cache() -> Result<()> {
        let input = temp_path("non-empty-input.xlsx");
        let output = temp_path("non-empty-output.xlsx");
        write_fixture(&input, "2")?;

        let error = clear_empty_formula_cache(&input, &output, "Summary", "A1")
            .expect_err("a non-empty formula result must be rejected");
        assert!(error.to_string().contains("non-empty formula cache"));
        assert!(!output.exists());
        fs::remove_file(input)?;
        Ok(())
    }
}

#[cfg(test)]
mod header_patch_tests {
    use super::{
        preflight_header_text_patch, read_header_inventory, replace_header_text,
        validate_header_text_patch, HeaderTextUpdate, HeaderVariant,
    };
    use anyhow::Result;
    use std::{
        collections::BTreeMap,
        fs::{self, File},
        io::{Read, Write},
        path::PathBuf,
    };
    use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

    const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

    fn temp_path(label: &str) -> PathBuf {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("clock after epoch")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "ooxml_header_patch_{}_{}_{}",
            std::process::id(),
            nonce,
            label
        ))
    }

    fn fixture_entries(nested_header: bool, signed: bool) -> BTreeMap<String, Vec<u8>> {
        let worksheet = if nested_header {
            format!(
                r#"<x:worksheet xmlns:x="{MAIN_NS}"><x:sheetData/><x:headerFooter><x:oddHeader><x:t>单位：万元</x:t></x:oddHeader></x:headerFooter></x:worksheet>"#
            )
        } else {
            format!(
                r#"<x:worksheet xmlns:x="{MAIN_NS}"><x:sheetData/><x:headerFooter differentOddEven="1" differentFirst="1"><x:oddHeader>&amp;L单位：万元&amp;C报表&lt;内容&gt;&amp;R右侧万元</x:oddHeader><x:evenHeader>&amp;L单位：万元</x:evenHeader><x:firstHeader>&amp;L保留</x:firstHeader></x:headerFooter></x:worksheet>"#
            )
        };
        let other_worksheet = format!(
            r#"<x:worksheet xmlns:x="{MAIN_NS}"><x:sheetData/><x:headerFooter><x:oddHeader>&amp;L单位：万元</x:oddHeader></x:headerFooter></x:worksheet>"#
        );
        let workbook = format!(
            r#"<x:workbook xmlns:x="{MAIN_NS}" xmlns:r="{REL_NS}"><x:sheets><x:sheet name="外汇报表" sheetId="4" r:id="rId9"/><x:sheet name="普通报表" sheetId="8" r:id="rId2"/></x:sheets></x:workbook>"#
        );
        let relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/fx.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/ordinary.xml"/></Relationships>"#
        );
        let mut entries = BTreeMap::from([
            (
                "_rels/.rels".to_owned(),
                format!(
                    r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
                )
                .into_bytes(),
            ),
            ("xl/workbook.xml".to_owned(), workbook.into_bytes()),
            (
                "xl/_rels/workbook.xml.rels".to_owned(),
                relationships.into_bytes(),
            ),
            ("xl/worksheets/fx.xml".to_owned(), worksheet.into_bytes()),
            (
                "xl/worksheets/ordinary.xml".to_owned(),
                other_worksheet.into_bytes(),
            ),
            ("custom/opaque.bin".to_owned(), vec![0, 1, 2, 255]),
        ]);
        if signed {
            entries.insert("_xmlsignatures/sig1.xml".to_owned(), b"signature".to_vec());
        }
        entries
    }

    fn write_fixture(
        path: &PathBuf,
        nested_header: bool,
        signed: bool,
    ) -> Result<BTreeMap<String, Vec<u8>>> {
        let entries = fixture_entries(nested_header, signed);
        let mut archive = ZipWriter::new(File::create(path)?);
        for (name, data) in &entries {
            archive.start_file(
                name,
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
            )?;
            archive.write_all(data)?;
        }
        archive.finish()?;
        Ok(entries)
    }

    fn read_entries(path: &PathBuf) -> Result<BTreeMap<String, Vec<u8>>> {
        let mut archive = ZipArchive::new(File::open(path)?)?;
        let mut entries = BTreeMap::new();
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            let mut data = Vec::new();
            entry.read_to_end(&mut data)?;
            entries.insert(entry.name().to_owned(), data);
        }
        Ok(entries)
    }

    fn update() -> HeaderTextUpdate {
        HeaderTextUpdate {
            sheet_name: "外汇报表".to_owned(),
            search: "万元".to_owned(),
            replacement: "万美元".to_owned(),
        }
    }

    #[test]
    fn patches_only_target_header_text_nodes_and_preserves_other_package_parts() -> Result<()> {
        let input = temp_path("input.xlsx");
        let output = temp_path("output.xlsx");
        let original = write_fixture(&input, false, false)?;

        let inventory = read_header_inventory(&input)?;
        assert_eq!(inventory.len(), 2);
        assert_eq!(inventory[0].sheet_name, "外汇报表");
        assert_eq!(inventory[0].headers[0].variant, HeaderVariant::Odd);
        assert_eq!(
            inventory[0].headers[0].text,
            "&L单位：万元&C报表<内容>&R右侧万元"
        );
        assert_eq!(inventory[0].headers.len(), 3);

        let report = preflight_header_text_patch(&input, &[update()])?;
        assert_eq!(report.modified_sheets, 1);
        assert_eq!(report.replacements, 3);
        let written = replace_header_text(&input, &output, &[update()])?;
        assert_eq!(written.replacements, 3);
        validate_header_text_patch(&input, &output, &[update()])?;

        let patched = read_entries(&output)?;
        for (name, bytes) in &original {
            if name == "xl/worksheets/fx.xml" {
                continue;
            }
            assert_eq!(patched[name], *bytes, "changed package part {name}");
        }
        let patched_inventory = read_header_inventory(&output)?;
        assert!(patched_inventory[0]
            .headers
            .iter()
            .all(|header| !header.text.contains("万元")));
        assert_eq!(patched_inventory[0].headers[2].text, "&L保留");
        assert_eq!(patched_inventory[1].headers[0].text, "&L单位：万元");

        fs::remove_file(input)?;
        fs::remove_file(output)?;
        Ok(())
    }

    #[test]
    fn resolves_workbook_and_worksheet_parts_from_package_relationships() -> Result<()> {
        let input = temp_path("custom-parts-input.xlsx");
        let output = temp_path("custom-parts-output.xlsx");
        let standard = fixture_entries(false, false);
        let workbook_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/fx.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/ordinary.xml"/></Relationships>"#
        );
        let entries = BTreeMap::from([
            (
                "_rels/.rels".to_owned(),
                format!(
                    r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="workbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/books/main.xml"/></Relationships>"#
                )
                .into_bytes(),
            ),
            (
                "custom/books/main.xml".to_owned(),
                standard["xl/workbook.xml"].clone(),
            ),
            (
                "custom/books/_rels/main.xml.rels".to_owned(),
                workbook_relationships.into_bytes(),
            ),
            (
                "custom/worksheets/fx.xml".to_owned(),
                standard["xl/worksheets/fx.xml"].clone(),
            ),
            (
                "custom/worksheets/ordinary.xml".to_owned(),
                standard["xl/worksheets/ordinary.xml"].clone(),
            ),
            ("custom/opaque.bin".to_owned(), vec![0, 1, 2, 255]),
        ]);
        let mut archive = ZipWriter::new(File::create(&input)?);
        for (name, data) in &entries {
            archive.start_file(
                name,
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
            )?;
            archive.write_all(data)?;
        }
        archive.finish()?;

        let report = replace_header_text(&input, &output, &[update()])?;
        assert_eq!(
            report.changed_worksheet_parts,
            vec!["custom/worksheets/fx.xml".to_owned()]
        );
        validate_header_text_patch(&input, &output, &[update()])?;
        let patched = read_entries(&output)?;
        for (name, bytes) in &entries {
            if name == "custom/worksheets/fx.xml" {
                continue;
            }
            assert_eq!(patched[name], *bytes, "changed package part {name}");
        }

        fs::remove_file(input)?;
        fs::remove_file(output)?;
        Ok(())
    }

    #[test]
    fn rejects_nested_header_markup_before_creating_output() -> Result<()> {
        let input = temp_path("nested-input.xlsx");
        let output = temp_path("nested-output.xlsx");
        write_fixture(&input, true, false)?;
        let error = preflight_header_text_patch(&input, &[update()])
            .expect_err("nested markup must be rejected");
        assert!(error.to_string().contains("nested markup"));
        assert!(!output.exists());
        fs::remove_file(input)?;
        Ok(())
    }

    #[test]
    fn refuses_signed_or_macro_packages_and_duplicate_updates() -> Result<()> {
        let signed = temp_path("signed.xlsx");
        let unsigned = temp_path("unsigned.xlsx");
        write_fixture(&signed, false, true)?;
        write_fixture(&unsigned, false, false)?;
        assert!(read_header_inventory(&signed)
            .expect_err("signed package must be rejected")
            .to_string()
            .contains("digital signatures"));
        assert!(
            preflight_header_text_patch(&unsigned, &[update(), update()])
                .expect_err("duplicate sheet updates must be rejected")
                .to_string()
                .contains("duplicate header text update")
        );
        fs::remove_file(signed)?;
        fs::remove_file(unsigned)?;
        Ok(())
    }
}

#[cfg(test)]
mod outline_copy_tests {
    use super::{copy_outline_levels, MAIN_NS};
    use anyhow::Result;
    use roxmltree::Document;
    use std::{
        collections::BTreeMap,
        fs::{self, File},
        io::{Read, Write},
        path::PathBuf,
    };
    use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

    const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

    fn temp_path(label: &str) -> PathBuf {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("clock after epoch")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "ooxml_outline_copy_{}_{}_{}",
            std::process::id(),
            nonce,
            label
        ))
    }

    fn fixture_entries(source: bool) -> BTreeMap<String, Vec<u8>> {
        let worksheet = if source {
            format!(
                r#"<s:worksheet xmlns:s="{MAIN_NS}"><s:sheetData><s:row r="2" outlineLevel="1"/></s:sheetData><s:cols><s:col min="2" max="4" width="11" customWidth="1" outlineLevel="2"/></s:cols></s:worksheet>"#
            )
        } else {
            format!(
                r#"<s:worksheet xmlns:s="{MAIN_NS}"><s:sheetData><s:row r="2"><s:c r="A2"><s:v>1</s:v></s:c></s:row></s:sheetData><s:cols><s:col min="1" max="5" width="10" customWidth="1"/></s:cols></s:worksheet>"#
            )
        };
        let workbook = format!(
            r#"<s:workbook xmlns:s="{MAIN_NS}" xmlns:r="{REL_NS}"><s:sheets><s:sheet name="Data" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>"#
        );
        let relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/data.xml"/></Relationships>"#
        );
        BTreeMap::from([
            (
                "_rels/.rels".to_owned(),
                format!(
                    r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
                )
                .into_bytes(),
            ),
            ("xl/workbook.xml".to_owned(), workbook.into_bytes()),
            (
                "xl/_rels/workbook.xml.rels".to_owned(),
                relationships.into_bytes(),
            ),
            ("xl/worksheets/data.xml".to_owned(), worksheet.into_bytes()),
            ("custom/opaque.bin".to_owned(), vec![0, 1, 2, 255]),
        ])
    }

    fn write_entries(path: &PathBuf, entries: &BTreeMap<String, Vec<u8>>) -> Result<()> {
        let mut archive = ZipWriter::new(File::create(path)?);
        for (name, data) in entries {
            archive.start_file(
                name,
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
            )?;
            archive.write_all(data)?;
        }
        archive.finish()?;
        Ok(())
    }

    fn read_entries(path: &PathBuf) -> Result<BTreeMap<String, Vec<u8>>> {
        let mut archive = ZipArchive::new(File::open(path)?)?;
        let mut entries = BTreeMap::new();
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes)?;
            entries.insert(entry.name().to_owned(), bytes);
        }
        Ok(entries)
    }

    #[test]
    fn splits_combined_output_column_ranges_around_imported_outline_levels() -> Result<()> {
        let source = temp_path("source.xlsx");
        let output = temp_path("output.xlsx");
        let source_entries = fixture_entries(true);
        let output_entries = fixture_entries(false);
        write_entries(&source, &source_entries)?;
        write_entries(&output, &output_entries)?;

        let copied = copy_outline_levels(&source, &output, "Data", "A1:E5", "Data", "A1")?;
        assert_eq!(copied, 4);

        let patched_entries = read_entries(&output)?;
        for (name, bytes) in &output_entries {
            if name != "xl/worksheets/data.xml" {
                assert_eq!(patched_entries[name], *bytes, "changed package part {name}");
            }
        }
        assert_eq!(read_entries(&source)?, source_entries);
        let worksheet = std::str::from_utf8(&patched_entries["xl/worksheets/data.xml"])?;
        let document = Document::parse(worksheet)?;
        let root = document.root_element();
        let row = root
            .descendants()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "row"
                    && node.attribute("r") == Some("2")
            })
            .expect("target row exists");
        assert_eq!(row.attribute("outlineLevel"), Some("1"));

        let cols = root
            .children()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "cols"
            })
            .expect("target columns exist");
        let ranges = cols
            .children()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "col"
            })
            .map(|node| {
                (
                    node.attribute("min").unwrap().parse::<u32>().unwrap(),
                    node.attribute("max").unwrap().parse::<u32>().unwrap(),
                    node.attribute("outlineLevel").map(str::to_owned),
                    node.attribute("width").map(str::to_owned),
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(
            ranges,
            vec![
                (1, 1, None, Some("10".to_owned())),
                (2, 4, Some("2".to_owned()), Some("10".to_owned())),
                (5, 5, None, Some("10".to_owned())),
            ]
        );

        fs::remove_file(source)?;
        fs::remove_file(output)?;
        Ok(())
    }
}
#[cfg(test)]
mod calculation_properties_tests {
    use super::{
        patch_calc_pr_xml, patch_workbook_calculation_properties,
        read_workbook_calculation_properties, WorkbookCalculationProperties, MAIN_NS,
        PACKAGE_REL_NS,
    };
    use anyhow::Result;
    use std::{
        fs::{self, File},
        io::Write,
        path::PathBuf,
    };
    use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

    fn temp_path(label: &str) -> PathBuf {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        std::env::temp_dir().join(format!(
            "ooxml_calcpr_{}_{}_{}",
            std::process::id(),
            nonce,
            label
        ))
    }

    #[test]
    fn patches_existing_flags_and_inserts_missing_calc_pr_with_namespace_prefix() -> Result<()> {
        let existing = format!(
            r#"<s:workbook xmlns:s="{MAIN_NS}"><s:calcPr calcMode="auto" forceFullCalc="0"/><s:sheets/></s:workbook>"#
        );
        let changed = patch_calc_pr_xml(
            existing.as_bytes(),
            WorkbookCalculationProperties {
                full_calc_on_load: Some(true),
                force_full_calc: Some(false),
            },
        )?;
        let text = String::from_utf8(changed)?;
        assert!(text.contains("calcMode=\"auto\""));
        assert!(text.contains("fullCalcOnLoad=\"1\""));
        assert!(text.contains("forceFullCalc=\"0\""));

        let absent = format!(r#"<s:workbook xmlns:s="{MAIN_NS}"><s:sheets/></s:workbook>"#);
        let inserted = patch_calc_pr_xml(
            absent.as_bytes(),
            WorkbookCalculationProperties {
                full_calc_on_load: Some(true),
                force_full_calc: Some(true),
            },
        )?;
        let document = roxmltree::Document::parse(std::str::from_utf8(&inserted)?)?;
        let calc = document
            .descendants()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "calcPr"
            })
            .unwrap();
        assert_eq!(calc.attribute("fullCalcOnLoad"), Some("1"));
        assert_eq!(calc.attribute("forceFullCalc"), Some("1"));
        Ok(())
    }

    #[test]
    fn package_patch_reads_back_flags_preserves_unspecified_flags_and_other_parts() -> Result<()> {
        let source = temp_path("source.xlsx");
        let output = temp_path("output.xlsx");
        let workbook = format!(
            r#"<x:workbook xmlns:x="{MAIN_NS}"><x:calcPr calcMode="auto" forceFullCalc="0"/><x:sheets/></x:workbook>"#
        );
        let mut zip = ZipWriter::new(File::create(&source)?);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        zip.start_file("_rels/.rels", options)?;
        zip.write_all(format!(r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).as_bytes())?;
        zip.start_file("xl/workbook.xml", options)?;
        zip.write_all(workbook.as_bytes())?;
        zip.start_file("xl/untouched.bin", options)?;
        zip.write_all(b"keep me")?;
        zip.finish()?;

        patch_workbook_calculation_properties(
            &source,
            &output,
            WorkbookCalculationProperties {
                full_calc_on_load: Some(true),
                force_full_calc: None,
            },
        )?;
        let actual = read_workbook_calculation_properties(&output)?;
        assert_eq!(actual.full_calc_on_load, Some(true));
        assert_eq!(actual.force_full_calc, Some(false));
        let mut archive = ZipArchive::new(File::open(&output)?)?;
        let mut untouched = String::new();
        use std::io::Read;
        archive
            .by_name("xl/untouched.bin")?
            .read_to_string(&mut untouched)?;
        assert_eq!(untouched, "keep me");
        let mut original = ZipArchive::new(File::open(&source)?)?;
        let mut original_workbook = Vec::new();
        original
            .by_name("xl/workbook.xml")?
            .read_to_end(&mut original_workbook)?;
        assert_eq!(original_workbook, workbook.as_bytes());
        fs::remove_file(source)?;
        fs::remove_file(output)?;
        Ok(())
    }
}

#[cfg(test)]
mod external_hyperlink_tests {
    use super::{
        patch_external_hyperlink, patch_external_hyperlinks, prepare_external_hyperlinks_parts,
        ExternalHyperlinkPatch, MAIN_NS, PACKAGE_REL_NS, REL_NS,
    };
    use anyhow::Result;
    use roxmltree::Document;
    use std::{
        collections::BTreeMap,
        fs::{self, File},
        io::{Read, Write},
        path::PathBuf,
    };
    use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

    fn temp_path(label: &str) -> PathBuf {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("clock after epoch")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "ooxml_hyperlink_{}_{}_{}",
            std::process::id(),
            nonce,
            label
        ))
    }

    fn fixture_entries() -> BTreeMap<String, Vec<u8>> {
        let workbook = format!(
            r#"<s:workbook xmlns:s="{MAIN_NS}" xmlns:r="{REL_NS}"><s:sheets><s:sheet name="审计" sheetId="1" r:id="rId1"/><s:sheet name="其他" sheetId="2" r:id="rId2"/></s:sheets></s:workbook>"#
        );
        let workbook_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheetA.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheetB.xml"/></Relationships>"#
        );
        let sheet_a = format!(
            r#"<s:worksheet xmlns:s="{MAIN_NS}" xmlns:r="{REL_NS}"><s:sheetData/><s:hyperlinks><s:hyperlink ref="A1" r:id="rId1" display="旧显示"/><s:hyperlink ref="B1" r:id="rId2" display="保留显示"/></s:hyperlinks></s:worksheet>"#
        );
        let sheet_a_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mnt/c/reports/old.xlsx"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///mnt/c/keep.xlsx" TargetMode="External"/></Relationships>"#
        );
        BTreeMap::from([
            (
                "_rels/.rels".to_owned(),
                format!(
                    r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
                )
                .into_bytes(),
            ),
            ("xl/workbook.xml".to_owned(), workbook.into_bytes()),
            (
                "xl/_rels/workbook.xml.rels".to_owned(),
                workbook_rels.into_bytes(),
            ),
            (
                "xl/worksheets/sheetA.xml".to_owned(),
                sheet_a.into_bytes(),
            ),
            (
                "xl/worksheets/sheetB.xml".to_owned(),
                format!(r#"<s:worksheet xmlns:s="{MAIN_NS}" xmlns:r="{REL_NS}"><s:sheetData/><s:hyperlinks><s:hyperlink ref="A1" r:id="h1"/></s:hyperlinks></s:worksheet>"#).into_bytes(),
            ),
            (
                "xl/worksheets/_rels/sheetA.xml.rels".to_owned(),
                sheet_a_rels.into_bytes(),
            ),
            (
                "xl/worksheets/_rels/sheetB.xml.rels".to_owned(),
                format!(
                    r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="h1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mnt/c/old.xlsx"/></Relationships>"#
                )
                .into_bytes(),
            ),
            ("custom/opaque.bin".to_owned(), vec![0, 1, 2, 255]),
        ])
    }

    fn write_entries(path: &PathBuf, entries: &BTreeMap<String, Vec<u8>>) -> Result<()> {
        let mut archive = ZipWriter::new(File::create(path)?);
        for (name, data) in entries {
            archive.start_file(
                name,
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
            )?;
            archive.write_all(data)?;
        }
        archive.finish()?;
        Ok(())
    }

    fn read_entries(path: &PathBuf) -> Result<BTreeMap<String, Vec<u8>>> {
        let mut archive = ZipArchive::new(File::open(path)?)?;
        let mut entries = BTreeMap::new();
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            let mut data = Vec::new();
            entry.read_to_end(&mut data)?;
            entries.insert(entry.name().to_owned(), data);
        }
        Ok(entries)
    }

    #[test]
    fn patches_one_external_link_and_preserves_other_package_parts_and_links() -> Result<()> {
        let source = temp_path("source.xlsx");
        let output = temp_path("output.xlsx");
        let original = fixture_entries();
        write_entries(&source, &original)?;
        let mut patch =
            ExternalHyperlinkPatch::new("审计", "$a$1", "file:///mnt/c/reports/a%20%26b.xlsx")?;
        patch.location = Some("摘要!A1".to_owned());
        patch.display = Some("审计 & 复核".to_owned());

        let report = patch_external_hyperlink(&source, &output, &patch)?;
        assert_eq!(report.sheet_name, "审计");
        assert_eq!(report.cell_reference, "A1");
        assert_eq!(report.worksheet_part, "xl/worksheets/sheetA.xml");
        assert_eq!(
            report.relationships_part,
            "xl/worksheets/_rels/sheetA.xml.rels"
        );
        assert_eq!(report.relationship_id, "rId1");
        assert_eq!(report.package_parts, original.len());

        let actual = read_entries(&output)?;
        assert_eq!(
            original.keys().collect::<Vec<_>>(),
            actual.keys().collect::<Vec<_>>()
        );
        let existing_output = fs::read(&output)?;
        assert!(patch_external_hyperlink(&source, &output, &patch).is_err());
        assert_eq!(fs::read(&output)?, existing_output);
        for (name, bytes) in &original {
            if name != "xl/worksheets/sheetA.xml" && name != "xl/worksheets/_rels/sheetA.xml.rels" {
                assert_eq!(actual[name], *bytes, "unexpected package change: {name}");
            }
        }
        let source_after = read_entries(&source)?;
        assert_eq!(
            source_after, original,
            "source package must remain unchanged"
        );

        let sheet_xml = std::str::from_utf8(&actual["xl/worksheets/sheetA.xml"])?;
        let sheet_doc = Document::parse(sheet_xml)?;
        let links = sheet_doc
            .descendants()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "hyperlink"
            })
            .collect::<Vec<_>>();
        let target_link = links
            .iter()
            .find(|node| node.attribute("ref") == Some("A1"))
            .unwrap();
        assert_eq!(target_link.attribute("display"), Some("审计 & 复核"));
        assert_eq!(target_link.attribute("location"), Some("摘要!A1"));
        assert_eq!(target_link.attribute((REL_NS, "id")), Some("rId1"));
        let untouched_link = links
            .iter()
            .find(|node| node.attribute("ref") == Some("B1"))
            .unwrap();
        assert_eq!(untouched_link.attribute("display"), Some("保留显示"));

        let rels_xml = std::str::from_utf8(&actual["xl/worksheets/_rels/sheetA.xml.rels"])?;
        let rels_doc = Document::parse(rels_xml)?;
        let relationships = rels_doc
            .descendants()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                    && node.tag_name().name() == "Relationship"
            })
            .collect::<Vec<_>>();
        let changed = relationships
            .iter()
            .find(|node| node.attribute("Id") == Some("rId1"))
            .unwrap();
        assert_eq!(
            changed.attribute("Target"),
            Some("file:///mnt/c/reports/a%20%26b.xlsx")
        );
        assert_eq!(changed.attribute("TargetMode"), Some("External"));
        let untouched = relationships
            .iter()
            .find(|node| node.attribute("Id") == Some("rId2"))
            .unwrap();
        assert_eq!(
            untouched.attribute("Target"),
            Some("file:///mnt/c/keep.xlsx")
        );
        assert_eq!(untouched.attribute("TargetMode"), Some("External"));

        fs::remove_file(source)?;
        fs::remove_file(output)?;
        Ok(())
    }

    #[test]
    fn rejects_missing_or_ambiguous_hyperlink_and_invalid_targets() -> Result<()> {
        let patch = ExternalHyperlinkPatch::new("审计", "A1", "file:///tmp/a.xlsx")?;
        let duplicate = format!(
            r#"<s:worksheet xmlns:s="{MAIN_NS}" xmlns:r="{REL_NS}"><s:hyperlinks><s:hyperlink ref="A1" r:id="rId1"/><s:hyperlink ref="A1" r:id="rId1"/></s:hyperlinks></s:worksheet>"#
        );
        let rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///tmp/old.xlsx" TargetMode="External"/></Relationships>"#
        );
        assert!(prepare_external_hyperlinks_parts(
            duplicate.as_bytes(),
            rels.as_bytes(),
            &[&patch]
        )
        .unwrap_err()
        .to_string()
        .contains("exactly one hyperlink"));

        let missing_relationship = format!(
            r#"<s:worksheet xmlns:s="{MAIN_NS}" xmlns:r="{REL_NS}"><s:hyperlinks><s:hyperlink ref="A1" r:id="missing"/></s:hyperlinks></s:worksheet>"#
        );
        assert!(prepare_external_hyperlinks_parts(
            missing_relationship.as_bytes(),
            rels.as_bytes(),
            &[&patch]
        )
        .unwrap_err()
        .to_string()
        .contains("missing or ambiguous"));
        assert!(ExternalHyperlinkPatch::new("审计", "A1", "relative/path.xlsx").is_err());
        assert!(ExternalHyperlinkPatch::new("审计", "A1", "file://server").is_err());
        Ok(())
    }

    #[test]
    fn batch_patches_multiple_links_once_and_rejects_a_missing_member_before_output() -> Result<()>
    {
        let source = temp_path("batch-source.xlsx");
        let output = temp_path("batch-output.xlsx");
        let original = fixture_entries();
        write_entries(&source, &original)?;

        let valid = ExternalHyperlinkPatch::new("审计", "A1", "file:///mnt/c/a.xlsx")?;
        let missing = ExternalHyperlinkPatch::new("审计", "C9", "file:///mnt/c/missing.xlsx")?;
        assert!(patch_external_hyperlinks(&source, &output, &[valid.clone(), missing]).is_err());
        assert!(!output.exists(), "failed batch must not create a candidate");
        assert_eq!(
            read_entries(&source)?,
            original,
            "source must remain unchanged"
        );

        let mut second = ExternalHyperlinkPatch::new("审计", "B1", "file:///mnt/c/b.xlsx")?;
        second.display = Some("第二个链接".to_owned());
        second.location = Some("结果!B1".to_owned());
        let third = ExternalHyperlinkPatch::new("其他", "A1", "file:///mnt/c/c.xlsx")?;
        let reports = patch_external_hyperlinks(&source, &output, &[valid, second, third])?;
        assert_eq!(reports.len(), 3);
        assert_eq!(reports[0].relationship_id, "rId1");
        assert_eq!(reports[1].relationship_id, "rId2");
        assert_eq!(reports[2].relationship_id, "h1");

        let actual = read_entries(&output)?;
        for (name, bytes) in &original {
            if name != "xl/worksheets/sheetA.xml"
                && name != "xl/worksheets/_rels/sheetA.xml.rels"
                && name != "xl/worksheets/sheetB.xml"
                && name != "xl/worksheets/_rels/sheetB.xml.rels"
            {
                assert_eq!(actual[name], *bytes, "unexpected package change: {name}");
            }
        }
        let worksheet = Document::parse(std::str::from_utf8(&actual["xl/worksheets/sheetA.xml"])?)?;
        let link_a1 = worksheet
            .descendants()
            .find(|node| node.is_element() && node.attribute("ref") == Some("A1"))
            .unwrap();
        let link_b1 = worksheet
            .descendants()
            .find(|node| node.is_element() && node.attribute("ref") == Some("B1"))
            .unwrap();
        assert_eq!(link_a1.attribute("display"), Some("旧显示"));
        assert_eq!(link_b1.attribute("display"), Some("第二个链接"));
        assert_eq!(link_b1.attribute("location"), Some("结果!B1"));

        let rels = Document::parse(std::str::from_utf8(
            &actual["xl/worksheets/_rels/sheetA.xml.rels"],
        )?)?;
        let target_a = rels
            .descendants()
            .find(|node| node.is_element() && node.attribute("Id") == Some("rId1"))
            .unwrap();
        let target_b = rels
            .descendants()
            .find(|node| node.is_element() && node.attribute("Id") == Some("rId2"))
            .unwrap();
        assert_eq!(target_a.attribute("Target"), Some("file:///mnt/c/a.xlsx"));
        assert_eq!(target_a.attribute("TargetMode"), Some("External"));
        assert_eq!(target_b.attribute("Target"), Some("file:///mnt/c/b.xlsx"));
        assert_eq!(target_b.attribute("TargetMode"), Some("External"));
        let rels_b = Document::parse(std::str::from_utf8(
            &actual["xl/worksheets/_rels/sheetB.xml.rels"],
        )?)?;
        let target_c = rels_b
            .descendants()
            .find(|node| node.is_element() && node.attribute("Id") == Some("h1"))
            .unwrap();
        assert_eq!(target_c.attribute("Target"), Some("file:///mnt/c/c.xlsx"));
        assert_eq!(target_c.attribute("TargetMode"), Some("External"));

        fs::remove_file(source)?;
        fs::remove_file(output)?;
        Ok(())
    }
}

#[cfg(test)]
mod formula_cache_recalculation_tests {
    use super::{clear_formula_caches_for_recalculation, MAIN_NS, PACKAGE_REL_NS, REL_NS};
    use anyhow::Result;
    use roxmltree::Document;
    use std::{
        fs::{self, File},
        io::{Read, Write},
        path::PathBuf,
    };
    use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

    fn temp_path(label: &str) -> PathBuf {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        std::env::temp_dir().join(format!(
            "ooxml_formula_cache_{}_{}_{}",
            std::process::id(),
            nonce,
            label
        ))
    }

    fn fixture(path: &PathBuf) -> Result<()> {
        let root_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<x:workbook xmlns:x="{MAIN_NS}" xmlns:r="{REL_NS}"><x:sheets><x:sheet name="Formula" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>"#
        );
        let workbook_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let worksheet = format!(
            r#"<x:worksheet xmlns:x="{MAIN_NS}"><x:sheetData><x:row r="1"><x:c r="A1" t="str"><x:f>1+1</x:f><x:v>0</x:v></x:c><x:c r="B1"><x:v>9</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
        );
        let mut zip = ZipWriter::new(File::create(path)?);
        for (name, content) in [
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", worksheet),
            ("custom/keep.txt", "keep unchanged".to_owned()),
        ] {
            zip.start_file(name, SimpleFileOptions::default())?;
            zip.write_all(content.as_bytes())?;
        }
        zip.finish()?;
        Ok(())
    }

    #[test]
    fn clears_formula_caches_only_and_preserves_source_and_other_parts() -> Result<()> {
        let source = temp_path("source.xlsx");
        let output = temp_path("cacheless.xlsx");
        fixture(&source)?;
        let report = clear_formula_caches_for_recalculation(&source, &output)?;
        assert_eq!(report.formula_cells, 1);
        assert_eq!(report.cache_values_removed, 1);
        assert_eq!(report.worksheet_parts, ["xl/worksheets/sheet1.xml"]);

        let mut source_zip = ZipArchive::new(File::open(&source)?)?;
        let mut source_xml = String::new();
        source_zip
            .by_name("xl/worksheets/sheet1.xml")?
            .read_to_string(&mut source_xml)?;
        assert!(source_xml.contains("<x:v>0</x:v>"));

        let mut output_zip = ZipArchive::new(File::open(&output)?)?;
        let mut output_xml = String::new();
        output_zip
            .by_name("xl/worksheets/sheet1.xml")?
            .read_to_string(&mut output_xml)?;
        let document = Document::parse(&output_xml)?;
        let formula = document
            .descendants()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "c"
                    && node.attribute("r") == Some("A1")
            })
            .unwrap();
        assert!(formula
            .children()
            .any(|node| node.is_element() && node.tag_name().name() == "f"));
        assert!(!formula
            .children()
            .any(|node| node.is_element() && node.tag_name().name() == "v"));
        assert_eq!(formula.attribute("t"), None);
        let mut kept = String::new();
        output_zip
            .by_name("custom/keep.txt")?
            .read_to_string(&mut kept)?;
        assert_eq!(kept, "keep unchanged");

        fs::remove_file(source)?;
        fs::remove_file(output)?;
        Ok(())
    }

    #[test]
    fn rejects_shared_or_annotated_formula_shapes_without_output() -> Result<()> {
        let xml = format!(
            r#"<x:worksheet xmlns:x="{MAIN_NS}"><x:sheetData><x:row r="1"><x:c r="A1"><x:f t="shared">1+1</x:f><x:v>0</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
        );
        let error = super::clear_formula_caches_for_recalculation_xml(xml.as_bytes())
            .expect_err("shared formula shape must be rejected");
        assert!(error.to_string().contains("unsupported shared/array"));
        Ok(())
    }
}

#[cfg(test)]
mod cell_region_copy_tests {
    use super::{
        copy_cell_region, quote_unquoted_sheet_references, CellRegionCopyRequest, MAIN_NS,
        PACKAGE_REL_NS, REL_NS,
    };
    use anyhow::Result;
    use std::{
        fs::{self, File},
        io::{Read, Write},
        path::PathBuf,
    };
    use zip::{write::SimpleFileOptions, ZipWriter};

    #[test]
    fn full_width_paren_sheet_names_get_quoted() {
        let formula =
            "IF(ISBLANK(存量单位贷款信息!C5),0,单位贷款发生额信息（隐藏发放）!C5+单位贷款发生额信息（隐藏回收）!C5)";
        let rewritten = quote_unquoted_sheet_references(formula);
        assert!(rewritten.contains("'存量单位贷款信息'!C5"), "{rewritten}");
        assert!(
            rewritten.contains("'单位贷款发生额信息（隐藏发放）'!C5"),
            "{rewritten}"
        );
        assert!(!rewritten.contains("（隐藏发放）!"), "{rewritten}");
        // Idempotent.
        assert_eq!(quote_unquoted_sheet_references(&rewritten), rewritten);
    }

    #[test]
    fn string_literals_and_quoted_names_stay_untouched() {
        let formula = "SUM(INDIRECT(\"单位贷款发生额信息!\"&ADDRESS(1,1)))+('存量单位贷款信息'!B5)+[1]外簿!C6+SUM('她''s'!D1)";
        let rewritten = quote_unquoted_sheet_references(formula);
        assert!(
            rewritten.contains("INDIRECT(\"单位贷款发生额信息!\""),
            "{rewritten}"
        );
        assert!(rewritten.contains("('存量单位贷款信息'!B5)"), "{rewritten}");
        assert!(rewritten.contains("[1]外簿!C6"), "{rewritten}");
        assert!(rewritten.contains("'她''s'!D1"), "{rewritten}");
    }

    #[test]
    fn worksheet_parts_normalize_formulas_in_place() -> Result<()> {
        let xml = format!(
            r#"<worksheet xmlns="{MAIN_NS}"><sheetData><row r="5"><c r="A5"><f>存量单位贷款信息!C5</f></c><c r="B5"><f>'存量单位贷款信息'!C6</f></c></row></sheetData></worksheet>"#
        );
        let normalized = super::normalize_worksheet_formula_references(xml.as_bytes())?;
        let text = String::from_utf8(normalized)?;
        assert!(text.contains("<f>'存量单位贷款信息'!C5</f>"), "{text}");
        assert!(text.contains("<f>'存量单位贷款信息'!C6</f>"), "{text}");
        Ok(())
    }

    #[test]
    fn missing_cells_sharing_anchor_keep_column_order() -> Result<()> {
        let sheet = format!(
            r#"<worksheet xmlns="{MAIN_NS}"><sheetData><row r="5"><c r="E5"/></row></sheetData></worksheet>"#
        );
        let formulas = [
            ("F5".to_owned(), "1".to_owned()),
            ("G5".to_owned(), "2".to_owned()),
        ]
        .into_iter()
        .collect();
        let (patched, _, _) = super::apply_region_formulas(sheet.as_bytes(), "F5:G5", &formulas)?;
        let xml = String::from_utf8(patched)?;
        assert!(xml.find("r=\"E5\"").unwrap() < xml.find("r=\"F5\"").unwrap());
        assert!(xml.find("r=\"F5\"").unwrap() < xml.find("r=\"G5\"").unwrap());
        Ok(())
    }

    fn temp_path(label: &str) -> PathBuf {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        std::env::temp_dir().join(format!(
            "ooxml_cell_region_{}_{}_{}.xlsx",
            std::process::id(),
            nonce,
            label
        ))
    }

    fn workbook(path: &PathBuf, worksheet: &str) -> Result<()> {
        let root_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let book = format!(
            r#"<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        );
        let book_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let sheet = format!(r#"<worksheet xmlns="{MAIN_NS}">{worksheet}</worksheet>"#);
        let types = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned();
        let mut zip = ZipWriter::new(File::create(path)?);
        for (name, content) in [
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", book),
            ("xl/_rels/workbook.xml.rels", book_rels),
            ("[Content_Types].xml", types),
            ("xl/worksheets/sheet1.xml", sheet),
        ] {
            zip.start_file(name, SimpleFileOptions::default())?;
            zip.write_all(content.as_bytes())?;
        }
        zip.finish()?;
        Ok(())
    }

    #[test]
    fn copies_formulas_and_drops_the_template_cache() -> Result<()> {
        let template = temp_path("template");
        let target = temp_path("target");
        workbook(
            &template,
            r#"<sheetData><row r="5"><c r="F5" t="str"><f>IF(B5=0,"","错")</f><v></v></c></row></sheetData>"#,
        )?;
        workbook(
            &target,
            r#"<sheetData><row r="5"><c r="F5" s="3" t="str"><v>模板缓存</v></c><c r="G5"><v>1</v></c></row></sheetData>"#,
        )?;

        let report = copy_cell_region(
            &template,
            &target,
            &[CellRegionCopyRequest {
                source_sheet: "Sheet1".to_owned(),
                target_sheet: "Sheet1".to_owned(),
                address: "F5:G5".to_owned(),
            }],
        )?;
        assert_eq!(report.cells_with_formula, 1);
        assert_eq!(report.cells_cleared, 1);

        let bytes = fs::read(&target)?;
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes))?;
        let mut xml = String::new();
        {
            use std::io::Read;
            zip.by_name("xl/worksheets/sheet1.xml")?
                .read_to_string(&mut xml)?;
        }
        // Formula copied, template cache not carried over, target style kept.
        assert!(xml.contains("<f>IF(B5=0,\"\",\"错\")</f>"), "{xml}");
        assert!(xml.contains(r#"<c r="F5" s="3">"#), "{xml}");
        assert!(!xml.contains("模板缓存"), "{xml}");
        // A cell the template does not define is blanked but retained.
        assert!(xml.contains(r#"<c r="G5"/>"#), "{xml}");
        assert!(!xml.contains("<v>1</v>"), "{xml}");

        let _ = fs::remove_file(&template);
        let _ = fs::remove_file(&target);
        Ok(())
    }

    #[test]
    fn copies_referenced_template_sheets_missing_from_the_target() -> Result<()> {
        let template = temp_path("tpl_sheets");
        let target = temp_path("target_sheets");
        // Template: Sheet1 with a formula referencing a sheet the target lacks,
        // plus that referenced sheet carrying static data.
        workbook_with_extra_sheet(
            &template,
            r#"<sheetData><row r="5"><c r="A5"><f>参照表!B1*2</f></c></row></sheetData>"#,
            "参照表",
            r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>21</v></c></row></sheetData>"#,
        )?;
        workbook(
            &target,
            r#"<sheetData><row r="5"><c r="A5"><v>0</v></c></row></sheetData>"#,
        )?;

        let report = copy_cell_region(
            &template,
            &target,
            &[CellRegionCopyRequest {
                source_sheet: "Sheet1".to_owned(),
                target_sheet: "Sheet1".to_owned(),
                address: "A5".to_owned(),
            }],
        )?;
        assert_eq!(report.cells_with_formula, 1);
        assert_eq!(report.sheets_copied, vec!["参照表".to_owned()]);

        let bytes = fs::read(&target)?;
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes))?;
        let mut workbook = String::new();
        {
            use std::io::Read;
            zip.by_name("xl/workbook.xml")?
                .read_to_string(&mut workbook)?;
        }
        assert!(
            workbook.contains("参照表"),
            "referenced sheet must be registered: {workbook}"
        );
        let parts = zip
            .file_names()
            .filter(|name| name.starts_with("xl/worksheets/sheet"))
            .map(str::to_owned)
            .collect::<Vec<_>>();
        assert_eq!(
            parts.len(),
            2,
            "template sheet part must be added: {parts:?}"
        );
        Ok(())
    }

    #[test]
    fn recursively_copies_formula_bearing_helper_sheet_dependencies() -> Result<()> {
        let template = temp_path("recursive_template");
        let target = temp_path("recursive_target");
        workbook_with_three_sheets(
            &template,
            r#"<sheetData><row r="5"><c r="A5"><f>参照表!D2</f><v>0</v></c></row></sheetData>"#,
            "参照表",
            r#"<sheetData><row r="2"><c r="D2"><f>基础参数!A1*2</f><v>0</v></c></row></sheetData>"#,
            "基础参数",
            r#"<sheetData><row r="1"><c r="A1"><v>21</v></c></row></sheetData>"#,
        )?;
        workbook(
            &target,
            r#"<sheetData><row r="5"><c r="A5"><v>0</v></c></row></sheetData>"#,
        )?;

        let report = copy_cell_region(
            &template,
            &target,
            &[CellRegionCopyRequest {
                source_sheet: "Sheet1".to_owned(),
                target_sheet: "Sheet1".to_owned(),
                address: "A5".to_owned(),
            }],
        )?;
        assert_eq!(report.cells_with_formula, 1);
        assert_eq!(report.sheets_copied.len(), 2);
        assert!(report.sheets_copied.iter().any(|sheet| sheet == "参照表"));
        assert!(report.sheets_copied.iter().any(|sheet| sheet == "基础参数"));

        let bytes = fs::read(&target)?;
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes))?;
        let mut workbook_xml = String::new();
        {
            use std::io::Read;
            zip.by_name("xl/workbook.xml")?
                .read_to_string(&mut workbook_xml)?;
        }
        assert!(workbook_xml.contains("参照表"), "{workbook_xml}");
        assert!(workbook_xml.contains("基础参数"), "{workbook_xml}");
        let formula_parts = zip
            .file_names()
            .filter(|name| name.starts_with("xl/worksheets/sheet"))
            .map(str::to_owned)
            .collect::<Vec<_>>();
        assert_eq!(formula_parts.len(), 3, "{formula_parts:?}");
        let all_parts = formula_parts
            .iter()
            .map(|part| {
                let mut xml = String::new();
                zip.by_name(part)?.read_to_string(&mut xml)?;
                Ok::<_, anyhow::Error>(xml)
            })
            .collect::<Result<Vec<_>>>()?
            .join("\n");
        assert!(all_parts.contains("'参照表'!D2"), "{all_parts}");
        assert!(all_parts.contains("'基础参数'!A1*2"), "{all_parts}");

        let _ = fs::remove_file(&template);
        let _ = fs::remove_file(&target);
        Ok(())
    }

    /// Two-sheet workbook whose second sheet is registered in workbook.xml,
    /// its relationships and [Content_Types].xml.
    fn workbook_with_extra_sheet(
        path: &PathBuf,
        first_sheet: &str,
        extra_name: &str,
        extra_sheet: &str,
    ) -> Result<()> {
        let root_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let book = format!(
            r#"<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="{extra_name}" sheetId="2" r:id="rId2"/></sheets></workbook>"#
        );
        let book_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#
        );
        let types = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned();
        let mut zip = ZipWriter::new(File::create(path)?);
        for (name, content) in [
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", book),
            ("xl/_rels/workbook.xml.rels", book_rels),
            ("[Content_Types].xml", types),
            (
                "xl/worksheets/sheet1.xml",
                format!(r#"<worksheet xmlns="{MAIN_NS}">{first_sheet}</worksheet>"#),
            ),
            (
                "xl/worksheets/sheet2.xml",
                format!(r#"<worksheet xmlns="{MAIN_NS}">{extra_sheet}</worksheet>"#),
            ),
        ] {
            zip.start_file(name, SimpleFileOptions::default())?;
            zip.write_all(content.as_bytes())?;
        }
        zip.finish()?;
        Ok(())
    }

    fn workbook_with_three_sheets(
        path: &PathBuf,
        first_sheet: &str,
        second_name: &str,
        second_sheet: &str,
        third_name: &str,
        third_sheet: &str,
    ) -> Result<()> {
        let root_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let book = format!(
            r#"<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="{second_name}" sheetId="2" r:id="rId2"/><sheet name="{third_name}" sheetId="3" r:id="rId3"/></sheets></workbook>"#
        );
        let book_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>"#
        );
        let types = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;
        let mut zip = ZipWriter::new(File::create(path)?);
        for (name, content) in [
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", book),
            ("xl/_rels/workbook.xml.rels", book_rels),
            ("[Content_Types].xml", types.to_owned()),
            (
                "xl/worksheets/sheet1.xml",
                format!(r#"<worksheet xmlns="{MAIN_NS}">{first_sheet}</worksheet>"#),
            ),
            (
                "xl/worksheets/sheet2.xml",
                format!(r#"<worksheet xmlns="{MAIN_NS}">{second_sheet}</worksheet>"#),
            ),
            (
                "xl/worksheets/sheet3.xml",
                format!(r#"<worksheet xmlns="{MAIN_NS}">{third_sheet}</worksheet>"#),
            ),
        ] {
            zip.start_file(name, SimpleFileOptions::default())?;
            zip.write_all(content.as_bytes())?;
        }
        zip.finish()?;
        Ok(())
    }

    #[test]
    fn creates_rows_and_cells_missing_from_the_target() -> Result<()> {
        let template = temp_path("template_missing");
        let target = temp_path("target_missing");
        workbook(
            &template,
            r#"<sheetData><row r="5"><c r="F5"><f>A5*2</f></c></row><row r="6"><c r="F6"><f>A6*2</f></c></row></sheetData>"#,
        )?;
        workbook(&target, r#"<sheetData></sheetData>"#)?;

        copy_cell_region(
            &template,
            &target,
            &[CellRegionCopyRequest {
                source_sheet: "Sheet1".to_owned(),
                target_sheet: "Sheet1".to_owned(),
                address: "F5:F6".to_owned(),
            }],
        )?;

        let bytes = fs::read(&target)?;
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes))?;
        let mut xml = String::new();
        {
            use std::io::Read;
            zip.by_name("xl/worksheets/sheet1.xml")?
                .read_to_string(&mut xml)?;
        }
        // Rows must stay siblings inside sheetData, never nested.
        assert!(
            xml.contains(r#"<row r="5"><c r="F5"><f>A5*2</f></c></row>"#),
            "{xml}"
        );
        assert!(
            xml.contains(r#"<row r="6"><c r="F6"><f>A6*2</f></c></row>"#),
            "{xml}"
        );
        assert!(xml.contains("</sheetData>"), "{xml}");

        let _ = fs::remove_file(&template);
        let _ = fs::remove_file(&target);
        Ok(())
    }
}

#[cfg(test)]
mod shared_formula_translation_tests {
    use super::translate_shared_formula;

    #[test]
    fn shifts_relative_references_and_keeps_absolute_ones() -> anyhow::Result<()> {
        // Real master from a bank template: K6 = "+B6" over K6:K37.
        assert_eq!(translate_shared_formula("+B6", (6, 11), (7, 11))?, "+B7");
        // SUM(C6) over L6:L25.
        assert_eq!(
            translate_shared_formula("SUM(C6)", (6, 12), (10, 12))?,
            "SUM(C10)"
        );
        // Absolute references must not move.
        assert_eq!(
            translate_shared_formula("$B$6+$B6+B$6", (6, 12), (9, 12))?,
            "$B$6+$B9+B$6"
        );
        // Horizontal move only shifts the column.
        assert_eq!(translate_shared_formula("A5", (5, 1), (5, 3))?, "C5");
        Ok(())
    }

    #[test]
    fn leaves_text_literals_function_names_and_names_alone() -> anyhow::Result<()> {
        // String literals are not references.
        assert_eq!(
            translate_shared_formula(r#"IF(A1=0,"B2","")"#, (1, 1), (2, 1))?,
            r#"IF(A2=0,"B2","")"#
        );
        // Function names with digits must not be rewritten.
        assert_eq!(
            translate_shared_formula("LOG10(A1)", (1, 1), (3, 1))?,
            "LOG10(A3)"
        );
        // Defined names that look like references stay intact.
        assert_eq!(
            translate_shared_formula("MyName1+A1", (1, 1), (2, 1))?,
            "MyName1+A2"
        );
        Ok(())
    }

    #[test]
    fn translates_real_indirect_formula_from_a_bank_template() -> anyhow::Result<()> {
        // N6 master = SUM(INDIRECT("单位贷款发生额信息!"&ADDRESS(ROW(E6),COLUMN(C6))))
        let master = r#"SUM(INDIRECT("单位贷款发生额信息!"&ADDRESS(ROW(E6),COLUMN(C6))))"#;
        let translated = translate_shared_formula(master, (6, 14), (9, 14))?;
        assert_eq!(
            translated,
            r#"SUM(INDIRECT("单位贷款发生额信息!"&ADDRESS(ROW(E9),COLUMN(C9))))"#
        );
        Ok(())
    }
}

#[cfg(test)]
mod sheet_reference_detection_tests {
    use super::sheet_references_in_formula;

    #[test]
    fn finds_plain_and_quoted_sheet_references() {
        assert_eq!(
            sheet_references_in_formula("INDEX(集中系统数据!$1:$1048576,1)"),
            vec!["集中系统数据".to_owned()]
        );
        assert_eq!(
            sheet_references_in_formula("VLOOKUP($D$3,参照表!$A:$B,2,0)"),
            vec!["参照表".to_owned()]
        );
        assert_eq!(
            sheet_references_in_formula("'My Sheet'!A1+'Other 2'!B2"),
            vec!["My Sheet".to_owned(), "Other 2".to_owned()]
        );
    }

    #[test]
    fn ignores_string_literals() {
        // The INDIRECT text must not be treated as a live sheet reference for
        // copy planning; the runtime string is resolved by Excel.
        assert!(sheet_references_in_formula(
            r#"SUM(INDIRECT("单位贷款发生额信息!"&ADDRESS(1,1)))"#
        )
        .is_empty());
    }
}

#[cfg(test)]
mod relationship_target_tests {
    use super::relative_target;

    #[test]
    fn resolves_sibling_and_nested_parts() {
        assert_eq!(
            relative_target("xl/worksheets/sheet2.xml", "xl/comments2.xml"),
            "../comments2.xml"
        );
        assert_eq!(
            relative_target("xl/worksheets/sheet2.xml", "xl/drawings/vmlDrawing2.vml"),
            "../drawings/vmlDrawing2.vml"
        );
        assert_eq!(
            relative_target("xl/worksheets/sheet2.xml", "xl/worksheets/sheet2.xml"),
            "sheet2.xml"
        );
    }
}
