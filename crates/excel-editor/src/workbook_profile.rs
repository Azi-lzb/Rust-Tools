//! Conservative XLSX package inventory used by the edit router.

use crate::{WorkbookFeature, WorkbookFormat, WorkbookProfile};
use anyhow::{Context, Result};
use roxmltree::Document;
use std::{collections::BTreeSet, fs::File, io::Read, path::Path};
use zip::ZipArchive;

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// Inspect the OOXML parts and worksheet structures that affect edit routing.
/// Unknown package parts set `inventory_complete` to false so the caller cannot
/// route a rewrite backend on an incomplete preservation inventory.
pub fn inspect_xlsx_profile(path: impl AsRef<Path>) -> Result<WorkbookProfile> {
    let path = path.as_ref();
    anyhow::ensure!(
        path.extension()
            .and_then(|value| value.to_str())
            .is_some_and(|value| value.eq_ignore_ascii_case("xlsx")),
        "the current package inspector accepts .xlsx only"
    );
    let mut archive =
        ZipArchive::new(File::open(path).with_context(|| format!("open {}", path.display()))?)?;
    let mut names = Vec::with_capacity(archive.len());
    let mut parts = BTreeSet::new();
    let mut features = BTreeSet::new();
    let mut inventory_complete = true;

    for index in 0..archive.len() {
        let mut entry = archive.by_index(index)?;
        if entry.is_dir() {
            continue;
        }
        let name = entry.name().to_owned();
        let lower = name.to_ascii_lowercase();
        names.push(name.clone());
        parts.insert(name.clone());

        if lower.ends_with(".rels") {
            features.insert(WorkbookFeature::Relationships);
        }
        if lower.starts_with("xl/comments") && lower.ends_with(".xml") {
            features.insert(WorkbookFeature::Comments);
        }
        if lower.starts_with("xl/charts/") {
            features.insert(WorkbookFeature::Charts);
        } else if lower.starts_with("xl/pivot") {
            features.insert(WorkbookFeature::PivotTables);
        } else if lower.starts_with("xl/slicers/") {
            features.insert(WorkbookFeature::Slicers);
        } else if lower.starts_with("xl/externallinks/") {
            features.insert(WorkbookFeature::ExternalLinks);
        } else if lower == "xl/vbaproject.bin" {
            features.insert(WorkbookFeature::Macros);
        } else if lower.starts_with("xl/activex/") {
            features.insert(WorkbookFeature::ActiveX);
        } else if lower.starts_with("xl/drawings/") {
            if lower.contains("vml") {
                // VML is the legacy comment shape container.
            } else {
                features.insert(WorkbookFeature::Drawings);
            }
        } else if lower.starts_with("xl/media/") {
            features.insert(WorkbookFeature::Images);
        } else if lower.starts_with("_xmlsignatures/") || lower.starts_with("xl/_xmlsignatures/") {
            features.insert(WorkbookFeature::DigitalSignatures);
        }

        if lower == "xl/workbook.xml" {
            let mut xml = String::new();
            entry
                .read_to_string(&mut xml)
                .context("read workbook.xml")?;
            let document = Document::parse(&xml).context("parse workbook.xml")?;
            if document.descendants().any(|node| {
                node.is_element()
                    && node.tag_name().name() == "definedNames"
                    && node.tag_name().namespace() == Some(MAIN_NS)
            }) {
                features.insert(WorkbookFeature::DefinedNames);
            }
        }

        if lower.starts_with("xl/worksheets/")
            && lower.ends_with(".xml")
            && !lower.contains("/_rels/")
        {
            let mut xml = String::new();
            entry
                .read_to_string(&mut xml)
                .with_context(|| format!("read {name}"))?;
            let document = Document::parse(&xml).with_context(|| format!("parse {name}"))?;
            for node in document
                .descendants()
                .filter(|node| node.is_element() && node.tag_name().namespace() == Some(MAIN_NS))
            {
                match node.tag_name().name() {
                    "mergeCells" => {
                        features.insert(WorkbookFeature::MergedCells);
                    }
                    "conditionalFormatting" => {
                        features.insert(WorkbookFeature::ConditionalFormatting);
                    }
                    "dataValidations" => {
                        features.insert(WorkbookFeature::DataValidations);
                    }
                    "sheetProtection" => {
                        features.insert(WorkbookFeature::SheetProtection);
                    }
                    "headerFooter" => {
                        features.insert(WorkbookFeature::HeadersFooters);
                    }
                    "pane" => {
                        features.insert(WorkbookFeature::FreezePanes);
                    }
                    "autoFilter" => {
                        features.insert(WorkbookFeature::AutoFilter);
                    }
                    "f" => {
                        features.insert(WorkbookFeature::Formulas);
                    }
                    _ => {}
                }
            }
        }

        if !known_part(&lower) {
            features.insert(WorkbookFeature::UnknownExtensions);
            inventory_complete = false;
        }
    }

    // Catch duplicate names in malformed archives before treating the inventory as complete.
    if names.len() != parts.len() {
        inventory_complete = false;
        features.insert(WorkbookFeature::UnknownExtensions);
    }

    Ok(WorkbookProfile {
        format: WorkbookFormat::Xlsx,
        inventory_complete,
        features_present: features,
    })
}

fn known_part(name: &str) -> bool {
    name == "[content_types].xml"
        || name.starts_with("_rels/")
        || name.starts_with("docprops/")
        || name == "xl/workbook.xml"
        || name == "xl/styles.xml"
        || name == "xl/sharedstrings.xml"
        || name == "xl/calcchain.xml"
        || name.starts_with("xl/theme/")
        || name.starts_with("xl/worksheets/")
        || name.starts_with("xl/comments")
        // openpyxl names legacy comment VML drawings `commentsDrawingN.vml`;
        // the VML marker is not necessarily the first segment after drawings/.
        || (name.starts_with("xl/drawings/") && name.ends_with(".vml"))
        || name.starts_with("xl/printersettings/")
        || name.ends_with(".rels")
}

#[cfg(test)]
mod tests {
    use super::{inspect_xlsx_profile, known_part};
    use crate::WorkbookFeature;
    use std::{fs, io};
    use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

    #[test]
    fn recognizes_openpyxl_legacy_comment_vml_part_name() {
        assert!(known_part("xl/drawings/commentsDrawing1.vml"));
    }

    #[test]
    fn ignores_zip_directory_entries_but_keeps_file_part_inventory() {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let source = std::env::temp_dir().join(format!(
            "excel_profile_directory_source_{}_{}.xlsx",
            std::process::id(),
            nonce
        ));
        let output = std::env::temp_dir().join(format!(
            "excel_profile_directory_output_{}_{}.xlsx",
            std::process::id(),
            nonce
        ));
        let mut workbook = rust_xlsxwriter::Workbook::new();
        workbook.add_worksheet().set_name("Data").unwrap();
        workbook.save(&source).unwrap();

        let mut input = ZipArchive::new(fs::File::open(&source).unwrap()).unwrap();
        let mut writer = ZipWriter::new(fs::File::create(&output).unwrap());
        for index in 0..input.len() {
            let mut entry = input.by_index(index).unwrap();
            if entry.is_dir() {
                continue;
            }
            let name = entry.name().to_owned();
            writer
                .start_file(name, SimpleFileOptions::default())
                .unwrap();
            io::copy(&mut entry, &mut writer).unwrap();
        }
        writer
            .add_directory("xl/drawings/", SimpleFileOptions::default())
            .unwrap();
        writer.finish().unwrap();

        let profile = inspect_xlsx_profile(&output).unwrap();
        assert!(profile.inventory_complete);
        assert!(!profile
            .features_present
            .contains(&WorkbookFeature::UnknownExtensions));

        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }
}
