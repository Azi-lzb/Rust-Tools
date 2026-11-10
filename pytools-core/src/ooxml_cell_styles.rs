//! Narrow OOXML cell-style reader for the 3-1 Rust print engine.
//!
//! calamine (via `excel_reader`) does not expose per-cell style metadata, so
//! this module reads exactly what the print engine needs straight from the
//! source package: `styles.xml` tables, per-worksheet cell `s` indices, and
//! `theme1.xml` color slots. Nothing else is parsed.
//!
//! Output is a hashable [`CellStyleInfo`] per source address, which the print
//! writer maps onto `rust_xlsxwriter::Format` objects with deduplication.

use anyhow::{Context, Result};
use std::collections::BTreeMap;
use std::io::Read;
use std::path::Path;

use zip::ZipArchive;

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_DOC_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// Legacy indexed-color palette (Excel 97 standard 64 entries).
const INDEXED_COLORS: [&str; 64] = [
    "FF000000", "FFFFFFFF", "FFFF0000", "FF00FF00", "FF0000FF", "FFFFFF00", "FFFF00FF", "FF00FFFF",
    "FF800000", "FF008000", "FF000080", "FF808000", "FF800080", "FF008080", "FFC0C0C0", "FF808080",
    "FF9999FF", "FF993366", "FFFFFFCC", "FFCCFFFF", "FF660066", "FF808000", "FF00AA00", "FFAA7F3C",
    "FF7F7F7F", "FFDDDDDD", "FFB3B3B3", "FF666666", "FF808080", "FFB36666", "FFB3B366", "FFB3B3B3",
    "FF66B366", "FF66B3B3", "FFB366B3", "FF6666B3", "FFB3993C", "FF3F3F3F", "FF5B5B5B", "FF7F7F7F",
    "FFB3B3B3", "FF33B333", "FF33B3B3", "FFB33333", "FF3333B3", "FFB333B3", "FF3333B3", "FF666666",
    "FFCC9999", "FFCC99CC", "FFCCCC99", "FFCCFFCC", "FFCCFFFF", "FFCC99CC", "FFCCCC66", "FF99CCCC",
    "FFCC99CC", "FFCC9999", "FF9999CC", "FFCCCC99", "FF99CCCC", "FFCC99CC", "FFCC9999", "FF3F3F3F",
];

#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
pub struct FontInfo {
    pub bold: bool,
    pub italic: bool,
    pub size: Option<String>,
    pub name: Option<String>,
    pub color: Option<String>,
    pub underline: Option<String>,
    pub strike: bool,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
pub struct FillInfo {
    pub pattern: Option<String>,
    pub fg_color: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
pub struct BorderInfo {
    pub left: Option<(Option<String>, Option<String>)>,
    pub right: Option<(Option<String>, Option<String>)>,
    pub top: Option<(Option<String>, Option<String>)>,
    pub bottom: Option<(Option<String>, Option<String>)>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
pub struct AlignInfo {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: bool,
}

/// Hashable style descriptor resolved for one source `xf` entry.
#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
pub struct CellStyleInfo {
    pub num_fmt: Option<String>,
    pub font: Option<FontInfo>,
    pub fill: Option<FillInfo>,
    pub border: Option<BorderInfo>,
    pub align: Option<AlignInfo>,
}

/// Per-workbook style tables parsed once and reused for every import.
#[derive(Debug, Default)]
pub struct SourceStyles {
    /// cellXfs entries with resolved numFmt codes
    xfs: Vec<CellXfEntry>,
    theme_slots: Vec<String>,
    /// worksheet name (case-preserving) -> package part path
    sheet_parts: BTreeMap<String, String>,
    tables: std::sync::OnceLock<StylesheetTablesSnapshot>,
}

#[derive(Clone, Debug, Default)]
struct StylesheetTablesSnapshot {
    fonts: Vec<FontRaw>,
    fills: Vec<FillRaw>,
    borders: Vec<BorderRaw>,
}

#[derive(Clone, Debug, Default)]
struct CellXfEntry {
    num_fmt_id: u32,
    num_fmt_code: Option<String>,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    align: Option<AlignInfo>,
    apply_number_format: bool,
    apply_font: bool,
    apply_fill: bool,
    apply_border: bool,
    apply_alignment: bool,
}

impl SourceStyles {
    /// Load and parse the style tables of a source workbook.
    pub fn load(path: &Path) -> Result<Self> {
        let file = std::fs::File::open(path)
            .with_context(|| format!("open source workbook for styles: {}", path.display()))?;
        let mut archive = zip::ZipArchive::new(file).context("open OOXML package for styles")?;
        let workbook_xml = read_zip_part(&mut archive, "xl/workbook.xml")?;
        let rels_xml = read_zip_part(&mut archive, "xl/_rels/workbook.xml.rels")?;
        let workbook_text =
            std::str::from_utf8(&workbook_xml).context("workbook.xml must be UTF-8")?;
        let rels_text = std::str::from_utf8(&rels_xml).context("workbook rels must be UTF-8")?;
        let sheet_parts = resolve_sheet_parts(workbook_text, rels_text)?;

        let theme_slots = read_zip_part(&mut archive, "xl/theme/theme1.xml")
            .ok()
            .and_then(|bytes| {
                std::str::from_utf8(&bytes)
                    .ok()
                    .map(|xml| parse_theme_slots(xml))
            })
            .unwrap_or_default();
        let styles_xml = read_zip_part(&mut archive, "xl/styles.xml")?;
        let styles_text = std::str::from_utf8(&styles_xml).context("styles.xml must be UTF-8")?;
        let (xfs, custom_fmts, tables) = parse_stylesheet(styles_text)?;
        let mut tables_snapshot = std::sync::OnceLock::new();
        tables_snapshot
            .set(StylesheetTablesSnapshot {
                fonts: tables.fonts,
                fills: tables.fills,
                borders: tables.borders,
            })
            .ok();
        Ok(Self {
            xfs: attach_num_fmt_codes(xfs, &custom_fmts),
            theme_slots,
            sheet_parts,
            tables: tables_snapshot,
        })
    }

    /// Read the `s` style indices of one worksheet's cells.
    pub fn cell_style_indices(
        &self,
        path: &Path,
        sheet_name: &str,
    ) -> Result<BTreeMap<String, u32>> {
        let part = self
            .sheet_parts
            .get(sheet_name)
            .with_context(|| format!("worksheet {sheet_name:?} has no package part for styles"))?;
        let file = std::fs::File::open(path)
            .with_context(|| format!("open source workbook for cell styles: {}", path.display()))?;
        let mut archive = zip::ZipArchive::new(file).context("reopen OOXML package")?;
        let xml = read_zip_part(&mut archive, part)?;
        let text = std::str::from_utf8(&xml).context("worksheet XML must be UTF-8")?;
        parse_cell_style_indices(text)
    }

    /// Resolve one style index into a hashable descriptor.
    pub fn descriptor(&self, style_index: u32) -> CellStyleInfo {
        let Some(xf) = self.xfs.get(style_index as usize) else {
            return CellStyleInfo::default();
        };
        // 注意：不用 apply* 属性门控。OOXML 规范里 apply* 只影响父样式继承
        // 语义，openpyxl/Excel 写出的 cellXfs 经常省略它们，而 numFmtId/
        // fontId/fillId/borderId 始终是该 xf 的权威格式来源。
        let num_fmt = if xf.num_fmt_id != 0 {
            xf.num_fmt_code.clone().filter(|code| !code.is_empty())
        } else {
            None
        };
        let font = if xf.font_id != 0 {
            self.font_at(xf.font_id)
        } else {
            None
        };
        let fill = if xf.fill_id != 0 {
            self.fill_at(xf.fill_id)
        } else {
            None
        };
        let border = if xf.border_id != 0 {
            self.border_at(xf.border_id)
        } else {
            None
        };
        let align = xf.align.clone();
        CellStyleInfo {
            num_fmt,
            font,
            fill,
            border,
            align,
        }
    }

    fn tables(&self) -> StylesheetTablesSnapshot {
        self.tables.get().cloned().unwrap_or_default()
    }

    fn font_at(&self, id: usize) -> Option<FontInfo> {
        let font = self.tables().fonts.get(id).cloned()?;
        Some(FontInfo {
            bold: font.bold,
            italic: font.italic,
            strike: font.strike,
            size: font.size.clone(),
            name: font.name.clone(),
            color: font
                .color
                .as_ref()
                .and_then(|color| self.resolve_color(color)),
            underline: font.underline.clone(),
        })
    }

    fn fill_at(&self, id: usize) -> Option<FillInfo> {
        let fill = self.tables().fills.get(id).cloned()?;
        let pattern = fill.pattern.clone().filter(|p| p != "none");
        let fg = fill
            .fg_color
            .as_ref()
            .and_then(|color| self.resolve_color(color));
        pattern.map(|pattern| FillInfo {
            pattern: Some(pattern),
            fg_color: fg,
        })
    }

    fn border_at(&self, id: usize) -> Option<BorderInfo> {
        let border = self.tables().borders.get(id).cloned()?;
        let edge = |edge: &Option<EdgeRaw>| {
            edge.as_ref().map(|edge| {
                (
                    edge.style.clone(),
                    edge.color
                        .as_ref()
                        .and_then(|color| self.resolve_color(color)),
                )
            })
        };
        let left = edge(&border.left);
        let right = edge(&border.right);
        let top = edge(&border.top);
        let bottom = edge(&border.bottom);
        if left.is_none() && right.is_none() && top.is_none() && bottom.is_none() {
            None
        } else {
            Some(BorderInfo {
                left,
                right,
                top,
                bottom,
            })
        }
    }

    fn resolve_color(&self, color: &ColorRaw) -> Option<String> {
        if let Some(rgb) = &color.rgb {
            return Some(normalize_rgb(rgb));
        }
        if let Some(indexed) = color.indexed {
            let index = indexed as usize;
            return INDEXED_COLORS
                .get(index)
                .map(|value| (*value).to_string())
                .or_else(|| INDEXED_COLORS.first().map(|value| (*value).to_string()));
        }
        if let Some(theme) = color.theme {
            let slot = self.theme_slots.get(theme as usize)?;
            return Some(apply_tint(slot, color.tint));
        }
        None
    }
}

#[derive(Clone, Debug, Default)]
struct FontRaw {
    bold: bool,
    italic: bool,
    strike: bool,
    underline: Option<String>,
    size: Option<String>,
    name: Option<String>,
    color: Option<ColorRaw>,
}

#[derive(Clone, Debug, Default)]
struct FillRaw {
    pattern: Option<String>,
    fg_color: Option<ColorRaw>,
}

#[derive(Clone, Debug, Default)]
struct EdgeRaw {
    style: Option<String>,
    color: Option<ColorRaw>,
}

#[derive(Clone, Debug, Default)]
struct BorderRaw {
    left: Option<EdgeRaw>,
    right: Option<EdgeRaw>,
    top: Option<EdgeRaw>,
    bottom: Option<EdgeRaw>,
}

#[derive(Clone, Debug, Default)]
struct ColorRaw {
    rgb: Option<String>,
    theme: Option<u32>,
    indexed: Option<u32>,
    tint: f64,
}

struct StylesheetTables {
    fonts: Vec<FontRaw>,
    fills: Vec<FillRaw>,
    borders: Vec<BorderRaw>,
}

fn parse_stylesheet(
    xml: &str,
) -> Result<(Vec<CellXfEntry>, BTreeMap<u32, String>, StylesheetTables)> {
    let doc = roxmltree::Document::parse(xml).context("parse styles.xml")?;
    let root = doc.root_element();
    let mut custom_fmts = BTreeMap::new();
    if let Some(node) = find_child(&root, "numFmts") {
        for fmt in element_children(node) {
            if node_name(&fmt) != "numFmt" {
                continue;
            }
            if let (Some(id), Some(code)) = (attr(&fmt, "numFmtId"), attr(&fmt, "formatCode")) {
                if let Ok(id) = id.parse::<u32>() {
                    custom_fmts.insert(id, code.to_owned());
                }
            }
        }
    }
    let tables = collect_style_tables(root);
    let mut xfs = Vec::new();
    if let Some(node) = find_child(&root, "cellXfs") {
        for xf in element_children(node) {
            if node_name(&xf) != "xf" {
                continue;
            }
            let num_fmt_id = attr(&xf, "numFmtId")
                .and_then(|v| v.parse().ok())
                .unwrap_or(0u32);
            let font_id = attr(&xf, "fontId")
                .and_then(|v| v.parse().ok())
                .unwrap_or(0usize);
            let fill_id = attr(&xf, "fillId")
                .and_then(|v| v.parse().ok())
                .unwrap_or(0usize);
            let border_id = attr(&xf, "borderId")
                .and_then(|v| v.parse().ok())
                .unwrap_or(0usize);
            let apply = |name: &str| {
                attr(&xf, name)
                    .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
                    .unwrap_or(false)
            };
            let mut align = None;
            if let Some(alignment) = xf
                .children()
                .find(|c| c.is_element() && node_name(c) == "alignment")
            {
                let horizontal = attr(&alignment, "horizontal").map(str::to_owned);
                let vertical = attr(&alignment, "vertical").map(str::to_owned);
                let wrap_text = attr(&alignment, "wrapText")
                    .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
                    .unwrap_or(false);
                if horizontal.is_some() || vertical.is_some() || wrap_text {
                    align = Some(AlignInfo {
                        horizontal,
                        vertical,
                        wrap_text,
                    });
                }
            }
            xfs.push(CellXfEntry {
                num_fmt_id,
                num_fmt_code: None,
                font_id,
                fill_id,
                border_id,
                align: align.clone(),
                apply_number_format: apply("applyNumberFormat"),
                apply_font: apply("applyFont"),
                apply_fill: apply("applyFill"),
                apply_border: apply("applyBorder"),
                apply_alignment: apply("applyAlignment") || align.is_some(),
            });
        }
    }
    Ok((xfs, custom_fmts, tables))
}

fn attach_num_fmt_codes(
    mut xfs: Vec<CellXfEntry>,
    custom_fmts: &BTreeMap<u32, String>,
) -> Vec<CellXfEntry> {
    for xf in &mut xfs {
        xf.num_fmt_code =
            builtin_num_fmt(xf.num_fmt_id).or_else(|| custom_fmts.get(&xf.num_fmt_id).cloned());
    }
    xfs
}

fn collect_style_tables(root: roxmltree::Node) -> StylesheetTables {
    let mut tables = StylesheetTables {
        fonts: Vec::new(),
        fills: Vec::new(),
        borders: Vec::new(),
    };
    let parse_color = |node: &roxmltree::Node| -> Option<ColorRaw> {
        let rgb = attr(&node, "rgb").map(str::to_owned);
        let theme = attr(&node, "theme").and_then(|v| v.parse().ok());
        let indexed = attr(&node, "indexed").and_then(|v| v.parse().ok());
        let tint = attr(&node, "tint")
            .and_then(|v| v.parse().ok())
            .unwrap_or(0.0);
        if rgb.is_none() && theme.is_none() && indexed.is_none() {
            None
        } else {
            Some(ColorRaw {
                rgb,
                theme,
                indexed,
                tint,
            })
        }
    };
    if let Some(fonts_node) = find_child(&root, "fonts") {
        for font in element_children(fonts_node) {
            if node_name(&font) != "font" {
                continue;
            }
            let mut raw = FontRaw::default();
            for child in element_children(font) {
                match node_name(&child) {
                    "b" => raw.bold = true,
                    "i" => raw.italic = true,
                    "strike" => raw.strike = true,
                    "u" => raw.underline = Some(attr(&child, "val").unwrap_or("single").to_owned()),
                    "sz" => raw.size = attr(&child, "val").map(str::to_owned),
                    "name" => raw.name = attr(&child, "val").map(str::to_owned),
                    "color" => raw.color = parse_color(&child),
                    _ => {}
                }
            }
            tables.fonts.push(raw);
        }
    }
    if let Some(fills_node) = find_child(&root, "fills") {
        for fill in element_children(fills_node) {
            if node_name(&fill) != "fill" {
                continue;
            }
            let mut raw = FillRaw::default();
            if let Some(pattern) = fill
                .children()
                .find(|c| c.is_element() && node_name(c) == "patternFill")
            {
                raw.pattern = attr(&pattern, "patternType").map(str::to_owned);
                if let Some(fg) = pattern
                    .children()
                    .find(|c| c.is_element() && node_name(c) == "fgColor")
                {
                    raw.fg_color = parse_color(&fg);
                }
            }
            tables.fills.push(raw);
        }
    }
    if let Some(borders_node) = find_child(&root, "borders") {
        for border in element_children(borders_node) {
            if node_name(&border) != "border" {
                continue;
            }
            let parse_edge = |name: &str| -> Option<EdgeRaw> {
                border
                    .children()
                    .find(|c| c.is_element() && node_name(c) == name)
                    .map(|edge| EdgeRaw {
                        style: attr(&edge, "style").map(str::to_owned),
                        color: edge
                            .children()
                            .find(|c| c.is_element() && node_name(c) == "color")
                            .and_then(|color| parse_color(&color)),
                    })
            };
            tables.borders.push(BorderRaw {
                left: parse_edge("left"),
                right: parse_edge("right"),
                top: parse_edge("top"),
                bottom: parse_edge("bottom"),
            });
        }
    }
    tables
}

impl StylesheetTables {
    fn font_raw(&self, id: usize) -> Option<FontRaw> {
        self.fonts.get(id).cloned()
    }
    fn fill_raw(&self, id: usize) -> Option<FillRaw> {
        self.fills.get(id).cloned()
    }
    fn border_raw(&self, id: usize) -> Option<BorderRaw> {
        self.borders.get(id).cloned()
    }
}

/// Built-in number format codes that 3-1 templates actually use.
fn builtin_num_fmt(id: u32) -> Option<String> {
    let code = match id {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "yyyy-m-d",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "yyyy-m-d h:mm",
        37 => "#,##0;(#,##0)",
        38 => "#,##0;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    };
    Some(code.to_owned())
}

fn parse_cell_style_indices(xml: &str) -> Result<BTreeMap<String, u32>> {
    let doc = roxmltree::Document::parse(xml).context("parse worksheet for cell styles")?;
    let mut out = BTreeMap::new();
    for cell in doc
        .root_element()
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "c")
    {
        let Some(address) = attr(&cell, "r") else {
            continue;
        };
        let Some(style) = attr(&cell, "s").and_then(|v| v.parse::<u32>().ok()) else {
            continue;
        };
        if style == 0 {
            continue;
        }
        out.insert(address.to_owned(), style);
    }
    Ok(out)
}

fn resolve_sheet_parts(workbook_xml: &str, rels_xml: &str) -> Result<BTreeMap<String, String>> {
    let wb_doc = roxmltree::Document::parse(workbook_xml).context("parse workbook.xml")?;
    let rels_doc = roxmltree::Document::parse(rels_xml).context("parse workbook rels")?;
    let mut rel_targets = BTreeMap::new();
    for rel in rels_doc
        .root()
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "Relationship")
    {
        if let (Some(id), Some(target)) = (attr(&rel, "Id"), attr(&rel, "Target")) {
            rel_targets.insert(id.to_owned(), target.to_owned());
        }
    }
    let mut parts = BTreeMap::new();
    let sheets_node = wb_doc
        .root_element()
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "sheets")
        .context("workbook.xml has no sheets element")?;
    for sheet in sheets_node
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "sheet")
    {
        let Some(name) = attr(&sheet, "name") else {
            continue;
        };
        let Some(rid) = sheet
            .attribute((REL_DOC_NS, "id"))
            .or_else(|| sheet.attribute(("r", "id")))
        else {
            continue;
        };
        let Some(target) = rel_targets.get(rid) else {
            continue;
        };
        let part = if target.starts_with('/') {
            target.trim_start_matches('/').to_owned()
        } else {
            format!("xl/{target}")
        };
        parts.insert(name.to_owned(), part);
    }
    Ok(parts)
}

fn parse_theme_slots(xml: &str) -> Vec<String> {
    let Ok(doc) = roxmltree::Document::parse(xml) else {
        return Vec::new();
    };
    // clrScheme order: dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink.
    // Theme color indices follow Excel's swapped order: 0=lt1, 1=dk1, 2=lt2, 3=dk2.
    let Some(clr_scheme) = doc
        .root_element()
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "clrScheme")
    else {
        return Vec::new();
    };
    let read_slot = |node: roxmltree::Node| -> Option<String> {
        for child in element_children(node) {
            if let Some(rgb) = attr(&child, "lastClr") {
                return Some(normalize_rgb(rgb));
            }
            if let Some(rgb) = attr(&child, "val") {
                if rgb.len() >= 6 {
                    return Some(normalize_rgb(rgb));
                }
            }
        }
        None
    };
    let mut slots = Vec::new();
    let names = [
        "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6",
    ];
    for name in names {
        if let Some(node) = clr_scheme
            .children()
            .find(|c| c.is_element() && node_name(c) == name)
        {
            slots.push(read_slot(node).unwrap_or_default());
        }
    }
    slots
}

fn apply_tint(rgb: &str, tint: f64) -> String {
    let normalized = normalize_rgb(rgb);
    if tint == 0.0 {
        return normalized;
    }
    if normalized.len() < 6 {
        return normalized;
    }
    let base = &normalized[normalized.len() - 6..];
    let mut out = String::from("FF");
    for index in (0..6).step_by(2) {
        let value = u16::from_str_radix(&base[index..index + 2], 16).unwrap_or(0);
        let value = if tint > 0.0 {
            (value as f64 * (1.0 - tint) + 255.0 * tint).round()
        } else {
            (value as f64 * (1.0 + tint)).round()
        };
        let value = value.clamp(0.0, 255.0) as u8;
        out.push_str(&format!("{value:02X}"));
    }
    out
}

fn normalize_rgb(rgb: &str) -> String {
    let value = rgb.trim_start_matches('#');
    if value.len() == 6 {
        format!("FF{value}")
    } else if value.len() == 8 {
        value.to_string()
    } else {
        rgb.to_string()
    }
}

fn read_zip_part(archive: &mut ZipArchive<std::fs::File>, name: &str) -> Result<Vec<u8>> {
    use std::io::Read;
    let mut file = archive
        .by_name(name)
        .with_context(|| format!("read OOXML part {name}"))?;
    let mut buffer = Vec::new();
    file.read_to_end(&mut buffer)?;
    Ok(buffer)
}

fn find_child<'a, 'doc>(
    node: &'a roxmltree::Node<'a, 'doc>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'doc>> {
    node.children()
        .find(|child| child.is_element() && child.tag_name().name() == name)
}

fn element_children<'a, 'doc>(
    node: roxmltree::Node<'a, 'doc>,
) -> impl Iterator<Item = roxmltree::Node<'a, 'doc>> {
    node.children().filter(|child| child.is_element())
}

fn node_name<'a, 'doc>(node: &roxmltree::Node<'a, 'doc>) -> &'a str {
    node.tag_name().name()
}

fn attr<'a, 'doc>(node: &roxmltree::Node<'a, 'doc>, name: &str) -> Option<&'a str> {
    node.attribute(name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builtin_num_fmt_covers_common_report_codes() {
        assert_eq!(builtin_num_fmt(9).as_deref(), Some("0%"));
        assert_eq!(builtin_num_fmt(14).as_deref(), Some("yyyy-m-d"));
        assert_eq!(builtin_num_fmt(39).as_deref(), Some("#,##0.00;(#,##0.00)"));
        assert_eq!(builtin_num_fmt(49).as_deref(), Some("@"));
        assert_eq!(builtin_num_fmt(200), None);
    }

    #[test]
    fn tint_lightens_and_darkens_channels() {
        assert_eq!(apply_tint("FFFFFF", 0.0), "FFFFFFFF");
        assert_eq!(apply_tint("000000", 0.5), "FF808080");
        assert_eq!(apply_tint("808080", -0.5), "FF404040");
    }

    #[test]
    fn normalize_rgb_handles_six_and_eight_digit_forms() {
        assert_eq!(normalize_rgb("#FF0000"), "FFFF0000");
        assert_eq!(normalize_rgb("FFFF0000"), "FFFF0000");
    }
}
