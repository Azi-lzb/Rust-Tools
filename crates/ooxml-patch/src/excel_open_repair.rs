//! Excel COM 打开失败的副本修复：全角括号表名裸引用 + 冻结窗格坐标矛盾。
//!
//! 移植自 3.open-data-audit 的 `base_audit/excel_open_repair.py`（根因分析与
//! 28/28 验证见 `test-output/s1f1_mode_parity_20261107/excel_open_root_cause.md`）：
//!
//! 1. 报送文件的公式直接裸引用含特殊字符（全角括号等）的表名。Excel 要求
//!    这类表名必须加单引号；存储公式解析失败 → 拒绝打开整个工作簿；WPS
//!    解析宽松能打开。载体：单元格 `<f>`、条件格式/数据验证 `<formula>`、
//!    图表 `<c:f>`、workbook.xml 的 `<definedNames>`。
//! 2. 部分文件的冻结窗格 `<pane ySplit="6" topLeftCell="C6" state="frozen"/>`
//!    自相矛盾（topLeftCell 行落在冻结区内），Excel 强制要求冻结态下
//!    topLeftCell 位于分割线右下方 → 拒绝打开。
//!
//! 修复只做两类无害规范化：给强制引号表名的裸引用补单引号（公式语义不变）；
//! 冻结窗格 topLeftCell 行号改为 ySplit+1。**只允许作用于一次性副本**，
//! 调用方必须保证传入路径不是原始报送文件。无问题的文件不做任何重写
//! （先廉价预检）；异常一律保持文件原样并返回未修复原因，不打断批处理。

use anyhow::{Context, Result};
use roxmltree::Document;
use std::{
    fs,
    io::{Read, Write as _},
    path::{Path, PathBuf},
};

/// 与参考实现 `_NEEDS_QUOTE` 一致：括号（全/半角）、空白（含全角空格）、
/// 连字符等。纯中文/字母数字表名不需要引号，不在此列。
fn char_needs_quote(c: char) -> bool {
    matches!(
        c,
        '（' | '）'
            | '('
            | ')'
            | ' '
            | '\t'
            | '\n'
            | '\x0B'
            | '\x0C'
            | '\r'
            | '\u{3000}'
            | '-'
            | '+'
            | '['
            | ']'
            | '{'
            | '}'
            | '\\'
            | '/'
            | ':'
            | ';'
            | '"'
            | '\''
            | '<'
            | '>'
            | ','
            | '?'
            | '~'
            | '！'
            | '？'
            | '＠'
            | '＃'
            | '￥'
            | '％'
            | '…'
            | '&'
            | '*'
            | '='
            | '|'
    )
}

fn name_needs_quotes(name: &str) -> bool {
    name.chars().any(char_needs_quote)
}

/// 与参考实现一致：只处理这些前缀下的 XML 部件与 workbook.xml。
fn is_repairable_part(name: &str) -> bool {
    name.ends_with(".xml")
        && (name.starts_with("xl/worksheets/")
            || name.starts_with("xl/charts/")
            || name.starts_with("xl/drawings/")
            || name == "xl/workbook.xml")
}

fn sheet_names(workbook_xml: &str) -> Vec<String> {
    let Ok(document) = Document::parse(workbook_xml) else {
        return Vec::new();
    };
    document
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().name() == "sheet"
                && node.tag_name().namespace().is_some()
        })
        .filter_map(|node| node.attribute("name").map(str::to_owned))
        .collect()
}

/// 把 `name!` 裸引用替换为 `'name'!`（名字里的单引号转成 `''`）。
/// 前一个字符是 `'` 时不重复处理（对应参考实现的 `(?<!')`）。
fn quote_bare_refs(text: &str, name: &str) -> (String, u64) {
    let needle = format!("{name}!");
    let quoted_name = name.replace('\'', "''");
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    let mut count = 0u64;
    while let Some(position) = rest.find(&needle) {
        out.push_str(&rest[..position]);
        if out.chars().last() == Some('\'') {
            out.push_str(&needle);
        } else {
            out.push('\'');
            out.push_str(&quoted_name);
            out.push_str("'!");
            count += 1;
        }
        rest = &rest[position + needle.len()..];
    }
    out.push_str(rest);
    (out, count)
}

/// 在一段文本里找到 `<tag ...>inner</tag>` 载体（`<f>`、`<formula>`、
/// `<c:f>`），对 inner 逐个补引号；返回改写后的文本与新增补引号数。
fn quote_carriers(
    text: &str,
    open_tag: &str,
    close_tag: &str,
    special: &[String],
) -> (String, u64) {
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    let mut total = 0u64;
    while let Some(start) = find_open_tag(rest, open_tag) {
        let after_open = start + open_tag.len();
        let Some(close_position) = rest[after_open..].find(close_tag) else {
            out.push_str(&rest[..after_open]);
            rest = &rest[after_open..];
            break;
        };
        let inner_end = after_open + close_position;
        out.push_str(&rest[..after_open]);
        let (rewritten, count) = rewrite_all_names(&rest[after_open..inner_end], special);
        out.push_str(&rewritten);
        total += count;
        out.push_str(close_tag);
        rest = &rest[inner_end + close_tag.len()..];
    }
    out.push_str(rest);
    (out, total)
}

/// `<f` 之后的字符必须是 `>` 或空白才算该载体开标签（避免匹配 `<formula`）。
fn find_open_tag(text: &str, open_tag: &str) -> Option<usize> {
    let mut from = 0usize;
    loop {
        let position = text[from..].find(open_tag)? + from;
        let next = text[position + open_tag.len()..].chars().next();
        if open_tag == "<c:f>" || open_tag == "</f>" {
            return Some(position);
        }
        match next {
            Some('>') | Some(' ') | Some('\t') | Some('\n') | Some('\r') => return Some(position),
            _ => from = position + 1,
        }
    }
}

fn rewrite_all_names(text: &str, special: &[String]) -> (String, u64) {
    let mut current = text.to_owned();
    let mut total = 0u64;
    for name in special {
        let (rewritten, count) = quote_bare_refs(&current, name);
        current = rewritten;
        total += count;
    }
    (current, total)
}

/// 冻结窗格修复：state=frozen/frozenSplit 且 topLeftCell 行号 ≤ ySplit 时
/// 行号改为 ySplit+1（列不变）。只处理自闭合 `<pane .../>`。
fn fix_frozen_panes(text: &str) -> (String, u64) {
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    let mut fixed = 0u64;
    while let Some(relative) = rest.find("<pane") {
        let absolute = relative;
        let tag_end = match rest[absolute..].find("/>") {
            Some(end) if !rest[absolute..absolute + end].contains('>') => absolute + end + 2,
            _ => {
                // 非自闭合 pane：整段原样拷到下一个 `>` 之后。
                match rest[absolute..].find('>') {
                    Some(end) => {
                        out.push_str(&rest[..absolute + end + 1]);
                        rest = &rest[absolute + end + 1..];
                        continue;
                    }
                    None => break,
                }
            }
        };
        let tag = &rest[absolute..tag_end];
        // 先把 `<pane` 之前的文本原样收进输出，避免截断文档头部。
        out.push_str(&rest[..absolute]);
        match repaired_pane_tag(tag) {
            Some(replacement) => {
                fixed += 1;
                out.push_str(&replacement);
            }
            None => out.push_str(tag),
        }
        rest = &rest[tag_end..];
    }
    out.push_str(rest);
    (out, fixed)
}

fn pane_attribute<'a>(tag: &'a str, name: &str) -> Option<&'a str> {
    let marker = format!("{name}=\"");
    let start = tag.find(&marker)? + marker.len();
    let end = start + tag[start..].find('"')?;
    Some(&tag[start..end])
}

fn repaired_pane_tag(tag: &str) -> Option<String> {
    let frozen = pane_attribute(tag, "state")
        .is_some_and(|state| state == "frozen" || state == "frozenSplit");
    if !frozen {
        return None;
    }
    let y_split: u64 = pane_attribute(tag, "ySplit")?.parse().ok()?;
    let top_left = pane_attribute(tag, "topLeftCell")?;
    let split = top_left
        .find(|c: char| c.is_ascii_digit())
        .unwrap_or(top_left.len());
    let (columns, rows) = top_left.split_at(split);
    let row: u64 = rows.parse().ok()?;
    if row > y_split {
        return None;
    }
    Some(tag.replace(
        &format!("topLeftCell=\"{top_left}\""),
        &format!("topLeftCell=\"{columns}{}\"", y_split + 1),
    ))
}

#[derive(Debug, Default, Clone, serde::Serialize)]
pub struct ExcelOpenRepairReport {
    pub changed: bool,
    pub quoted: u64,
    pub pane_fixed: u64,
    pub reason: String,
}

/// 廉价预检：无裸引用嫌疑且无矛盾窗格的文件不重写。
fn needs_repair(parts: &[(String, Vec<u8>)], special: &[String]) -> bool {
    for (name, data) in parts {
        if !is_repairable_part(name) {
            continue;
        }
        let Ok(text) = std::str::from_utf8(data) else {
            continue;
        };
        if rewrite_all_names(text, special).1 > 0 {
            return true;
        }
        if name.starts_with("xl/worksheets/") && text.contains("<pane ") {
            let (_, fixed) = fix_frozen_panes(text);
            if fixed > 0 {
                return true;
            }
        }
    }
    false
}

/// 修复传入的工作簿副本使其可被 Excel 打开；返回统计。
///
/// 只处理 .xlsx/.xlsm（ZIP 容器，必须含 `xl/workbook.xml`）；其余格式原样
/// 跳过。修复先写临时文件再原子替换；任何异常都返回 `changed=false` 且保持
/// 文件原样。**调用方必须保证 `path` 是一次性副本**，本函数不做来源检查。
pub fn repair_excel_com_open_issues(path: &Path) -> ExcelOpenRepairReport {
    let mut report = ExcelOpenRepairReport::default();
    let temporary: PathBuf;
    match repair_inner(path, &mut report) {
        Ok(temp) => temporary = temp,
        Err(error) => {
            report.reason = format!("{error:#}");
            return report;
        }
    }
    match fs::rename(&temporary, path) {
        Ok(()) => {
            report.changed = true;
        }
        Err(error) => {
            let _ = fs::remove_file(&temporary);
            report.reason = format!("atomic replace failed: {error}");
        }
    }
    report
}

fn repair_inner(path: &Path, report: &mut ExcelOpenRepairReport) -> Result<PathBuf> {
    let file = fs::File::open(path).with_context(|| format!("open {}", path.display()))?;
    let mut archive = zip::ZipArchive::new(file).with_context(|| "read ZIP container")?;
    let names: Vec<String> = (0..archive.len())
        .filter_map(|index| {
            archive
                .by_index(index)
                .ok()
                .map(|entry| entry.name().to_owned())
        })
        .collect();
    if !names.iter().any(|name| name == "xl/workbook.xml") {
        report.reason = "非 OOXML 工作簿".to_owned();
        return Err(anyhow::anyhow!(report.reason.clone()));
    }
    let workbook_xml = {
        let mut buffer = String::new();
        archive
            .by_name("xl/workbook.xml")?
            .read_to_string(&mut buffer)
            .context("read xl/workbook.xml")?;
        buffer
    };
    let special: Vec<String> = sheet_names(&workbook_xml)
        .into_iter()
        .filter(|name| !name.is_empty() && name_needs_quotes(name))
        .collect();

    let mut parts: Vec<(String, Vec<u8>)> = Vec::with_capacity(names.len());
    for name in &names {
        let mut data = Vec::new();
        archive
            .by_name(name)?
            .read_to_end(&mut data)
            .with_context(|| format!("read {name}"))?;
        parts.push((name.clone(), data));
    }
    drop(archive);

    if !needs_repair(&parts, &special) {
        report.reason = "未发现需要修复的内容".to_owned();
        anyhow::bail!(report.reason.clone());
    }

    let temporary = path.with_file_name(format!(
        "{}.excel-repair.tmp{}",
        path.file_stem()
            .and_then(|stem| stem.to_str())
            .unwrap_or("workbook"),
        path.extension()
            .and_then(|extension| extension.to_str())
            .unwrap_or_default()
    ));
    let file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&temporary)
        .with_context(|| format!("create repair temp {}", temporary.display()))?;
    let mut writer = zip::ZipWriter::new(file);
    let write_result: Result<()> = (|| {
        for (name, data) in &parts {
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            writer.start_file(name.as_str(), options)?;
            if !is_repairable_part(name) {
                writer.write_all(data)?;
                continue;
            }
            let Ok(text) = std::str::from_utf8(data) else {
                writer.write_all(data)?;
                continue;
            };
            let mut rewritten = text.to_owned();
            if !special.is_empty() {
                let mut total = 0u64;
                let (next, count) = quote_carriers(&rewritten, "<f", "</f>", &special);
                total += count;
                rewritten = next;
                let (next, count) = quote_carriers(&rewritten, "<formula", "</formula>", &special);
                total += count;
                rewritten = next;
                let (next, count) = quote_carriers(&rewritten, "<c:f>", "</c:f>", &special);
                total += count;
                rewritten = next;
                let (names_block, count) = quote_defined_names(&rewritten, &special);
                total += count;
                rewritten = names_block;
                report.quoted += total;
            }
            let (pane_text, pane_fixed) = fix_frozen_panes(&rewritten);
            report.pane_fixed += pane_fixed;
            writer.write_all(pane_text.as_bytes())?;
        }
        Ok(())
    })();
    let finish_result = writer.finish().context("finish repaired ZIP");
    match (write_result, finish_result) {
        (Ok(()), Ok(_)) => Ok(temporary),
        (Err(error), _) | (_, Err(error)) => {
            let _ = fs::remove_file(&temporary);
            report.reason = format!("{error:#}");
            Err(anyhow::anyhow!(report.reason.clone()))
        }
    }
}

/// 参考实现对 `<definedNames>` 块内容整体补引号。
fn quote_defined_names(text: &str, special: &[String]) -> (String, u64) {
    let Some(start) = text.find("<definedNames>") else {
        return (text.to_owned(), 0);
    };
    let Some(end) = text[start..].find("</definedNames>") else {
        return (text.to_owned(), 0);
    };
    let inner_start = start + "<definedNames>".len();
    let inner_end = start + end;
    let (rewritten, count) = rewrite_all_names(&text[inner_start..inner_end], special);
    let mut out = String::with_capacity(text.len());
    out.push_str(&text[..inner_start]);
    out.push_str(&rewritten);
    out.push_str(&text[inner_end..]);
    (out, count)
}

#[cfg(test)]
mod excel_open_repair_tests {
    use super::*;
    use std::io::Write as _;

    const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
    const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

    fn workbook_xml(sheet_name: &str, defined_names: &str) -> String {
        format!(
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/><sheet name="个人贷款发生额信息（隐藏发放）" sheetId="2" r:id="rId2" state="veryHidden"/></sheets>{defined_names}</workbook>"#
        )
    }

    fn write_workbook(path: &Path, parts: &[(&str, Vec<u8>)]) {
        let file = fs::File::create(path).unwrap();
        let mut writer = zip::ZipWriter::new(file);
        for (name, data) in parts {
            writer
                .start_file(*name, zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(data).unwrap();
        }
        writer.finish().unwrap();
    }

    fn base_parts(
        sheet_name: &str,
        worksheet: &str,
        defined_names: &str,
    ) -> Vec<(String, Vec<u8>)> {
        vec![
            (
                "[Content_Types].xml".to_owned(),
                CONTENT_TYPES.as_bytes().to_vec(),
            ),
            ("_rels/.rels".to_owned(), ROOT_RELS.as_bytes().to_vec()),
            (
                "xl/workbook.xml".to_owned(),
                workbook_xml(sheet_name, defined_names).into_bytes(),
            ),
            (
                "xl/worksheets/sheet1.xml".to_owned(),
                worksheet.as_bytes().to_vec(),
            ),
        ]
    }

    fn part_bytes(path: &Path, name: &str) -> Vec<u8> {
        let mut archive = zip::ZipArchive::new(fs::File::open(path).unwrap()).unwrap();
        let mut data = Vec::new();
        archive
            .by_name(name)
            .unwrap()
            .read_to_end(&mut data)
            .unwrap();
        data
    }

    #[test]
    fn quotes_all_four_formula_carriers() {
        let dir = std::env::temp_dir().join(format!("ooxml-repair-test-{}", std::process::id()));
        fs::create_dir_all(&dir).unwrap();
        let path = dir.join("carriers.xlsx");
        let worksheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>个人贷款发生额信息（隐藏发放）!A1+1</f><v>0</v></c></row></sheetData><conditionalFormatting sqref="B1"><cfRule type="expression"><formula>存量个人贷款信息!B1&gt;0</formula></cfRule></conditionalFormatting></worksheet>"#;
        let chart = r#"<chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:ser><c:f>个人贷款发生额信息（隐藏发放）!A1:A9</c:f></c:ser></chartSpace>"#;
        let mut parts = base_parts(
            "存量个人贷款信息",
            worksheet,
            r#"<definedNames><definedName name="X">个人贷款发生额信息（隐藏发放）!$A$1</definedName></definedNames>"#,
        );
        parts.push(("xl/charts/chart1.xml".to_owned(), chart.as_bytes().to_vec()));
        let written: Vec<(&str, Vec<u8>)> = parts
            .iter()
            .map(|(name, data)| (name.as_str(), data.clone()))
            .collect();
        write_workbook(&path, &written);

        let report = repair_excel_com_open_issues(&path);
        assert!(report.changed, "{report:?}");
        // 单元格 1 + 图表 1 + definedNames 1；纯中文条件格式不改。
        assert_eq!(report.quoted, 3, "{report:?}");
        assert_eq!(report.pane_fixed, 0);
        let sheet = String::from_utf8(part_bytes(&path, "xl/worksheets/sheet1.xml")).unwrap();
        assert!(
            sheet.contains("<f>'个人贷款发生额信息（隐藏发放）'!A1+1</f>"),
            "{sheet}"
        );
        assert!(
            sheet.contains("<formula>存量个人贷款信息!B1&gt;0</formula>"),
            "{sheet}"
        );
        let chart_out = String::from_utf8(part_bytes(&path, "xl/charts/chart1.xml")).unwrap();
        assert!(
            chart_out.contains("<c:f>'个人贷款发生额信息（隐藏发放）'!A1:A9</c:f>"),
            "{chart_out}"
        );
        let book = String::from_utf8(part_bytes(&path, "xl/workbook.xml")).unwrap();
        assert!(
            book.contains("'个人贷款发生额信息（隐藏发放）'!$A$1"),
            "{book}"
        );
        fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn fixes_contradictory_frozen_pane() {
        let dir = std::env::temp_dir().join(format!("ooxml-repair-pane-{}", std::process::id()));
        fs::create_dir_all(&dir).unwrap();
        let path = dir.join("pane.xlsx");
        let worksheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="6" topLeftCell="C6" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetData/></worksheet>"#;
        let binding = base_parts("数据", worksheet, "");
        let written: Vec<(&str, Vec<u8>)> = binding
            .iter()
            .map(|(name, data)| (name.as_str(), data.clone()))
            .collect();
        write_workbook(&path, &written);

        let report = repair_excel_com_open_issues(&path);
        assert!(report.changed, "{report:?}");
        assert_eq!(report.pane_fixed, 1);
        assert_eq!(report.quoted, 0);
        let sheet = String::from_utf8(part_bytes(&path, "xl/worksheets/sheet1.xml")).unwrap();
        assert!(sheet.contains(r#"topLeftCell="C7""#), "{sheet}");
        assert!(sheet.contains(r#"ySplit="6""#), "{sheet}");
        // 窗格之前的文档头必须原样保留（回归：曾有实现把头部截掉）。
        assert!(sheet.starts_with("<worksheet"), "{sheet}");
        assert!(
            sheet.contains(r#"<sheetView workbookViewId="0">"#),
            "{sheet}"
        );
        assert!(
            sheet.contains(r#"<pane xSplit="2" ySplit="6" topLeftCell="C7""#),
            "{sheet}"
        );
        assert!(sheet.ends_with("<sheetData/></worksheet>"), "{sheet}");
        fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn clean_files_and_non_zip_stay_untouched() {
        let dir = std::env::temp_dir().join(format!("ooxml-repair-clean-{}", std::process::id()));
        fs::create_dir_all(&dir).unwrap();
        // 干净文件：字节不变。
        let clean = dir.join("clean.xlsx");
        let worksheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="6" topLeftCell="C7" state="frozen"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><f>存量个人贷款信息!A1</f></c></row></sheetData></worksheet>"#;
        let binding = base_parts("存量个人贷款信息", worksheet, "");
        let written: Vec<(&str, Vec<u8>)> = binding
            .iter()
            .map(|(name, data)| (name.as_str(), data.clone()))
            .collect();
        write_workbook(&clean, &written);
        let before = fs::read(&clean).unwrap();
        let report = repair_excel_com_open_issues(&clean);
        assert!(!report.changed, "{report:?}");
        assert_eq!(report.quoted, 0);
        assert_eq!(
            fs::read(&clean).unwrap(),
            before,
            "clean file must stay byte-identical"
        );

        // 非 ZIP：原样跳过。
        let garbage = dir.join("garbage.xlsx");
        fs::write(&garbage, b"not a zip container").unwrap();
        let garbage_before = fs::read(&garbage).unwrap();
        let report = repair_excel_com_open_issues(&garbage);
        assert!(!report.changed, "{report:?}");
        assert_eq!(fs::read(&garbage).unwrap(), garbage_before);
        fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn xlsm_binary_parts_pass_through_and_repair_is_idempotent() {
        let dir = std::env::temp_dir().join(format!("ooxml-repair-xlsm-{}", std::process::id()));
        fs::create_dir_all(&dir).unwrap();
        let path = dir.join("macro.xlsm");
        let vba: Vec<u8> = (0u8..=255).cycle().take(4096).collect();
        let worksheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>个人贷款发生额信息（隐藏发放）!A1</f></c></row></sheetData></worksheet>"#;
        let mut parts = base_parts("存量个人贷款信息", worksheet, "");
        parts.push(("xl/vbaProject.bin".to_owned(), vba.clone()));
        let written: Vec<(&str, Vec<u8>)> = parts
            .iter()
            .map(|(name, data)| (name.as_str(), data.clone()))
            .collect();
        write_workbook(&path, &written);

        let report = repair_excel_com_open_issues(&path);
        assert!(report.changed, "{report:?}");
        assert_eq!(report.quoted, 1);
        assert_eq!(
            part_bytes(&path, "xl/vbaProject.bin"),
            vba,
            "vbaProject must stay byte-identical"
        );
        // 幂等：已加引号后再次修复不再改写。
        let report = repair_excel_com_open_issues(&path);
        assert!(!report.changed, "{report:?}");
        fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn already_quoted_and_apostrophe_names_are_handled_once() {
        let text = "('个人贷款发生额信息（隐藏发放）'!A1)+个人贷款发生额信息（隐藏发放）!A2";
        let (rewritten, count) = quote_bare_refs(text, "个人贷款发生额信息（隐藏发放）");
        assert_eq!(count, 1, "{rewritten}");
        assert_eq!(
            rewritten,
            "('个人贷款发生额信息（隐藏发放）'!A1)+'个人贷款发生额信息（隐藏发放）'!A2"
        );
        // 名字含单引号时按 OOXML 规则转义为两个单引号。
        let (rewritten, count) = quote_bare_refs("O'Brien!A1", "O'Brien");
        assert_eq!(count, 1);
        assert_eq!(rewritten, "'O''Brien'!A1");
    }
}
