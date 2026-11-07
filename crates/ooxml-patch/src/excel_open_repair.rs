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

/// 剥离公式中指向本地已存在表的 `[N]` 外部引用前缀；返回改写文本与次数。
/// 形态：`[1]表名!`、`[1]'表名'!`（索引数字任意、表名可能带引号）。
/// 仅当 N 后跟的表名存在于 `local_sheet_names` 时剥离（指向真正外部簿的不动）。
fn strip_external_refs(text: &str, local_sheet_names: &[String]) -> (String, u64) {
    if local_sheet_names.is_empty() {
        return (text.to_owned(), 0);
    }
    // 找所有 `[N]` 模式（N 为纯数字），检查后续是否跟本地表名 + `!`
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    let mut count = 0u64;
    while let Some(bracket_start) = rest.find('[') {
        // 先推入 [ 之前的文本
        out.push_str(&rest[..bracket_start]);
        // 找到 `]` 结束的 [N]
        let Some(bracket_end_rel) = rest[bracket_start..].find(']') else {
            out.push_str(rest);
            break;
        };
        let bracket_end = bracket_start + bracket_end_rel;
        let index_text = &rest[bracket_start + 1..bracket_end];
        if index_text.is_empty() || !index_text.chars().all(|c| c.is_ascii_digit()) {
            // 非 [数字] 格式，原样推入 [ 并继续扫描
            out.push_str(&rest[..bracket_start + 1]);
            rest = &rest[bracket_start + 1..];
            continue;
        }
        // 检查 ] 后是否跟本地表名 + !
        let after_bracket = &rest[bracket_end + 1..];
        let mut matched = false;
        for name in local_sheet_names {
            let escaped = name.replace(std::char::from_u32(39).unwrap(), "''");
            let quoted_form = format!("'{}'!", escaped);
            let bare_form = format!("{}!", name);
            if after_bracket.starts_with(&quoted_form) {
                out.push_str(&quoted_form);
                rest = &rest[bracket_end + 1 + quoted_form.len()..];
                matched = true;
                count += 1;
                break;
            } else if after_bracket.starts_with(&bare_form) {
                out.push_str(&bare_form);
                rest = &rest[bracket_end + 1 + bare_form.len()..];
                matched = true;
                count += 1;
                break;
            }
        }
        if !matched {
            // [N] 后面不跟本地表名，原样保留（out 已有前置文本，只需推入 [N] 本身）
            out.push_str(&rest[bracket_start..bracket_end + 1]);
            rest = &rest[bracket_end + 1..];
        }
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
    pub external_refs_stripped: u64,
    pub reason: String,
}

/// 对整个文本执行 [N] 外部引用前缀剥离；仅剥离指向本地表名的引用。
///
/// 走 roxmltree 解码（生成侧可能把表名写成 `&#23384;` 这类数字字符引用），
/// 命中后用公式内容转义写回，保证输出仍是合法 XML。
fn strip_text_external_refs(text: &str, local_sheet_names: &[String]) -> (String, u64) {
    if local_sheet_names.is_empty() {
        return (text.to_owned(), 0);
    }
    let Ok(document) = Document::parse(text) else {
        return (text.to_owned(), 0);
    };
    let mut edits: Vec<(std::ops::Range<usize>, Vec<u8>)> = Vec::new();
    let mut total = 0u64;
    for node in document.descendants().filter(|node| {
        node.is_element()
            && matches!(
                node.tag_name().name(),
                "f" | "formula" | "c:f" | "definedName"
            )
    }) {
        let Some(original) = node.text() else {
            continue;
        };
        let (rewritten, count) = strip_external_refs(original, local_sheet_names);
        if count == 0 {
            continue;
        }
        total += count;
        let range = node.range();
        let open_end = match text[range.clone()].find('>') {
            Some(offset) => offset,
            None => continue,
        };
        let open_tag = text[range.start..range.start + open_end + 1].to_owned();
        let replacement = format!(
            "{open_tag}{}</{}>",
            escape_formula_content(&rewritten),
            node.tag_name().name()
        );
        edits.push((range, replacement.into_bytes()));
    }
    if edits.is_empty() {
        return (text.to_owned(), 0);
    }
    let mut result = text.as_bytes().to_vec();
    for (range, replacement) in edits.into_iter().rev() {
        result.splice(range, replacement);
    }
    (String::from_utf8_lossy(&result).into_owned(), total)
}

/// 公式内容最小转义：只处理 `& < >`，单双引号保持原样。
fn escape_formula_content(text: &str) -> String {
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

/// 廉价预检：无裸引用嫌疑且无矛盾窗格的文件不重写。
fn needs_repair(
    parts: &[(String, Vec<u8>)],
    special: &[String],
    local_sheet_names: &[String],
) -> bool {
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
        let (_, stripped) = strip_text_external_refs(text, local_sheet_names);
        if stripped > 0 {
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

/// 模板外部工作簿引用扫描结果。
#[derive(Debug, Default, Clone, serde::Serialize)]
pub struct ExternalReferenceScan {
    /// `xl/externalLinks/*` 部件名（存在即代表工作簿带外链定义）。
    pub external_link_parts: Vec<String>,
    /// 公式中 `[N]` 前缀出现次数（`<f>` / `<formula>` / `<c:f>` / `<definedName>`）。
    pub formula_hits: u64,
    /// 命中样本（最多 3 条），用于错误信息定位。
    pub samples: Vec<String>,
}

impl ExternalReferenceScan {
    pub fn is_clean(&self) -> bool {
        self.external_link_parts.is_empty() && self.formula_hits == 0
    }

    pub fn describe(&self) -> String {
        let mut parts = Vec::new();
        if !self.external_link_parts.is_empty() {
            parts.push(format!(
                "外部链接部件 {} 个（{}）",
                self.external_link_parts.len(),
                self.external_link_parts.join("、")
            ));
        }
        if self.formula_hits > 0 {
            let mut detail = format!("公式含 [N] 外部工作簿引用 {} 处", self.formula_hits);
            if !self.samples.is_empty() {
                detail.push_str(&format!("，例如：{}", self.samples.join("；")));
            }
            parts.push(detail);
        }
        if parts.is_empty() {
            "未发现外部工作簿引用".to_owned()
        } else {
            parts.join("；")
        }
    }
}

/// 扫描工作簿是否带外部工作簿引用（模板入口 fail-closed 检查用）。
///
/// 判定两类证据：`xl/externalLinks/` 部件存在；四类公式载体内出现 `[N]`
/// 前缀（无论 N 后跟的表名是否在本地——带前缀即引用外部工作簿）。
/// 非 ZIP / 非 OOXML 直接报错，由调用方决定如何处理。
pub fn scan_external_workbook_references(path: &Path) -> Result<ExternalReferenceScan> {
    let file = fs::File::open(path).with_context(|| format!("open {}", path.display()))?;
    let mut archive = zip::ZipArchive::new(file).with_context(|| "read ZIP container")?;
    let mut scan = ExternalReferenceScan::default();
    let names: Vec<String> = (0..archive.len())
        .filter_map(|index| {
            archive
                .by_index(index)
                .ok()
                .map(|entry| entry.name().to_owned())
        })
        .collect();
    for name in &names {
        if name.starts_with("xl/externalLinks/") && !name.contains("/_rels/") {
            scan.external_link_parts.push(name.clone());
        }
    }
    for name in names.iter().filter(|name| is_repairable_part(name)) {
        let Ok(mut entry) = archive.by_name(name) else {
            continue;
        };
        let mut data = Vec::new();
        if entry.read_to_end(&mut data).is_err() {
            continue;
        }
        let Ok(text) = std::str::from_utf8(&data) else {
            continue;
        };
        let Ok(document) = Document::parse(text) else {
            continue;
        };
        for node in document.descendants().filter(|node| {
            node.is_element()
                && matches!(
                    node.tag_name().name(),
                    "f" | "formula" | "c:f" | "definedName"
                )
        }) {
            let Some(original) = node.text() else {
                continue;
            };
            let hits = count_external_prefixes(original);
            if hits == 0 {
                continue;
            }
            scan.formula_hits += hits;
            if scan.samples.len() < 3 {
                if let Some(sample) = first_external_prefix(original) {
                    scan.samples.push(sample);
                }
            }
        }
    }
    Ok(scan)
}

/// 统计文本中 `[数字]` 前缀出现次数。
fn count_external_prefixes(text: &str) -> u64 {
    let mut count = 0u64;
    let mut rest = text;
    while let Some(start) = rest.find('[') {
        let Some(end_rel) = rest[start..].find(']') else {
            break;
        };
        let index_text = &rest[start + 1..start + end_rel];
        if !index_text.is_empty() && index_text.chars().all(|c| c.is_ascii_digit()) {
            count += 1;
        }
        rest = &rest[start + end_rel + 1..];
    }
    count
}

fn first_external_prefix(text: &str) -> Option<String> {
    let mut rest = text;
    while let Some(start) = rest.find('[') {
        let Some(end_rel) = rest[start..].find(']') else {
            return None;
        };
        let bracket_end = start + end_rel;
        let index_text = &rest[start + 1..bracket_end];
        if !index_text.is_empty() && index_text.chars().all(|c| c.is_ascii_digit()) {
            let tail = &rest[bracket_end + 1..];
            let tail_len = tail
                .find(|c: char| {
                    !(c.is_alphanumeric() || c == '_' || c == '\'' || c == '（' || c == '）')
                })
                .unwrap_or(tail.len().min(24));
            let snippet: String = rest[start..bracket_end + 1 + tail_len.min(24)]
                .chars()
                .take(40)
                .collect();
            return Some(snippet);
        }
        rest = &rest[bracket_end + 1..];
    }
    None
}

/// 修复传入的工作簿副本使其可被 Excel 打开；返回统计。
///
/// 只处理 .xlsx/.xlsm（ZIP 容器，必须含 `xl/workbook.xml`）；其余格式原样
/// 跳过。修复先写临时文件再原子替换；任何异常都返回 `changed=false` 且保持
/// 文件原样。**调用方必须保证 `path` 是一次性副本**，本函数不做来源检查。
pub fn repair_excel_com_open_issues(path: &Path) -> ExcelOpenRepairReport {
    let mut report = ExcelOpenRepairReport {
        external_refs_stripped: 0,
        ..Default::default()
    };
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
    let all_sheet_names = sheet_names(&workbook_xml);
    let special: Vec<String> = all_sheet_names
        .clone()
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

    // 本地表名 = 所有 <sheet> 名称（含纯中文、带括号等）
    let local_sheet_names: Vec<String> = all_sheet_names
        .iter()
        .filter(|n| !n.is_empty())
        .cloned()
        .collect();
    if !needs_repair(&parts, &special, &local_sheet_names) {
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
            // 剥离指向本地已存在表的 [N] 外部引用前缀（与表名是否需引号无关）
            let (stripped_text, stripped_count) =
                strip_text_external_refs(&rewritten, &local_sheet_names);
            rewritten = stripped_text;
            report.external_refs_stripped += stripped_count;
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
    #[test]
    fn strips_dangling_external_refs_to_local_sheets() {
        // 带全角括号表名引用的 [N] 前缀 → 剥离 [1] 保留表名
        let text = "SUM([1]单位贷款发生额信息（隐藏发放）!A1)+[1]'存量单位贷款信息'!B2";
        let local = vec![
            "单位贷款发生额信息（隐藏发放）".to_owned(),
            "存量单位贷款信息".to_owned(),
        ];
        let (rewritten, count) = strip_external_refs(text, &local);
        assert_eq!(count, 2);
        assert!(
            rewritten.contains("单位贷款发生额信息（隐藏发放）!A1"),
            "{rewritten}"
        );
        assert!(rewritten.contains("'存量单位贷款信息'!B2"), "{rewritten}");
        assert!(!rewritten.contains("[1]"), "{rewritten}");
        // 幂等
        let (r2, c2) = strip_external_refs(&rewritten, &local);
        assert_eq!(c2, 0);
        assert_eq!(r2, rewritten);
    }

    #[test]
    fn external_refs_to_non_local_sheets_are_not_stripped() {
        let text = "[99]外部簿!A1+[1]本地表!B2";
        let local = vec!["本地表".to_owned()];
        let (rewritten, count) = strip_external_refs(text, &local);
        assert_eq!(count, 1);
        assert!(rewritten.contains("[99]外部簿!A1"), "{rewritten}");
        assert!(rewritten.contains("本地表!B2"), "{rewritten}");
        assert!(!rewritten.contains("[1]本地表"), "{rewritten}");
    }

    #[test]
    fn scan_flags_external_prefixes_and_external_link_parts() {
        let dir = std::env::temp_dir().join(format!("ooxml-scan-ext-{}", std::process::id()));
        fs::create_dir_all(&dir).unwrap();
        let path = dir.join("with-external.xlsx");
        let worksheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>[1]存量单位贷款信息!C5</f><v>0</v></c></row></sheetData></worksheet>"#;
        let binding = base_parts("校验", worksheet, "");
        let written: Vec<(&str, Vec<u8>)> = binding
            .iter()
            .map(|(name, data)| (name.as_str(), data.clone()))
            .collect();
        write_workbook(&path, &written);
        let scan = scan_external_workbook_references(&path).unwrap();
        assert!(!scan.is_clean(), "{scan:?}");
        assert!(scan.formula_hits >= 1, "{scan:?}");
        assert!(!scan.samples.is_empty(), "{scan:?}");
        assert!(scan.describe().contains("[N]"), "{}", scan.describe());
        fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn scan_accepts_local_only_formulas() {
        let dir = std::env::temp_dir().join(format!("ooxml-scan-local-{}", std::process::id()));
        fs::create_dir_all(&dir).unwrap();
        let path = dir.join("local-only.xlsx");
        let worksheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>SUM(存量单位贷款信息!C5:C10)</f><v>0</v></c></row></sheetData></worksheet>"#;
        let binding = base_parts("校验", worksheet, "");
        let written: Vec<(&str, Vec<u8>)> = binding
            .iter()
            .map(|(name, data)| (name.as_str(), data.clone()))
            .collect();
        write_workbook(&path, &written);
        let scan = scan_external_workbook_references(&path).unwrap();
        assert!(scan.is_clean(), "{scan:?}");
        fs::remove_dir_all(&dir).unwrap();
    }

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
