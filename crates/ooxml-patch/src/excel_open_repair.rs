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
    collections::BTreeSet,
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

/// 外部工作簿引用剥离结果。
#[derive(Debug, Default, Clone)]
struct ExternalStripOutcome {
    /// 改写后的公式文本。
    text: String,
    /// 剥离的前缀数量。
    stripped: u64,
    /// 剥离后指向"本副本不存在同名工作表"的表名（去重、有序）。
    missing_sheets: Vec<String>,
}

/// 工作簿容器扩展名白名单（路径式外部引用的判据）。
fn looks_like_workbook_file(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    lower.ends_with(".xlsx")
        || lower.ends_with(".xlsm")
        || lower.ends_with(".xlsb")
        || lower.ends_with(".xls")
}

fn is_all_digits(text: &str) -> bool {
    !text.is_empty() && text.chars().all(|c| c.is_ascii_digit())
}

/// 解析引号内的一段外部引用（已去掉两侧引号），返回被引用的工作表名。
///
/// 覆盖两种形态：
/// - 索引式 `[N]表名`（N 对应 workbook.xml 的 externalReference 序号）；
/// - 路径式 `C:\文件夹\[工作簿.xlsx]表名` 或 `[工作簿.xlsx]表名`。
/// 本簿内引用（如 `表名`）返回 `None`，调用方原样保留。
fn sheet_name_from_quoted_inner(inner: &str) -> Option<String> {
    let chars: Vec<char> = inner.chars().collect();
    if chars.first() == Some(&'[') {
        // 必须从开头就是 `[...]`：索引式，或省略路径的簿名式。
        let close = chars.iter().position(|c| *c == ']')?;
        let content: String = chars[1..close].iter().collect();
        let sheet: String = chars[close + 1..].iter().collect();
        if is_all_digits(&content) || looks_like_workbook_file(&content) {
            return (!sheet.is_empty()).then_some(sheet);
        }
        return None;
    }
    // 路径式：`路径\[工作簿.xlsx]表名`
    let open = chars.iter().position(|c| *c == '[')?;
    let close = chars[open..].iter().position(|c| *c == ']')? + open;
    let content: String = chars[open + 1..close].iter().collect();
    if !looks_like_workbook_file(&content) {
        return None;
    }
    let sheet: String = chars[close + 1..].iter().collect();
    (!sheet.is_empty()).then_some(sheet)
}

/// 从 `chars[index]`（指向 `'`）尝试吃掉一个引号式外部引用。
/// 返回 (消耗的字符数, 工作表名)。消耗数包含末尾的 `!`。
fn try_quoted_external(chars: &[char], index: usize) -> Option<(usize, String)> {
    let mut cursor = index + 1;
    let mut inner = String::new();
    loop {
        let current = *chars.get(cursor)?;
        if current == '\'' {
            // `''` 是转义的单引号，继续读；单个 `'` 是收尾。
            if chars.get(cursor + 1) == Some(&'\'') {
                inner.push('\'');
                cursor += 2;
                continue;
            }
            break;
        }
        inner.push(current);
        cursor += 1;
    }
    if chars.get(cursor + 1) != Some(&'!') {
        return None;
    }
    let sheet = sheet_name_from_quoted_inner(&inner)?;
    Some((cursor + 2 - index, sheet))
}

/// 从 `chars[index]`（指向 `[`）尝试吃掉一个裸式外部引用。
/// 返回 (消耗的字符数, 工作表名)。消耗数包含末尾的 `!`。
fn try_bare_external(chars: &[char], index: usize) -> Option<(usize, String)> {
    // 结构化引用（如 `Table1[列]`）不是外部引用：`[` 紧跟标识符时跳过。
    if index > 0 {
        let previous = chars[index - 1];
        if previous.is_alphanumeric() || previous == '_' || previous == ']' {
            return None;
        }
    }
    let close = chars[index..].iter().position(|c| *c == ']')? + index;
    let content: String = chars[index + 1..close].iter().collect();
    if !(is_all_digits(&content) || looks_like_workbook_file(&content)) {
        return None;
    }
    // `[N]'表名'!`：表名按 Excel 规则带引号。
    if chars.get(close + 1) == Some(&'\'') {
        let mut quoted_cursor = close + 2;
        let mut quoted_sheet = String::new();
        loop {
            let current = *chars.get(quoted_cursor)?;
            if current == '\'' {
                if chars.get(quoted_cursor + 1) == Some(&'\'') {
                    quoted_sheet.push('\'');
                    quoted_cursor += 2;
                    continue;
                }
                break;
            }
            quoted_sheet.push(current);
            quoted_cursor += 1;
        }
        if chars.get(quoted_cursor + 1) != Some(&'!') || quoted_sheet.is_empty() {
            return None;
        }
        return Some((quoted_cursor + 2 - index, quoted_sheet));
    }
    let mut cursor = close + 1;
    let start = cursor;
    while let Some(current) = chars.get(cursor) {
        if *current == '!' {
            break;
        }
        // 裸表名不允许空白、引号、括号等（Excel 对这类名字强制引号）。
        if current.is_whitespace()
            || matches!(*current, '\'' | '[' | ']' | '<' | '>' | '(' | ')' | '"')
        {
            return None;
        }
        cursor += 1;
    }
    if chars.get(cursor) != Some(&'!') || cursor == start {
        return None;
    }
    let sheet: String = chars[start..cursor].iter().collect();
    Some((cursor + 1 - index, sheet))
}

/// 剥离公式里的外部工作簿前缀，返回改写文本、剥离数与缺失表名。
///
/// 按业务口径**无条件剥离**（副本只允许工作簿内引用）：
/// - 表名在本副本存在同名工作表 → 公式指向本地表，正常重算；
/// - 不存在 → 公式自然成为 `#REF!`，按审核异常口径计入，不阻断流程。
///
/// 覆盖四种实测形态（`local_sheets` 只用于统计缺失表名，不决定是否剥离）：
/// ① `'C:\外部文件夹\[工作簿.xlsx]集中系统数据'!$A$1` → `'集中系统数据'!$A$1`
/// ② `'[工作簿.xlsx]集中系统数据'!$A$1` → `'集中系统数据'!$A$1`
/// ③ `[工作簿.xlsx]集中系统数据!$A$1` → `集中系统数据!$A$1`
/// ④ `[1]集中系统数据!$A$1` / `[1]'集中系统数据'!$A$1` → 去前缀
///
/// 双引号字符串字面量内的内容一律不动；本簿内引用（`表名!A1`）不动。
fn strip_external_refs(text: &str, local_sheets: &[String]) -> ExternalStripOutcome {
    let chars: Vec<char> = text.chars().collect();
    let mut out = String::with_capacity(text.len());
    let mut outcome = ExternalStripOutcome::default();
    let mut missing: BTreeSet<String> = BTreeSet::new();
    let mut index = 0usize;
    let mut in_string = false;
    while index < chars.len() {
        let current = chars[index];
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
            if let Some((consumed, sheet)) = try_quoted_external(&chars, index) {
                out.push('\'');
                out.push_str(&sheet.replace('\'', "''"));
                out.push('\'');
                outcome.stripped += 1;
                if !local_sheets.iter().any(|local| local == &sheet) {
                    missing.insert(sheet);
                }
                // 已消费到 `!` 之前；让下一轮把 `!` 原样写出。
                index += consumed - 1;
                continue;
            }
        } else if current == '[' {
            if let Some((consumed, sheet)) = try_bare_external(&chars, index) {
                out.push_str(&sheet);
                outcome.stripped += 1;
                if !local_sheets.iter().any(|local| local == &sheet) {
                    missing.insert(sheet);
                }
                index += consumed - 1;
                continue;
            }
        }
        out.push(current);
        index += 1;
    }
    outcome.text = out;
    outcome.missing_sheets = missing.into_iter().collect();
    outcome
}

/// 文本里是否仍残留任何外部工作簿引用（索引式或路径式，引号或裸式）。
///
/// **必须与剥离共用同一解析器**：只查 `[N]` 会在路径式引用残留时误判"无残留"，
/// 进而删除 externalLink 定义，把原本靠外链存活的引用变成新的悬空引用。
fn has_external_refs(text: &str) -> bool {
    let chars: Vec<char> = text.chars().collect();
    let mut index = 0usize;
    let mut in_string = false;
    while index < chars.len() {
        let current = chars[index];
        if in_string {
            if current == '"' {
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
            if try_quoted_external(&chars, index).is_some() {
                return true;
            }
        } else if current == '[' && try_bare_external(&chars, index).is_some() {
            return true;
        }
        index += 1;
    }
    false
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
    /// 剥离后指向"本副本无同名工作表"的表名（公式将按 `#REF!` 计入审核异常）。
    pub missing_sheet_refs: Vec<String>,
    /// 是否已彻底断开外链定义（清 externalReferences / 关系 / Content_Types / 部件）。
    pub external_links_dropped: bool,
    pub external_links_removed: u64,
    pub reason: String,
}

/// 对整个部件执行外部引用剥离。
///
/// 走 roxmltree 解码后再处理：生成侧可能把中文表名写成 `&#23384;` 这类数字
/// 字符引用，解码后即还原为字面文本，匹配不再需要 NCR 弹性正则；写回时用
/// 公式内容最少转义（只处理 `& < >`）即可保持合法 XML。
fn strip_text_external_refs(
    text: &str,
    local_sheet_names: &[String],
) -> (String, u64, Vec<String>) {
    let Ok(document) = Document::parse(text) else {
        return (text.to_owned(), 0, Vec::new());
    };
    let mut edits: Vec<(std::ops::Range<usize>, Vec<u8>)> = Vec::new();
    let mut total = 0u64;
    let mut missing: BTreeSet<String> = BTreeSet::new();
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
        let outcome = strip_external_refs(original, local_sheet_names);
        if outcome.stripped == 0 {
            continue;
        }
        total += outcome.stripped;
        missing.extend(outcome.missing_sheets);
        let range = node.range();
        let open_end = match text[range.clone()].find('>') {
            Some(offset) => offset,
            None => continue,
        };
        let open_tag = text[range.start..range.start + open_end + 1].to_owned();
        let replacement = format!(
            "{open_tag}{}</{}>",
            escape_formula_content(&outcome.text),
            node.tag_name().name()
        );
        edits.push((range, replacement.into_bytes()));
    }
    if edits.is_empty() {
        return (text.to_owned(), 0, Vec::new());
    }
    let mut result = text.as_bytes().to_vec();
    for (range, replacement) in edits.into_iter().rev() {
        result.splice(range, replacement);
    }
    (
        String::from_utf8_lossy(&result).into_owned(),
        total,
        missing.into_iter().collect(),
    )
}

/// 部件里是否仍残留任何外部工作簿引用（剥离后的残留判定）。
fn part_has_external_refs(text: &str) -> bool {
    let Ok(document) = Document::parse(text) else {
        // 解析失败时按"有残留"保守处理：宁可不删外链，也不制造新的悬空引用。
        return true;
    };
    document
        .descendants()
        .filter(|node| {
            node.is_element()
                && matches!(
                    node.tag_name().name(),
                    "f" | "formula" | "c:f" | "definedName"
                )
        })
        .filter_map(|node| node.text())
        .any(has_external_refs)
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
        // 剥离是无条件的（副本只允许工作簿内引用），因此有外部引用就要修。
        if part_has_external_refs(text) {
            return true;
        }
        // 外链定义本身也属于待清理内容（剥离后若已无残留会整体断开）。
        if matches!(
            name.as_str(),
            "xl/workbook.xml" | "[Content_Types].xml" | "xl/_rels/workbook.xml.rels"
        ) && (text.contains("<externalReferences") || text.contains("externalLink"))
        {
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

/// 统计文本中的外部工作簿引用数量（四种形态统一口径）。
fn count_external_prefixes(text: &str) -> u64 {
    let chars: Vec<char> = text.chars().collect();
    let mut count = 0u64;
    let mut index = 0usize;
    let mut in_string = false;
    while index < chars.len() {
        let current = chars[index];
        if in_string {
            if current == '"' {
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
            if try_quoted_external(&chars, index).is_some() {
                count += 1;
                index += 1;
                continue;
            }
        } else if current == '[' && try_bare_external(&chars, index).is_some() {
            count += 1;
        }
        index += 1;
    }
    count
}

/// 取第一条外部引用的原文片段用于错误信息定位。
fn first_external_prefix(text: &str) -> Option<String> {
    let chars: Vec<char> = text.chars().collect();
    let mut index = 0usize;
    let mut in_string = false;
    while index < chars.len() {
        let current = chars[index];
        if in_string {
            if current == '"' {
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
        let consumed = if current == '\'' {
            try_quoted_external(&chars, index).map(|(consumed, _)| consumed)
        } else if current == '[' {
            try_bare_external(&chars, index).map(|(consumed, _)| consumed)
        } else {
            None
        };
        if let Some(consumed) = consumed {
            return Some(
                chars[index..(index + consumed).min(chars.len())]
                    .iter()
                    .collect(),
            );
        }
        index += 1;
    }
    None
}

/// 彻底断开外链：清 `<externalReferences>`、删关系与 Content_Types 声明、
/// 剔除 `xl/externalLinks/` 全部部件。
///
/// 调用方必须先确认全包已无任何外部引用残留（`has_external_refs`）；仍残留时
/// 删除外链定义会把这些引用变成新的悬空引用。返回移除的关系条数。
fn drop_external_link_parts(parts: &mut Vec<(String, Vec<u8>)>) -> u64 {
    let mut removed = 0u64;
    for (name, data) in parts.iter_mut() {
        match name.as_str() {
            "xl/workbook.xml" => {
                let Ok(text) = std::str::from_utf8(data) else {
                    continue;
                };
                let mut rewritten = text.to_owned();
                // 成对标签与自闭合两种变体；以第一个 `>` 判定开标签边界，
                // 再看其前一字符是否为 `/` 区分自闭合。
                while let Some(start) = rewritten.find("<externalReferences") {
                    let after_name = start + "<externalReferences".len();
                    let open_end = match rewritten[after_name..].find('>') {
                        Some(offset) => after_name + offset,
                        None => break,
                    };
                    if rewritten.as_bytes().get(open_end.wrapping_sub(1)) == Some(&b'/') {
                        rewritten.replace_range(start..open_end + 1, "");
                        continue;
                    }
                    let end = match rewritten[open_end..].find("</externalReferences>") {
                        Some(offset) => open_end + offset + "</externalReferences>".len(),
                        None => break,
                    };
                    rewritten.replace_range(start..end, "");
                }
                *data = rewritten.into_bytes();
            }
            "xl/_rels/workbook.xml.rels" => {
                let Ok(text) = std::str::from_utf8(data) else {
                    continue;
                };
                let (rewritten, count) = remove_relationships_of_type(text, "externalLink");
                removed += count;
                *data = rewritten.into_bytes();
            }
            "[Content_Types].xml" => {
                let Ok(text) = std::str::from_utf8(data) else {
                    continue;
                };
                let mut rewritten = text.to_owned();
                let mut cursor = 0usize;
                while let Some(start) = rewritten[cursor..].find("<Override") {
                    let start = cursor + start;
                    let Some(end_rel) = rewritten[start..].find("/>") else {
                        break;
                    };
                    let end = start + end_rel + 2;
                    let entry = &rewritten[start..end];
                    if entry.contains("/xl/externalLinks/") {
                        rewritten.replace_range(start..end, "");
                        cursor = start;
                    } else {
                        cursor = end;
                    }
                }
                *data = rewritten.into_bytes();
            }
            _ => {}
        }
    }
    let before = parts.len();
    parts.retain(|(name, _)| !name.starts_with("xl/externalLinks/"));
    removed += (before - parts.len()) as u64;
    removed
}

/// 从 rels 里删掉指定 Type 结尾的关系条目，返回删除条数。
fn remove_relationships_of_type(text: &str, type_suffix: &str) -> (String, u64) {
    let mut rewritten = text.to_owned();
    let mut removed = 0u64;
    let mut cursor = 0usize;
    while let Some(offset) = rewritten[cursor..].find("<Relationship") {
        let start = cursor + offset;
        // 必须是 `<Relationship` 元素，不能命中根元素 `<Relationships>`。
        let after = rewritten
            .as_bytes()
            .get(start + "<Relationship".len())
            .copied();
        if !matches!(
            after,
            Some(b' ') | Some(b'\t') | Some(b'\r') | Some(b'\n') | Some(b'/') | Some(b'>')
        ) {
            cursor = start + "<Relationship".len();
            continue;
        }
        let Some(open_end_rel) = rewritten[start..].find('>') else {
            break;
        };
        let open_end = start + open_end_rel;
        let self_closing = rewritten.as_bytes().get(open_end.wrapping_sub(1)) == Some(&b'/');
        let end = if self_closing {
            open_end + 1
        } else {
            match rewritten[open_end..].find("</Relationship>") {
                Some(close_rel) => open_end + close_rel + "</Relationship>".len(),
                None => break,
            }
        };
        let entry = &rewritten[start..end];
        if entry.contains(type_suffix) {
            rewritten.replace_range(start..end, "");
            removed += 1;
            cursor = start;
        } else {
            cursor = end;
        }
    }
    (rewritten, removed)
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
            // 剥离外部工作簿前缀（无条件：副本只允许工作簿内引用）。
            // roxmltree 解码后处理，天然兼容 openpyxl 写出的 NCR 实体。
            let (stripped_text, stripped_count, missing) =
                strip_text_external_refs(&rewritten, &local_sheet_names);
            rewritten = stripped_text;
            report.external_refs_stripped += stripped_count;
            for name in missing {
                if !report.missing_sheet_refs.contains(&name) {
                    report.missing_sheet_refs.push(name);
                }
            }
            let (pane_text, pane_fixed) = fix_frozen_panes(&rewritten);
            report.pane_fixed += pane_fixed;
            writer.write_all(pane_text.as_bytes())?;
        }
        Ok(())
    })();
    let finish_result = writer.finish().context("finish repaired ZIP");
    match (write_result, finish_result) {
        (Ok(()), Ok(_)) => {}
        (Err(error), _) | (_, Err(error)) => {
            let _ = fs::remove_file(&temporary);
            report.reason = format!("{error:#}");
            return Err(anyhow::anyhow!(report.reason.clone()));
        }
    }
    // 剥离后确认全包再无任何外部引用（索引式或路径式）才彻底断开外链；
    // 仍有残留时保持外链定义原样，避免制造新的悬空引用。
    if report.external_refs_stripped > 0 {
        let mut repaired: Vec<(String, Vec<u8>)> = Vec::new();
        {
            let file = fs::File::open(&temporary)?;
            let mut archive = zip::ZipArchive::new(file)?;
            for index in 0..archive.len() {
                let mut entry = archive.by_index(index)?;
                let name = entry.name().to_owned();
                let mut data = Vec::new();
                entry.read_to_end(&mut data)?;
                repaired.push((name, data));
            }
        }
        let still_dangling = repaired.iter().any(|(name, data)| {
            is_repairable_part(name)
                && std::str::from_utf8(data)
                    .ok()
                    .is_some_and(part_has_external_refs)
        });
        if !still_dangling {
            let removed = drop_external_link_parts(&mut repaired);
            report.external_links_dropped = true;
            report.external_links_removed = removed;
            let file = fs::OpenOptions::new()
                .write(true)
                .truncate(true)
                .open(&temporary)
                .with_context(|| format!("rewrite repaired ZIP {}", temporary.display()))?;
            let mut writer = zip::ZipWriter::new(file);
            for (name, data) in &repaired {
                let options = zip::write::SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Deflated);
                writer.start_file(name.as_str(), options)?;
                writer.write_all(data)?;
            }
            writer.finish().context("finish external-link cleanup")?;
        }
    }
    Ok(temporary)
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
    fn strips_all_four_external_reference_shapes() {
        let cases = [
            // ① 路径式 + 引号 + 完整路径
            (
                r"'C:\外部文件夹\[工作簿.xlsx]集中系统数据'!$A$1",
                "'集中系统数据'!$A$1",
            ),
            // ② 路径式 + 引号 + 仅簿名
            ("'[工作簿.xlsx]集中系统数据'!$A$1", "'集中系统数据'!$A$1"),
            // ③ 路径式 + 无引号
            ("[工作簿.xlsx]集中系统数据!$A$1", "集中系统数据!$A$1"),
            // ④ 索引式（裸 / 引号）
            ("[1]集中系统数据!$A$1", "集中系统数据!$A$1"),
            ("[2]'集中系统数据'!$A$1", "集中系统数据!$A$1"),
        ];
        for (input, expected) in cases {
            let outcome = strip_external_refs(input, &["集中系统数据".to_owned()]);
            assert_eq!(outcome.stripped, 1, "{input} -> {}", outcome.text);
            assert_eq!(outcome.text, expected, "input {input}");
            assert!(outcome.missing_sheets.is_empty(), "{input}");
        }
    }

    #[test]
    fn strips_inside_functions_and_keeps_local_references() {
        // 函数内出现 + 本簿内引用不动 + 双引号字符串不动
        let text = r#"IFERROR(INDEX('C:\[簿.xlsm]数据'!$A:$A,1),"数据!不是引用")+本地表!B2"#;
        let outcome = strip_external_refs(text, &["数据".to_owned(), "本地表".to_owned()]);
        assert_eq!(outcome.stripped, 1, "{}", outcome.text);
        assert!(
            outcome.text.contains("INDEX('数据'!$A:$A,1)"),
            "{}",
            outcome.text
        );
        assert!(outcome.text.contains("本地表!B2"), "{}", outcome.text);
        assert!(
            outcome.text.contains(r#""数据!不是引用""#),
            "{}",
            outcome.text
        );
    }

    #[test]
    fn strips_workbook_names_with_spaces_and_all_extensions() {
        for extension in ["xlsx", "xlsm", "xlsb", "xls"] {
            let input = format!("'[某 工作簿.{extension}]某 表'!B2");
            let outcome = strip_external_refs(&input, &["某 表".to_owned()]);
            assert_eq!(outcome.stripped, 1, "{input} -> {}", outcome.text);
            assert_eq!(outcome.text, "'某 表'!B2", "{input}");
        }
    }

    #[test]
    fn reports_missing_local_sheets_after_strip() {
        // 副本未导入外部表时：照样剥离，并记录缺失表名（公式将按 #REF! 计入异常）
        let text = "[1]集中系统数据!A1+[1]参照表!B2";
        let outcome = strip_external_refs(text, &["其他表".to_owned()]);
        assert_eq!(outcome.stripped, 2, "{}", outcome.text);
        assert_eq!(
            outcome.missing_sheets,
            vec!["参照表".to_owned(), "集中系统数据".to_owned()]
        );
        assert_eq!(outcome.text, "集中系统数据!A1+参照表!B2");
    }

    #[test]
    fn strip_is_idempotent_and_detection_matches_strip() {
        let text = r#"'C:\a\[簿.xlsx]数据'!A1+[2]参照表!B2"#;
        let local = vec!["数据".to_owned(), "参照表".to_owned()];
        let first = strip_external_refs(text, &local);
        assert_eq!(first.stripped, 2);
        assert!(!has_external_refs(&first.text), "{}", first.text);
        let second = strip_external_refs(&first.text, &local);
        assert_eq!(second.stripped, 0);
        assert_eq!(second.text, first.text);
        // 检测与剥离同源：残留的单引号式路径引用也必须被识别
        assert!(has_external_refs("'[簿.xlsx]表'!A1"));
        assert!(!has_external_refs("本地表!A1"));
    }

    #[test]
    fn drops_external_link_definitions_after_strip() {
        let mut parts = vec![
            (
                "xl/workbook.xml".to_owned(),
                br#"<workbook><sheets/><externalReferences><externalReference r:id="rId8"/></externalReferences><definedNames/></workbook>"#.to_vec(),
            ),
            (
                "xl/_rels/workbook.xml.rels".to_owned(),
                br#"<Relationships xmlns="x"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>"#.to_vec(),
            ),
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="x"/><Override PartName="/xl/workbook.xml" ContentType="y"/></Types>"#.to_vec(),
            ),
            (
                "xl/externalLinks/externalLink1.xml".to_owned(),
                b"<externalLink/>".to_vec(),
            ),
        ];
        let removed = drop_external_link_parts(&mut parts);
        assert!(removed >= 2, "removed={removed}");
        let workbook = std::str::from_utf8(
            &parts
                .iter()
                .find(|(n, _)| n == "xl/workbook.xml")
                .unwrap()
                .1,
        )
        .unwrap();
        assert!(!workbook.contains("externalReferences"), "{workbook}");
        assert!(workbook.contains("<definedNames/>"), "{workbook}");
        let rels = std::str::from_utf8(
            &parts
                .iter()
                .find(|(n, _)| n.ends_with("workbook.xml.rels"))
                .unwrap()
                .1,
        )
        .unwrap();
        assert!(!rels.contains("externalLink"), "{rels}");
        assert!(rels.contains("worksheets/sheet1.xml"), "{rels}");
        // 根元素必须保留
        assert!(
            rels.trim_start().starts_with("<?xml") || rels.contains("<Relationships"),
            "{rels}"
        );
        let types = std::str::from_utf8(
            &parts
                .iter()
                .find(|(n, _)| n == "[Content_Types].xml")
                .unwrap()
                .1,
        )
        .unwrap();
        assert!(!types.contains("externalLinks"), "{types}");
        assert!(types.contains("/xl/workbook.xml"), "{types}");
        assert!(
            !parts
                .iter()
                .any(|(n, _)| n.starts_with("xl/externalLinks/")),
            "external link parts must be removed"
        );
    }

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
