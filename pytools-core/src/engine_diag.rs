//! 统一引擎诊断日志（第一阶段：报表打印功能接入）。
//!
//! 目标：让业务功能在调用 Excel 引擎（calamine / rust_xlsxwriter / umya /
//! openpyxl / OOXML Patch）时复用统一的结构化诊断信息，包含引擎、操作阶段、
//! 结果、上下文（文件/工作表/配置行号/耗时）与原始错误信息。
//!
//! 约束：本模块只产生与收集事件，由界面层统一显示；不改变引擎选择、回退
//! 策略、文件写入与业务结果；不记录工作簿单元格内容；时间戳使用本机当前
//! 时间（格式沿用旧版 PyTools：`%Y-%m-%d %H:%M:%S`）。

use chrono::Local;
use serde::Serialize;
use std::fmt;

/// 当前本机时间前缀，格式 `[YYYY-MM-DD HH:MM:SS]`。运行时实时取值，不含写死的时间。
pub fn now_prefix() -> String {
    format!("[{}]", Local::now().format("%Y-%m-%d %H:%M:%S"))
}

/// 将子进程回传的 Unix 秒转换为同样的本机时间前缀（保留事件实际发生时刻）。
pub fn prefix_from_epoch(seconds: i64) -> String {
    match chrono::DateTime::from_timestamp(seconds, 0) {
        Some(utc) => format!(
            "[{}]",
            utc.with_timezone(&Local).format("%Y-%m-%d %H:%M:%S")
        ),
        None => now_prefix(),
    }
}

/// 统一的操作阶段。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum DiagStage {
    ReadConfig,
    ReadSource,
    CopySheet,
    ModifyWorkbook,
    SaveOutput,
}

impl fmt::Display for DiagStage {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let text = match self {
            Self::ReadConfig => "读取配置",
            Self::ReadSource => "读取源文件",
            Self::CopySheet => "复制工作表",
            Self::ModifyWorkbook => "修改工作簿",
            Self::SaveOutput => "保存输出",
        };
        write!(formatter, "{text}")
    }
}

/// 统一的操作结果。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum DiagOutcome {
    Started,
    Ok,
    Skipped,
    Failed,
}

impl fmt::Display for DiagOutcome {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let text = match self {
            Self::Started => "开始",
            Self::Ok => "成功",
            Self::Skipped => "跳过",
            Self::Failed => "失败",
        };
        write!(formatter, "{text}")
    }
}

/// 失败事件的显示标签（阶段 + 失败），例如“读取失败”“保存失败”。
#[must_use]
pub fn stage_failure_label(stage: DiagStage) -> &'static str {
    match stage {
        DiagStage::ReadConfig => "配置失败",
        DiagStage::ReadSource => "读取失败",
        DiagStage::CopySheet => "复制失败",
        DiagStage::ModifyWorkbook => "修改失败",
        DiagStage::SaveOutput => "保存失败",
    }
}

/// 一条结构化诊断事件。`timestamp` 在事件创建（或子进程回传）时生成，
/// 结束汇总显示时沿用该时刻，不使用汇总时的当前时间。
#[derive(Clone, Debug, Serialize)]
pub struct DiagEvent {
    pub timestamp: String,
    pub timestamp_rfc3339: String,
    pub engine: String,
    pub stage: DiagStage,
    pub outcome: DiagOutcome,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub row_no: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub elapsed_ms: Option<u128>,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub details: Option<serde_json::Value>,
}

impl DiagEvent {
    #[must_use]
    pub fn new(engine: &str, stage: DiagStage, outcome: DiagOutcome) -> Self {
        Self {
            timestamp: now_prefix(),
            timestamp_rfc3339: Local::now().to_rfc3339(),
            engine: engine.to_owned(),
            stage,
            outcome,
            row_no: None,
            file: None,
            sheet: None,
            elapsed_ms: None,
            message: String::new(),
            details: None,
        }
    }

    /// Existing adapters persist this independently of terminal verbosity.
    #[must_use]
    pub fn diagnostic_details(&self) -> serde_json::Value {
        let mut details = self
            .details
            .as_ref()
            .and_then(serde_json::Value::as_object)
            .cloned()
            .unwrap_or_default();
        if !self.message.is_empty() {
            details.insert(
                "diagnostic".to_owned(),
                serde_json::Value::String(self.message.clone()),
            );
        }
        serde_json::Value::Object(details)
    }

    #[must_use]
    pub fn with_details(mut self, details: serde_json::Value) -> Self {
        self.details = Some(details);
        self
    }

    #[must_use]
    pub fn with_row(mut self, row: usize) -> Self {
        self.row_no = Some(row);
        self
    }

    #[must_use]
    pub fn with_file(mut self, file: impl Into<String>) -> Self {
        self.file = Some(file.into());
        self
    }

    #[must_use]
    pub fn with_sheet(mut self, sheet: impl Into<String>) -> Self {
        self.sheet = Some(sheet.into());
        self
    }

    #[must_use]
    pub fn with_elapsed_ms(mut self, elapsed_ms: u128) -> Self {
        self.elapsed_ms = Some(elapsed_ms);
        self
    }

    #[must_use]
    pub fn with_message(mut self, message: impl Into<String>) -> Self {
        self.message = message.into();
        self
    }

    /// 用子进程回传的 Unix 秒覆盖事件时间（保留实际发生时刻）。
    #[must_use]
    pub fn with_epoch(mut self, seconds: Option<i64>) -> Self {
        if let Some(seconds) = seconds {
            self.timestamp = prefix_from_epoch(seconds);
            if let Some(time) = chrono::DateTime::from_timestamp(seconds, 0) {
                self.timestamp_rfc3339 = time.with_timezone(&Local).to_rfc3339();
            }
        }
        self
    }

    /// 单行显示。失败事件形如：
    /// `[2026-10-28 16:04:48] [读取失败][umya] 第 38 行：源工作簿 xxx.xlsx，具体错误：……`
    #[must_use]
    pub fn line(&self) -> String {
        let mut line = format!("{} ", self.timestamp);
        match self.outcome {
            DiagOutcome::Failed => {
                line.push_str(&format!(
                    "[{}][{}]",
                    stage_failure_label(self.stage),
                    self.engine
                ));
            }
            DiagOutcome::Skipped => {
                line.push_str(&format!("[跳过][{}]", self.engine));
            }
            DiagOutcome::Ok => {
                line.push_str(&format!("[{}成功][{}]", self.stage, self.engine));
            }
            DiagOutcome::Started => {
                line.push_str(&format!("[{}开始][{}]", self.stage, self.engine));
            }
        }
        match self.row_no {
            Some(row) => line.push_str(&format!(" 第 {row} 行：")),
            None => line.push('：'),
        }
        let mut details: Vec<String> = Vec::new();
        match (self.stage, self.file.as_deref()) {
            (DiagStage::ReadSource, Some(file)) => {
                details.push(format!("源工作簿 {file}"));
            }
            (DiagStage::SaveOutput, Some(file)) => {
                details.push(format!("输出文件 {file}"));
            }
            (_, Some(file)) => {
                details.push(format!("文件 {file}"));
            }
            (_, None) => {}
        }
        if let Some(sheet) = &self.sheet {
            details.push(format!("工作表 {sheet}"));
        }
        if let Some(elapsed_ms) = self.elapsed_ms {
            details.push(format!("耗时 {:.1} 秒", elapsed_ms as f64 / 1000.0));
        }
        if !details.is_empty() {
            line.push_str(&details.join("，"));
        }
        if !self.message.is_empty() {
            if details.is_empty() {
                line.push_str("具体错误：");
            } else {
                line.push_str("，具体错误：");
            }
            line.push_str(&self.message);
        }
        line
    }

    /// 中文提示只保留定位和已知类别，避免把引擎错误中的路径或业务值原样倾倒到终端。
    #[must_use]
    pub fn user_line(&self) -> String {
        let level = match self.outcome {
            DiagOutcome::Failed => "ERROR",
            DiagOutcome::Skipped => "WARN",
            DiagOutcome::Ok | DiagOutcome::Started => "INFO",
        };
        let mut line = format!(
            "{} [{}] {}",
            self.timestamp,
            level,
            match self.outcome {
                DiagOutcome::Failed => "失败",
                DiagOutcome::Skipped => "跳过",
                _ => "处理完成",
            }
        );
        if let Some(row) = self.row_no {
            line.push_str(&format!("；配置第 {row} 行"));
        }
        if let Some(file) = &self.file {
            let name = file.rsplit(['/', '\\']).next().unwrap_or(file);
            line.push_str(&format!("；文件 {name}"));
        }
        if let Some(sheet) = &self.sheet {
            line.push_str(&format!("；工作表 {sheet}"));
        }
        if matches!(self.outcome, DiagOutcome::Failed | DiagOutcome::Skipped) {
            let text = self.message.to_lowercase();
            let reason = if text.contains("默认格式") || text.contains("默认行高") {
                "源格式未能完整读取，输出使用默认格式，请检查打印结果".to_owned()
            } else if text.contains("工作表")
                && (text.contains("缺失") || text.contains("不存在") || text.contains("missing"))
            {
                "工作表不存在，请检查名称".to_owned()
            } else if text.contains("不存在")
                || text.contains("not found")
                || text.contains("no such file")
            {
                "文件或配置项不存在，请检查路径".to_owned()
            } else if text.contains("permission")
                || text.contains("权限")
                || text.contains("被占用")
            {
                "文件被占用或没有访问权限，请关闭占用程序并检查权限".to_owned()
            } else if text.contains("区域") || text.contains("range") {
                "区域配置或复制失败，请检查区域和工作表".to_owned()
            } else if text.contains("格式") || text.contains("unsupported") {
                "文件格式或配置不受支持，请检查输入".to_owned()
            } else if self.outcome == DiagOutcome::Skipped {
                format!("{}阶段未执行，请检查对应文件和配置", self.stage)
            } else {
                format!("{}阶段失败，请检查对应文件、配置和访问权限", self.stage)
            };
            line.push_str(&format!("；{reason}"));
        }
        line
    }
}

/// 事件收集器。业务层在引擎调用点推送事件，界面层在结束时统一显示。
#[derive(Clone, Debug, Default)]
pub struct DiagLog {
    events: Vec<DiagEvent>,
}

impl DiagLog {
    pub fn push(&mut self, event: DiagEvent) {
        self.events.push(event);
    }

    #[must_use]
    pub fn events(&self) -> &[DiagEvent] {
        &self.events
    }

    pub fn into_events(self) -> Vec<DiagEvent> {
        self.events
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn timestamp_prefix_matches_pytools_format() {
        let prefix = now_prefix();
        let bytes = prefix.as_bytes();
        assert_eq!(bytes[0], b'[', "{prefix}");
        assert_eq!(bytes[5], b'-', "{prefix}");
        assert_eq!(bytes[8], b'-', "{prefix}");
        assert_eq!(bytes[11], b' ', "{prefix}");
        assert_eq!(bytes[14], b':', "{prefix}");
        assert_eq!(bytes[17], b':', "{prefix}");
        assert_eq!(bytes[20], b']', "{prefix}");
        assert_eq!(prefix.len(), 21, "{prefix}");
    }

    #[test]
    fn failure_line_follows_unified_shape() {
        let event = DiagEvent::new(
            "umya-spreadsheet",
            DiagStage::ReadSource,
            DiagOutcome::Failed,
        )
        .with_row(38)
        .with_file("src.xlsx")
        .with_message("zip 错误：文件头缺失");
        let line = event.line();
        assert!(line.contains("[读取失败][umya-spreadsheet]"), "{line}");
        assert!(line.contains("第 38 行："), "{line}");
        assert!(line.contains("源工作簿 src.xlsx"), "{line}");
        assert!(line.contains("具体错误：zip 错误：文件头缺失"), "{line}");
        assert!(line.starts_with('['), "{line}");
    }

    #[test]
    fn detailed_diagnostics_do_not_expand_user_line() {
        let event = DiagEvent::new(
            "rust_xlsxwriter",
            DiagStage::ModifyWorkbook,
            DiagOutcome::Skipped,
        )
        .with_file("source.xlsx")
        .with_message("XML parse error: source style unavailable; 输出使用默认格式")
        .with_details(serde_json::json!({"stage_timings": {"source_read_ms": 12}}));
        let user = event.user_line();
        assert!(user.contains("默认格式"));
        assert!(user.contains("source.xlsx"));
        assert!(!user.contains("XML"));
        assert!(!user.contains("rust_xlsxwriter"));
        let details = event.diagnostic_details();
        assert_eq!(details["stage_timings"]["source_read_ms"], 12);
        assert!(details["diagnostic"]
            .as_str()
            .unwrap()
            .contains("XML parse error"));
    }

    #[test]
    fn epoch_prefix_converts_to_local_time() {
        // 2026-10-28 08:04:30 UTC；本地时区偏移在断言中只检查格式而非数值。
        let prefix = prefix_from_epoch(1_792_253_070);
        assert!(prefix.starts_with('[') && prefix.ends_with(']'), "{prefix}");
        assert_eq!(prefix.len(), 21, "{prefix}");
    }
}
