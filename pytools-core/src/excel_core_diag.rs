//! Excel-Core 统一诊断适配层（第二阶段）。
//!
//! 把 `excel-editor` 已有的执行记录（`EditExecutionReport`：引擎、逐后端尝试、
//! 结果、耗时、原始消息）与 `ExcelModifierError` 转换成统一诊断事件，供各业务
//! 功能复用。**不新建重复的引擎选择或回退机制**，只做记录到事件的映射；
//! 转换过程不改变任何执行结果。
//!
//! 读取侧（excel-reader / calamine）与新建侧（excel-writer / rust_xlsxwriter）
//! 没有内建执行记录，这里提供轻量构造器，由调用方在关键阶段各记一条，
//! 避免逐单元格写日志。

use crate::engine_diag::{DiagEvent, DiagOutcome, DiagStage};
use excel_editor::{BackendAttemptOutcome, EditExecutionReport, EngineId, ExcelModifierError};

/// 引擎名统一为诊断使用的可读文本。
#[must_use]
pub fn engine_label(engine: EngineId) -> &'static str {
    match engine {
        EngineId::OoxmlPatch => "ooxml-patch",
        EngineId::UmyaSpreadsheet => "umya-spreadsheet",
        EngineId::OpenpyxlWorker => "openpyxl",
    }
}

/// 把一次成功的 ExcelModifier 执行转换为诊断事件：
/// 每个成功/被拒的后端尝试各一条，另加一条整体保存成功事件。
#[must_use]
pub fn events_from_execution(
    report: &EditExecutionReport,
    stage: DiagStage,
    row_no: Option<usize>,
) -> Vec<DiagEvent> {
    let mut events = Vec::with_capacity(report.attempts.len() + 1);
    for attempt in &report.attempts {
        let outcome = match attempt.outcome {
            BackendAttemptOutcome::Succeeded => DiagOutcome::Ok,
            BackendAttemptOutcome::RejectedBeforeWrite => DiagOutcome::Skipped,
            BackendAttemptOutcome::Failed | BackendAttemptOutcome::ValidationFailed => {
                DiagOutcome::Failed
            }
        };
        let mut event = DiagEvent::new(engine_label(attempt.engine), stage, outcome)
            .with_file(report.input.to_string_lossy().into_owned())
            .with_elapsed_ms(attempt.elapsed_ms);
        if let Some(row) = row_no {
            event = event.with_row(row);
        }
        if !attempt.message.is_empty() {
            event = event.with_message(attempt.message.clone());
        }
        events.push(event);
    }
    let mut saved = DiagEvent::new(
        engine_label(report.engine),
        DiagStage::SaveOutput,
        DiagOutcome::Ok,
    )
    .with_file(report.output.to_string_lossy().into_owned())
    .with_elapsed_ms(report.elapsed_ms);
    if let Some(row) = row_no {
        saved = saved.with_row(row);
    }
    events.push(saved);
    events
}

/// 把 ExcelModifier 失败转换为诊断事件：保留错误类别与原始消息，
/// 并逐条附带后端尝试记录（便于用户判断是能力不足还是文件问题）。
#[must_use]
pub fn events_from_modifier_error(
    error: &ExcelModifierError,
    row_no: Option<usize>,
) -> Vec<DiagEvent> {
    let mut events = Vec::with_capacity(error.attempts.len() + 1);
    let engine = error.engine.map(engine_label).unwrap_or("excel-modifier");
    let mut head = DiagEvent::new(engine, DiagStage::ModifyWorkbook, DiagOutcome::Failed)
        .with_message(format!("{:?}: {}", error.kind, error.message));
    if let Some(row) = row_no {
        head = head.with_row(row);
    }
    events.push(head);
    for attempt in &error.attempts {
        let outcome = match attempt.outcome {
            BackendAttemptOutcome::Succeeded => DiagOutcome::Ok,
            BackendAttemptOutcome::RejectedBeforeWrite => DiagOutcome::Skipped,
            BackendAttemptOutcome::Failed | BackendAttemptOutcome::ValidationFailed => {
                DiagOutcome::Failed
            }
        };
        let mut event = DiagEvent::new(
            engine_label(attempt.engine),
            DiagStage::ModifyWorkbook,
            outcome,
        )
        .with_elapsed_ms(attempt.elapsed_ms);
        if let Some(row) = row_no {
            event = event.with_row(row);
        }
        if !attempt.message.is_empty() {
            event = event.with_message(attempt.message.clone());
        }
        events.push(event);
    }
    events
}

/// 读取阶段事件（excel-reader / calamine 没有内建执行记录，由调用方记录）。
#[must_use]
pub fn read_event(
    file: &std::path::Path,
    row_no: Option<usize>,
    elapsed_ms: u128,
    error: Option<&str>,
) -> DiagEvent {
    let mut event = DiagEvent::new(
        "calamine",
        DiagStage::ReadSource,
        if error.is_some() {
            DiagOutcome::Failed
        } else {
            DiagOutcome::Ok
        },
    )
    .with_file(file.to_string_lossy().into_owned())
    .with_elapsed_ms(elapsed_ms);
    if let Some(row) = row_no {
        event = event.with_row(row);
    }
    if let Some(message) = error {
        event = event.with_message(message);
    }
    event
}

/// 新建输出事件（excel-writer / rust_xlsxwriter）。
#[must_use]
pub fn write_event(
    file: &std::path::Path,
    row_no: Option<usize>,
    elapsed_ms: u128,
    error: Option<&str>,
) -> DiagEvent {
    let mut event = DiagEvent::new(
        "rust_xlsxwriter",
        DiagStage::SaveOutput,
        if error.is_some() {
            DiagOutcome::Failed
        } else {
            DiagOutcome::Ok
        },
    )
    .with_file(file.to_string_lossy().into_owned())
    .with_elapsed_ms(elapsed_ms);
    if let Some(row) = row_no {
        event = event.with_row(row);
    }
    if let Some(message) = error {
        event = event.with_message(message);
    }
    event
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_editor::{BackendAttempt, EditExecutionReport};
    use std::path::PathBuf;

    fn report_with(outcome: BackendAttemptOutcome) -> EditExecutionReport {
        EditExecutionReport {
            engine: EngineId::UmyaSpreadsheet,
            input: PathBuf::from("source.xlsx"),
            output: PathBuf::from("target.xlsx"),
            elapsed_ms: 120,
            route_decisions: Vec::new(),
            attempts: vec![BackendAttempt {
                engine: EngineId::UmyaSpreadsheet,
                outcome,
                message: "sample message".to_owned(),
                elapsed_ms: 80,
            }],
        }
    }

    #[test]
    fn successful_execution_yields_attempt_and_save_events() {
        let events = events_from_execution(
            &report_with(BackendAttemptOutcome::Succeeded),
            DiagStage::ModifyWorkbook,
            Some(7),
        );
        assert_eq!(events.len(), 2);
        assert_eq!(events[0].outcome, DiagOutcome::Ok);
        assert_eq!(events[0].engine, "umya-spreadsheet");
        assert_eq!(events[0].row_no, Some(7));
        assert_eq!(events[1].stage, DiagStage::SaveOutput);
        assert!(events[1].file.as_deref() == Some("target.xlsx"));
    }

    #[test]
    fn rejected_attempt_maps_to_skipped() {
        let events = events_from_execution(
            &report_with(BackendAttemptOutcome::RejectedBeforeWrite),
            DiagStage::ModifyWorkbook,
            None,
        );
        assert_eq!(events[0].outcome, DiagOutcome::Skipped);
    }

    #[test]
    fn modifier_error_keeps_original_message_and_attempts() {
        let error = ExcelModifierError {
            kind: excel_editor::ExcelModifierErrorKind::NoEligibleBackend,
            engine: None,
            message: "no backend matched the plan".to_owned(),
            route_decisions: Vec::new(),
            attempts: Vec::new(),
        };
        let events = events_from_modifier_error(&error, Some(3));
        assert_eq!(events.len(), 1);
        assert_eq!(events[0].outcome, DiagOutcome::Failed);
        assert!(events[0].message.contains("NoEligibleBackend"));
        assert!(events[0].message.contains("no backend matched the plan"));
        assert_eq!(events[0].row_no, Some(3));
    }

    #[test]
    fn read_and_write_events_flag_errors() {
        let ok = read_event(std::path::Path::new("a.xlsx"), None, 5, None);
        assert_eq!(ok.outcome, DiagOutcome::Ok);
        let failed = write_event(
            std::path::Path::new("b.xlsx"),
            Some(2),
            9,
            Some("permission denied"),
        );
        assert_eq!(failed.outcome, DiagOutcome::Failed);
        assert!(failed.message.contains("permission denied"));
    }
}
