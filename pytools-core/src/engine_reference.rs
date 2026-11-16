//! Display-only, correctness-gated references. Never selects an engine.
use serde::Deserialize;
use std::{io::Write, sync::OnceLock};

#[derive(Default, Deserialize)]
struct References {
    #[serde(default)]
    cases: Vec<Case>,
}
#[derive(Deserialize)]
struct Case {
    context: String,
    baseline: String,
    #[serde(default)]
    routes: Vec<Route>,
}
#[derive(Default, Deserialize)]
struct Route {
    engine: String,
    status: String,
    #[serde(default)]
    files: usize,
    #[serde(default)]
    verified_outputs: usize,
    total_e2e_ms: Option<f64>,
}

fn data() -> &'static References {
    static DATA: OnceLock<References> = OnceLock::new();
    DATA.get_or_init(|| {
        serde_json::from_str(include_str!("engine_reference_data.json")).unwrap_or_default()
    })
}

fn valid_time(route: &Route) -> Option<f64> {
    route.total_e2e_ms.filter(|time| {
        route.status == "PASS"
            && route.files == 50
            && route.verified_outputs == 50
            && time.is_finite()
            && *time > 0.0
    })
}

fn lookup(references: &References, context: &str, engine: &str) -> String {
    let Some(case) = references.cases.iter().find(|c| c.context == context) else {
        return "待测".into();
    };
    let Some(route) = case.routes.iter().find(|r| r.engine == engine) else {
        return "待测".into();
    };
    let Some(time) = valid_time(route) else {
        return if matches!(route.status.as_str(), "NOT_TESTED" | "NOT_APPLICABLE") {
            "待测".into()
        } else {
            "样本未通过".into()
        };
    };
    let Some(base) = case
        .routes
        .iter()
        .find(|r| r.engine == case.baseline)
        .and_then(valid_time)
    else {
        return "基准未完成".into();
    };
    if engine == case.baseline {
        "1.00×，基准".into()
    } else {
        format!("约 {:.2}×", time / base)
    }
}

pub fn hint(context: &str, engine: &str) -> String {
    if !cfg!(windows) {
        return "当前平台未测".into();
    }
    lookup(data(), context, engine)
}

pub fn conversion_hint(operation: &str, engine: &str) -> String {
    match operation {
        "files.office_conversion.excel" => format!(
            "XLSX {} / XLS {}",
            hint("4-4-xlsx", engine),
            hint("4-4-xls", engine)
        ),
        "files.office_conversion.word" => format!("DOC {}；DOCX 待测", hint("4-4-word", engine)),
        _ => hint("4-7", engine),
    }
}

pub fn current_policy_hint(context: &str) -> Option<String> {
    use crate::excel_edit_router::{RouteMode, ROUTE_ENV};
    let mode: RouteMode = std::env::var(ROUTE_ENV).ok()?.parse().ok()?;
    let engine = match mode {
        RouteMode::OpenpyxlOnly => "openpyxl",
        RouteMode::UmyaOnly => "umya",
        RouteMode::BetterOfficeOnly => "betteroffice_modify",
        RouteMode::RustNativeOnly => "rustnative",
        RouteMode::AsposeOnly => "aspose",
        RouteMode::Auto => "route_auto",
    };
    Some(format!(
        "{}；参考耗时 {}",
        mode.as_str(),
        hint(context, engine)
    ))
}

pub fn write_note(output: &mut impl Write, context: &str) -> std::io::Result<()> {
    let sample = if matches!(context, "5-7" | "7-1") {
        "50 次独立任务"
    } else {
        "50 文件批次"
    };
    let detail = if context == "3-1" {
        "，打印 mode 1"
    } else {
        ""
    };
    let format = match context {
        "4-4-excel" => "XLS/XLSX",
        "4-4-word" => "DOC",
        _ => "XLSX",
    };
    writeln!(
        output,
        "参考：Windows Release，{format} 小型合成样本，{sample}单轮初测{detail}；耗时比越小越快。"
    )?;
    writeln!(
        output,
        "按完整处理总耗时比较；仅供参考，实际耗时随文件内容变化。"
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn embedded_catalog_has_every_picker_and_only_completed_50_file_ratios() {
        let references: References =
            serde_json::from_str(include_str!("engine_reference_data.json")).unwrap();
        let expected = [
            "1-1", "1-5", "2-1", "2-2", "2-3", "3-1", "4-4-xlsx", "4-4-xls", "4-4-word", "4-5",
            "4-7", "5-1", "5-2", "5-3", "5-4", "5-6", "5-7", "7-1",
        ];
        assert_eq!(references.cases.len(), expected.len());
        for context in expected {
            let case = references
                .cases
                .iter()
                .find(|c| c.context == context)
                .unwrap();
            assert!(case
                .routes
                .iter()
                .find(|r| r.engine == case.baseline)
                .and_then(valid_time)
                .is_some());
            let mut engines = std::collections::BTreeSet::new();
            for route in &case.routes {
                assert!(engines.insert(&route.engine));
                if route.status == "PASS" {
                    assert!(valid_time(route).is_some());
                } else {
                    assert!(!lookup(&references, context, &route.engine).contains('×'));
                }
            }
        }
    }

    #[test]
    fn refuses_failed_partial_zero_and_unrelated_references() {
        let references: References = serde_json::from_value(serde_json::json!({"cases":[{
            "context":"3-1", "baseline":"openpyxl", "routes":[
                {"engine":"openpyxl","status":"PASS","files":50,"verified_outputs":50,"total_e2e_ms":1000},
                {"engine":"rust","status":"PASS","files":50,"verified_outputs":50,"total_e2e_ms":500},
                {"engine":"aspose","status":"FAILED","files":50,"verified_outputs":50,"total_e2e_ms":10},
                {"engine":"partial","status":"PASS","files":50,"verified_outputs":49,"total_e2e_ms":10},
                {"engine":"zero","status":"PASS","files":50,"verified_outputs":50,"total_e2e_ms":0}
            ]}]})).unwrap();
        assert_eq!(lookup(&references, "3-1", "rust"), "约 0.50×");
        for engine in ["aspose", "partial", "zero"] {
            assert!(!lookup(&references, "3-1", engine).contains('×'));
        }
        assert_eq!(lookup(&references, "5-4", "rust"), "待测");
    }

    #[test]
    fn missing_or_invalid_baseline_never_produces_a_ratio() {
        let references: References = serde_json::from_value(serde_json::json!({"cases":[{
        "context":"4-7", "baseline":"auto", "routes":[
            {"engine":"auto","status":"FAILED","total_e2e_ms":1},
            {"engine":"aspose","status":"PASS","files":50,"verified_outputs":50,"total_e2e_ms":100}
        ]}]}))
        .unwrap();
        assert_eq!(lookup(&references, "4-7", "aspose"), "基准未完成");
    }
}
