//! Display-only, correctness-gated references. Never selects an engine.
use serde::Deserialize;
use std::{io::Write, sync::OnceLock};

#[derive(Default, Deserialize)]
struct References {
    #[serde(default)]
    reference_date: String,
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
            && route.files > 0
            && route.verified_outputs == route.files
            && time.is_finite()
            && *time > 0.0
    })
}

fn baseline_name(case: &Case) -> &str {
    match case.baseline.as_str() {
        "openpyxl" => "Openpyxl",
        "com" => "Excel/WPS COM（本次测 Excel）",
        "auto" => match case.context.as_str() {
            "5-4" => "原 OOXML 页眉路径",
            "4-7" => "原 OOXML 修复路径",
            "4-4-xlsx" | "4-4-xls" | "4-4-word" => "原 Excel/WPS COM 路径",
            _ => "原自动路径",
        },
        other => other,
    }
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
        .find(|r| r.engine == case.baseline && r.files == route.files)
        .and_then(valid_time)
    else {
        return "基准未完成".into();
    };
    if engine == case.baseline {
        "1.00×，基准".into()
    } else {
        format!("约 {:.2}×，相对{}", time / base, baseline_name(case))
    }
}

pub fn hint(context: &str, engine: &str) -> String {
    if !cfg!(windows) {
        return "当前平台未测".into();
    }
    lookup(data(), context, engine)
}

pub fn measured_route(context: &str, engine: &str) -> bool {
    cfg!(windows)
        && data()
            .cases
            .iter()
            .find(|case| case.context == context)
            .and_then(|case| case.routes.iter().find(|route| route.engine == engine))
            .and_then(valid_time)
            .is_some()
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
    let configured = std::env::var(ROUTE_ENV).ok()?;
    let Ok(mode) = configured.parse::<RouteMode>() else {
        return Some("编辑策略配置无效；执行时拒绝，不自动改用默认引擎".into());
    };
    let engine = match mode {
        RouteMode::OpenpyxlOnly => "openpyxl",
        RouteMode::UmyaOnly => "umya",
        RouteMode::BetterOfficeOnly => "betteroffice_modify",
        RouteMode::RustNativeOnly => "rustnative",
        RouteMode::AsposeOnly => "aspose",
        RouteMode::Auto => "route_auto",
    };
    Some(format!(
        "{}：{}；参考耗时 {}",
        mode.as_str(),
        modification_policy_description(mode),
        hint(context, engine)
    ))
}

pub fn modification_policy_description(mode: crate::excel_edit_router::RouteMode) -> &'static str {
    use crate::excel_edit_router::RouteMode;
    match mode {
        RouteMode::OpenpyxlOnly => "Openpyxl，不切换其他引擎",
        RouteMode::UmyaOnly => "UMYA，不切换其他引擎",
        RouteMode::BetterOfficeOnly => "BetterOffice，不切换其他引擎",
        RouteMode::RustNativeOnly => "RustNative OOXML，不切换其他引擎",
        RouteMode::AsposeOnly => "Aspose .NET，保留评估水印，不切换其他引擎",
        RouteMode::Auto => {
            "按修改计划选择 RustNative OOXML、UMYA 或 Openpyxl；满足安全条件时才回退 Openpyxl"
        }
    }
}

/// Display the existing per-feature path, without selecting or probing an engine.
fn xlsx_modification_hint(context: &str) -> &'static str {
    match context {
        "1-1" => "Calamine 读取 XLS/XLSX 源；Rust 新建 XLSX；Openpyxl 修改既有 XLSX 目标",
        "2-3" => "新建 XLSX：Rust；追加既有 XLSX：Openpyxl；Windows 既有 XLS：Excel/WPS COM",
        "5-4" => "XLSX：OOXML 定点页眉修改；就地修改允许安全回退 Openpyxl",
        "5-6" => "新建 XLSX：Rust；追加既有目标：Openpyxl",
        _ => "XLSX 修改：Openpyxl",
    }
}

pub fn selected_xlsx_modification_hint(context: &str) -> String {
    current_policy_hint(context).map_or_else(
        || xlsx_modification_hint(context).to_owned(),
        |policy| format!("按已配置的 XLSX 策略处理（{policy}）"),
    )
}

pub fn default_modification_hint(context: &str) -> String {
    let xlsx = selected_xlsx_modification_hint(context);
    let legacy = if matches!(context, "2-1" | "2-2" | "4-5" | "5-1" | "5-2" | "5-3" | "5-4" | "5-7") {
        if cfg!(windows) {
            "；XLS：功能专用 Excel/WPS COM"
        } else {
            "；XLS 修改不支持"
        }
    } else {
        ""
    };
    format!("{xlsx}{legacy}")
}

pub fn default_conversion_hint(operation: &str) -> &'static str {
    match (operation, cfg!(windows), cfg!(target_os = "linux")) {
        ("files.excel_open_repair", true, _) => "XLSX：OOXML 修复；XLS：Excel/WPS COM 转换后修复",
        ("files.excel_open_repair", _, true) => {
            "XLSX/XLSM：OOXML 修复；XLS 先用 4-4 的 Aspose .NET 转换"
        }
        ("files.excel_open_repair", _, _) => "XLSX：OOXML 修复；XLS 转换不可用",
        (_, true, _) => "Excel/WPS COM（优先 Excel，按原规则使用 WPS）",
        ("files.office_conversion.excel", _, true) => "Aspose .NET（仅 XLS → XLSX）",
        _ => "当前平台无可用的原 Office 转换路径",
    }
}

pub fn write_note(output: &mut impl Write, context: &str) -> std::io::Result<()> {
    let references = data();
    let key = match context {
        "4-4-excel" => "4-4-xlsx",
        "4-4-word" => "4-4-word",
        _ => context,
    };
    let count = references
        .cases
        .iter()
        .find(|case| case.context == key)
        .and_then(|case| {
            case.routes
                .iter()
                .find(|route| route.engine == case.baseline && valid_time(route).is_some())
        })
        .map(|route| route.files);
    let Some(count) = count else {
        return writeln!(output, "当前功能没有通过校验的可比耗时基准。");
    };
    let sample = if matches!(context, "5-7" | "7-1") {
        format!("{count} 次独立任务")
    } else {
        format!("{count} 文件批次")
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
        "实测参考（{}）：Windows Release，{format} 小型合成样本，{sample}单轮{detail}。",
        references.reference_date
    )?;
    writeln!(
        output,
        "耗时倍数 = 该选项总耗时 ÷ 基准总耗时（含启动，不含人工操作）；0.50× 用时一半，2.00× 用时两倍。仅供参考。"
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn embedded_catalog_has_every_picker_and_only_completed_ratios() {
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
            assert!(case.routes.iter().any(|r| r.engine == case.baseline));
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
                {"engine":"different_batch","status":"PASS","files":5,"verified_outputs":5,"total_e2e_ms":10},
                {"engine":"zero","status":"PASS","files":50,"verified_outputs":50,"total_e2e_ms":0}
            ]}]})).unwrap();
        assert_eq!(lookup(&references, "3-1", "rust"), "约 0.50×，相对Openpyxl");
        for engine in ["aspose", "partial", "different_batch", "zero"] {
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
