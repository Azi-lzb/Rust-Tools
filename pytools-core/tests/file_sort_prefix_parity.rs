use pytools_core::file_sort_prefix::{
    build_file_sort_rules, match_file_sort_rule, normalize_sort_prefix, prefixed_file_name,
    sort_prefix_digits, sort_prefix_width, strip_sort_prefix, FileSortConfigRow,
};
use serde_json::{json, Value};

fn fixture() -> Value {
    serde_json::from_str(include_str!("fixtures/file_sort_prefix_cases.json"))
        .expect("valid file sort prefix cases")
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/file_sort_prefix_python_baseline.json"
    ))
    .expect("valid Python helper baseline")
}

#[test]
fn filename_sort_prefix_helpers_match_python_baseline() {
    let cases = fixture();
    let baseline = baseline();

    let digits = cases["digit_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(|value| json!(sort_prefix_digits(value.as_str().unwrap()).unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(json!(digits), baseline["sort_prefix_digits"]);

    let widths = cases["width_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(|value| json!(sort_prefix_width(value.as_str().unwrap())))
        .collect::<Vec<_>>();
    assert_eq!(json!(widths), baseline["sort_prefix_width"]);

    let normalized = cases["normalize_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(|case| {
            json!(normalize_sort_prefix(
                case["value"].as_str().unwrap(),
                case["width"].as_u64().unwrap() as usize
            )
            .unwrap())
        })
        .collect::<Vec<_>>();
    assert_eq!(json!(normalized), baseline["normalize_sort_prefix"]);

    let stripped = cases["strip_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(|value| json!(strip_sort_prefix(value.as_str().unwrap())))
        .collect::<Vec<_>>();
    assert_eq!(json!(stripped), baseline["strip_sort_prefix"]);

    let rows = cases["config_rows"]
        .as_array()
        .unwrap()
        .iter()
        .map(|row| FileSortConfigRow {
            row_no: row["row_no"].as_u64().unwrap() as usize,
            enabled: row["enabled"].as_str().unwrap().to_owned(),
            sort_prefix: row["sort_prefix"].as_str().unwrap().to_owned(),
            keywords: row["keywords"].as_str().unwrap().to_owned(),
            remark: row["remark"].as_str().unwrap().to_owned(),
        })
        .collect::<Vec<_>>();
    let rules = build_file_sort_rules(&rows).unwrap();
    assert_eq!(serde_json::to_value(&rules).unwrap(), baseline["rules"]);

    let matches = cases["match_names"]
        .as_array()
        .unwrap()
        .iter()
        .map(|name| {
            match_file_sort_rule(name.as_str().unwrap(), &rules)
                .map(|rule| json!({"row_no": rule.row_no, "prefix": rule.prefix}))
                .unwrap_or(Value::Null)
        })
        .collect::<Vec<_>>();
    assert_eq!(json!(matches), baseline["matches"]);

    let prefixed_names = cases["prefix_names"]
        .as_array()
        .unwrap()
        .iter()
        .map(|case| {
            json!(prefixed_file_name(
                case["file_name"].as_str().unwrap(),
                case["prefix"].as_str().unwrap()
            ))
        })
        .collect::<Vec<_>>();
    assert_eq!(json!(prefixed_names), baseline["prefixed_names"]);
}
