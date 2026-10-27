use pytools_core::{
    config::ExtractTask,
    sheet_extract::{
        compact, extract_configured_values, find_matched_sheets, find_matched_sheets_for_task,
        name_matches_exclude, name_matches_keyword_options, normalize_range_spec,
        output_path_for_task, parse_area_indexes, parse_indexes, sanitize_sheet_name,
        split_keyword_groups, stitch_area_values, token_to_index, unique_sheet_name,
        workbook_matches, CellMatrix,
    },
};
use serde_json::{json, Value};
use std::path::Path;

fn matrix(value: &Value) -> CellMatrix {
    value
        .as_array()
        .unwrap()
        .iter()
        .map(|row| {
            row.as_array()
                .unwrap()
                .iter()
                .map(|cell| (!cell.is_null()).then(|| cell.clone()))
                .collect()
        })
        .collect()
}

fn strings(value: &Value) -> Vec<String> {
    value
        .as_array()
        .unwrap()
        .iter()
        .map(|item| item.as_str().unwrap().to_owned())
        .collect()
}

fn extract_task(value: &Value) -> ExtractTask {
    ExtractTask {
        row_no: 0,
        workbook_keywords: value["workbook_keywords"]
            .as_str()
            .unwrap_or_default()
            .to_owned(),
        sheet_keywords: value["sheet_keywords"]
            .as_str()
            .unwrap_or_default()
            .to_owned(),
        exclude_keywords: value["exclude_keywords"]
            .as_str()
            .unwrap_or_default()
            .to_owned(),
        area_spec: value["area_spec"].as_str().unwrap_or_default().to_owned(),
        stitch_mode: value["stitch_mode"].as_str().unwrap_or_default().to_owned(),
        output_folder: value["output_folder"]
            .as_str()
            .unwrap_or_default()
            .to_owned(),
        output_name: value["output_name"].as_str().unwrap_or_default().to_owned(),
        target_sheet: value["target_sheet"]
            .as_str()
            .unwrap_or_default()
            .to_owned(),
        remark: String::new(),
    }
}

#[test]
fn extraction_helpers_match_python_baseline_vectors() {
    let cases: Value =
        serde_json::from_str(include_str!("../tests/fixtures/sheet_extract_cases.json"))
            .expect("valid sheet extraction cases");
    let baseline: Value = serde_json::from_str(include_str!(
        "../probe/reports/sheet_extract_python_baseline.json"
    ))
    .expect("valid Python extraction helper baseline");

    let range_helpers = cases["range_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(|case| match case["op"].as_str().unwrap() {
            "normalize" => json!(normalize_range_spec(case["value"].as_str().unwrap())),
            "token" => json!(token_to_index(
                case["value"].as_str().unwrap(),
                case["is_row"].as_bool().unwrap()
            )),
            "parse" => json!(parse_indexes(
                case["value"].as_str().unwrap(),
                case["max"].as_u64().unwrap() as usize,
                case["is_row"].as_bool().unwrap()
            )),
            op => panic!("unknown range operation {op}"),
        })
        .collect::<Vec<_>>();

    let keyword_helpers = cases["keyword_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(|case| match case["op"].as_str().unwrap() {
            "split" => json!(split_keyword_groups(case["value"].as_str().unwrap())),
            "match" => json!(name_matches_keyword_options(
                case["name"].as_str().unwrap(),
                case["value"].as_str().unwrap()
            )),
            "exclude" => json!(name_matches_exclude(
                case["name"].as_str().unwrap(),
                case["value"].as_str().unwrap()
            )),
            op => panic!("unknown keyword operation {op}"),
        })
        .collect::<Vec<_>>();

    let sheet_names = strings(&cases["sheet_selection"]["sheet_names"]);
    let keywords = strings(&cases["sheet_selection"]["keywords"]);
    let exact_selection = find_matched_sheets(&sheet_names, &keywords)
        .into_iter()
        .map(|index| sheet_names[index].clone())
        .collect::<Vec<_>>();

    let task_case = &cases["task_sheet_selection"];
    let task_sheet_names = strings(&task_case["sheet_names"]);
    let task_selection = find_matched_sheets_for_task(
        &task_sheet_names,
        task_case["sheet_keywords"].as_str().unwrap(),
        task_case["exclude_keywords"].as_str().unwrap(),
    )
    .into_iter()
    .map(|index| task_sheet_names[index].clone())
    .collect::<Vec<_>>();

    let workbook_matches = cases["workbook_matches"]
        .as_array()
        .unwrap()
        .iter()
        .map(|case| {
            let task = extract_task(&json!({"workbook_keywords": case["keywords"]}));
            workbook_matches(case["name"].as_str().unwrap(), &task)
        })
        .collect::<Vec<_>>();

    let area_parsing = cases["area_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(
            |area| match parse_area_indexes(area.as_str().unwrap(), 3, 3) {
                Some((rows, columns)) => json!([rows, columns]),
                None => Value::Null,
            },
        )
        .collect::<Vec<_>>();

    let data = matrix(&cases["matrix"]);
    let configured_values = cases["configured_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(|case| {
            extract_configured_values(
                &data,
                case["area_spec"].as_str().unwrap(),
                case["stitch_mode"].as_str().unwrap(),
            )
        })
        .collect::<Vec<_>>();
    let stitching = cases["stitch_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(|case| {
            let areas = case["areas"]
                .as_array()
                .unwrap()
                .iter()
                .map(matrix)
                .collect::<Vec<_>>();
            stitch_area_values(&areas, case["mode"].as_str().unwrap())
        })
        .collect::<Vec<_>>();
    let compaction = cases["compact_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(|value| compact(&matrix(value)))
        .collect::<Vec<_>>();

    let names_case = &cases["sheet_name_cases"];
    let sanitized = names_case["sanitize"]
        .as_array()
        .unwrap()
        .iter()
        .map(|name| sanitize_sheet_name(name.as_str().unwrap()))
        .collect::<Vec<_>>();
    let unique = names_case["unique"]
        .as_array()
        .unwrap()
        .iter()
        .map(|case| {
            unique_sheet_name(
                &strings(&case["used"]),
                case["base"].as_str().unwrap(),
                "123456",
            )
        })
        .collect::<Vec<_>>();

    let output_paths = cases["output_path_cases"]
        .as_array()
        .unwrap()
        .iter()
        .map(|case| {
            let task = extract_task(&case["task"]);
            output_path_for_task(
                &task,
                Path::new(case["output_dir"].as_str().unwrap()),
                case["default_name"].as_str().unwrap(),
            )
            .to_string_lossy()
            .replace('\\', "/")
        })
        .collect::<Vec<_>>();

    let actual = json!({
        "range_helpers": range_helpers,
        "keyword_helpers": keyword_helpers,
        "exact_sheet_selection": exact_selection,
        "task_sheet_selection": task_selection,
        "workbook_matches": workbook_matches,
        "area_parsing": area_parsing,
        "configured_values": configured_values,
        "stitching": stitching,
        "compaction": compaction,
        "sheet_names": {"sanitize": sanitized, "unique": unique},
        "output_paths": output_paths,
    });
    assert_eq!(baseline["python_baseline"]["status"], "保留");
    assert_eq!(actual, baseline["result"]);
}

#[test]
fn unique_sheet_names_avoid_case_insensitive_excel_collisions() {
    assert_eq!(
        unique_sheet_name(&["Report".to_owned()], "report", "123456"),
        "report_2"
    );
}
