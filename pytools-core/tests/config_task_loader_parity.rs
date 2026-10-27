use pytools_core::config::{
    load_dedup_tasks, load_extract_tasks, load_foreign_banks, load_institution_mapping,
    load_print_tasks, DedupTask, ExtractTask, PrintTask,
};
use serde_json::{json, Value};
use std::path::{Path, PathBuf};

fn normalized_task_path(path: &Option<PathBuf>, base: &Path) -> Option<String> {
    path.as_ref().map(|path| {
        let path = without_windows_extended_prefix(path);
        let base = without_windows_extended_prefix(base);
        let relative = path
            .strip_prefix(&base)
            .unwrap_or_else(|_| panic!("{} is not under {}", path.display(), base.display()));
        format!("$BASE/{}", relative.to_string_lossy().replace('\\', "/"))
    })
}

fn without_windows_extended_prefix(path: &Path) -> PathBuf {
    let path_text = path.to_string_lossy();
    if let Some(rest) = path_text.strip_prefix("\\\\?\\UNC\\") {
        PathBuf::from(format!("\\\\{rest}"))
    } else if let Some(rest) = path_text.strip_prefix("\\\\?\\") {
        PathBuf::from(rest)
    } else {
        path.to_path_buf()
    }
}

fn dedup_json(tasks: &[DedupTask], base: &Path) -> Value {
    Value::Array(
        tasks
            .iter()
            .map(|task| {
                json!({
                    "row_no": task.row_no,
                    "enabled": task.enabled,
                    "task_name": task.task_name,
                    "source_wb": normalized_task_path(&task.source_wb, base),
                    "source_ws": task.source_ws,
                    "key_cols": task.key_cols,
                    "target_wb": normalized_task_path(&task.target_wb, base),
                    "target_ws": task.target_ws,
                    "write_mode": task.write_mode,
                })
            })
            .collect(),
    )
}

fn print_json(tasks: &[PrintTask], base: &Path) -> Value {
    Value::Array(
        tasks
            .iter()
            .map(|task| {
                json!({
                    "row_no": task.row_no,
                    "enabled": task.enabled,
                    "mode": task.mode,
                    "source_wb": normalized_task_path(&task.source_wb, base),
                    "source_ws": task.source_ws,
                    "source_ranges": task.source_ranges,
                    "target_wb": normalized_task_path(&task.target_wb, base),
                    "target_ws": task.target_ws,
                    "fit_wide": task.fit_wide,
                    "fit_tall": task.fit_tall,
                    "orientation": task.orientation,
                    "center_h": task.center_h,
                    "center_v": task.center_v,
                    "drop_comments": task.drop_comments,
                    "hide_zero": task.hide_zero,
                    "remark": task.remark,
                })
            })
            .collect(),
    )
}

fn extract_json(tasks: &[ExtractTask]) -> Value {
    Value::Array(
        tasks
            .iter()
            .map(|task| {
                let mut value = json!({
                    "row_no": task.row_no,
                    "workbook_keywords": task.workbook_keywords,
                    "sheet_keywords": task.sheet_keywords,
                    "exclude_keywords": task.exclude_keywords,
                    "area_spec": task.area_spec,
                    "stitch_mode": task.stitch_mode,
                    "output_folder": task.output_folder,
                    "output_name": task.output_name,
                    "remark": task.remark,
                });
                if !task.target_sheet.is_empty() {
                    value["target_sheet"] = json!(task.target_sheet);
                }
                value
            })
            .collect(),
    )
}

#[test]
fn task_config_loaders_match_python_baseline() {
    let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let base = fixture_dir.parent().unwrap();
    let baseline: Value = serde_json::from_str(include_str!(
        "../probe/reports/config_task_loaders_python_baseline.json"
    ))
    .expect("valid Python config task loader baseline");
    let main = fixture_dir.join("config_task_loaders.xlsx");
    let legacy = fixture_dir.join("config_task_loaders_legacy.xlsx");
    let missing = fixture_dir.join("config_task_loaders_missing.xlsx");

    let actual = json!({
        "dedup_tasks": dedup_json(&load_dedup_tasks(&main).unwrap(), base),
        "print_tasks": print_json(&load_print_tasks(&main).unwrap(), base),
        "extract_tasks": extract_json(&load_extract_tasks(&main).unwrap()),
        "foreign_banks": load_foreign_banks(&main).unwrap(),
        "institution_mapping": load_institution_mapping(&main).unwrap(),
        "legacy_dedup_tasks": dedup_json(&load_dedup_tasks(&legacy).unwrap(), base),
        "missing_sheets": {
            "dedup_tasks": dedup_json(&load_dedup_tasks(&missing).unwrap(), base),
            "print_tasks": print_json(&load_print_tasks(&missing).unwrap(), base),
            "extract_tasks": extract_json(&load_extract_tasks(&missing).unwrap()),
            "foreign_banks": load_foreign_banks(&missing).unwrap(),
            "institution_mapping": load_institution_mapping(&missing).unwrap(),
        }
    });

    assert_eq!(baseline["python_baseline"]["status"], "保留");
    assert_eq!(actual, baseline["result"]);
}
