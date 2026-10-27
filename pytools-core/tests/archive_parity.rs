use pytools_core::archive_ops::{
    archive_by_date_copy, archive_by_date_preview, archive_by_type_copy, archive_by_type_preview,
    execute_archive_plan, ArchiveResult,
};
use serde_json::{json, Value};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};
use std::{fs, io::Read};

fn cases() -> Value {
    serde_json::from_str(include_str!("fixtures/archive_cases.json")).unwrap()
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/archive_python_baseline.json"
    ))
    .unwrap()
}

fn read_xlsx_part(path: &Path, part_name: &str) -> String {
    let file = fs::File::open(path).unwrap();
    let mut archive = zip::ZipArchive::new(file).unwrap();
    let mut part = archive.by_name(part_name).unwrap();
    let mut xml = String::new();
    part.read_to_string(&mut xml).unwrap();
    xml
}

fn spreadsheet_element(node: roxmltree::Node<'_, '_>, local_name: &str) -> bool {
    node.tag_name().namespace() == Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        && node.tag_name().name() == local_name
}

fn assert_python_archive_report_format(path: &Path) {
    let worksheet = read_xlsx_part(path, "xl/worksheets/sheet1.xml");
    let worksheet = roxmltree::Document::parse(&worksheet).unwrap();
    let header = worksheet
        .descendants()
        .find(|node| spreadsheet_element(*node, "c") && node.attribute("r") == Some("A1"))
        .expect("archive report header cell A1");
    let style_id = header
        .attribute("s")
        .unwrap_or("0")
        .parse::<usize>()
        .unwrap();

    let styles = read_xlsx_part(path, "xl/styles.xml");
    let styles = roxmltree::Document::parse(&styles).unwrap();
    let cell_xfs = styles
        .descendants()
        .find(|node| spreadsheet_element(*node, "cellXfs"))
        .expect("cellXfs collection");
    let formats = cell_xfs
        .children()
        .filter(|node| spreadsheet_element(*node, "xf"))
        .collect::<Vec<_>>();
    let header_format = formats[style_id];
    let alignment = header_format
        .children()
        .find(|node| spreadsheet_element(*node, "alignment"))
        .expect("archive report header alignment");
    assert_eq!(alignment.attribute("horizontal"), Some("center"));
    assert_eq!(alignment.attribute("vertical"), Some("top"));
    let border_id = header_format
        .attribute("borderId")
        .unwrap_or("0")
        .parse::<usize>()
        .unwrap();
    let borders = styles
        .descendants()
        .find(|node| spreadsheet_element(*node, "borders"))
        .expect("borders collection");
    let borders = borders
        .children()
        .filter(|node| spreadsheet_element(*node, "border"))
        .collect::<Vec<_>>();
    let border = borders[border_id];
    for side_name in ["left", "right", "top", "bottom"] {
        let side = border
            .children()
            .find(|node| spreadsheet_element(*node, side_name))
            .expect("header border side");
        assert_eq!(side.attribute("style"), Some("thin"));
        if let Some(color) = side
            .children()
            .find(|node| spreadsheet_element(*node, "color"))
        {
            let automatic_black = color.attribute("auto") == Some("1");
            let theme_black = color.attribute("theme") == Some("1")
                && color
                    .attribute("tint")
                    .map(|tint| tint.parse::<f64>().is_ok_and(|value| value == 0.0))
                    .unwrap_or(true);
            assert!(
                automatic_black || theme_black,
                "archive report {side_name} border should use Excel's default black"
            );
        }
    }
    let font_id = formats[style_id]
        .attribute("fontId")
        .unwrap_or("0")
        .parse::<usize>()
        .unwrap();
    let fonts = styles
        .descendants()
        .find(|node| spreadsheet_element(*node, "fonts"))
        .expect("fonts collection");
    let fonts = fonts
        .children()
        .filter(|node| spreadsheet_element(*node, "font"))
        .collect::<Vec<_>>();
    assert!(
        fonts[font_id]
            .children()
            .any(|node| { spreadsheet_element(node, "b") && node.attribute("val") != Some("0") }),
        "archive report headers should be bold like the Python output"
    );

    let margins = worksheet
        .descendants()
        .find(|node| spreadsheet_element(*node, "pageMargins"))
        .expect("pageMargins element");
    for (name, expected) in [
        ("left", 0.75),
        ("right", 0.75),
        ("top", 1.0),
        ("bottom", 1.0),
        ("header", 0.5),
        ("footer", 0.5),
    ] {
        let actual = margins.attribute(name).unwrap().parse::<f64>().unwrap();
        assert!(
            (actual - expected).abs() < 1e-9,
            "archive report {name} margin should match Python output"
        );
    }
}

struct TempDir(PathBuf);

impl TempDir {
    fn new(label: &str) -> Self {
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "pytools-archive-{label}-{}-{unique}",
            std::process::id()
        ));
        fs::create_dir_all(&path).unwrap();
        Self(path)
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn setup(root: &Path, cases: &Value, existing_target: Option<&str>) -> (PathBuf, Vec<PathBuf>) {
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).unwrap();
    for name in cases["input_names"].as_array().unwrap() {
        let relative = name.as_str().unwrap();
        let path = input_dir.join(relative);
        fs::create_dir_all(path.parent().unwrap()).unwrap();
        let bytes = if path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .starts_with("~$")
        {
            b"temporary".to_vec()
        } else {
            format!("fixture:{relative}").into_bytes()
        };
        fs::write(path, bytes).unwrap();
    }
    let missing = input_dir.join(cases["missing_name"].as_str().unwrap());
    if let Some(relative) = existing_target {
        let target = root.join("target").join(relative);
        fs::create_dir_all(target.parent().unwrap()).unwrap();
        fs::write(target, b"occupied target").unwrap();
    }
    (input_dir.clone(), vec![input_dir, missing])
}

fn relative(path: &str, root: &Path) -> String {
    if path.is_empty() {
        return String::new();
    }
    Path::new(path)
        .strip_prefix(root)
        .unwrap_or_else(|error| panic!("{} is outside {}: {error}", path, root.display()))
        .to_string_lossy()
        .replace('\\', "/")
}

fn preview_rows(result: &ArchiveResult, root: &Path) -> Vec<Value> {
    result
        .rows
        .iter()
        .map(|row| {
            json!({
                "source": relative(&row.source_path, root),
                "file_name": row.file_name,
                "mode": row.mode,
                "category": row.category,
                "target_dir": relative(&row.target_dir, root),
                "target": relative(&row.target_path, root),
                "status": row.status,
                "note": row.note,
            })
        })
        .collect()
}

fn copy_rows(result: &ArchiveResult, root: &Path) -> Vec<Value> {
    result
        .rows
        .iter()
        .map(|row| {
            json!({
                "source": relative(&row.source_path, root),
                "planned_target": relative(&row.planned_target_path, root),
                "actual_target": if row.status == "已复制" { relative(&row.target_path, root) } else { String::new() },
                "status": row.status,
                "note": row.note,
            })
        })
        .collect()
}

fn assert_preview(name: &str, result: &ArchiveResult, root: &Path, baseline: &Value) {
    let expected = &baseline[name];
    let stats = &expected["stats"];
    assert_eq!(
        result.files,
        stats["files"].as_u64().unwrap() as usize,
        "{name} files"
    );
    assert_eq!(
        result.invalid,
        stats["invalid"].as_u64().unwrap() as usize,
        "{name} invalid"
    );
    assert_eq!(
        result.rows.len(),
        stats["planned"].as_u64().unwrap() as usize,
        "{name} planned"
    );
    let skipped_extensions = result
        .rows
        .iter()
        .filter(|row| {
            matches!(
                row.status.as_str(),
                "跳过：后缀被排除" | "跳过：后缀不在归档范围"
            )
        })
        .count();
    assert_eq!(
        skipped_extensions,
        stats["skipped_ext"].as_u64().unwrap() as usize,
        "{name} extension filter"
    );
    assert_eq!(
        preview_rows(result, root),
        expected["rows"].as_array().unwrap().clone(),
        "{name} rows"
    );
    assert!(result.output.is_file(), "{name} report must be written");
    assert_python_archive_report_format(&result.output);
}

fn assert_copy(name: &str, result: &ArchiveResult, root: &Path, baseline: &Value) {
    let expected = &baseline[name];
    let stats = &expected["stats"];
    assert_eq!(
        result.copied,
        stats["copied"].as_u64().unwrap() as usize,
        "{name} copied"
    );
    assert_eq!(
        result.skipped,
        stats["skip"].as_u64().unwrap() as usize,
        "{name} skipped"
    );
    assert_eq!(
        result.failed,
        stats["failed"].as_u64().unwrap() as usize,
        "{name} failed"
    );
    assert_eq!(
        result.rows.len(),
        stats["rows"].as_u64().unwrap() as usize,
        "{name} rows"
    );
    assert_eq!(
        copy_rows(result, root),
        expected["rows"].as_array().unwrap().clone(),
        "{name} report rows"
    );
    assert!(result.output.is_file(), "{name} report must be written");
    assert_python_archive_report_format(&result.output);
    for row in result.rows.iter().filter(|row| row.status == "已复制") {
        let source = Path::new(&row.source_path);
        let target = Path::new(&row.target_path);
        assert_eq!(
            fs::read(source).unwrap(),
            fs::read(target).unwrap(),
            "{name} copied bytes"
        );
    }
}

#[test]
fn archive_previews_and_copy_workflows_match_python_baseline() {
    let cases = cases();
    let baseline = baseline();
    let config =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/archive_type_config.xlsx");

    let root = TempDir::new("date-preview");
    let (_input, paths) = setup(&root.0, &cases, cases["existing_targets"]["date"].as_str());
    let result =
        archive_by_date_preview(&paths, &root.0.join("target"), &root.0.join("out"), None).unwrap();
    assert_preview("date_preview", &result, &root.0, &baseline);

    let root = TempDir::new("date-filtered-preview");
    let (_input, paths) = setup(&root.0, &cases, None);
    let result = archive_by_date_preview(
        &paths,
        &root.0.join("target"),
        &root.0.join("out"),
        Some(&config),
    )
    .unwrap();
    assert_preview("date_filtered_preview", &result, &root.0, &baseline);

    let root = TempDir::new("type-preview");
    let (_input, paths) = setup(&root.0, &cases, cases["existing_targets"]["type"].as_str());
    let result =
        archive_by_type_preview(&paths, &root.0.join("target"), &root.0.join("out"), &config)
            .unwrap();
    assert_preview("type_preview", &result, &root.0, &baseline);

    let root = TempDir::new("date-copy");
    let (input, _paths) = setup(&root.0, &cases, cases["existing_targets"]["date"].as_str());
    let result =
        archive_by_date_copy(&[input], &root.0.join("target"), &root.0.join("out"), None).unwrap();
    assert_copy("date_copy", &result, &root.0, &baseline);

    let root = TempDir::new("type-copy");
    let (input, _paths) = setup(&root.0, &cases, cases["existing_targets"]["type"].as_str());
    let result = archive_by_type_copy(
        &[input],
        &root.0.join("target"),
        &root.0.join("out"),
        &config,
    )
    .unwrap();
    assert_copy("type_copy", &result, &root.0, &baseline);

    let root = TempDir::new("plan-copy");
    let (_input, paths) = setup(&root.0, &cases, cases["existing_targets"]["date"].as_str());
    let plan = archive_by_date_preview(&paths, &root.0.join("target"), &root.0.join("plan"), None)
        .unwrap();
    let result = execute_archive_plan(&plan.output, &root.0.join("out")).unwrap();
    assert_copy("plan_copy", &result, &root.0, &baseline);
}

#[test]
fn date_archive_cli_accepts_the_python_compatible_extension_filter_config() {
    let cases = cases();
    let config =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/archive_type_config.xlsx");
    let root = TempDir::new("date-preview-cli");
    let (_input, paths) = setup(&root.0, &cases, None);
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("archive-date-preview")
        .arg("--target-root")
        .arg(root.0.join("target"))
        .arg("--config")
        .arg(config)
        .arg("--output-dir")
        .arg(root.0.join("out"))
        .args(paths)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(report["files"], 12);
    assert_eq!(report["invalid"], 1);
    assert_eq!(report["rows"].as_array().unwrap().len(), 13);
    assert_eq!(
        report["rows"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|row| row["status"].as_str().unwrap().starts_with("跳过：后缀"))
            .count(),
        2
    );
    assert!(Path::new(report["output"].as_str().unwrap()).is_file());
}

#[test]
fn date_archive_preview_uses_python_windows_path_order_before_target_planning() {
    let root = TempDir::new("date-windows-path-order");
    let input = root.0.join("input");
    for (directory, file_name) in [
        ("a", "20261026_report.xlsx"),
        ("a", "20261026_report_1.xlsx"),
        ("B", "20261026_report.xlsx"),
    ] {
        let path = input.join(directory).join(file_name);
        fs::create_dir_all(path.parent().unwrap()).unwrap();
        fs::write(path, b"fixture").unwrap();
    }

    let result = archive_by_date_preview(
        std::slice::from_ref(&input),
        &root.0.join("target"),
        &root.0.join("out"),
        None,
    )
    .unwrap();

    let sources = result
        .rows
        .iter()
        .map(|row| relative(&row.source_path, &input))
        .collect::<Vec<_>>();
    let expected_sources = if cfg!(windows) {
        [
            "a/20261026_report.xlsx",
            "a/20261026_report_1.xlsx",
            "B/20261026_report.xlsx",
        ]
    } else {
        [
            "B/20261026_report.xlsx",
            "a/20261026_report.xlsx",
            "a/20261026_report_1.xlsx",
        ]
    };
    assert_eq!(sources, expected_sources);

    let targets = result
        .rows
        .iter()
        .map(|row| relative(&row.target_path, &root.0))
        .collect::<Vec<_>>();
    let expected_targets = if cfg!(windows) {
        [
            "target/202610/20261026_report.xlsx",
            "target/202610/20261026_report_1.xlsx",
            "target/202610/20261026_report_2.xlsx",
        ]
    } else {
        [
            "target/202610/20261026_report.xlsx",
            "target/202610/20261026_report_1.xlsx",
            "target/202610/20261026_report_1_1.xlsx",
        ]
    };
    assert_eq!(targets, expected_targets);
}

#[test]
fn date_archive_preview_uses_native_path_case_rules_for_target_collisions() {
    let root = TempDir::new("date-target-case-collision");
    let input = root.0.join("input");
    for (directory, file_name) in [("a", "20261026_report.xlsx"), ("b", "20261026_REPORT.xlsx")] {
        let path = input.join(directory).join(file_name);
        fs::create_dir_all(path.parent().unwrap()).unwrap();
        fs::write(path, b"fixture").unwrap();
    }

    let result = archive_by_date_preview(
        std::slice::from_ref(&input),
        &root.0.join("target"),
        &root.0.join("out"),
        None,
    )
    .unwrap();
    let targets = result
        .rows
        .iter()
        .map(|row| relative(&row.target_path, &root.0))
        .collect::<Vec<_>>();
    let expected = if cfg!(windows) {
        [
            "target/202610/20261026_report.xlsx",
            "target/202610/20261026_REPORT_1.xlsx",
        ]
    } else {
        [
            "target/202610/20261026_report.xlsx",
            "target/202610/20261026_REPORT.xlsx",
        ]
    };
    assert_eq!(targets, expected);
}

#[test]
fn date_archive_preview_preserves_dots_inside_the_file_stem_when_suffixing() {
    let root = TempDir::new("date-target-multidot-name");
    let input = root.0.join("input");
    for directory in ["a", "B"] {
        let path = input.join(directory).join("20261026_report.v2.xlsx");
        fs::create_dir_all(path.parent().unwrap()).unwrap();
        fs::write(path, b"fixture").unwrap();
    }

    let result = archive_by_date_preview(
        std::slice::from_ref(&input),
        &root.0.join("target"),
        &root.0.join("out"),
        None,
    )
    .unwrap();
    let targets = result
        .rows
        .iter()
        .map(|row| relative(&row.target_path, &root.0))
        .collect::<Vec<_>>();
    assert_eq!(
        targets,
        [
            "target/202610/20261026_report.v2.xlsx",
            "target/202610/20261026_report.v2_1.xlsx",
        ]
    );
}
