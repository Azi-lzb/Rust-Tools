use serde_json::Value;
use std::{fs, process::Command};

#[test]
fn structural_edit_keeps_supported_features_and_formula_references_are_updated() {
    let root = std::env::temp_dir().join(format!("betteroffice-audit-{}", std::process::id()));
    fs::create_dir_all(&root).unwrap();
    let run = Command::new(env!("CARGO_BIN_EXE_betteroffice-modification-audit"))
        .arg("audit")
        .arg(&root)
        .output()
        .unwrap();
    assert!(
        run.status.success(),
        "{}",
        String::from_utf8_lossy(&run.stderr)
    );
    let report: Value =
        serde_json::from_slice(&fs::read(root.join("audit.json")).unwrap()).unwrap();
    let structure = &report["structure"];
    assert_eq!(structure["chart_count"], 1);
    assert_eq!(structure["conditional_format_count"], 1);
    assert_eq!(structure["merge_count"], 1);
    assert_eq!(structure["a1_has_style"], true);
    assert_eq!(structure["column_width_e"], 21.0);
    assert_eq!(structure["row_height_2"], 27.0);
    assert!(structure["defined_names"]
        .as_array()
        .unwrap()
        .iter()
        .any(|n| n["name"] == "DataRange" && n["formula"] == "Data!$A$1:$B$4"));
    assert_eq!(structure["custom_xml_preserved"], true);
    assert!(report["print_area_probe"]["defined_names"]
        .as_array()
        .unwrap()
        .iter()
        .any(|n| n["name"] == "_xlnm.Print_Area"));
    let formulas = &report["formula"];
    assert_eq!(formulas["deleted_reference_formula"]["formula"], "A1+#REF!");
    assert_eq!(
        formulas["deleted_reference_formula"]["cached_value"]["value"],
        "#REF!"
    );
    assert_eq!(
        formulas["a1_dependent_formula_after_delete_b"]["formula"],
        "A1*2"
    );
    assert_eq!(
        formulas["a1_dependent_formula_after_delete_b"]["cached_value"]["value"],
        30.0
    );
}
