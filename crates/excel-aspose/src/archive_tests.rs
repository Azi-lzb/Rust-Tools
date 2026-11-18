use super::*;

#[test]
fn business_roster_is_exact_and_never_filters_names_or_blanks() {
    let expected = vec!["Data".to_owned(), "Blank".to_owned(), "Evaluation Warning".to_owned()];
    verify_business_roster(&expected, &expected).unwrap();
    assert!(verify_business_roster(&expected[..2], &expected).is_err());
    assert!(verify_business_roster(&expected, &expected[..2]).is_err());
    assert!(verify_business_roster(&[], &[]).is_err());
    let mut reordered = expected.clone();
    reordered.swap(0, 1);
    assert!(verify_business_roster(&reordered, &expected).is_err());
}

fn directory() -> PathBuf {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let root = std::env::temp_dir().join(format!(
        "aspose-archive-test-{}-{stamp}",
        std::process::id()
    ));
    fs::create_dir(&root).unwrap();
    root
}

fn evidence(root: &Path) -> Value {
    let folder = root.join("raw-sdk-save");
    fs::create_dir(&folder).unwrap();
    let path = folder.join("raw.xlsx");
    fs::write(&path, b"synthetic exact Save bytes").unwrap();
    let hash = hex_fingerprint(&path).unwrap();
    let manifest = path.with_extension("xlsx.json");
    fs::write(
        &manifest,
        serde_json::to_vec(&json!({
            "stage":"sdk_save_before_worker_preservation", "run_id":"run", "request_id":"task",
            "archive_sha256":hash
        }))
        .unwrap(),
    )
    .unwrap();
    json!({"path":path,"manifest_path":manifest,"sha256":hash,"stage":"sdk_save_before_worker_preservation"})
}

#[test]
fn absent_legacy_archive_is_unknown_not_fabricated() {
    assert!(
        verify_sdk_archive(&Value::Null, &directory(), "run", "task")
            .unwrap()
            .is_none()
    );
}

#[test]
fn matching_save_archive_and_manifest_are_accepted() {
    let root = directory();
    let value = evidence(&root);
    let accepted = verify_sdk_archive(&value, &root, "run", "task")
        .unwrap()
        .unwrap();
    assert_eq!(accepted.sha256, hex_fingerprint(&accepted.path).unwrap());
}

#[test]
fn changed_archive_is_rejected_without_modifying_evidence() {
    let root = directory();
    let value = evidence(&root);
    let path = Path::new(value["path"].as_str().unwrap());
    fs::write(path, b"changed after Save").unwrap();
    assert!(verify_sdk_archive(&value, &root, "run", "task").is_err());
    assert_eq!(fs::read(path).unwrap(), b"changed after Save");
}

#[test]
fn worker_preserved_stage_cannot_be_reported_as_raw_sdk_stage() {
    let root = directory();
    let mut value = evidence(&root);
    value["stage"] = json!("after_preservation");
    assert!(verify_sdk_archive(&value, &root, "run", "task").is_err());
}

#[test]
fn manifest_for_another_request_or_run_is_rejected() {
    let root = directory();
    let value = evidence(&root);
    assert!(verify_sdk_archive(&value, &root, "run", "other").is_err());
    assert!(verify_sdk_archive(&value, &root, "other", "task").is_err());
}

#[test]
fn archive_outside_configured_directory_is_rejected() {
    let root = directory();
    let value = evidence(&root);
    let elsewhere = directory();
    fs::create_dir(elsewhere.join("raw-sdk-save")).unwrap();
    assert!(verify_sdk_archive(&value, &elsewhere, "run", "task").is_err());
}

#[test]
fn input_audit_captures_each_request_input_and_marks_business_scope_incomplete() {
    let root = directory();
    let source = root.join("source.xlsx");
    {
        let mut zip = zip::ZipWriter::new(File::create(&source).unwrap());
        zip.start_file("xl/workbook.xml", zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Data"/><sheet name="User Sheet"/></sheets></workbook>"#).unwrap();
        zip.finish().unwrap();
    }
    let binary = root.join("legacy.xls");
    fs::write(&binary, b"synthetic BIFF placeholder, not a business file").unwrap();
    let before = fs::read(&source).unwrap();
    let record = write_input_trace(
        &root,
        "run",
        "task",
        "Modify",
        [source.as_path(), binary.as_path()].into_iter(),
    )
    .unwrap();
    let trace: Value = serde_json::from_slice(&fs::read(&record).unwrap()).unwrap();
    assert!(record.is_absolute());
    assert_eq!(trace["logical_business_inputs_complete"], false);
    assert_eq!(trace["inputs"].as_array().unwrap().len(), 2);
    assert_eq!(trace["inputs"][0]["sheets"], json!(["Data", "User Sheet"]));
    assert_eq!(trace["inputs"][1]["sheets_status"], "unknown");
    assert!(trace["inputs"][1]["sheets"].is_null());
    assert_eq!(fs::read(&source).unwrap(), before);
    let another = write_input_trace(
        &root,
        "run",
        "task",
        "Modify",
        [source.as_path()].into_iter(),
    )
    .unwrap();
    assert_ne!(record, another);
    assert_eq!(
        trace,
        serde_json::from_slice::<Value>(&fs::read(record).unwrap()).unwrap()
    );
}
