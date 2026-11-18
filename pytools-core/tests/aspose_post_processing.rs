//! Isolated synthetic tests for Aspose output post-processing: raw SDK output
//! archiving, input worksheet tracing, and exact-name `Evaluation Warning`
//! worksheet removal with its safety conditions. No worker is involved; the
//! candidates are synthetic packages with SDK-like evaluation sheets.

use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};

use pytools_core::aspose_post::{
    archive_raw_candidate, remove_evaluation_warning_sheet, trace_input_workbook,
    write_trace_record, EvaluationPostOutcome, InputWorkbookTrace, EVALUATION_SHEET_NAME,
};

fn root(tag: &str) -> PathBuf {
    let dir = std::env::temp_dir().join(format!(
        "aspose-post-{tag}-{}-{}",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    fs::create_dir_all(&dir).unwrap();
    dir
}

fn book(path: &Path, sheets: &[(&str, &str)]) -> PathBuf {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    for (name, formula) in sheets {
        let sheet = workbook.add_worksheet();
        sheet.set_name(*name).unwrap();
        if !formula.is_empty() {
            sheet.write_formula(0, 0, *formula).unwrap();
        } else {
            sheet.write_string(0, 0, "value").unwrap();
        }
    }
    workbook.save(path).unwrap();
    path.to_path_buf()
}

fn trace(path: &Path, sheets: Option<Vec<String>>) -> InputWorkbookTrace {
    InputWorkbookTrace {
        path: path.to_path_buf(),
        sha256: Some("deadbeef".to_owned()),
        sheets,
    }
}

fn sheet_names(path: &Path) -> Vec<String> {
    excel_reader::read_workbook(path)
        .unwrap()
        .sheets
        .into_iter()
        .map(|sheet| sheet.name)
        .collect()
}

fn sha(path: &Path) -> String {
    pytools_core::file_commit::sha256_file(path).unwrap()
}

#[test]
fn exact_name_with_business_values_is_never_deleted() {
    let dir = root("exact");
    let input = book(&dir.join("input.xlsx"), &[("Data", "")]);
    let candidate = book(
        &dir.join("candidate.xlsx"),
        &[("Data", ""), (EVALUATION_SHEET_NAME, "")],
    );
    let outcome = remove_evaluation_warning_sheet(&candidate, &[trace_input_workbook(&input)]);
    assert!(matches!(outcome, EvaluationPostOutcome::UnsafeSkipped { .. }));
    assert_eq!(sheet_names(&candidate), ["Data", EVALUATION_SHEET_NAME]);
    // Input untouched.
    assert_eq!(sheet_names(&input), ["Data"]);
    fs::remove_dir_all(&dir).unwrap();
}

#[test]
fn not_present_and_case_or_near_names_do_not_match() {
    let dir = root("names");
    let input = book(&dir.join("input.xlsx"), &[("Data", "")]);
    let candidate = book(
        &dir.join("candidate.xlsx"),
        &[
            ("Data", ""),
            ("evaluation warning", ""),
            ("Evaluation Warning 2", ""),
            ("EvaluationWarning", ""),
        ],
    );
    let before = sha(&candidate);
    let outcome = remove_evaluation_warning_sheet(&candidate, &[trace_input_workbook(&input)]);
    assert_eq!(outcome, EvaluationPostOutcome::NotPresent);
    assert_eq!(sha(&candidate), before, "candidate must stay untouched");
    assert_eq!(sheet_names(&candidate).len(), 4);
    fs::remove_dir_all(&dir).unwrap();
}

#[test]
fn input_conflict_or_unknown_or_empty_inputs_skip() {
    let dir = root("conflict");
    let plain = book(&dir.join("plain.xlsx"), &[("Data", "")]);
    let conflicted = book(
        &dir.join("conflicted.xlsx"),
        &[("Data", ""), (EVALUATION_SHEET_NAME, "")],
    );
    let candidate = book(
        &dir.join("candidate.xlsx"),
        &[("Data", ""), (EVALUATION_SHEET_NAME, "")],
    );
    let before = sha(&candidate);

    let outcome = remove_evaluation_warning_sheet(
        &candidate,
        &[trace_input_workbook(&plain), trace_input_workbook(&conflicted)],
    );
    assert_eq!(
        outcome,
        EvaluationPostOutcome::ConflictSkipped {
            input: conflicted.clone()
        }
    );

    let unknown = trace(&plain, None);
    assert_eq!(
        remove_evaluation_warning_sheet(&candidate, &[unknown.clone()]),
        EvaluationPostOutcome::InputUnknown { input: plain.clone() }
    );

    assert_eq!(
        remove_evaluation_warning_sheet(&candidate, &[]),
        EvaluationPostOutcome::NoInputSourceUnverified
    );
    assert_eq!(sha(&candidate), before, "all skips keep the file untouched");
    fs::remove_dir_all(&dir).unwrap();
}

#[test]
fn cross_sheet_and_three_dee_references_block_removal() {
    let dir = root("refs");
    let input = book(&dir.join("input.xlsx"), &[("Data", "")]);
    for (tag, formula) in [
        ("cross", "='Evaluation Warning'!A1"),
        ("threedee", "=SUM('Data:Evaluation Warning'!A1)"),
    ] {
        let candidate = book(
            &dir.join(format!("candidate-{tag}.xlsx")),
            &[
                ("Data", formula),
                ("Other", ""),
                (EVALUATION_SHEET_NAME, ""),
            ],
        );
        let before = sha(&candidate);
        let outcome = remove_evaluation_warning_sheet(&candidate, &[trace_input_workbook(&input)]);
        match outcome {
            EvaluationPostOutcome::UnsafeSkipped { .. } => {}
            other => panic!("{tag}: expected UnsafeSkipped, got {other:?}"),
        }
        assert_eq!(sha(&candidate), before);
    }
    fs::remove_dir_all(&dir).unwrap();
}

#[test]
fn all_remaining_sheets_hidden_blocks_removal() {
    let dir = root("hidden");
    let input = book(&dir.join("input.xlsx"), &[("Data", "")]);
    let candidate = book(
        &dir.join("candidate.xlsx"),
        &[("Data", ""), (EVALUATION_SHEET_NAME, "")],
    );
    // Hide the only other sheet via workbook.xml state attribute.
    patch_zip_entry(&candidate, "xl/workbook.xml", |xml| {
        xml.replacen(
            "<sheet name=\"Data\"",
            "<sheet name=\"Data\" state=\"hidden\"",
            1,
        )
    });
    let outcome = remove_evaluation_warning_sheet(&candidate, &[trace_input_workbook(&input)]);
    assert!(matches!(outcome, EvaluationPostOutcome::UnsafeSkipped { .. }));
    fs::remove_dir_all(&dir).unwrap();
}

#[test]
fn read_only_legacy_check_preserves_active_view_names_and_calc_chain() {
    let dir = root("active-readonly");
    let input = book(&dir.join("input.xlsx"), &[("Data", ""), ("After", "")]);
    let candidate = book(&dir.join("candidate.xlsx"),
        &[("Data", "=After!A1"), (EVALUATION_SHEET_NAME, ""), ("After", "")]);
    patch_zip_entry(&candidate, "xl/workbook.xml", |xml| {
        let defined = "<definedNames><definedName name=\"LocalOnWarning\" localSheetId=\"1\">$A$1</definedName><definedName name=\"LocalOnAfter\" localSheetId=\"2\">$B$1</definedName></definedNames>";
        xml.replace("<sheets>", &format!("{defined}<sheets>"))
    });
    add_calc_chain(&candidate, &[1, 2, 3]);
    let before = sha(&candidate);
    assert!(matches!(remove_evaluation_warning_sheet(&candidate, &[trace_input_workbook(&input)]), EvaluationPostOutcome::UnsafeSkipped { .. }));
    assert_eq!(sha(&candidate), before);
    assert_eq!(sheet_names(&candidate), ["Data", EVALUATION_SHEET_NAME, "After"]);
    let xml = read_zip_entry(&candidate, "xl/workbook.xml");
    roxmltree::Document::parse(&xml).unwrap();
    assert!(xml.contains("localSheetId=\"2\""));
    assert!(read_zip_entry(&candidate, "xl/calcChain.xml").contains("i=\"3\""));
    fs::remove_dir_all(&dir).unwrap();
}

#[test]
fn read_only_roster_verification_keeps_business_blanks_and_exact_names() {
    use pytools_core::aspose_post::verify_sanitized_business_sheets;
    let dir = root("business-roster");
    let path = dir.join("candidate.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    for name in ["Data", "Blank", EVALUATION_SHEET_NAME, "evaluation warning 2"] {
        let sheet = workbook.add_worksheet();
        sheet.set_name(name).unwrap();
        if name != "Blank" { sheet.write_string(0, 0, "business content").unwrap(); }
    }
    workbook.save(&path).unwrap();
    let before = sha(&path);
    let names = sheet_names(&path);
    assert_eq!(verify_sanitized_business_sheets(&path, &names).unwrap(), EvaluationPostOutcome::NotPresent);
    assert!(verify_sanitized_business_sheets(&path, &names[..2]).is_err());
    assert!(verify_sanitized_business_sheets(&path, &[]).is_err());
    assert_eq!(sha(&path), before);
    fs::remove_dir_all(&dir).unwrap();
}

#[test]
fn empty_exact_named_business_sheet_is_not_a_deletion_signal() {
    let dir = root("blank-namesake");
    let input = book(&dir.join("input.xlsx"), &[("Data", "")]);
    let candidate = dir.join("candidate.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_name("Data").unwrap();
    workbook.add_worksheet().set_name(EVALUATION_SHEET_NAME).unwrap();
    workbook.save(&candidate).unwrap();
    let before = sha(&candidate);
    assert!(matches!(remove_evaluation_warning_sheet(&candidate, &[trace_input_workbook(&input)]), EvaluationPostOutcome::UnsafeSkipped { .. }));
    assert_eq!(sha(&candidate), before);
    assert_eq!(sheet_names(&candidate), ["Data", EVALUATION_SHEET_NAME]);
    fs::remove_dir_all(&dir).unwrap();
}

#[test]
fn archive_preserves_raw_copy_hash_and_avoids_collisions() {
    let dir = root("archive");
    let diagnostics = dir.join("diag");
    let first = book(&dir.join("candidate.xlsx"), &[("Data", "")]);
    let archive_one = archive_raw_candidate(&first, &diagnostics).unwrap();
    assert!(archive_one.path.is_file());
    assert_eq!(archive_one.sha256, sha(&first));
    // A second, different candidate must not overwrite the earlier archive.
    let second = book(&dir.join("candidate.xlsx"), &[("Data2", "")]);
    let archive_two = archive_raw_candidate(&second, &diagnostics).unwrap();
    assert_ne!(archive_one.path, archive_two.path);
    assert_eq!(sha(&archive_one.path), archive_one.sha256, "earlier archive intact");
    assert_eq!(archive_two.sha256, sha(&second));
    // Trace record written beside the archive with input sheet list or unknown.
    let input = book(&dir.join("input.xlsx"), &[("Data", "")]);
    let traced = trace_input_workbook(&input);
    assert_eq!(traced.sheets.as_deref(), Some(&["Data".to_owned()][..]));
    let unknown = trace(&dir.join("missing.xlsx"), None);
    let record = write_trace_record(
        &archive_two,
        "run-1",
        "request-1",
        "aspose_worker (test)",
        true,
        &[traced, unknown],
        &EvaluationPostOutcome::Processed,
    )
    .unwrap();
    let value: serde_json::Value =
        serde_json::from_slice(&fs::read(&record).unwrap()).unwrap();
    assert_eq!(value["evaluation_sheet_post_processing"], "processed");
    assert_eq!(value["inputs"][0]["sheets"][0], "Data");
    assert!(value["inputs"][1]["sheets"].is_null(), "unknown stays null");
    fs::remove_dir_all(&dir).unwrap();
}

/// Rewrite one zip entry in place (test scaffolding only).
fn patch_zip_entry(path: &Path, entry: &str, edit: impl FnOnce(String) -> String) {
    let mut edit = Some(edit);
    let source = fs::read(path).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(source)).unwrap();
    let mut items = Vec::new();
    for index in 0..zip.len() {
        let mut file = zip.by_index(index).unwrap();
        let name = file.name().to_owned();
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes).unwrap();
        items.push((name, bytes));
    }
    let out = fs::File::create(path).unwrap();
    let mut writer = zip::ZipWriter::new(out);
    for (name, bytes) in items {
        if name == entry {
            let text = String::from_utf8(bytes).unwrap();
            let edit = edit.take().expect("patch entry appears once");
            writer
                .start_file(
                    name,
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .unwrap();
            std::io::Write::write_all(&mut writer, edit(text).as_bytes()).unwrap();
        } else {
            writer
                .start_file(
                    name.clone(),
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .unwrap();
            std::io::Write::write_all(&mut writer, &bytes).unwrap();
        }
    }
    writer.finish().unwrap();
}

fn read_zip_entry(path: &Path, entry: &str) -> String {
    let mut zip = zip::ZipArchive::new(fs::File::open(path).unwrap()).unwrap();
    let mut text = String::new();
    zip.by_name(entry).unwrap().read_to_string(&mut text).unwrap();
    text
}

/// Insert a synthetic calcChain part with entries for 1-based sheet indices.
fn add_calc_chain(path: &Path, sheet_indices: &[u64]) {
    let entries: String = sheet_indices
        .iter()
        .map(|i| format!("<c r=\"A1\" i=\"{i}\"/>"))
        .collect();
    let chain = format!("<?xml version=\"1.0\"?><calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{}\">{entries}</calcChain>", sheet_indices.len());
    let source = fs::read(path).unwrap();
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(source)).unwrap();
    let mut items = Vec::new();
    for index in 0..zip.len() {
        let mut file = zip.by_index(index).unwrap();
        let name = file.name().to_owned();
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes).unwrap();
        items.push((name, bytes));
    }
    let out = fs::File::create(path).unwrap();
    let mut writer = zip::ZipWriter::new(out);
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for (name, bytes) in items {
        if name == "[Content_Types].xml" {
            let text = String::from_utf8(bytes.clone()).unwrap();
            let added = text.replace(
                "</Types>",
                "<Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/></Types>",
            );
            writer.start_file(name, options).unwrap();
            std::io::Write::write_all(&mut writer, added.as_bytes()).unwrap();
        } else if name == "xl/_rels/workbook.xml.rels" {
            let text = String::from_utf8(bytes.clone()).unwrap();
            let added = text.replace(
                "</Relationships>",
                "<Relationship Id=\"rIdCalc\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.xml\"/></Relationships>",
            );
            writer.start_file(name, options).unwrap();
            std::io::Write::write_all(&mut writer, added.as_bytes()).unwrap();
        } else {
            writer.start_file(name, options).unwrap();
            std::io::Write::write_all(&mut writer, &bytes).unwrap();
        }
    }
    writer.start_file("xl/calcChain.xml", options).unwrap();
    std::io::Write::write_all(&mut writer, chain.as_bytes()).unwrap();
    writer.finish().unwrap();
}
