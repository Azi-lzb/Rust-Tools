//! LibreOffice is reserved for optional validation recalculation.
//! Format conversion requests must fail before creating a candidate.

use anyhow::Result;
use pytools_core::conversion;
use std::{
    fs,
    path::PathBuf,
    time::{SystemTime, UNIX_EPOCH},
};

fn test_root() -> Result<PathBuf> {
    let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    let path = std::env::temp_dir().join(format!(
        "rust-tools-excel-lo-disabled-{}-{nonce}",
        std::process::id()
    ));
    fs::create_dir(&path)?;
    Ok(path)
}

#[test]
fn explicit_libreoffice_office_conversion_rejects_word_targets_without_output() -> Result<()> {
    let root = test_root()?;
    for (extension, target) in [("doc", "docx")] {
        let input = root.join(format!("input.{extension}"));
        let output = root.join(format!("output-{extension}.{target}"));
        fs::write(&input, b"isolated route-only fixture")?;

        let error = conversion::convert_one(&input, &output, target, "lo")
            .expect_err("Word targets must still reject the LO engine");
        let message = format!("{error:#}");
        assert!(
            message.contains("LibreOffice 不用于格式转换"),
            "unexpected route error for .{extension} -> .{target}: {message}"
        );
        assert!(
            !output.exists(),
            "rejected route must not create a candidate"
        );
    }
    fs::remove_dir_all(&root)?;
    Ok(())
}

#[test]
fn explicit_libreoffice_spreadsheet_conversion_is_rejected_before_creating_output() -> Result<()> {
    let root = test_root()?;
    let input = root.join("input.xls");
    let output = root.join("output.xlsx");
    let original = b"isolated route-only fixture";
    fs::write(&input, original)?;
    let error = conversion::convert_one(&input, &output, "xlsx", "lo")
        .expect_err("LibreOffice is not a spreadsheet conversion route");
    let message = format!("{error:#}");
    assert!(
        message.contains("LibreOffice 不用于格式转换") && message.contains("4-4 的 Aspose"),
        "spreadsheet rejection should direct users to the supported route: {message}"
    );
    assert_eq!(fs::read(&input)?, original);
    assert!(
        !output.exists(),
        "failed conversion must not create a candidate"
    );
    fs::remove_dir_all(&root)?;
    Ok(())
}

#[test]
fn native_value_rebuild_is_rejected_for_legacy_excel_format_conversion() -> Result<()> {
    let root = test_root()?;
    let input = root.join("input.xls");
    let output = root.join("output.xlsx");
    fs::write(&input, b"isolated route-only fixture")?;

    let error = conversion::convert_one(&input, &output, "xlsx", "native")
        .expect_err("legacy workbook conversion must not use value-only rebuild");
    let message = format!("{error:#}");
    assert!(
        message.contains("Excel workbook format conversion requires pywin32"),
        "unexpected native-route error: {message}"
    );
    assert!(!output.exists());
    fs::remove_dir_all(&root)?;
    Ok(())
}
