//! Office conversion is temporarily restricted to pywin32 with Microsoft
//! Excel or WPS COM. These checks ensure explicit LibreOffice requests fail
//! before creating an output or starting an external process.

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
fn explicit_libreoffice_office_conversion_is_rejected_without_output() -> Result<()> {
    let root = test_root()?;
    for (extension, target) in [("xls", "xlsx"), ("xlsx", "xlsx"), ("doc", "docx")] {
        let input = root.join(format!("input.{extension}"));
        let output = root.join(format!("output-{extension}.{target}"));
        fs::write(&input, b"isolated route-only fixture")?;

        let error = conversion::convert_one(&input, &output, target, "lo")
            .expect_err("Office conversion must reject the paused LO engine");
        let message = format!("{error:#}");
        assert!(
            message.contains("LibreOffice conversion is paused"),
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
