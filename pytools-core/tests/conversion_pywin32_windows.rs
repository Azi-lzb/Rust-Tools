#![cfg(windows)]

use anyhow::Result;
use pytools_core::conversion::{
    convert_excel_files_with_progress, convert_word_files_with_progress,
};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

struct TempDir(PathBuf);
impl TempDir {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let path = std::env::temp_dir().join(format!(
            "rust-tools-com-integration-{}-{nonce}",
            std::process::id()
        ));
        fs::create_dir(&path)?;
        Ok(Self(path))
    }
    fn path(&self) -> &Path {
        &self.0
    }
}
impl Drop for TempDir {
    fn drop(&mut self) {
        let parent = std::env::temp_dir().canonicalize();
        let target = self.0.canonicalize();
        if let (Ok(parent), Ok(target)) = (parent, target) {
            let is_ours = target.parent() == Some(parent.as_path())
                && target.file_name().is_some_and(|name| {
                    name.to_string_lossy()
                        .starts_with("rust-tools-com-integration-")
                });
            if is_ours {
                let _ = fs::remove_dir_all(target);
            }
        }
    }
}

#[test]
#[ignore = "requires locally installed Excel or WPS COM and the bundled pywin32 runtime"]
fn automatic_xls_conversion_uses_pywin32_batch_session_and_preserves_sources() -> Result<()> {
    let root = TempDir::new()?;
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/timeline/wide_summary_elapsed_duration.xls");
    let first = root.path().join("first.xls");
    let second = root.path().join("second.xls");
    for path in [&first, &second] {
        fs::copy(&fixture, path)?;
    }
    let before = [&first, &second]
        .iter()
        .map(|path| fs::read(path))
        .collect::<std::io::Result<Vec<_>>>()?;
    let mut events = Vec::new();
    let output = root.path().join("output");
    let run = convert_excel_files_with_progress(
        &[first.clone(), second.clone()],
        Some(&output),
        "xlsx",
        "auto",
        |event| events.push(event.clone()),
    )?;

    assert_eq!(run.converted, 2, "{:?}", run.items);
    assert_eq!(run.failed, 0, "{:?}", run.items);
    let details = run
        .items
        .iter()
        .map(|item| serde_json::from_str::<Value>(&item.detail))
        .collect::<std::result::Result<Vec<_>, _>>()?;
    assert!(details.iter().all(|detail| detail["engine"] == "pywin32"));
    assert!(details
        .iter()
        .all(|detail| detail["com_bridge"] == "pywin32"));
    assert!(details.iter().all(|detail| {
        detail["validation"]["verified"] == true
            && detail["validation"]["level"] == "excel_com_sheet_order_and_formula_cells"
    }));
    let worker_process_ids = details
        .iter()
        .map(|detail| detail["com_worker_process_id"].as_u64())
        .collect::<Vec<_>>();
    assert!(worker_process_ids[0].is_some());
    assert_eq!(worker_process_ids[0], worker_process_ids[1]);
    let office_process_ids = details
        .iter()
        .map(|detail| detail["com_process_id"].as_u64())
        .collect::<Vec<_>>();
    if office_process_ids.iter().all(Option::is_some) {
        assert_eq!(office_process_ids[0], office_process_ids[1]);
    }
    assert_eq!(
        [&first, &second]
            .iter()
            .map(|path| fs::read(path))
            .collect::<std::io::Result<Vec<_>>>()?,
        before
    );
    assert!(events.iter().any(|event| {
        event.actual_engine.as_deref() == Some("Microsoft Excel COM (pywin32)")
            || event.actual_engine.as_deref() == Some("WPS 表格 COM (pywin32)")
    }));
    Ok(())
}

/// DOC / WPS / WPT all route through the same pywin32 Word worker and are
/// validated by comparing the source and candidate body text.
#[test]
#[ignore = "requires locally installed Word or WPS Writer COM and the bundled pywin32 runtime"]
fn automatic_doc_wps_wpt_conversion_uses_pywin32_and_preserves_sources() -> Result<()> {
    let root = TempDir::new()?;
    let doc = root.path().join("probe-doc.doc");
    let wps = root.path().join("probe-wps.wps");
    let wpt = root.path().join("probe-wpt.wpt");
    if !make_word_fixtures(&doc, &wps, &wpt)? {
        eprintln!("skipping: neither Word nor WPS Writer COM can create the fixtures");
        return Ok(());
    }
    let sources = [doc.clone(), wps.clone(), wpt.clone()];
    let before = sources
        .iter()
        .map(|path| fs::read(path))
        .collect::<std::io::Result<Vec<_>>>()?;
    let mut events = Vec::new();
    let output = root.path().join("output");
    let run = convert_word_files_with_progress(&sources, Some(&output), "docx", "auto", |event| {
        events.push(event.clone())
    })?;

    assert_eq!(run.converted, 3, "{:?}", run.items);
    assert_eq!(run.failed, 0, "{:?}", run.items);
    let details = run
        .items
        .iter()
        .map(|item| serde_json::from_str::<Value>(&item.detail))
        .collect::<std::result::Result<Vec<_>, _>>()?;
    assert!(details.iter().all(|detail| detail["engine"] == "pywin32"));
    assert!(details
        .iter()
        .all(|detail| detail["com_bridge"] == "pywin32"));
    assert!(details
        .iter()
        .all(|detail| { detail["validation"]["level"] == "docx_package_and_text_content" }));
    for (index, (extension, stem)) in [
        ("doc", "probe-doc"),
        ("wps", "probe-wps"),
        ("wpt", "probe-wpt"),
    ]
    .iter()
    .enumerate()
    {
        assert_eq!(
            details[index]["validation"]["source_extension"], *extension,
            "worker must report the source extension it opened"
        );
        let converted = output.join(format!("{stem}.docx"));
        assert!(
            converted.is_file(),
            "missing converted candidate {}",
            converted.display()
        );
        assert_eq!(
            fs::read(&sources[index])?,
            before[index],
            "source .{extension} must not be modified"
        );
    }
    assert!(events.iter().any(|event| {
        event
            .actual_engine
            .as_deref()
            .is_some_and(|engine| engine.contains("COM (pywin32)"))
    }));
    Ok(())
}

/// Build readable DOC/WPS/WPT fixtures through the locally installed
/// Office/WPS Writer. Returns `false` when no Writer can create them.
fn make_word_fixtures(doc: &Path, wps: &Path, wpt: &Path) -> Result<bool> {
    let script = format!(
        "$ErrorActionPreference='Stop'\n\
function New-Writer {{\n\
  foreach ($n in @('kwps.Application','KWPS.Application','Word.Application')) {{\n\
    try {{ return (New-Object -ComObject $n) }} catch {{}}\n\
  }}\n\
  return $null\n\
}}\n\
function Save-Text($app, $path, $format) {{\n\
  $d=$app.Documents.Add()\n\
  try {{\n\
    $d.Content.Text='pywin32 word conversion probe'\n\
    $d.SaveAs2($path,$format)\n\
  }} finally {{ $d.Close($false) }}\n\
}}\n\
$app=New-Writer\n\
if ($null -eq $app) {{ exit 3 }}\n\
try {{\n\
  $app.Visible=$false\n\
  try {{ $app.DisplayAlerts=0 }} catch {{}}\n\
  Save-Text $app '{}' 0\n\
  Save-Text $app '{}' 0\n\
  Save-Text $app '{}' 1\n\
}} finally {{ $app.Quit() }}\n",
        doc.to_string_lossy().replace('\'', "''"),
        wps.to_string_lossy().replace('\'', "''"),
        wpt.to_string_lossy().replace('\'', "''"),
    );
    let script_path = doc.with_extension("make-word.ps1");
    fs::write(&script_path, script)?;
    let output = Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
        ])
        .arg(&script_path)
        .output()?;
    let _ = fs::remove_file(&script_path);
    if !output.status.success() {
        return Ok(false);
    }
    Ok([doc, wps, wpt].iter().all(|path| path.is_file()))
}
