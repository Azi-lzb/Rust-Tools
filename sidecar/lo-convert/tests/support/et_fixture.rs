use anyhow::{ensure, Context, Result};
use std::{
    fs,
    path::{Path, PathBuf},
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

pub struct TempDirGuard(pub PathBuf);

impl Drop for TempDirGuard {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

pub fn temp_test_dir(prefix: &str) -> Result<TempDirGuard> {
    let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    let path = std::env::temp_dir().join(format!("{prefix}_{}_{}", std::process::id(), nonce));
    fs::create_dir_all(&path)?;
    Ok(TempDirGuard(path))
}

pub fn copy_and_seed_excel8_et_named_fixture(root: &Path, fixture: &Path) -> Result<PathBuf> {
    let input = root.join("input.et");
    fs::copy(fixture, &input).context("copy test fixture into isolated temp directory")?;
    let before = file_format(&input)?;
    ensure!(
        before == 56,
        "expected this explicitly mislabeled CFB/XLS fixture to report FileFormat 56, got {before}"
    );
    seed_test_data(&input)?;
    let after = file_format(&input)?;
    ensure!(
        after == 56,
        "WPS changed the test fixture format unexpectedly: FileFormat {after}"
    );
    Ok(input)
}

#[allow(dead_code)] // Used by the integration target; the binary unit-test build includes this module too.
pub fn copy_native_et_fixture(root: &Path, fixture: &Path) -> Result<PathBuf> {
    let input = root.join("input.et");
    fs::copy(fixture, &input).context("copy native ET fixture into isolated temp directory")?;
    let format = file_format(&input)?;
    ensure!(
        format == 65521,
        "fixture is not verified as native ET: KET reports FileFormat {format}, expected 65521"
    );
    Ok(input)
}

fn file_format(path: &Path) -> Result<i32> {
    let script_path = path.with_extension("file-format.ps1");
    let script = format!(
        "$ErrorActionPreference='Stop'; $app=$null; $book=$null;\ntry {{ $app=New-Object -ComObject 'KET.Application'; $app.Visible=$false; $app.DisplayAlerts=$false; $book=$app.Workbooks.Open('{}',0,$true); [Console]::Out.WriteLine([string]$book.FileFormat) }} finally {{ if ($null -ne $book) {{ $book.Close($false) }}; if ($null -ne $app) {{ $app.Quit() }} }}\n",
        ps_quote(path)
    );
    fs::write(&script_path, script)?;
    let result = Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
        ])
        .arg(&script_path)
        .output()
        .context("start KET COM to inspect test fixture file format")?;
    let _ = fs::remove_file(script_path);
    ensure!(
        result.status.success(),
        "KET could not inspect {}: {}",
        path.display(),
        String::from_utf8_lossy(&result.stderr)
    );
    let format = String::from_utf8_lossy(&result.stdout)
        .trim()
        .parse::<i32>()
        .with_context(|| format!("parse KET FileFormat output for {}", path.display()))?;
    Ok(format)
}

fn seed_test_data(path: &Path) -> Result<()> {
    let script_path = path.with_extension("seed-data.ps1");
    let script = format!(
        "$ErrorActionPreference='Stop'; $app=$null; $book=$null;\ntry {{ $app=New-Object -ComObject 'KET.Application'; $app.Visible=$false; $app.DisplayAlerts=$false; $book=$app.Workbooks.Open('{}',0,$false); $sheet=$book.Worksheets.Item(1); $sheet.Name='Data'; $sheet.Cells.Item(1,1).Value2=12; $sheet.Cells.Item(2,1).Value2=3; $sheet.Cells.Item(1,2).Formula='=SUM(A1:A2)'; $book.Save() }} finally {{ if ($null -ne $book) {{ $book.Close($false) }}; if ($null -ne $app) {{ $app.Quit() }} }}\n",
        ps_quote(path)
    );
    fs::write(&script_path, script)?;
    let result = Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
        ])
        .arg(&script_path)
        .output()
        .context("start KET COM to add deterministic test values")?;
    let _ = fs::remove_file(script_path);
    ensure!(
        result.status.success(),
        "KET could not seed test values: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    Ok(())
}

fn ps_quote(path: &Path) -> String {
    path.to_string_lossy().replace('\'', "''")
}
