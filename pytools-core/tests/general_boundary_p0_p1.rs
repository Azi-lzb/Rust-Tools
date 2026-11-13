//! Synthetic P0/P1 safety regressions. No business originals or release artifacts.
use anyhow::Result;
use pytools_core::{archive_ops::*, file_commit::*};
use rust_xlsxwriter::Workbook;
use std::{
    fs,
    path::{Path, PathBuf},
    process::Command,
};
struct Root(PathBuf);
impl Root {
    fn new() -> Self {
        let p = std::env::temp_dir().join(format!(
            "general-edge-{}-{:?}-{}",
            std::process::id(),
            std::thread::current().id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::create_dir_all(&p).unwrap();
        Self(p)
    }
    fn p(&self, n: &str) -> PathBuf {
        self.0.join(n)
    }
}
impl Drop for Root {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}
fn book(p: &Path, v: &str) -> Result<()> {
    let mut b = Workbook::new();
    b.add_worksheet().write_string(0, 0, v)?;
    b.save(p)?;
    Ok(())
}
#[test]
fn candidate_alias_rejected_without_any_byte_changes() -> Result<()> {
    let d = Root::new();
    let p = d.p("book.xlsx");
    book(&p, "source")?;
    fs::create_dir(d.p("child"))?;
    let before = fs::read(&p)?;
    for alias in [p.clone(), d.p("child/../book.xlsx")] {
        assert!(replace_target_from_candidate(&p, &alias, &sha256_file(&p)?).is_err());
        assert!(create_target_from_candidate(&p, &alias).is_err());
        assert_eq!(fs::read(&p)?, before);
    }
    assert_eq!(fs::read_dir(&d.0)?.count(), 2);
    Ok(())
}
#[test]
fn invalid_new_candidate_never_publishes_or_leaves_stage() -> Result<()> {
    let d = Root::new();
    let c = d.p("candidate.xlsx");
    fs::write(&c, b"broken ZIP")?;
    let before = fs::read(&c)?;
    assert!(create_target_from_candidate(&c, &d.p("result.xlsx")).is_err());
    assert!(!d.p("result.xlsx").exists());
    assert_eq!(fs::read(&c)?, before);
    assert_eq!(fs::read_dir(&d.0)?.count(), 1);
    Ok(())
}
#[cfg(windows)]
#[test]
fn locked_target_commit_fails_and_preserves_both_books() -> Result<()> {
    use std::os::windows::fs::OpenOptionsExt;
    let d = Root::new();
    let c = d.p("candidate.xlsx");
    let t = d.p("target.xlsx");
    book(&c, "new")?;
    book(&t, "old")?;
    let cb = fs::read(&c)?;
    let tb = fs::read(&t)?;
    let hash = sha256_file(&t)?;
    let lock = fs::OpenOptions::new().read(true).share_mode(1).open(&t)?;
    assert!(replace_target_from_candidate(&c, &t, &hash).is_err());
    assert_eq!(fs::read(&t)?, tb);
    assert_eq!(fs::read(&c)?, cb);
    drop(lock);
    Ok(())
}
#[cfg(windows)]
#[test]
fn readonly_target_commit_refuses_and_keeps_original() -> Result<()> {
    let d = Root::new();
    let c = d.p("candidate.xlsx");
    let t = d.p("target.xlsx");
    book(&c, "new")?;
    book(&t, "old")?;
    let tb = fs::read(&t)?;
    let cb = fs::read(&c)?;
    let hash = sha256_file(&t)?;
    let perms = fs::metadata(&t)?.permissions();
    let mut ro = perms.clone();
    ro.set_readonly(true);
    fs::set_permissions(&t, ro)?;
    let result = replace_target_from_candidate(&c, &t, &hash);
    fs::set_permissions(&t, perms)?;
    assert!(
        result.is_err(),
        "read-only target must not be silently replaced"
    );
    assert_eq!(fs::read(&t)?, tb);
    assert_eq!(fs::read(&c)?, cb);
    Ok(())
}
#[test]
fn archive_preview_is_readonly_and_late_collision_keeps_existing_file() -> Result<()> {
    let d = Root::new();
    let s = d.p("20260101_source.xlsx");
    fs::write(&s, b"synthetic-source")?;
    let preview = archive_by_date_preview(&[s.clone()], &d.p("target"), &d.p("reports"), None)?;
    assert_eq!(preview.copied, 0);
    assert!(!d.p("target").exists());
    assert_eq!(fs::read(&s)?, b"synthetic-source");
    let target = PathBuf::from(&preview.rows[0].target_path);
    fs::create_dir_all(target.parent().unwrap())?;
    fs::write(&target, b"preexisting")?;
    let result = execute_archive_plan(&preview.output, &d.p("execution"))?;
    assert_eq!(result.copied, 1);
    assert_eq!(result.failed, 0);
    assert_eq!(fs::read(&target)?, b"preexisting");
    assert_ne!(result.rows[0].target_path, target.to_string_lossy());
    assert_eq!(fs::read(&result.rows[0].target_path)?, b"synthetic-source");
    assert_eq!(fs::read(&s)?, b"synthetic-source");
    Ok(())
}
#[test]
fn archive_plan_rejects_source_changed_after_preview() -> Result<()> {
    let d = Root::new();
    let s = d.p("20260101_source.xlsx");
    fs::write(&s, b"planned bytes")?;
    let preview = archive_by_date_preview(&[s.clone()], &d.p("target"), &d.p("reports"), None)?;
    fs::write(&s, b"changed after planning")?;
    let result = execute_archive_plan(&preview.output, &d.p("execution"))?;
    assert_eq!(fs::read(&s)?, b"changed after planning");
    assert_eq!(
        result.copied, 0,
        "stale archive plan must not copy changed source data"
    );
    assert!(!Path::new(&result.rows[0].target_path).exists());
    Ok(())
}
#[test]
fn archive_plan_refuses_source_as_its_own_destination() -> Result<()> {
    let d = Root::new();
    let s = d.p("source.xlsx");
    fs::write(&s, b"original bytes")?;
    let plan = d.p("plan.xlsx");
    let mut b = Workbook::new();
    let w = b.add_worksheet();
    w.set_name("归档计划")?;
    for (i, v) in ["源文件路径", "目标文件路径", "状态"].iter().enumerate() {
        w.write_string(0, i as u16, *v)?;
    }
    w.write_string(1, 0, s.to_string_lossy().as_ref())?;
    w.write_string(1, 1, s.to_string_lossy().as_ref())?;
    w.write_string(1, 2, "待归档")?;
    b.save(&plan)?;
    let result = execute_archive_plan(&plan, &d.p("execution"))?;
    assert_eq!(fs::read(&s)?, b"original bytes");
    assert_eq!(
        result.copied, 0,
        "source/target identity must be rejected, not silently suffixed and copied"
    );
    assert!(!d.p("source_1.xlsx").exists());
    Ok(())
}
#[test]
fn archive_nested_destination_does_not_reingest_previous_archive_files() -> Result<()> {
    let d = Root::new();
    let input = d.p("input");
    let target = input.join("archive");
    fs::create_dir_all(&target)?;
    let source = input.join("20260101_source.xlsx");
    let old = target.join("20250101_previous.xlsx");
    fs::write(&source, b"new source")?;
    fs::write(&old, b"previous archive")?;
    let preview = archive_by_date_preview(&[input], &target, &d.p("reports"), None)?;
    assert_eq!(fs::read(&source)?, b"new source");
    assert_eq!(fs::read(&old)?, b"previous archive");
    assert!(
        !preview
            .rows
            .iter()
            .any(|row| Path::new(&row.source_path).starts_with(&target)),
        "nested destination must not become its own future source set"
    );
    Ok(())
}
#[test]
#[ignore = "requires packaged Windows Python/openpyxl/pywin32 and Excel COM; execute explicitly with --include-ignored"]
fn synthetic_python_cli_and_print_boundaries() -> Result<()> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).parent().unwrap();
    let python = root.join("dist/python/python.exe");
    assert!(
        python.exists(),
        "this explicit runtime test requires the packaged Python; do not silently skip"
    );
    let out = Command::new(python)
        .args(["-X", "utf8", "-I"])
        .arg(root.join("pytools-core/tests/python/test_general_boundaries.py"))
        .env("RUSTTOOLS_BOUNDARY_CLI", env!("CARGO_BIN_EXE_RustTools"))
        .env("RUSTTOOLS_TEST_COM", "1")
        .output()?;
    println!("{}", String::from_utf8_lossy(&out.stdout));
    println!("{}", String::from_utf8_lossy(&out.stderr));
    assert!(out.status.success(), "Python boundary assertions failed");
    Ok(())
}
