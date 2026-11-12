use std::{
    fs, io,
    path::PathBuf,
    process::Output,
    sync::atomic::{AtomicU64, Ordering},
};

static CAPTURE_ID: AtomicU64 = AtomicU64::new(0);

/// Opt-in capture of real child-process output from synthetic CLI tests.
pub fn capture(label: &str, output: &Output) -> io::Result<()> {
    let Some(root) = std::env::var_os("RUSTTOOLS_LOGGING_EVIDENCE_DIR") else {
        return Ok(());
    };
    let thread = std::thread::current();
    let test = thread.name().unwrap_or("unnamed-test");
    let safe = |value: &str| {
        value
            .chars()
            .map(|c| {
                if c.is_ascii_alphanumeric() || c == '-' || c == '_' {
                    c
                } else {
                    '_'
                }
            })
            .collect::<String>()
    };
    let directory = PathBuf::from(root)
        .join("terminal-captures")
        .join(safe(test))
        .join(format!(
            "{}-{}",
            safe(label),
            CAPTURE_ID.fetch_add(1, Ordering::Relaxed)
        ));
    fs::create_dir_all(&directory)?;
    fs::write(directory.join("stdout.txt"), &output.stdout)?;
    fs::write(directory.join("stderr.txt"), &output.stderr)?;
    fs::write(directory.join("exit.txt"), output.status.to_string())?;
    Ok(())
}
