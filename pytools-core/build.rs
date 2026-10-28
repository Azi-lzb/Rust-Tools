//! Embeds a build timestamp so the running executable can show which build
//! the user is actually looking at. The raw UTC value is baked in by Cargo and
//! converted to local time at startup with chrono.

use std::time::{SystemTime, UNIX_EPOCH};

fn main() {
    println!("cargo:rerun-if-changed=src");
    println!("cargo:rerun-if-env-changed=PYTOOLS_BUILD_UTC");
    if std::env::var_os("PYTOOLS_BUILD_UTC").is_none() {
        let seconds = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|duration| duration.as_secs())
            .unwrap_or_default();
        println!("cargo:rustc-env=PYTOOLS_BUILD_UTC={seconds}");
    }
}
