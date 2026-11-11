//! Failure-safety fixture generator: corrupt / fake-extension / empty / OLE
//! header damage variants built from a valid source XLS. No Excel involved.

use anyhow::{ensure, Result};
use std::path::{Path, PathBuf};

pub fn run(source: &Path, out_dir: &Path) -> Result<()> {
    ensure!(
        source.is_file(),
        "source does not exist: {}",
        source.display()
    );
    std::fs::create_dir_all(out_dir)?;
    let bytes = std::fs::read(source)?;
    ensure!(
        bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]),
        "source is not an OLE2 compound file (XLS expected)"
    );

    let mut written: Vec<(String, PathBuf, usize)> = Vec::new();

    // 1) Truncated at ~60%: parser must reject, never panic, never emit output.
    let cut = bytes.len() * 6 / 10;
    written.push(write(
        out_dir,
        "corrupt-truncated.xls",
        &bytes[..cut],
    )?);

    // 2) Whole-file XOR corruption (valid OLE header size, garbage content).
    let xored: Vec<u8> = bytes.iter().map(|b| b ^ 0x5A).collect();
    written.push(write(out_dir, "corrupt-xor.xls", &xored)?);

    // 3) OLE header damaged at byte 0.
    let mut header = bytes.clone();
    header[0] = b'X';
    written.push(write(out_dir, "corrupt-ole-header.xls", &header)?);

    // 4) Fake extension: PNG content named .xls.
    let mut png = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    png.extend_from_slice(&[0; 64]);
    written.push(write(out_dir, "fake-png.xls", &png)?);

    // 5) Fake extension: real XLSX bytes (rxls round-trip of the source) named .xls.
    let source_workbook = rxls::Workbook::open(&bytes)?;
    let xlsx_bytes = source_workbook.to_xlsx();
    written.push(write(out_dir, "fake-xlsx-renamed.xls", &xlsx_bytes)?);

    // 6) Empty file.
    written.push(write(out_dir, "empty.xls", &[])?);

    // 7) OLE file with a directory entry pointing at a nonexistent stream:
    //    corrupt the FAT/MiniFAT region by zeroing the middle.
    let mut mid = bytes.clone();
    let start = bytes.len() / 3;
    let end = (bytes.len() * 2) / 3;
    for byte in &mut mid[start..end] {
        *byte = 0;
    }
    written.push(write(out_dir, "corrupt-holes.xls", &mid)?);

    eprintln!("failure fixtures written to {}:", out_dir.display());
    for (name, path, size) in written {
        eprintln!("  {name} ({} bytes) -> {}", size, path.display());
    }
    Ok(())
}

fn write(out_dir: &Path, name: &str, bytes: &[u8]) -> Result<(String, PathBuf, usize)> {
    let path = out_dir.join(name);
    std::fs::write(&path, bytes)?;
    Ok((name.to_owned(), path, bytes.len()))
}
