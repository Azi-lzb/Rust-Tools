use std::{
    io,
    path::{Path, PathBuf},
};

/// Prepare a path for native file APIs and isolated file workers. This does
/// not require the leaf to exist or resolve symlinks. Do not use this for
/// Office COM, whose path contract is different.
pub fn path_for_external_io(path: &Path) -> io::Result<PathBuf> {
    #[cfg(windows)]
    {
        use std::os::windows::ffi::{OsStrExt, OsStringExt};
        let absolute = std::path::absolute(path)?;
        let raw: Vec<u16> = absolute.as_os_str().encode_wide().collect();
        if raw.contains(&0) || raw.starts_with(&[92, 92, 46, 92]) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "invalid filesystem path",
            ));
        }
        if raw.starts_with(&[92, 92, 63, 92]) {
            return Ok(absolute);
        }
        let mut extended: Vec<u16> = if raw.starts_with(&[92, 92]) {
            r"\\?\UNC\".encode_utf16().collect()
        } else {
            r"\\?\".encode_utf16().collect()
        };
        extended.extend_from_slice(if raw.starts_with(&[92, 92]) {
            &raw[2..]
        } else {
            &raw
        });
        Ok(PathBuf::from(std::ffi::OsString::from_wide(&extended)))
    }
    #[cfg(not(windows))]
    {
        Ok(path.to_path_buf())
    }
}

#[cfg(all(test, windows))]
mod tests {
    use super::*;

    #[test]
    fn external_paths_cover_relative_unicode_unc_and_existing_prefixes() {
        let relative = Path::new(r"不存在的目录\..\业务 空格\新文件.xlsx");
        let result = path_for_external_io(relative).unwrap();
        assert!(result.is_absolute());
        assert!(result.as_os_str().to_string_lossy().starts_with(r"\\?\"));
        assert_eq!(path_for_external_io(&result).unwrap(), result);
        assert_eq!(
            path_for_external_io(Path::new(r"\\server\share\业务\new.xlsx")).unwrap(),
            PathBuf::from(r"\\?\UNC\server\share\业务\new.xlsx")
        );
        assert_eq!(
            path_for_external_io(Path::new(r"\\?\UNC\server\share\new.xlsx")).unwrap(),
            PathBuf::from(r"\\?\UNC\server\share\new.xlsx")
        );
        assert!(path_for_external_io(Path::new("bad\0path")).is_err());
        assert!(path_for_external_io(Path::new(r"\\.\PhysicalDrive0")).is_err());
    }
}
