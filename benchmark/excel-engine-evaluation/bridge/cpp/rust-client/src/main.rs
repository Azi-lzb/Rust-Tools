use std::{
    env,
    ffi::{c_void, CStr, CString},
    os::raw::c_char,
    path::Path,
};

#[link(name = "xlbridge")]
extern "C" {
    fn xl_open(path: *const c_char, xlsx: i32) -> *mut c_void;
    fn xl_append_rows(
        handle: *mut c_void,
        labels: *const *const c_char,
        values: *const f64,
        count: usize,
    ) -> i32;
    fn xl_save(handle: *mut c_void, path: *const c_char) -> i32;
    fn xl_verify_last_row(handle: *mut c_void, label: *const c_char, value: f64) -> i32;
    fn xl_close(handle: *mut c_void);
    fn xl_last_error() -> *const c_char;
}

fn error_text() -> String {
    unsafe {
        let p = xl_last_error();
        if p.is_null() {
            "unknown bridge error".into()
        } else {
            CStr::from_ptr(p).to_string_lossy().into_owned()
        }
    }
}

fn main() -> Result<(), String> {
    let args: Vec<String> = env::args().collect();
    if args.len() != 4 {
        return Err("usage: libxl-rust-ffi-prototype INPUT.xlsx OUTPUT.xlsx ROWS".into());
    }
    let input = CString::new(args[1].as_bytes()).map_err(|e| e.to_string())?;
    let output = CString::new(args[2].as_bytes()).map_err(|e| e.to_string())?;
    let rows: usize = args[3]
        .parse()
        .map_err(|e| format!("invalid row count: {e}"))?;
    let start = std::time::Instant::now();
    let handle = unsafe { xl_open(input.as_ptr(), 1) };
    if handle.is_null() {
        return Err(format!("open: {}", error_text()));
    }
    let labels: Vec<CString> = (0..rows)
        .map(|i| CString::new(format!("ffi-row-{i}")).unwrap())
        .collect();
    let label_ptrs: Vec<*const c_char> = labels.iter().map(|s| s.as_ptr()).collect();
    let values: Vec<f64> = (0..rows).map(|i| i as f64).collect();
    let status = unsafe { xl_append_rows(handle, label_ptrs.as_ptr(), values.as_ptr(), rows) };
    if status != 0 {
        unsafe { xl_close(handle) };
        return Err(format!("append: {}", error_text()));
    }
    let save_status = unsafe { xl_save(handle, output.as_ptr()) };
    if save_status != 0 {
        unsafe { xl_close(handle) };
        return Err(format!("save: {}", error_text()));
    }
    unsafe { xl_close(handle) };
    if !Path::new(&args[2]).is_file() {
        return Err("save returned success but output file is missing".into());
    }
    let verify = unsafe { xl_open(output.as_ptr(), 1) };
    if verify.is_null() {
        return Err(format!("reopen output: {}", error_text()));
    }
    let expected = CString::new(format!("ffi-row-{}", rows - 1)).unwrap();
    let verify_status = unsafe { xl_verify_last_row(verify, expected.as_ptr(), (rows - 1) as f64) };
    if verify_status != 0 {
        unsafe { xl_close(verify) };
        return Err(format!("verify output: {}", error_text()));
    }
    unsafe { xl_close(verify) };
    println!(
        "PASS engine=libxl ffi_rows={rows} reopen=PASS total_ms={:.3} output_bytes={}",
        start.elapsed().as_secs_f64() * 1000.0,
        std::fs::metadata(&args[2])
            .map_err(|e| e.to_string())?
            .len()
    );
    Ok(())
}
