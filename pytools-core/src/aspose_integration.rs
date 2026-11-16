//! Explicit per-run calculation selection. Auto retains existing routing.
use anyhow::{bail, ensure, Context, Result};
use excel_aspose::{AsposeWorkerBackend, CalcEngine, WorkbookRequest, WorkerConfig};
use sha2::{Digest, Sha256};
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    time::{Duration, SystemTime, UNIX_EPOCH},
};

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum ValidationCalcMode {
    #[default]
    Auto,
    Com,
    BetterOffice,
    Aspose,
}
impl ValidationCalcMode {
    pub fn parse(value: &str) -> Result<Self> {
        match value {
            "auto" => Ok(Self::Auto),
            "com" => Ok(Self::Com),
            "betteroffice" => Ok(Self::BetterOffice),
            "aspose" => Ok(Self::Aspose),
            _ => bail!("--calc-engine 仅支持 auto、com、betteroffice、aspose"),
        }
    }
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Com => "com",
            Self::BetterOffice => "betteroffice",
            Self::Aspose => "aspose",
        }
    }
    pub fn label(self) -> &'static str {
        match self {
            Self::Auto => "自动（现有规则）",
            Self::Com => "Excel/WPS COM",
            Self::BetterOffice => "betteroffice (pure Rust)",
            Self::Aspose => "Aspose.Cells .NET（评估模式）",
        }
    }
    pub fn preflight(self) -> Result<()> {
        match self {
            Self::Aspose => {
                ensure!(
                    cfg!(windows),
                    "Aspose 菜单接入当前仅验证 Windows，其他平台路由不变"
                );
                worker_command()?;
            }
            Self::Com => ensure!(
                cfg!(windows),
                "强制 Excel/WPS COM 需要 Windows；未切换到其他引擎"
            ),
            _ => {}
        }
        Ok(())
    }
}

#[derive(Clone, Debug, Default)]
pub struct ValidationExecution {
    pub mode: ValidationCalcMode,
    pub run_id: Option<String>,
    pub diagnostic_dir: Option<PathBuf>,
}

fn checked_file(value: std::ffi::OsString, name: &str) -> Result<PathBuf> {
    let path = PathBuf::from(value);
    ensure!(
        path.is_file(),
        "{name} 指向的程序或文件不存在：{}",
        path.display()
    );
    fs::canonicalize(&path).with_context(|| format!("定位 {name} 失败"))
}
pub(crate) fn worker_command() -> Result<(PathBuf, Vec<String>)> {
    if let Some(exe) = std::env::var_os("RUSTTOOLS_ASPOSE_WORKER_EXE") {
        return Ok((
            checked_file(exe, "RUSTTOOLS_ASPOSE_WORKER_EXE")?,
            vec!["--stdio".to_owned()],
        ));
    }
    let dotnet = std::env::var_os("RUSTTOOLS_ASPOSE_DOTNET");
    let dll = std::env::var_os("RUSTTOOLS_ASPOSE_WORKER_DLL");
    if dotnet.is_some() || dll.is_some() {
        let dotnet = checked_file(
            dotnet
                .context("需要同时设置 RUSTTOOLS_ASPOSE_DOTNET 和 RUSTTOOLS_ASPOSE_WORKER_DLL")?,
            "RUSTTOOLS_ASPOSE_DOTNET",
        )?;
        let dll = checked_file(
            dll.context("需要同时设置 RUSTTOOLS_ASPOSE_DOTNET 和 RUSTTOOLS_ASPOSE_WORKER_DLL")?,
            "RUSTTOOLS_ASPOSE_WORKER_DLL",
        )?;
        return Ok((
            dotnet,
            vec![dll.to_string_lossy().into_owned(), "--stdio".to_owned()],
        ));
    }
    let packaged = std::env::current_exe()?
        .parent()
        .context("无法定位程序目录")?
        .join("workers/aspose/AsposeWorker.exe");
    if packaged.is_file() {
        return Ok((packaged, vec!["--stdio".to_owned()]));
    }
    bail!("未找到 Aspose .NET Worker。请设置 RUSTTOOLS_ASPOSE_WORKER_EXE，或同时设置 RUSTTOOLS_ASPOSE_DOTNET 和 RUSTTOOLS_ASPOSE_WORKER_DLL；本次未回退其他计算引擎")
}

pub(crate) struct AsposeCalculationSession {
    engine: Box<dyn CalcEngine + Send>,
    run_id: String,
    diagnostic_dir: PathBuf,
}
impl AsposeCalculationSession {
    pub(crate) fn start(execution: &ValidationExecution, working_copy: &Path) -> Result<Self> {
        ensure!(
            std::env::var_os("PYTOOLS_CONVERT_ENGINES_DISABLED").is_none(),
            "外部计算引擎已禁用；未启动 Aspose"
        );
        let (executable, arguments) = worker_command()?;
        let run_id = execution.run_id.clone().unwrap_or_else(|| {
            format!(
                "aspose-validation-{}-{}",
                std::process::id(),
                SystemTime::now()
                    .duration_since(UNIX_EPOCH)
                    .unwrap_or_default()
                    .as_nanos()
            )
        });
        let diagnostic_dir = execution.diagnostic_dir.clone().unwrap_or_else(|| {
            working_copy
                .parent()
                .unwrap_or(Path::new("."))
                .join("aspose-diagnostics")
        });
        let engine = Box::new(AsposeWorkerBackend::new(WorkerConfig {
            executable,
            arguments,
            diagnostics_dir: diagnostic_dir.clone(),
            startup_timeout: Duration::from_secs(30),
            request_timeout: Duration::from_secs(180),
        }));
        Ok(Self {
            engine,
            run_id,
            diagnostic_dir,
        })
    }
    pub(crate) fn calculate_copy(
        &mut self,
        working_copy: &Path,
        regions: &[crate::conversion_engine::NamedValidationRegionRequest],
    ) -> Result<serde_json::Value> {
        ensure!(
            working_copy
                .extension()
                .and_then(|x| x.to_str())
                .is_some_and(|x| x.eq_ignore_ascii_case("xlsx")),
            "Aspose 校验计算暂仅支持 XLSX 副本；不会转换或丢弃 XLSM 宏"
        );
        let working_copy = fs::canonicalize(working_copy)?;
        let source = excel_reader::read_workbook(&working_copy)?;
        let before = crate::file_commit::sha256_file(&working_copy)?;
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let candidate = working_copy.with_file_name(format!(
            ".aspose-candidate-{}-{nonce}.xlsx",
            std::process::id()
        ));
        let result = self
            .engine
            .calculate(&WorkbookRequest {
                run_id: self.run_id.clone(),
                input: working_copy.clone(),
                candidate: candidate.clone(),
            })
            .context("Aspose 计算失败；未切换到其他引擎，校验副本未提交")?;
        let output = excel_reader::read_workbook(&candidate).context("Aspose 候选缓存无法回读")?;
        validate_recalculated_workbook(&source, &output)?;
        crate::conversion_engine::ensure_aspose_validation_functions_supported(&output, regions)?;
        let commit =
            crate::file_commit::replace_target_from_candidate(&candidate, &working_copy, &before)
                .context("Aspose 候选未通过快照检查或安全提交失败")?;
        let diagnostic = serde_json::json!({"run_id":self.run_id,"engine":"Aspose.Cells .NET","input":working_copy,"candidate":candidate,"outcome":result,"commit":commit});
        // A bounded per-run file: concurrent jobs must not interleave JSONL,
        // and caller-supplied run IDs must never become filesystem paths.
        let log = self.diagnostic_dir.join(format!(
            "host-calculations-{:x}.jsonl",
            Sha256::digest(self.run_id.as_bytes())
        ));
        if let Err(error) = (|| -> Result<()> {
            let mut f = fs::OpenOptions::new().create(true).append(true).open(log)?;
            serde_json::to_writer(&mut f, &diagnostic)?;
            writeln!(f)?;
            Ok(())
        })() {
            eprintln!("Aspose 诊断记录写入失败：{error}");
        }
        Ok(diagnostic)
    }
}

fn formula_key(formula: &str, names: &[&str]) -> String {
    // Normalize harmless quotes around simple known sheet names, only outside
    // Excel string literals. Absolute/relative references are never normalized.
    let mut out = String::new();
    let mut segment = String::new();
    let mut in_string = false;
    let flush = |segment: &mut String, out: &mut String| {
        for name in names {
            if name
                .chars()
                .next()
                .is_some_and(|c| c.is_alphabetic() || c == '_')
                && name.chars().all(|c| c.is_alphanumeric() || c == '_')
            {
                *segment = segment.replace(&format!("'{name}'!"), &format!("{name}!"));
            }
        }
        out.push_str(segment);
        segment.clear();
    };
    for c in formula.trim_start_matches('=').chars() {
        if c == '"' {
            if !in_string {
                flush(&mut segment, &mut out);
            }
            out.push(c);
            in_string = !in_string;
        } else if in_string {
            out.push(c);
        } else {
            segment.push(c);
        }
    }
    flush(&mut segment, &mut out);
    out
}
pub(crate) fn validate_recalculated_workbook(
    source: &excel_reader::WorkbookDump,
    output: &excel_reader::WorkbookDump,
) -> Result<()> {
    let names: Vec<&str> = source.sheets.iter().map(|s| s.name.as_str()).collect();
    ensure!(
        output.sheets.len() >= source.sheets.len(),
        "Aspose 计算改变了原工作表数量"
    );
    for (a, b) in source.sheets.iter().zip(&output.sheets) {
        ensure!(
            a.name == b.name && a.formulas.len() == b.formulas.len(),
            "Aspose 计算改变工作表顺序或公式数量：{}",
            a.name
        );
        for (cell, formula) in &a.formulas {
            ensure!(
                b.formulas
                    .get(cell)
                    .is_some_and(|f| formula_key(formula, &names) == formula_key(f, &names)),
                "Aspose 计算改变公式：{}!{cell}",
                a.name
            );
        }
        ensure!(
            a.merges == b.merges && a.comments == b.comments,
            "Aspose 计算改变合并区域或批注：{}",
            a.name
        );
        for (cell, value) in &a.values {
            if !a.formulas.contains_key(cell) {
                ensure!(
                    !b.formulas.contains_key(cell)
                        && b.values
                            .get(cell)
                            .is_some_and(|v| crate::conversion_engine::cell_values_equal(value, v)),
                    "Aspose 计算改变非公式数据：{}!{cell}",
                    a.name
                );
            }
        }
    }
    ensure!(
        output.sheets[source.sheets.len()..]
            .iter()
            .all(|s| s.name == "Evaluation Warning"),
        "Aspose 计算新增了非评估工作表"
    );
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn explicit_modes_are_strict() {
        assert_eq!(
            ValidationCalcMode::parse("aspose").unwrap(),
            ValidationCalcMode::Aspose
        );
        assert!(ValidationCalcMode::parse("typo").is_err());
        assert_eq!(ValidationCalcMode::default(), ValidationCalcMode::Auto);
    }
    #[test]
    fn formula_quotes_preserve_literal_and_reference_semantics() {
        assert_eq!(
            formula_key("='数据'!A1", &["数据"]),
            formula_key("=数据!A1", &["数据"])
        );
        assert_ne!(
            formula_key("=\"'数据'!A1\"", &["数据"]),
            formula_key("=\"数据!A1\"", &["数据"])
        );
        assert_ne!(
            formula_key("='数据'!$A$1", &["数据"]),
            formula_key("=数据!A1", &["数据"])
        );
    }
    #[test]
    fn recalculation_must_preserve_nonformula_data_and_formula_text() -> Result<()> {
        let root = std::env::temp_dir().join(format!(
            "aspose-calc-validation-{}-{}",
            std::process::id(),
            SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos()
        ));
        fs::create_dir_all(&root)?;
        let write = |name: &str, data: f64, formula: &str| -> Result<excel_reader::WorkbookDump> {
            let mut book = rust_xlsxwriter::Workbook::new();
            let sheet = book.add_worksheet();
            sheet.set_name("数据")?;
            sheet.write_number(0, 0, data)?;
            sheet.write_formula(0, 1, formula)?;
            let path = root.join(name);
            book.save(&path)?;
            excel_reader::read_workbook(path)
        };
        let source = write("source.xlsx", 10.0, "=A1+1")?;
        assert!(
            validate_recalculated_workbook(&source, &write("same.xlsx", 10.0, "=A1+1")?).is_ok()
        );
        assert!(
            validate_recalculated_workbook(&source, &write("wrong-data.xlsx", 20.0, "=A1+1")?)
                .is_err()
        );
        assert!(validate_recalculated_workbook(
            &source,
            &write("wrong-formula.xlsx", 10.0, "=A1+2")?
        )
        .is_err());
        Ok(())
    }
}
