//! Explicit Aspose.Words conversion, isolated from the spreadsheet Worker.
use anyhow::{bail, ensure, Context, Result};
use excel_aspose::{
    AsposeDocumentBackend, CandidateDocument, DocumentConverterEngine, WorkbookRequest,
    WorkerConfig,
};
use sha2::{Digest, Sha256};
use std::{
    fs,
    io::{Read, Write},
    path::{Path, PathBuf},
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};

fn command() -> Result<(PathBuf, Vec<String>)> {
    if let Some(exe) = std::env::var_os("RUSTTOOLS_ASPOSE_WORDS_EXE") {
        let exe = PathBuf::from(exe);
        ensure!(exe.is_file(), "RUSTTOOLS_ASPOSE_WORDS_EXE 不存在");
        return Ok((exe, vec!["--stdio".into()]));
    }
    match (
        std::env::var_os("RUSTTOOLS_ASPOSE_DOTNET"),
        std::env::var_os("RUSTTOOLS_ASPOSE_WORDS_DLL"),
    ) {
        (Some(exe), Some(dll)) => {
            let exe = PathBuf::from(exe);
            let dll = PathBuf::from(dll);
            ensure!(
                exe.is_file() && dll.is_file(),
                "Aspose.Words 的 DOTNET 或 DLL 路径不存在"
            );
            Ok((
                exe,
                vec![dll.to_string_lossy().into_owned(), "--stdio".into()],
            ))
        }
        _ => {
            let exe = std::env::current_exe()?
                .parent()
                .context("missing executable directory")?
                .join(if cfg!(windows) {
                    "workers/aspose-words/AsposeWordsWorker.exe"
                } else {
                    "workers/aspose-words/AsposeWordsWorker"
                });
            ensure!(exe.is_file(), "未配置 Aspose.Words Worker；请设置 RUSTTOOLS_ASPOSE_WORDS_EXE，或 RUSTTOOLS_ASPOSE_DOTNET + RUSTTOOLS_ASPOSE_WORDS_DLL");
            Ok((exe, vec!["--stdio".into()]))
        }
    }
}
pub fn preflight() -> Result<()> {
    ensure!(
        cfg!(any(windows, target_os = "linux")),
        "Aspose.Words Worker 仅支持 Windows 和 Linux；未回退其他引擎"
    );
    ensure!(
        std::env::var_os("PYTOOLS_CONVERT_ENGINES_DISABLED").is_none(),
        "外部引擎已禁用"
    );
    command()?;
    Ok(())
}
pub(crate) struct AsposeDocumentSession {
    engine: Box<dyn DocumentConverterEngine + Send>,
    run_id: String,
    diagnostics: PathBuf,
}
impl AsposeDocumentSession {
    fn new() -> Result<Self> {
        preflight()?;
        let (executable, arguments) = command()?;
        let stamp = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let run_id = format!("aspose-words-{}-{stamp}", std::process::id());
        let diagnostics = crate::run_logging::default_log_dir()
            .join("aspose-words")
            .join(&run_id);
        let engine = Box::new(AsposeDocumentBackend::new(WorkerConfig {
            executable,
            arguments,
            diagnostics_dir: diagnostics.clone(),
            startup_timeout: Duration::from_secs(30),
            request_timeout: Duration::from_secs(180),
        }));
        Ok(Self {
            engine,
            run_id,
            diagnostics,
        })
    }
    fn convert(&mut self, input: &Path, output: &Path) -> Result<String> {
        let clock = Instant::now();
        let input = fs::canonicalize(input)?;
        ensure!(
            input
                .extension()
                .is_some_and(|e| e.eq_ignore_ascii_case("doc") || e.eq_ignore_ascii_case("docx")),
            "Aspose.Words 仅启用 DOC/DOCX；WPS/WPT 未验证，不回退"
        );
        ensure!(!output.exists(), "拒绝覆盖 Word 转换目标");
        let hash = crate::file_commit::sha256_file(&input)?;
        let parent = output
            .parent()
            .filter(|p| !p.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        fs::create_dir_all(parent)?;
        let parent = fs::canonicalize(parent)?;
        let stamp = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let scratch = parent.join(format!(".aspose-words-{}-{stamp}", std::process::id()));
        fs::create_dir(&scratch)?;
        struct Cleanup(PathBuf);
        impl Drop for Cleanup {
            fn drop(&mut self) {
                let _ = fs::remove_file(self.0.join("candidate.docx"));
                let _ = fs::remove_dir(&self.0);
            }
        }
        let _cleanup = Cleanup(scratch.clone());
        let candidate = scratch.join("candidate.docx");
        let result = self.engine.convert_document(&WorkbookRequest {
            run_id: self.run_id.clone(),
            input: input.clone(),
            candidate: candidate.clone(),
        })?;
        let _review_origin = crate::review_output::ReviewOriginScope::enter(&input, &hash);
        let _review_destination = crate::review_output::ReviewDestinationScope::enter(output);
        validate_document(&result).map_err(|error| {
            crate::review_output::retain_validation_error(&input, &candidate, &self.run_id, error)
        })?;
        crate::file_commit::verify_planning_snapshot(&input, Some(&hash))?;
        let commit = crate::file_commit::create_target_from_candidate_with_validator(
            &candidate,
            output,
            |path, expected_hash, _role| {
                ensure!(
                    crate::file_commit::sha256_file(path)? == expected_hash,
                    "DOCX candidate changed before publication"
                );
                let mut archive = zip::ZipArchive::new(fs::File::open(path)?)?;
                let mut xml = String::new();
                archive
                    .by_name("word/document.xml")?
                    .read_to_string(&mut xml)?;
                roxmltree::Document::parse(&xml)?;
                Ok(())
            },
        )?;
        let report = serde_json::json!({"status":"converted","engine":"aspose_words","run_id":self.run_id,"evaluation_output":true,"result":result,"source_sha256":hash,"validation":{"verified":true,"level":"paragraph_order_text_headers_footers_images_table_section_counts"},"commit":commit,"end_to_end_ms":clock.elapsed().as_secs_f64()*1000.0});
        fs::create_dir_all(&self.diagnostics)?;
        let mut log = fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(self.diagnostics.join("host.jsonl"))?;
        writeln!(log, "{report}")?;
        Ok(serde_json::to_string(&report)?)
    }
}
fn paragraph_text(node: roxmltree::Node<'_, '_>) -> String {
    let mut text = String::new();
    for n in node.descendants() {
        match n.tag_name().name() {
            "t" | "instrText" => text.push_str(n.text().unwrap_or_default()),
            "tab" => text.push('\t'),
            "br" | "cr" => text.push('\u{b}'),
            _ => {}
        }
    }
    text
}
fn digest(text: &str) -> String {
    format!("{:X}", Sha256::digest(text.as_bytes()))
}
fn evaluation_paragraph(text: &str) -> bool {
    // Read-only comparison exclusions, never physically removed from output.
    matches!(text, "Evaluation Only. Created with Aspose.Words. Copyright 2003-2026 Aspose Pty Ltd." |
        "Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License  HYPERLINK \"https://products.aspose.com/words/temporary-license/\" https://products.aspose.com/words/temporary-license/" |
        "Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License https://products.aspose.com/words/temporary-license/" |
        "This document was truncated here because it was created in the Evaluation Mode.")
}
fn validate_paragraphs(actual: Vec<String>, expected: &[String], ordered: bool) -> Result<()> {
    let mut hashes = Vec::new();
    for text in actual {
        // A truncation marker is a capability failure, never an acceptable watermark.
        ensure!(
            !text.contains("document was truncated"),
            "Aspose.Words 评估输出被截断，未发布"
        );
        if !evaluation_paragraph(&text) {
            hashes.push(digest(&text));
        }
    }
    let mut expected = expected.to_vec();
    if !ordered {
        hashes.sort();
        expected.sort();
    }
    ensure!(
        hashes == expected,
        "Aspose.Words 段落文本或顺序不一致（评估标记之外），未发布"
    );
    Ok(())
}
fn validate_document(result: &CandidateDocument) -> Result<()> {
    ensure!(
        result.source_shapes == result.source_image_hashes.len() as u64,
        "Aspose.Words 非图片绘图对象尚未完成保真验证，未发布"
    );
    let mut zip = zip::ZipArchive::new(fs::File::open(&result.path)?)?;
    let mut body = String::new();
    zip.by_name("word/document.xml")?
        .read_to_string(&mut body)?;
    let doc = roxmltree::Document::parse(&body)?;
    let paragraphs = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "p")
        .map(paragraph_text)
        .collect();
    validate_paragraphs(paragraphs, &result.source_paragraph_hashes, true)?;
    ensure!(
        doc.descendants()
            .filter(|n| n.has_tag_name((
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "tbl"
            )))
            .count() as u64
            == result.source_tables,
        "Aspose.Words 表格数量变化"
    );
    ensure!(
        doc.descendants()
            .filter(|n| n.has_tag_name((
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "sectPr"
            )))
            .count() as u64
            == result.source_sections,
        "Aspose.Words 节数量变化"
    );
    let names: Vec<_> = zip.file_names().map(str::to_owned).collect();
    let mut auxiliary = Vec::new();
    let mut images = Vec::new();
    for name in names {
        if (name.starts_with("word/header") || name.starts_with("word/footer"))
            && name.ends_with(".xml")
        {
            let mut xml = String::new();
            zip.by_name(&name)?.read_to_string(&mut xml)?;
            let doc = roxmltree::Document::parse(&xml)?;
            auxiliary.extend(
                doc.descendants()
                    .filter(|n| n.is_element() && n.tag_name().name() == "p")
                    .map(paragraph_text),
            );
        } else if name.starts_with("word/media/") {
            let mut data = Vec::new();
            zip.by_name(&name)?.read_to_end(&mut data)?;
            images.push(format!("{:X}", Sha256::digest(data)));
        }
    }
    validate_paragraphs(auxiliary, &result.source_auxiliary_paragraph_hashes, false)?;
    for image in &result.source_image_hashes {
        ensure!(images.contains(image), "Aspose.Words 图片内容变化");
    }
    Ok(())
}
pub(crate) fn convert(
    input: &Path,
    output: &Path,
    runtime: &mut crate::conversion_engine::ConversionRuntime,
) -> Result<String> {
    if runtime.aspose_document_converter.is_none() {
        runtime.aspose_document_converter = Some(AsposeDocumentSession::new()?);
    }
    runtime
        .aspose_document_converter
        .as_mut()
        .context("missing document session")?
        .convert(input, output)
}
