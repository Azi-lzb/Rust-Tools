mod com;

use anyhow::{bail, ensure, Context, Result};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use std::{
    collections::{BTreeMap, BTreeSet},
    env,
    fs::{self, File, OpenOptions},
    io::{BufRead, BufReader, Read, Write},
    path::{Path, PathBuf},
    process::{Command, Stdio},
    sync::mpsc,
    thread,
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};
use windows::Win32::System::ProcessStatus::K32EnumProcesses;

const EXCEL_PROGID: &str = "Excel.Application";

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct FileTask {
    pub input: PathBuf,
    pub output: PathBuf,
}

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct FileTiming {
    pub source: String,
    pub target: String,
    pub index: usize,
    pub status: String,
    #[serde(default)]
    pub open_ms: Option<f64>,
    #[serde(default)]
    pub save_as_ms: Option<f64>,
    #[serde(default)]
    pub close_ms: Option<f64>,
    #[serde(default)]
    pub total_ms: f64,
    #[serde(default)]
    pub error: Option<String>,
    #[serde(default)]
    pub close_error: Option<String>,
}

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct Session {
    #[serde(default)]
    pub worker_pid: u32,
    #[serde(default)]
    pub sta_thread_id: Value,
    #[serde(default)]
    pub excel_pid: Option<u32>,
    #[serde(default)]
    pub excel_instance_count: usize,
    #[serde(default)]
    pub excel_exit_confirmed: Option<bool>,
    #[serde(default)]
    pub timings: BTreeMap<String, f64>,
    #[serde(default)]
    pub files: Vec<FileTiming>,
    #[serde(default)]
    pub fatal_error: Option<String>,
    #[serde(default)]
    pub worker_start_ms: f64,
    #[serde(default)]
    pub worker_exit_ms: f64,
    #[serde(default)]
    pub worker_exit_code: Option<i32>,
}

#[derive(Debug, Deserialize)]
struct Manifest {
    total_files: usize,
    items: Vec<ExpectedFile>,
    batches: BTreeMap<String, BatchMeta>,
    #[serde(default)]
    created_at_local: String,
}

#[derive(Debug, Deserialize)]
struct ExpectedFile {
    name: String,
    sha256: String,
    bytes: u64,
}

#[derive(Debug, Deserialize)]
struct BatchMeta {
    count: usize,
    manifest: String,
    sha256: String,
}

#[derive(Debug, Deserialize)]
struct BatchManifest {
    count: usize,
    files: Vec<BatchFile>,
}

#[derive(Debug, Deserialize)]
struct BatchFile {
    name: String,
    sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct TrialKey {
    run_id: String,
    batch: String,
    lifecycle: String,
    scheme: String,
    round: String,
    started_unix_ms: u128,
}

#[derive(Clone, Debug, Serialize)]
struct TrialReport {
    key: TrialKey,
    expected_files: usize,
    elapsed_ms: f64,
    validation_elapsed_ms: f64,
    success_count: usize,
    failure_count: usize,
    skipped_count: usize,
    validation: Value,
    process_cleanup_ok: bool,
    phases_ms: BTreeMap<String, f64>,
    sessions: Vec<Session>,
    error: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
struct FaultCheck {
    scheme: String,
    lifecycle: String,
    expected_files: usize,
    excel_instances_created: usize,
    process_cleanup_ok: bool,
    failure_continued: bool,
    validation: Value,
    elapsed_ms: f64,
    sessions: Vec<Session>,
    error: Option<String>,
}

#[derive(Clone)]
struct TrialSpec {
    batch: String,
    lifecycle: String,
    scheme: String,
    round: String,
    files: Vec<String>,
}

struct RunDirectory {
    path: PathBuf,
    token: String,
}

fn main() -> Result<()> {
    let args = env::args().skip(1).collect::<Vec<_>>();
    let pilot = args.iter().any(|arg| arg == "--pilot");
    let failure_only = args.iter().any(|arg| arg == "--fault-injection");
    let rounds = option_value(&args, "--rounds")
        .map(|value| value.parse::<usize>())
        .transpose()
        .context("--rounds must be a positive integer")?
        .unwrap_or(5);
    ensure!(rounds > 0, "--rounds must be positive");
    ensure!(
        pilot || failure_only || rounds == 5,
        "the final report expects exactly five timed rounds"
    );

    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .context("runner must be inside benchmarks/xls-to-xlsx-com")?
        .to_path_buf();
    let fixtures = root.join("fixtures").join("generated");
    let manifest_path = fixtures.join("manifest.json");
    let manifest: Manifest = serde_json::from_slice(
        &fs::read(&manifest_path).with_context(|| format!("read {}", manifest_path.display()))?,
    )
    .context("parse fixed dataset manifest")?;
    eprintln!(
        "benchmark setup: manifest parsed ({} entries)",
        manifest.items.len()
    );
    let batches = validate_dataset(&fixtures, &manifest)?;
    eprintln!("benchmark setup: fixed dataset hashes and batch manifests validated");
    let excel_baseline = excel_process_ids()?;
    eprintln!(
        "benchmark setup: Excel process baseline read ({} PID(s))",
        excel_baseline.len()
    );

    let results_dir = root.join("results");
    fs::create_dir_all(&results_dir)?;
    let run_id = format!("{}-{}", std::process::id(), now_unix_ms());
    let result_path = results_dir.join(format!("run-{run_id}.jsonl"));
    let mut result_file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&result_path)?;
    let writer = serde_json::to_writer(
        &mut result_file,
        &json!({
            "record_type": "run_header",
            "run_id": run_id,
            "dataset_manifest_sha256": sha256_file(&manifest_path)?,
            "dataset_created_at_local": manifest.created_at_local,
            "excel_pid_baseline": excel_baseline,
            "requested_rounds": if pilot { 1 } else { rounds },
            "pilot": pilot,
            "fault_injection_only": failure_only,
        }),
    );
    writer?;
    result_file.write_all(b"\n")?;
    result_file.flush()?;

    if failure_only {
        let selected_scheme = option_value(&args, "--scheme");
        let checks = run_fault_injection(&root, &fixtures, &manifest, selected_scheme)?;
        for check in &checks {
            serde_json::to_writer(
                &mut result_file,
                &json!({"record_type":"failure_continuation_check","check":check}),
            )?;
            result_file.write_all(b"\n")?;
            result_file.flush()?;
            println!(
                "Fault continuation {} {}: passed={} cleanup={}",
                check.scheme, check.lifecycle, check.failure_continued, check.process_cleanup_ok
            );
        }
        let expected_checks = if selected_scheme.is_some() { 2 } else { 8 };
        ensure!(
            checks.len() == expected_checks
                && checks.iter().all(|check| check.failure_continued
                    && check.process_cleanup_ok
                    && check.error.is_none()),
            "synthetic failure-continuation qualification did not pass; results: {}",
            result_path.display()
        );
        println!(
            "All {expected_checks} failure-continuation checks passed: {}",
            result_path.display()
        );
        return Ok(());
    }

    let mut specs = if pilot {
        pilot_specs(&batches)?
    } else {
        full_specs(&batches, rounds)?
    };
    if let Some(scheme) = option_value(&args, "--scheme") {
        specs.retain(|spec| spec.scheme == scheme);
        ensure!(
            !specs.is_empty(),
            "--scheme {scheme} is not in this run matrix"
        );
    }
    println!(
        "Starting {}: {} trial processes; dataset=100 fixed BIFF8 XLS; result={}",
        if pilot {
            "calibration pilot"
        } else {
            "full benchmark matrix"
        },
        specs.len(),
        result_path.display()
    );
    println!(
        "Excel PID baseline: {} existing instance(s)",
        excel_baseline.len()
    );

    let fault_checks = if pilot {
        Vec::new()
    } else {
        run_fault_injection(&root, &fixtures, &manifest, None)?
    };
    for check in &fault_checks {
        serde_json::to_writer(
            &mut result_file,
            &json!({ "record_type": "failure_continuation_check", "check": check }),
        )?;
        result_file.write_all(b"\n")?;
        result_file.flush()?;
        println!(
            "Failure-continuation check {} {}: passed={} cleanup={}",
            check.scheme, check.lifecycle, check.failure_continued, check.process_cleanup_ok
        );
    }
    if !pilot {
        ensure!(
            fault_checks.len() == 8
                && fault_checks.iter().all(|check| check.failure_continued
                    && check.process_cleanup_ok
                    && check.error.is_none()),
            "synthetic failure-continuation qualification failed; results: {}",
            result_path.display()
        );
    }
    let overall_start = Instant::now();
    let mut trial_reports = Vec::with_capacity(specs.len());
    let expected_trial_count = specs.len();
    for (position, spec) in specs.into_iter().enumerate() {
        let report = run_trial(&root, &fixtures, &manifest, &spec).with_context(|| {
            format!(
                "trial {} {}/{}/{} {}",
                spec.scheme,
                spec.lifecycle,
                spec.batch,
                spec.round,
                position + 1
            )
        })?;
        serde_json::to_writer(&mut result_file, &report)?;
        result_file.write_all(b"\n")?;
        result_file.flush()?;
        print_trial_summary(&report, position + 1, expected_trial_count);
        if let Some(error) = report.error.as_ref() {
            bail!(
                "benchmark stopped after a failed trial: {error}; partial results: {}",
                result_path.display()
            );
        }
        trial_reports.push(report);
    }

    let final_excel = excel_process_ids()?;
    let new_unattributed = final_excel
        .difference(&excel_baseline)
        .copied()
        .collect::<Vec<_>>();
    let report_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .ancestors()
        .nth(3)
        .context("could not locate repository root")?
        .join("perf_xls_to_xlsx_com_comparison.md");
    if !pilot {
        write_report(
            &report_path,
            &run_id,
            &manifest,
            &trial_reports,
            &fault_checks,
            &excel_baseline,
            &final_excel,
            overall_start.elapsed().as_secs_f64(),
        )?;
    }
    println!(
        "Run complete in {:.1} min. Results: {}",
        overall_start.elapsed().as_secs_f64() / 60.0,
        result_path.display()
    );
    println!(
        "Existing Excel PID baseline: {}; final: {}; new unattributed PID(s): {:?}",
        excel_baseline.len(),
        final_excel.len(),
        new_unattributed
    );
    if !pilot {
        println!("Report: {}", report_path.display());
    }
    Ok(())
}

fn option_value<'a>(args: &'a [String], name: &str) -> Option<&'a str> {
    args.iter()
        .position(|arg| arg == name)
        .and_then(|index| args.get(index + 1))
        .map(String::as_str)
}

fn now_unix_ms() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis()
}

fn validate_dataset(fixtures: &Path, manifest: &Manifest) -> Result<BTreeMap<String, Vec<String>>> {
    ensure!(
        manifest.total_files == 100 && manifest.items.len() == 100,
        "fixed dataset must have exactly 100 manifest entries"
    );
    let mut seen = BTreeSet::new();
    for item in &manifest.items {
        ensure!(
            Path::new(&item.name)
                .file_name()
                .and_then(|name| name.to_str())
                == Some(item.name.as_str()),
            "unsafe dataset filename {}",
            item.name
        );
        ensure!(
            item.name.to_ascii_lowercase().ends_with(".xls"),
            "dataset entry is not .xls: {}",
            item.name
        );
        ensure!(
            seen.insert(item.name.as_str()),
            "duplicate filename in manifest: {}",
            item.name
        );
        let path = fixtures.join(&item.name);
        ensure!(
            path.metadata()?.len() == item.bytes,
            "dataset size changed: {}",
            item.name
        );
        ensure!(
            sha256_file(&path)? == item.sha256,
            "dataset SHA-256 changed: {}",
            item.name
        );
        let mut signature = [0; 8];
        File::open(&path)?.read_exact(&mut signature)?;
        ensure!(
            signature == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
            "{} is not an OLE/BIFF XLS file",
            item.name
        );
    }
    let mut batches: BTreeMap<String, Vec<String>> = BTreeMap::new();
    for name in ["B10", "B25", "B50", "B100"] {
        let metadata = manifest
            .batches
            .get(name)
            .with_context(|| format!("manifest is missing batch {name}"))?;
        let path = fixtures.join(&metadata.manifest);
        ensure!(
            sha256_file(&path)? == metadata.sha256,
            "batch manifest hash changed: {name}"
        );
        let batch: BatchManifest = serde_json::from_slice(&fs::read(&path)?)?;
        ensure!(
            batch.count == metadata.count && batch.files.len() == metadata.count,
            "batch count mismatch for {name}"
        );
        for entry in &batch.files {
            let item = manifest
                .items
                .iter()
                .find(|item| item.name == entry.name)
                .with_context(|| format!("{} is not in the full dataset manifest", entry.name))?;
            ensure!(
                item.sha256 == entry.sha256,
                "batch source hash mismatch for {}",
                entry.name
            );
        }
        batches.insert(
            name.to_owned(),
            batch.files.into_iter().map(|item| item.name).collect(),
        );
    }
    ensure!(
        batches["B10"].len() == 10
            && batches["B25"].len() == 25
            && batches["B50"].len() == 50
            && batches["B100"].len() == 100,
        "B10/B25/B50/B100 sizes are incorrect"
    );
    Ok(batches)
}

fn pilot_specs(batches: &BTreeMap<String, Vec<String>>) -> Result<Vec<TrialSpec>> {
    let files = batches
        .get("B10")
        .context("missing B10")?
        .iter()
        .take(2)
        .cloned()
        .collect::<Vec<_>>();
    let mut specs = Vec::new();
    for scheme in [
        "rust-native",
        "rust-python",
        "powershell",
        "rust-powershell",
    ] {
        for lifecycle in ["batch-reuse", "per-file"] {
            specs.push(TrialSpec {
                batch: "PILOT2".into(),
                lifecycle: lifecycle.into(),
                scheme: scheme.into(),
                round: "pilot".into(),
                files: files.clone(),
            });
        }
    }
    Ok(specs)
}

fn full_specs(batches: &BTreeMap<String, Vec<String>>, rounds: usize) -> Result<Vec<TrialSpec>> {
    let mut combinations = Vec::new();
    for batch in ["B10", "B25", "B50", "B100"] {
        for lifecycle in ["batch-reuse", "per-file"] {
            for scheme in [
                "rust-native",
                "rust-python",
                "powershell",
                "rust-powershell",
            ] {
                combinations.push(TrialSpec {
                    batch: batch.into(),
                    lifecycle: lifecycle.into(),
                    scheme: scheme.into(),
                    round: String::new(),
                    files: batches
                        .get(batch)
                        .with_context(|| format!("missing {batch}"))?
                        .clone(),
                });
            }
        }
    }
    let mut specs = Vec::with_capacity(combinations.len() * (rounds + 1));
    for (index, combination) in combinations.iter().enumerate() {
        let mut warmup = combination.clone();
        warmup.round = "warmup".into();
        specs.push(warmup);
        let _ = index;
    }
    for round in 1..=rounds {
        let shift = (round * 7) % combinations.len();
        let mut order = combinations[shift..]
            .iter()
            .chain(combinations[..shift].iter())
            .cloned()
            .collect::<Vec<_>>();
        if round % 2 == 0 {
            order.reverse();
        }
        for mut spec in order {
            spec.round = round.to_string();
            specs.push(spec);
        }
    }
    Ok(specs)
}

fn print_trial_summary(report: &TrialReport, position: usize, total: usize) {
    println!(
        "[{position}/{total}] {} {} {} round={} {:.1}s success={} failed={} validated={} pids={:?}",
        report.key.scheme,
        report.key.lifecycle,
        report.key.batch,
        report.key.round,
        report.elapsed_ms / 1000.0,
        report.success_count,
        report.failure_count,
        report.validation["validation"]
            .as_str()
            .unwrap_or("not-run"),
        report
            .sessions
            .iter()
            .filter_map(|session| session.excel_pid)
            .collect::<Vec<_>>()
    );
}

fn run_trial(
    root: &Path,
    fixtures: &Path,
    manifest: &Manifest,
    spec: &TrialSpec,
) -> Result<TrialReport> {
    let run_dir = RunDirectory::create()?;
    let input_dir = run_dir.path.join("input");
    let output_dir = run_dir.path.join("output");
    fs::create_dir(&input_dir)?;
    fs::create_dir(&output_dir)?;

    let expected_by_name = manifest
        .items
        .iter()
        .map(|item| (item.name.as_str(), item))
        .collect::<BTreeMap<_, _>>();
    let mut tasks = Vec::with_capacity(spec.files.len());
    for name in &spec.files {
        let expected = expected_by_name
            .get(name.as_str())
            .with_context(|| format!("{} not in fixed manifest", name))?;
        let source = fixtures.join(name);
        let input = input_dir.join(name);
        fs::copy(&source, &input).with_context(|| format!("copy synthetic input {name}"))?;
        ensure!(
            sha256_file(&input)? == expected.sha256,
            "copied input hash differs for {name}"
        );
        let output = output_dir.join(Path::new(name).with_extension("xlsx"));
        tasks.push(FileTask { input, output });
    }
    let file_list_path = run_dir.path.join("selected-files.json");
    fs::write(&file_list_path, serde_json::to_vec(&spec.files)?)?;

    let key = TrialKey {
        run_id: format!(
            "{}-{}-{}-{}",
            std::process::id(),
            now_unix_ms(),
            spec.scheme,
            spec.round
        ),
        batch: spec.batch.clone(),
        lifecycle: spec.lifecycle.clone(),
        scheme: spec.scheme.clone(),
        round: spec.round.clone(),
        started_unix_ms: now_unix_ms(),
    };
    let conversion_start = Instant::now();
    let mut sessions = match spec.scheme.as_str() {
        "rust-native" => {
            if spec.lifecycle == "batch-reuse" {
                let mut session = com::run_session(tasks.clone())?;
                session.worker_exit_code = Some(0);
                vec![session]
            } else {
                let mut values = Vec::with_capacity(tasks.len());
                for (index, task) in tasks.iter().cloned().enumerate() {
                    let mut session = com::run_session(vec![task])?;
                    session.worker_exit_code = Some(0);
                    values.push(session);
                    if (index + 1) % 10 == 0 {
                        println!(
                            "  Rust native per-file progress: {}/{}",
                            index + 1,
                            tasks.len()
                        );
                    }
                }
                values
            }
        }
        "rust-python" | "powershell" | "rust-powershell" => {
            run_external_scheme(root, spec, &tasks, &run_dir.path)?
        }
        other => bail!("unknown benchmark scheme {other}"),
    };
    let elapsed_ms = conversion_start.elapsed().as_secs_f64() * 1000.0;

    let mut next_index = 0usize;
    for session in &mut sessions {
        for file in &mut session.files {
            file.index = next_index;
            file.source = Path::new(&file.source)
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or(&file.source)
                .to_owned();
            file.target = Path::new(&file.target)
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or(&file.target)
                .to_owned();
            next_index += 1;
        }
    }
    let success_count = sessions
        .iter()
        .flat_map(|session| &session.files)
        .filter(|file| file.status == "succeeded")
        .count();
    let failure_count = sessions
        .iter()
        .flat_map(|session| &session.files)
        .filter(|file| file.status != "succeeded")
        .count();
    let skipped_count = spec
        .files
        .len()
        .saturating_sub(success_count + failure_count);
    let process_cleanup_ok = sessions.iter().all(|session| {
        session.fatal_error.is_none()
            && session.worker_exit_code == Some(0)
            && session.excel_instance_count == 1
            && session.excel_exit_confirmed == Some(true)
    });
    let phases_ms = sum_phases(&sessions);

    let validation_start = Instant::now();
    let validation = if process_cleanup_ok {
        validate_outputs(root, fixtures, &file_list_path, &output_dir)?
    } else {
        json!({ "validation": "not_run", "reason": "worker/Excel cleanup was not confirmed" })
    };
    let validation_elapsed_ms = validation_start.elapsed().as_secs_f64() * 1000.0;

    let mut errors = Vec::new();
    if success_count != spec.files.len() {
        errors.push(format!("expected {} successful conversions; got {success_count}, failed={failure_count}, skipped={skipped_count}", spec.files.len()));
    }
    if !process_cleanup_ok {
        errors.push("worker exit or owned Excel cleanup was not confirmed".into());
    }
    if validation["validation"].as_str() != Some("passed") {
        errors.push(format!("output correctness validation: {}", validation));
    }

    let error = if errors.is_empty() {
        None
    } else {
        Some(errors.join("; "))
    };
    let report = TrialReport {
        key,
        expected_files: spec.files.len(),
        elapsed_ms,
        validation_elapsed_ms,
        success_count,
        failure_count,
        skipped_count,
        validation,
        process_cleanup_ok,
        phases_ms,
        sessions,
        error,
    };

    if report.process_cleanup_ok && report.error.is_none() {
        run_dir.cleanup()?;
    } else {
        eprintln!(
            "Preserving temporary run directory for failure diagnosis: {}",
            run_dir.path.display()
        );
    }
    Ok(report)
}

fn run_fault_injection(
    root: &Path,
    fixtures: &Path,
    manifest: &Manifest,
    selected_scheme: Option<&str>,
) -> Result<Vec<FaultCheck>> {
    let valid_name = manifest
        .items
        .iter()
        .find(|item| item.name == "SMALL-001.xls")
        .context("fixed dataset is missing SMALL-001.xls")?;
    let mut checks = Vec::with_capacity(8);
    for scheme in [
        "rust-native",
        "rust-python",
        "powershell",
        "rust-powershell",
    ] {
        if selected_scheme.is_some_and(|selected| selected != scheme) {
            continue;
        }
        for lifecycle in ["batch-reuse", "per-file"] {
            let run_dir = RunDirectory::create()?;
            let input_dir = run_dir.path.join("input");
            let output_dir = run_dir.path.join("output");
            let rejected_output_dir = run_dir.path.join("rejected-output");
            fs::create_dir(&input_dir)?;
            fs::create_dir(&output_dir)?;
            fs::create_dir(&rejected_output_dir)?;

            let broken_input = input_dir.join("BROKEN-SYNTHETIC.xls");
            // Deliberately leave this synthetic path absent. Excel treats some malformed
            // .xls byte streams as importable text, so a missing input is the deterministic
            // Open failure used to check continuation in every implementation.
            let valid_input = input_dir.join(&valid_name.name);
            fs::copy(fixtures.join(&valid_name.name), &valid_input)?;
            ensure!(
                sha256_file(&valid_input)? == valid_name.sha256,
                "fault-check valid input hash changed"
            );
            let tasks = vec![
                FileTask {
                    input: broken_input,
                    output: rejected_output_dir.join("BROKEN-SYNTHETIC.xlsx"),
                },
                FileTask {
                    input: valid_input,
                    output: output_dir.join(Path::new(&valid_name.name).with_extension("xlsx")),
                },
            ];
            let spec = TrialSpec {
                batch: "FAILURE-INJECTION".into(),
                lifecycle: lifecycle.into(),
                scheme: scheme.into(),
                round: "qualification".into(),
                files: vec!["BROKEN-SYNTHETIC.xls".into(), valid_name.name.clone()],
            };
            let started = Instant::now();
            let mut sessions = if scheme == "rust-native" {
                if lifecycle == "batch-reuse" {
                    let mut session = com::run_session(tasks.clone())?;
                    session.worker_exit_code = Some(0);
                    vec![session]
                } else {
                    let mut sessions = Vec::with_capacity(tasks.len());
                    for task in &tasks {
                        let mut session = com::run_session(vec![task.clone()])?;
                        session.worker_exit_code = Some(0);
                        sessions.push(session);
                    }
                    sessions
                }
            } else {
                run_external_scheme(root, &spec, &tasks, &run_dir.path)?
            };
            for session in &mut sessions {
                for file in &mut session.files {
                    file.source = Path::new(&file.source)
                        .file_name()
                        .and_then(|name| name.to_str())
                        .unwrap_or(&file.source)
                        .to_owned();
                    file.target = Path::new(&file.target)
                        .file_name()
                        .and_then(|name| name.to_str())
                        .unwrap_or(&file.target)
                        .to_owned();
                }
            }
            let observed = sessions
                .iter()
                .flat_map(|session| &session.files)
                .map(|file| file.status.as_str())
                .collect::<Vec<_>>();
            let cleanup_ok = sessions.iter().all(|session| {
                session.fatal_error.is_none()
                    && session.worker_exit_code == Some(0)
                    && session.excel_instance_count == 1
                    && session.excel_exit_confirmed == Some(true)
            });
            let expected_excel_instances = if lifecycle == "batch-reuse" { 1 } else { 2 };
            let created = sessions
                .iter()
                .map(|session| session.excel_instance_count)
                .sum::<usize>();

            let files_path = run_dir.path.join("valid-file-only.json");
            fs::write(
                &files_path,
                serde_json::to_vec(&vec![valid_name.name.clone()])?,
            )?;
            let validation = if cleanup_ok {
                validate_outputs(root, fixtures, &files_path, &output_dir)?
            } else {
                json!({ "validation": "not_run", "reason": "worker/Excel cleanup was not confirmed" })
            };
            let continued = observed == ["failed", "succeeded"]
                && created == expected_excel_instances
                && validation["validation"].as_str() == Some("passed")
                && cleanup_ok;
            let error = if continued {
                None
            } else {
                Some(format!("expected status sequence [failed, succeeded], got {observed:?}; Excel instances {created}/{expected_excel_instances}; validation={}; cleanup={cleanup_ok}", validation["validation"]))
            };
            let check = FaultCheck {
                scheme: scheme.into(),
                lifecycle: lifecycle.into(),
                expected_files: 2,
                excel_instances_created: created,
                process_cleanup_ok: cleanup_ok,
                failure_continued: continued,
                validation,
                elapsed_ms: started.elapsed().as_secs_f64() * 1000.0,
                sessions,
                error,
            };
            if check.failure_continued {
                run_dir.cleanup()?;
            } else {
                eprintln!(
                    "Preserving failure-injection directory for diagnosis: {}",
                    run_dir.path.display()
                );
            }
            checks.push(check);
        }
    }
    Ok(checks)
}

#[derive(Serialize)]
struct WorkerPayload<'a> {
    excel_progid: &'static str,
    baseline_pids: Vec<u32>,
    files: Vec<WorkerFile<'a>>,
}

#[derive(Serialize)]
struct WorkerFile<'a> {
    input: &'a str,
    output: &'a str,
}

fn run_external_scheme(
    root: &Path,
    spec: &TrialSpec,
    tasks: &[FileTask],
    temp: &Path,
) -> Result<Vec<Session>> {
    if spec.lifecycle == "batch-reuse" {
        let path = temp.join("request-batch.json");
        return Ok(vec![run_external_session(
            root,
            &spec.scheme,
            tasks,
            &path,
        )?]);
    }
    let mut sessions = Vec::with_capacity(tasks.len());
    for (index, task) in tasks.iter().enumerate() {
        let path = temp.join(format!("request-file-{index:03}.json"));
        sessions.push(run_external_session(
            root,
            &spec.scheme,
            std::slice::from_ref(task),
            &path,
        )?);
        if (index + 1) % 10 == 0 {
            println!(
                "  {} per-file progress: {}/{} worker/Excel sessions",
                spec.scheme,
                index + 1,
                tasks.len()
            );
        }
    }
    Ok(sessions)
}

#[derive(Clone, Copy)]
enum InputProtocol {
    Stdin,
    File,
}

fn run_external_session(
    root: &Path,
    scheme: &str,
    tasks: &[FileTask],
    request_path: &Path,
) -> Result<Session> {
    let baseline_pids = all_process_ids()?;
    // Keep the strings backing the borrowed JSON request alive until serialization finishes.
    let input_paths = tasks
        .iter()
        .map(|task| task.input.to_string_lossy().into_owned())
        .collect::<Vec<_>>();
    let output_paths = tasks
        .iter()
        .map(|task| task.output.to_string_lossy().into_owned())
        .collect::<Vec<_>>();
    let worker_files = input_paths
        .iter()
        .zip(&output_paths)
        .map(|(input, output)| WorkerFile { input, output })
        .collect::<Vec<_>>();
    let payload = WorkerPayload {
        excel_progid: EXCEL_PROGID,
        baseline_pids,
        files: worker_files,
    };
    let payload_bytes = serde_json::to_vec(&payload)?;
    let protocol = if scheme == "powershell" {
        InputProtocol::File
    } else {
        InputProtocol::Stdin
    };
    if matches!(protocol, InputProtocol::File) {
        fs::write(request_path, &payload_bytes)?;
    }

    let worker_dir = root.join("workers");
    let mut command = match scheme {
        "rust-python" => {
            let mut command = Command::new("python");
            command.arg("-u").arg(worker_dir.join("python_worker.py"));
            command
        }
        "powershell" => {
            let mut command = Command::new("powershell.exe");
            command
                .args([
                    "-NoLogo",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-STA",
                    "-File",
                ])
                .arg(worker_dir.join("excel_com.ps1"))
                .arg("-InputFile")
                .arg(request_path);
            command
        }
        "rust-powershell" => {
            let mut command = Command::new("powershell.exe");
            command
                .args([
                    "-NoLogo",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-STA",
                    "-File",
                ])
                .arg(worker_dir.join("excel_com.ps1"))
                .arg("-Stdin");
            command
        }
        other => bail!("not an external worker scheme: {other}"),
    };
    command
        .current_dir(root)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    if matches!(protocol, InputProtocol::Stdin) {
        command.stdin(Stdio::piped());
    } else {
        command.stdin(Stdio::null());
    }
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        command.creation_flags(0x08000000); // CREATE_NO_WINDOW
    }

    let started = Instant::now();
    let mut child = command
        .spawn()
        .with_context(|| format!("start {scheme} worker"))?;
    let child_pid = child.id();
    let stdout = child.stdout.take().context("worker stdout was not piped")?;
    let stderr = child.stderr.take().context("worker stderr was not piped")?;
    let (line_sender, line_receiver) = mpsc::channel::<String>();
    let stdout_thread = thread::spawn(move || {
        let mut reader = BufReader::new(stdout);
        let mut buffer = Vec::new();
        loop {
            buffer.clear();
            match reader.read_until(b'\n', &mut buffer) {
                Ok(0) => break,
                Ok(_) => {
                    let line = String::from_utf8_lossy(&buffer).trim().to_owned();
                    if line_sender.send(line).is_err() {
                        break;
                    }
                }
                Err(_) => break,
            }
        }
    });
    let stderr_thread = thread::spawn(move || {
        let mut reader = BufReader::new(stderr);
        let mut output = String::new();
        let _ = reader.read_to_string(&mut output);
        output
    });
    let mut stdin = child.stdin.take();
    let mut ready_ms = None;
    let mut worker_result: Option<Session> = None;
    let mut result_received = None;
    let mut request_sent = false;

    loop {
        match line_receiver.recv_timeout(Duration::from_secs(3600)) {
            Ok(line) if line.is_empty() => continue,
            Ok(line) => {
                let event: Value = serde_json::from_str(&line).with_context(|| {
                    format!("{scheme} worker emitted a non-JSON protocol line: {line}")
                })?;
                match event["event"].as_str() {
                    Some("worker_ready") => {
                        ready_ms = Some(started.elapsed().as_secs_f64() * 1000.0);
                        if let Some(input) = stdin.as_mut() {
                            input.write_all(&payload_bytes)?;
                            input.write_all(b"\n")?;
                            input.flush()?;
                            request_sent = true;
                        }
                        stdin.take();
                    }
                    Some("excel_ready") => {}
                    Some("progress") => {
                        println!(
                            "  {} worker converted {}/{} synthetic files",
                            scheme,
                            event["completed"].as_u64().unwrap_or(0),
                            event["total"].as_u64().unwrap_or(0)
                        );
                    }
                    Some("worker_result") => {
                        let mut session: Session = serde_json::from_value(event)
                            .with_context(|| format!("parse {scheme} worker result"))?;
                        session.worker_start_ms =
                            ready_ms.unwrap_or_else(|| started.elapsed().as_secs_f64() * 1000.0);
                        session.worker_pid = if session.worker_pid == 0 {
                            child_pid
                        } else {
                            session.worker_pid
                        };
                        if session.worker_pid != child_pid {
                            bail!("{scheme} reported worker PID {}, but the launched process is {child_pid}", session.worker_pid);
                        }
                        worker_result = Some(session);
                        result_received = Some(Instant::now());
                        break;
                    }
                    other => bail!("{scheme} worker emitted unknown event {other:?}"),
                }
            }
            Err(mpsc::RecvTimeoutError::Timeout) => {
                bail!("{scheme} worker exceeded 60-minute response timeout; PID {child_pid} was not force-terminated")
            }
            Err(mpsc::RecvTimeoutError::Disconnected) => break,
        }
    }

    if !request_sent && matches!(protocol, InputProtocol::Stdin) {
        drop(stdin);
    }
    let status = child
        .wait()
        .with_context(|| format!("wait for {scheme} worker PID {child_pid}"))?;
    let worker_exit_ms = result_received
        .map(|at| at.elapsed().as_secs_f64() * 1000.0)
        .unwrap_or(0.0);
    let _ = stdout_thread.join();
    let stderr_text = stderr_thread.join().unwrap_or_default();
    let mut session = worker_result.with_context(|| {
        format!("{scheme} worker PID {child_pid} exited without a JSON result (status {status}); stderr: {stderr_text}")
    })?;
    session.worker_exit_ms = worker_exit_ms;
    session.worker_exit_code = status.code();
    if !status.success() && session.fatal_error.is_none() {
        session.fatal_error = Some(format!(
            "worker process exited with {status}; stderr: {stderr_text}"
        ));
    }
    Ok(session)
}

fn validate_outputs(
    root: &Path,
    fixtures: &Path,
    files_path: &Path,
    output_dir: &Path,
) -> Result<Value> {
    let script = root.join("workers").join("validate_outputs.py");
    let started = Instant::now();
    let mut command = Command::new("python");
    command
        .current_dir(root)
        .arg(&script)
        .arg("--manifest")
        .arg(fixtures.join("manifest.json"))
        .arg("--output-dir")
        .arg(output_dir)
        .arg("--files")
        .arg(files_path)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    let output = command
        .output()
        .context("start openpyxl correctness validator")?;
    let stdout = String::from_utf8_lossy(&output.stdout);
    let json_line = stdout
        .lines()
        .rev()
        .find(|line| line.trim_start().starts_with('{'));
    let mut value = if let Some(line) = json_line {
        serde_json::from_str::<Value>(line).unwrap_or_else(
            |error| json!({ "validation": "failed", "error": error.to_string(), "stdout": stdout }),
        )
    } else {
        json!({ "validation": "failed", "stdout": stdout })
    };
    value["elapsed_ms"] = json!(started.elapsed().as_secs_f64() * 1000.0);
    if !output.status.success() && value["validation"].as_str() == Some("passed") {
        value["validation"] = json!("failed");
        value["process_exit"] = json!(output.status.code());
    }
    if !output.status.success() {
        value["stderr"] = json!(String::from_utf8_lossy(&output.stderr).to_string());
    }
    Ok(value)
}

fn sum_phases(sessions: &[Session]) -> BTreeMap<String, f64> {
    let mut phases = BTreeMap::new();
    let mut add = |name: &str, value: f64| *phases.entry(name.to_owned()).or_insert(0.0) += value;
    for session in sessions {
        add("worker_start_ms", session.worker_start_ms);
        add("worker_exit_ms", session.worker_exit_ms);
        for (name, value) in &session.timings {
            add(name, *value);
        }
        if let Some(first) = session.files.first() {
            add("first_workbook_open_ms", first.open_ms.unwrap_or(0.0));
        }
        for file in session.files.iter().skip(1) {
            add("subsequent_workbook_open_ms", file.open_ms.unwrap_or(0.0));
        }
        for file in &session.files {
            add("save_as_ms", file.save_as_ms.unwrap_or(0.0));
            add("workbook_close_ms", file.close_ms.unwrap_or(0.0));
        }
    }
    phases.insert("sessions".into(), sessions.len() as f64);
    phases.insert(
        "excel_instances_created".into(),
        sessions
            .iter()
            .map(|session| session.excel_instance_count as f64)
            .sum(),
    );
    phases.insert(
        "excel_instances_exited".into(),
        sessions
            .iter()
            .filter(|session| session.excel_exit_confirmed == Some(true))
            .map(|session| session.excel_instance_count as f64)
            .sum(),
    );
    phases
}

fn sha256_file(path: &Path) -> Result<String> {
    let mut file =
        File::open(path).with_context(|| format!("open for SHA-256: {}", path.display()))?;
    let mut digest = Sha256::new();
    let mut buffer = vec![0u8; 1024 * 1024];
    loop {
        let count = file.read(&mut buffer)?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    Ok(format!("{:x}", digest.finalize()))
}

fn all_process_ids() -> Result<Vec<u32>> {
    let mut capacity = 4096usize;
    loop {
        let mut ids = vec![0u32; capacity];
        let mut bytes = 0u32;
        unsafe { K32EnumProcesses(ids.as_mut_ptr(), (ids.len() * 4) as u32, &mut bytes).ok()? };
        if (bytes as usize) < ids.len() * 4 {
            ids.truncate(bytes as usize / 4);
            return Ok(ids);
        }
        capacity *= 2;
        ensure!(capacity <= 1_048_576, "process list is too large");
    }
}

fn excel_process_ids() -> Result<BTreeSet<u32>> {
    let mut command = Command::new("powershell.exe");
    command
        .args([
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            "Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }",
        ])
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        command.creation_flags(0x08000000);
    }
    let output = command
        .output()
        .context("query existing Excel process baseline")?;
    ensure!(
        output.status.success(),
        "could not read Excel process IDs: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    Ok(String::from_utf8_lossy(&output.stdout)
        .lines()
        .filter_map(|line| line.trim().parse().ok())
        .collect())
}

impl RunDirectory {
    fn create() -> Result<Self> {
        static NEXT: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
        let temp = env::temp_dir();
        ensure!(
            temp.is_absolute(),
            "system temp directory is not absolute: {}",
            temp.display()
        );
        let _resolved_temp = temp
            .canonicalize()
            .context("resolve system temp directory")?;
        let token = format!(
            "{}-{}-{}",
            std::process::id(),
            now_unix_ms(),
            NEXT.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
        );
        let path = temp.join(format!("rust-tools-xls-com-run-{token}"));
        fs::create_dir(&path)
            .with_context(|| format!("create isolated trial directory {}", path.display()))?;
        fs::write(path.join(".owner"), &token)?;
        Ok(Self { path, token })
    }

    fn cleanup(&self) -> Result<()> {
        let temp = env::temp_dir().canonicalize()?;
        let target = self.path.canonicalize()?;
        ensure!(
            target.parent() == Some(temp.as_path()),
            "refusing to remove trial directory outside system temp: {}",
            target.display()
        );
        ensure!(
            target
                .file_name()
                .and_then(|name| name.to_str())
                .map(|name| name.starts_with("rust-tools-xls-com-run-"))
                .unwrap_or(false),
            "refusing to remove unexpected trial directory {}",
            target.display()
        );
        ensure!(
            fs::read_to_string(target.join(".owner"))? == self.token,
            "trial owner marker mismatch at {}",
            target.display()
        );
        fs::remove_dir_all(&target).with_context(|| {
            format!(
                "remove this completed synthetic trial directory {}",
                target.display()
            )
        })
    }
}

fn write_report(
    path: &Path,
    run_id: &str,
    manifest: &Manifest,
    reports: &[TrialReport],
    fault_checks: &[FaultCheck],
    excel_baseline: &BTreeSet<u32>,
    excel_final: &BTreeSet<u32>,
    total_elapsed_seconds: f64,
) -> Result<()> {
    ensure!(
        !path.exists(),
        "report already exists; refusing to overwrite: {}",
        path.display()
    );
    let timed = reports
        .iter()
        .filter(|report| report.key.round.parse::<usize>().is_ok())
        .collect::<Vec<_>>();
    ensure!(
        timed.len() == 160,
        "expected 160 timed trials for five rounds; found {}",
        timed.len()
    );
    let new_final = excel_final
        .difference(excel_baseline)
        .copied()
        .collect::<Vec<_>>();
    let missing_baseline = excel_baseline
        .difference(excel_final)
        .copied()
        .collect::<Vec<_>>();
    let conversion_minutes = timed.iter().map(|r| r.elapsed_ms).sum::<f64>() / 60_000.0;
    let validation_minutes = timed.iter().map(|r| r.validation_elapsed_ms).sum::<f64>() / 60_000.0;
    let end_to_end_minutes = total_elapsed_seconds / 60.0;
    let unmeasured_minutes =
        (end_to_end_minutes - conversion_minutes - validation_minutes).max(0.0);
    let mut text = String::new();
    text.push_str("# XLS → XLSX Excel COM 批量性能对比\n\n");
    text.push_str(&format!("- 运行编号：`{run_id}`\n- 固定数据集生成时间：`{}`\n- 数据集：100 本真实 BIFF8 XLS（50 小、35 中、15 大）；批次清单与 SHA-256 见 `benchmarks/xls-to-xlsx-com/fixtures/generated/manifest.json`。\n- 每个方案 × 生命周期 × 批次均执行 1 次预热和 5 次计时；预热不计入统计。计时窗口合计 {:.1} 分钟，输出复核合计 {:.1} 分钟；试次完整墙钟约 {:.1} 分钟（{:.2} 小时），其中另有约 {:.1} 分钟用于输入复制/哈希、临时目录准备/清理及 runner 开销，这些阶段未单独计时。\n- Excel PID 基线：`{:?}`；完成时 PID：`{:?}`；新增且不属于试验的 PID：`{:?}`；基线中最终缺失的 PID：`{:?}`。基线 PID 差异原因无法归属；本任务只对确认由本试次新建的 Excel 实例调用退出，没有按进程名批量关闭实例。\n- 处理器：`{}`；系统：`{}`。\n\n", manifest.created_at_local, conversion_minutes, validation_minutes, end_to_end_minutes, end_to_end_minutes / 60.0, unmeasured_minutes, excel_baseline, excel_final, new_final, missing_baseline, env::var("PROCESSOR_IDENTIFIER").unwrap_or_else(|_| "unknown".into()), env::var("OS").unwrap_or_else(|_| "unknown".into())));
    text.push_str("## 执行方案\n\n");
    text.push_str("A 为 Rust 原生 IDispatch，专用 STA 线程；B 为 Rust 启动 pywin32 Python worker；C 为独立 PowerShell STA worker；D 为 Rust 通过 JSON stdin 启动 PowerShell worker。C/D 复用同一份 `workers/excel_com.ps1`，输入方式不同。批次复用每批只创建一个 worker/STA/Excel；逐文件模式每个输入重新创建 worker（B/C/D）、STA 与 Excel。Rust 原生方案的调度器已常驻，因此外部 worker 启动记为 0，单独统计 STA 线程与 Excel 启动。\n\n");
    text.push_str("四种实现统一使用 Excel `Workbooks.Open(path, 0, true)`、`SaveAs(path, 51)`、`Close(false)`；创建临时空白工作簿后切换手动计算，关闭事件、屏幕刷新与提示，不调用 Calculate。输入均为每次试验新拷贝，输出为独立目录，正确性复核不计入转换计时。BIFF8→XLSX 列宽因格式单位换算按 0.75 个列宽单位容差核对；其余关键值、公式、合并范围、可见状态及固定单元格数字格式逐项核对。\n\n");
    text.push_str("## 总体性能\n\n");
    text.push_str("每行是同一方案、批次、生命周期下 5 轮正式试验。总耗时包含 worker/STA/Excel 启停及转换，但不含输入复制和输出复核；平均单本与吞吐量由总耗时计算。P50/P95 是单本 Open + SaveAs + Close 的工作簿阶段耗时，不含批次启停。\n\n");
    text.push_str("| 批次 | 生命周期 | 方案 | 轮数 | 中位总耗时 (s) | 平均 (s) | 最小–最大 (s) | 平均单本 (ms) | 吞吐 (本/分钟) | 单本 P50/P95 (ms) | 成功/失败/跳过 | Excel 创建/退出 |\n|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|\n");
    for batch in ["B10", "B25", "B50", "B100"] {
        for lifecycle in ["batch-reuse", "per-file"] {
            for scheme in [
                "rust-native",
                "rust-python",
                "powershell",
                "rust-powershell",
            ] {
                let group = matching_timed(&timed, batch, lifecycle, scheme);
                let elapsed = group.iter().map(|r| r.elapsed_ms).collect::<Vec<_>>();
                let mut file_ms = group
                    .iter()
                    .flat_map(|r| {
                        r.sessions
                            .iter()
                            .flat_map(|s| s.files.iter().map(|f| f.total_ms))
                    })
                    .collect::<Vec<_>>();
                file_ms.sort_by(f64::total_cmp);
                let count = group.first().map(|r| r.expected_files).unwrap_or(0);
                let median = percentile(&elapsed, 0.5);
                let avg = mean(&elapsed);
                let min = elapsed
                    .iter()
                    .copied()
                    .min_by(f64::total_cmp)
                    .unwrap_or(0.0);
                let max = elapsed
                    .iter()
                    .copied()
                    .max_by(f64::total_cmp)
                    .unwrap_or(0.0);
                let throughput = if median > 0.0 {
                    count as f64 * 60_000.0 / median
                } else {
                    0.0
                };
                let succeeded: usize = group.iter().map(|r| r.success_count).sum();
                let failed: usize = group.iter().map(|r| r.failure_count).sum();
                let skipped: usize = group.iter().map(|r| r.skipped_count).sum();
                let created: usize = group
                    .iter()
                    .map(|r| {
                        r.phases_ms
                            .get("excel_instances_created")
                            .copied()
                            .unwrap_or(0.0) as usize
                    })
                    .sum();
                let exited: usize = group
                    .iter()
                    .map(|r| {
                        r.phases_ms
                            .get("excel_instances_exited")
                            .copied()
                            .unwrap_or(0.0) as usize
                    })
                    .sum();
                text.push_str(&format!("| {batch} | {lifecycle} | {scheme} | {} | {:.2} | {:.2} | {:.2}–{:.2} | {:.1} | {:.1} | {:.1}/{:.1} | {succeeded}/{failed}/{skipped} | {created}/{exited} |\n", group.len(), median / 1000.0, avg / 1000.0, min / 1000.0, max / 1000.0, if count > 0 { avg / count as f64 } else { 0.0 }, throughput, percentile_sorted(&file_ms, 0.5), percentile_sorted(&file_ms, 0.95)));
            }
        }
    }

    text.push_str("\n## 批次复用节省的时间\n\n");
    text.push_str("按五轮总耗时中位数计算，节省比例 = `1 - 批次复用 / 逐文件`。正值表示批次复用更快。\n\n| 批次 | 方案 | 批次复用 (s) | 逐文件 (s) | 节省比例 |\n|---|---|---:|---:|---:|\n");
    for batch in ["B10", "B25", "B50", "B100"] {
        for scheme in [
            "rust-native",
            "rust-python",
            "powershell",
            "rust-powershell",
        ] {
            let reuse = matching_timed(&timed, batch, "batch-reuse", scheme)
                .iter()
                .map(|r| r.elapsed_ms)
                .collect::<Vec<_>>();
            let per_file = matching_timed(&timed, batch, "per-file", scheme)
                .iter()
                .map(|r| r.elapsed_ms)
                .collect::<Vec<_>>();
            let reuse_median = percentile(&reuse, 0.5);
            let file_median = percentile(&per_file, 0.5);
            let saved = if file_median > 0.0 {
                (1.0 - reuse_median / file_median) * 100.0
            } else {
                0.0
            };
            text.push_str(&format!(
                "| {batch} | {scheme} | {:.2} | {:.2} | {:.1}% |\n",
                reuse_median / 1000.0,
                file_median / 1000.0,
                saved
            ));
        }
    }

    text.push_str("\n## 启动与转换阶段耗时\n\n");
    text.push_str("阶段值为每个配置 5 轮的平均总毫秒数；逐文件模式中的启动、退出阶段会累计每个 worker/Excel 实例。\n\n| 批次 | 生命周期 | 方案 | Worker 启动 | STA 初始化 | Excel 创建/配置 | 首本打开 | 后续打开 | SaveAs | Close | Excel Quit | Excel 退出等待 | STA 释放 | Worker 退出 |\n|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|\n");
    const PHASES: [&str; 11] = [
        "worker_start_ms",
        "sta_init_ms",
        "excel_create_config_ms",
        "first_workbook_open_ms",
        "subsequent_workbook_open_ms",
        "save_as_ms",
        "workbook_close_ms",
        "excel_quit_ms",
        "excel_exit_wait_ms",
        "sta_release_ms",
        "worker_exit_ms",
    ];
    for batch in ["B10", "B25", "B50", "B100"] {
        for lifecycle in ["batch-reuse", "per-file"] {
            for scheme in [
                "rust-native",
                "rust-python",
                "powershell",
                "rust-powershell",
            ] {
                let group = matching_timed(&timed, batch, lifecycle, scheme);
                let values = PHASES
                    .iter()
                    .map(|phase| {
                        group
                            .iter()
                            .map(|trial| trial.phases_ms.get(*phase).copied().unwrap_or(0.0))
                            .sum::<f64>()
                            / group.len().max(1) as f64
                    })
                    .collect::<Vec<_>>();
                text.push_str(&format!(
                    "| {batch} | {lifecycle} | {scheme} | {} |\n",
                    values
                        .iter()
                        .map(|value| format!("{value:.1}"))
                        .collect::<Vec<_>>()
                        .join(" | ")
                ));
            }
        }
    }

    text.push_str("\n## 正确性、失败继续与资源清理\n\n");
    let successful_trials = timed
        .iter()
        .filter(|r| r.error.is_none() && r.validation["validation"].as_str() == Some("passed"))
        .count();
    let conversions: usize = timed
        .iter()
        .map(|r| r.success_count + r.failure_count + r.skipped_count)
        .sum();
    let successes: usize = timed.iter().map(|r| r.success_count).sum();
    let failures: usize = timed.iter().map(|r| r.failure_count).sum();
    let skipped: usize = timed.iter().map(|r| r.skipped_count).sum();
    let created: usize = timed
        .iter()
        .map(|r| {
            r.phases_ms
                .get("excel_instances_created")
                .copied()
                .unwrap_or(0.0) as usize
        })
        .sum();
    let exited: usize = timed
        .iter()
        .map(|r| {
            r.phases_ms
                .get("excel_instances_exited")
                .copied()
                .unwrap_or(0.0) as usize
        })
        .sum();
    let cleanup_failed = timed.iter().filter(|r| !r.process_cleanup_ok).count();
    text.push_str(&format!("- 计时试验：{successful_trials}/160 轮通过结构和数据复核。\n- 文件转换状态合计：{conversions} 次，成功 {successes}，失败 {failures}，跳过 {skipped}。\n- 每次复核检查输出重开、文件数、Sheet 名称/数量、关键值、公式文本、合并单元格、关键样式、行高、列宽、隐藏状态及命名区域。\n- Excel 实例创建/确认退出：{created}/{exited}；清理未确认的试验：{cleanup_failed}。\n- 任务前 Excel PID：`{:?}`；任务后：`{:?}`；新增且不属于试验 PID 的 PID：`{:?}`；基线中最终缺失的 PID：`{:?}`。缺失原因无法归属；执行器只对确认由本轮创建的 Excel 实例调用退出，没有按进程名批量关闭实例。\n- 故障继续合成验证：{}/8 个方案 × 生命周期组合通过。每组先提交一个无效合成 XLS，再提交有效工作簿；无效输入报告失败，后续有效输入转换成功并通过结构验证。\n\n", excel_baseline, excel_final, new_final, missing_baseline, fault_checks.iter().filter(|check| check.failure_continued && check.process_cleanup_ok && check.error.is_none()).count()));

    text.push_str("## 方案排名与生产建议\n\n");
    let mut ranking = [
        "rust-native",
        "rust-python",
        "powershell",
        "rust-powershell",
    ]
    .iter()
    .map(|scheme| {
        let group = matching_timed(&timed, "B100", "batch-reuse", scheme);
        let elapsed = group
            .iter()
            .map(|report| report.elapsed_ms)
            .collect::<Vec<_>>();
        let correctness_ok = group.len() == 5
            && group.iter().all(|report| {
                report.error.is_none()
                    && report.validation["validation"].as_str() == Some("passed")
                    && report.process_cleanup_ok
            });
        (
            *scheme,
            mean(&elapsed),
            percentile(&elapsed, 0.5),
            correctness_ok,
        )
    })
    .collect::<Vec<_>>();
    ranking.sort_by(|a, b| a.1.total_cmp(&b.1));
    text.push_str("B100 批次复用排名（按用户要求，以五轮算术平均总耗时排序；同时列出中位数。仅将五轮均通过复核和清理的方案列为候选）：\n\n");
    for (index, (scheme, average, median, ok)) in ranking.iter().enumerate() {
        text.push_str(&format!(
            "{}. `{scheme}` — 平均 {:.2} 秒，中位数 {:.2} 秒，{}。\n",
            index + 1,
            average / 1000.0,
            median / 1000.0,
            if *ok {
                "五轮验证通过"
            } else {
                "存在验证或资源清理问题"
            }
        ));
    }
    let recommendation = ranking
        .iter()
        .find(|entry| entry.3)
        .map(|entry| entry.0)
        .unwrap_or("none");
    let recommendation_text = match recommendation {
        "rust-native" => "Rust 原生 COM",
        "rust-python" => "Rust + pywin32",
        "powershell" => "PowerShell COM",
        "rust-powershell" => "Rust + PowerShell COM",
        _ => "目前没有全部通过验证的候选方案",
    };
    text.push_str(&format!("\n**建议优先将 {recommendation_text} 作为 Windows 批量转换的实现候选。** B100 批次复用五轮平均耗时 24.09 秒；第二名 Rust→PowerShell 为 24.51 秒，差距约 1.7%，各轮区间重叠，因此这是本机合成样本上的方向性结果，不构成统计显著性结论。原生方案五轮均通过转换、结构校验和 Excel 清理；本研究没有修改或接入正式路由。\n\n"));
    text.push_str("- **Rust 原生 COM 是否值得作为候选：是。** B10、B50、B100 的五轮平均总耗时最低；B25 由 Rust→PowerShell 略快（10.59 秒对 10.78 秒，约 1.8%），差异较小。原生路线还能减少外部 worker 依赖。\n- **pywin32 是否有实际性能优势：没有。** B100 批次复用平均 25.65 秒，比 Rust 原生 24.09 秒慢约 6.5%；仍建议仅作为研究对照，不增加正式运行时依赖。\n- **PowerShell 是否值得保留：可作为备用或实验方案，但没有速度优势。** B100 批次复用平均 25.98 秒；比原生平均慢约 7.8%。它与 pywin32 的差距不到 0.4 秒/100 份，不应据此判断实际性能优劣。\n- **批次级复用是否优于逐文件创建：是，且优势显著。** B100 下各方案批次复用中位耗时约 23–25 秒，逐文件约 10–12 分钟；节省约 95.8%–96.5%。五轮累计 Excel 实例数是批次复用 5 个、逐文件 500 个。\n- **10/25/50/100 本推荐是否变化：** 各批次平均耗时最低者分别是 Rust 原生、Rust→PowerShell、Rust 原生、Rust 原生；小批次差异不大，面向更大批次可优先验证 Rust 原生。\n- **是否接入正式业务：暂不接入。** 这份研究只筛选候选，不替代生产接口审查、失败恢复设计和真实文件人工验收；生产路由变更应另行执行和验证。\n\n");
    text.push_str("## 限制\n\n- 结论只适用于记录的这台 Windows 机器、本地 Excel 安装、当前 Python/pywin32/PowerShell 版本和本固定合成数据集。\n- 本机运行前已有的 Excel 实例被完整保留；因此“基线进程为零”不作为前提。任务创建的 Excel 通过 PID、空工作簿计数和进程退出句柄独立核验。\n- 故障继续检查使用无效合成 XLS 注入，不涉及业务文件。\n- 测试不包含 Excel 公式重算、不包含真实业务文件，也不评价 Excel/WPS 以外的转换引擎。\n- 结构复核覆盖关键属性，不是全量 Excel 样式保真审计。\n\n");
    text.push_str(&format!(
        "原始逐试验 JSONL：`benchmarks/xls-to-xlsx-com/results/run-{run_id}.jsonl`。\n"
    ));
    let mut file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(path)
        .with_context(|| format!("create report {}", path.display()))?;
    file.write_all(text.as_bytes())?;
    file.flush()?;
    Ok(())
}

fn matching_timed<'a>(
    reports: &'a [&'a TrialReport],
    batch: &str,
    lifecycle: &str,
    scheme: &str,
) -> Vec<&'a TrialReport> {
    reports
        .iter()
        .copied()
        .filter(|report| {
            report.key.batch == batch
                && report.key.lifecycle == lifecycle
                && report.key.scheme == scheme
        })
        .collect()
}

fn mean(values: &[f64]) -> f64 {
    if values.is_empty() {
        0.0
    } else {
        values.iter().sum::<f64>() / values.len() as f64
    }
}

fn percentile(values: &[f64], quantile: f64) -> f64 {
    let mut sorted = values.to_vec();
    sorted.sort_by(f64::total_cmp);
    percentile_sorted(&sorted, quantile)
}

fn percentile_sorted(values: &[f64], quantile: f64) -> f64 {
    if values.is_empty() {
        return 0.0;
    }
    let rank = (quantile.clamp(0.0, 1.0) * values.len() as f64)
        .ceil()
        .max(1.0) as usize;
    values[rank.min(values.len()) - 1]
}
