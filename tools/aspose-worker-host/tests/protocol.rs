use aspose_worker_host::*;
use std::{
    fs,
    path::PathBuf,
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};

fn setup(mode: &str) -> (AsposeWorkerBackend, WorkbookRequest, PathBuf) {
    let root = std::env::var("ASPOSE_TEST_ROOT")
        .map(PathBuf::from)
        .unwrap_or_else(|_| std::env::temp_dir().join("rusttools-aspose-worker-protocol"))
        .join(format!(
            "{mode}-{}",
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    fs::create_dir_all(&root).unwrap();
    fs::write(root.join("source.xls"), b"synthetic protocol stub input").unwrap();
    let source = fs::canonicalize(root.join("source.xls")).unwrap();
    let root = fs::canonicalize(root).unwrap();
    let engine = AsposeWorkerBackend::new(WorkerConfig {
        executable: std::env::var("ASPOSE_TEST_PYTHON")
            .unwrap_or_else(|_| {
                if cfg!(windows) {
                    "py".into()
                } else {
                    "python3".into()
                }
            })
            .into(),
        arguments: vec![
            PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("tests/fake_worker.py")
                .display()
                .to_string(),
            mode.into(),
            root.join("requests.txt").display().to_string(),
        ],
        diagnostics_dir: root.join("logs"),
        startup_timeout: Duration::from_millis(500),
        request_timeout: Duration::from_millis(200),
    });
    (
        engine,
        WorkbookRequest {
            run_id: "protocol-regression".into(),
            input: source,
            candidate: root.join("candidate.xlsx"),
        },
        root,
    )
}
#[test]
fn resident_health_and_request() {
    let (mut e, q, _) = setup("normal");
    let pid = e.health_check().unwrap();
    let out = e.calculate(&q).unwrap();
    assert_eq!(pid, out.worker_pid);
    assert!(out.evaluation_output);
    assert_eq!(e.generations(), 1);
    assert_eq!(
        fs::read(&q.input).unwrap(),
        b"synthetic protocol stub input"
    );
}
#[test]
fn target_conflict_preserves_both_files_without_starting_worker() {
    let (mut e, q, _) = setup("normal");
    fs::write(&q.candidate, b"existing target").unwrap();
    assert!(e.calculate(&q).is_err());
    assert_eq!(e.generations(), 0);
    assert_eq!(fs::read(&q.candidate).unwrap(), b"existing target");
    assert_eq!(
        fs::read(&q.input).unwrap(),
        b"synthetic protocol stub input"
    );
}
#[test]
fn nonfinite_number_rejected_before_serialization() {
    let (mut e, q, _) = setup("normal");
    assert!(e
        .modify(&ModifyRequest {
            workbook: q,
            operations: vec![WorkbookOperation::SetValue {
                sheet: "Data".into(),
                cell: "A1".into(),
                value: CellValue::Number(f64::NAN)
            }]
        })
        .is_err());
    assert_eq!(e.generations(), 0);
}
#[test]
fn response_identity_error_is_not_retried() {
    failure("wrong_id");
}
#[test]
fn corrupt_candidate_is_not_accepted() {
    failure("bad_candidate");
}
#[test]
fn crash_partial_candidate_is_not_accepted_or_retried() {
    failure("crash");
}
fn failure(mode: &str) {
    let (mut e, q, root) = setup(mode);
    assert!(e.calculate(&q).is_err());
    assert_eq!(
        fs::read_to_string(root.join("requests.txt")).unwrap(),
        "Calculate\n"
    );
    assert_eq!(
        fs::read(&q.input).unwrap(),
        b"synthetic protocol stub input"
    );
}
#[test]
fn source_snapshot_change_rejects_candidate() {
    let (mut e, q, _) = setup("source_changed");
    assert!(e.calculate(&q).unwrap_err().0.contains("Source changed"));
}
#[test]
fn startup_timeout_reaps_owned_child() {
    let (mut e, _, _) = setup("startup_timeout");
    let start = Instant::now();
    assert!(e.health_check().is_err());
    assert!(start.elapsed() < Duration::from_secs(4));
}
#[test]
fn request_timeout_does_not_replay() {
    let (mut e, q, root) = setup("request_timeout");
    let start = Instant::now();
    assert!(e.calculate(&q).is_err());
    assert!(start.elapsed() < Duration::from_secs(4));
    assert_eq!(
        fs::read_to_string(root.join("requests.txt")).unwrap(),
        "Calculate\n"
    );
}
#[test]
fn blocked_large_write_also_has_timeout() {
    let (mut e, q, _) = setup("write_timeout");
    let start = Instant::now();
    let req = ModifyRequest {
        workbook: q,
        operations: vec![WorkbookOperation::SetValue {
            sheet: "Data".into(),
            cell: "A1".into(),
            value: CellValue::Text("x".repeat(700_000)),
        }],
    };
    assert!(e.modify(&req).is_err());
    assert!(start.elapsed() < Duration::from_secs(4));
}
#[test]
fn draining_response_is_delivered_before_next_worker_starts() {
    let (mut e, q, root) = setup("drain");
    let a = e.calculate(&q).unwrap();
    let next = WorkbookRequest {
        run_id: q.run_id,
        input: q.input,
        candidate: root.join("next.xlsx"),
    };
    let b = e.calculate(&next).unwrap();
    assert!(a.draining && b.draining);
    assert_eq!(e.generations(), 2);
    assert_ne!(a.worker_pid, b.worker_pid);
}
