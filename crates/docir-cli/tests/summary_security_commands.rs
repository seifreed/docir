use serde_json::Value;
use std::path::PathBuf;
use std::process::Command;

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf()
}

#[test]
fn summary_command_prints_expected_sections() {
    let root = repo_root();
    let fixture = root.join("fixtures/ooxml/minimal.docx");
    let bin = env!("CARGO_BIN_EXE_docir");

    let output = Command::new(bin)
        .arg("summary")
        .arg(&fixture)
        .output()
        .expect("run summary command");
    assert!(
        output.status.success(),
        "summary command failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Document Summary"));
    assert!(stdout.contains("Format:"));
    assert!(stdout.contains("Structure:"));
    assert!(stdout.contains("Text Statistics:"));
    assert!(stdout.contains("Security:"));
}

#[test]
fn summary_with_metrics_prints_metrics_section() {
    let root = repo_root();
    let fixture = root.join("fixtures/ooxml/rich.docx");
    let bin = env!("CARGO_BIN_EXE_docir");

    let output = Command::new(bin)
        .arg("--metrics")
        .arg("summary")
        .arg(&fixture)
        .output()
        .expect("run summary --metrics command");
    assert!(
        output.status.success(),
        "summary --metrics command failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Parse Metrics (ms):"));
    assert!(stdout.contains("Content Types:"));
    assert!(stdout.contains("Normalization:"));
}

#[test]
fn security_command_json_mode_outputs_machine_report() {
    let root = repo_root();
    let fixture = root.join("fixtures/ooxml/minimal.xlsx");
    let bin = env!("CARGO_BIN_EXE_docir");

    let output = Command::new(bin)
        .arg("security")
        .arg(&fixture)
        .arg("--json")
        .output()
        .expect("run security --json command");
    assert!(
        output.status.success(),
        "security json command failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    let parsed: Value = serde_json::from_str(&stdout).expect("security json output should parse");
    assert_eq!(
        parsed.get("file").and_then(Value::as_str),
        Some(fixture.to_string_lossy().as_ref())
    );
    assert!(parsed.get("threat_level").is_some());
    assert!(parsed.get("findings_count").is_some());
    assert!(parsed.get("findings").and_then(Value::as_array).is_some());
}

#[test]
fn security_command_human_mode_prints_report() {
    let root = repo_root();
    let fixture = root.join("fixtures/ooxml/minimal.pptx");
    let bin = env!("CARGO_BIN_EXE_docir");

    let output = Command::new(bin)
        .arg("security")
        .arg(&fixture)
        .arg("--verbose")
        .output()
        .expect("run security command");
    assert!(
        output.status.success(),
        "security command failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Security Analysis Report"));
    assert!(stdout.contains("Threat Level:"));
    assert!(stdout.contains("Security Features Detected:"));
}

#[test]
fn security_verbose_with_findings_prints_findings_and_recommendation() {
    let root = repo_root();
    let fixture = root.join("fixtures/rtf/rich.rtf");
    let bin = env!("CARGO_BIN_EXE_docir");

    let output = Command::new(bin)
        .arg("security")
        .arg(&fixture)
        .arg("--verbose")
        .output()
        .expect("run security --verbose command for rich rtf");
    assert!(
        output.status.success(),
        "security --verbose command failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Findings ("));
    assert!(
        stdout.contains("RECOMMENDATION: This document contains potentially dangerous content.")
    );
    assert!(stdout.contains("Location:"));
}
