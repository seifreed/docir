use serde_json::Value;
use std::fs;
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

fn bin_path() -> &'static str {
    env!("CARGO_BIN_EXE_docir")
}

#[test]
fn parse_command_writes_json_output_file() {
    let root = repo_root();
    let fixture = root.join("fixtures/ooxml/minimal.docx");
    let tmp = tempfile::tempdir().expect("tempdir");
    let out = tmp.path().join("parse.json");

    let output = Command::new(bin_path())
        .arg("parse")
        .arg(&fixture)
        .arg("--format")
        .arg("json")
        .arg("--output")
        .arg(&out)
        .output()
        .expect("run parse command");
    assert!(
        output.status.success(),
        "parse failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(out.exists(), "parse output file should exist");

    let json = fs::read_to_string(&out).expect("read parse output");
    let parsed: Value = serde_json::from_str(&json).expect("parse json output");
    assert_eq!(parsed.get("type").and_then(Value::as_str), Some("Document"));
    assert!(parsed.get("children").is_some());
}

#[test]
fn query_command_writes_matches_json_with_filters() {
    let root = repo_root();
    let fixture = root.join("fixtures/ooxml/minimal.docx");
    let tmp = tempfile::tempdir().expect("tempdir");
    let out = tmp.path().join("query.json");

    let output = Command::new(bin_path())
        .arg("query")
        .arg(&fixture)
        .arg("--format")
        .arg("docx")
        .arg("--output")
        .arg(&out)
        .arg("--pretty")
        .output()
        .expect("run query command");
    assert!(
        output.status.success(),
        "query failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(out.exists(), "query output file should exist");

    let json = fs::read_to_string(&out).expect("read query output");
    let parsed: Value = serde_json::from_str(&json).expect("query json output");
    let matches = parsed
        .get("matches")
        .and_then(Value::as_array)
        .expect("matches array");
    assert!(
        !matches.is_empty(),
        "query should return at least one match"
    );
    assert!(matches[0].get("node_id").is_some());
    assert!(matches[0].get("node_type").is_some());
}

#[test]
fn grep_command_writes_matches_json() {
    let root = repo_root();
    let fixture = root.join("fixtures/ooxml/minimal.docx");
    let tmp = tempfile::tempdir().expect("tempdir");
    let out = tmp.path().join("grep.json");

    let output = Command::new(bin_path())
        .arg("grep")
        .arg(&fixture)
        .arg("Hello")
        .arg("--node-type")
        .arg("Paragraph")
        .arg("--format")
        .arg("docx")
        .arg("--output")
        .arg(&out)
        .arg("--pretty")
        .output()
        .expect("run grep command");
    assert!(
        output.status.success(),
        "grep failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(out.exists(), "grep output file should exist");

    let json = fs::read_to_string(&out).expect("read grep output");
    let parsed: Value = serde_json::from_str(&json).expect("grep json output");
    assert!(
        parsed.get("matches").and_then(Value::as_array).is_some(),
        "grep should emit matches array"
    );
}

#[test]
fn extract_command_by_type_writes_nodes_json() {
    let root = repo_root();
    let fixture = root.join("fixtures/ooxml/minimal.docx");
    let tmp = tempfile::tempdir().expect("tempdir");
    let out = tmp.path().join("extract.json");

    let output = Command::new(bin_path())
        .arg("extract")
        .arg(&fixture)
        .arg("--node-type")
        .arg("Paragraph")
        .arg("--output")
        .arg(&out)
        .arg("--pretty")
        .output()
        .expect("run extract command");
    assert!(
        output.status.success(),
        "extract failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(out.exists(), "extract output file should exist");

    let json = fs::read_to_string(&out).expect("read extract output");
    let parsed: Value = serde_json::from_str(&json).expect("extract json output");
    let nodes = parsed
        .get("nodes")
        .and_then(Value::as_array)
        .expect("nodes array");
    assert!(
        !nodes.is_empty(),
        "extract should include at least one node"
    );
}

#[test]
fn dump_node_command_outputs_root_document_node() {
    let root = repo_root();
    let fixture = root.join("fixtures/ooxml/minimal.docx");

    let output = Command::new(bin_path())
        .arg("dump-node")
        .arg(&fixture)
        .arg("--node-id")
        .arg("1")
        .arg("--format")
        .arg("json")
        .output()
        .expect("run dump-node command");
    assert!(
        output.status.success(),
        "dump-node failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    let parsed: Value = serde_json::from_str(&stdout).expect("dump-node json output");
    assert_eq!(parsed.get("type").and_then(Value::as_str), Some("Document"));
}
