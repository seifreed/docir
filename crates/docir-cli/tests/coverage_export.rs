use std::fs;
use std::path::PathBuf;
use std::process::Command;

#[derive(serde::Deserialize)]
struct Manifest {
    fixtures: Vec<FixtureEntry>,
}

#[derive(serde::Deserialize)]
struct FixtureEntry {
    path: String,
}

#[test]
fn coverage_exports_json_and_csv() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf();
    let manifest_path = root.join("fixtures/manifest.json");
    let data = fs::read_to_string(&manifest_path).expect("read manifest");
    let manifest: Manifest = serde_json::from_str(&data).expect("parse manifest");

    let bin = env!("CARGO_BIN_EXE_docir");
    let tmp = tempfile::tempdir().expect("tempdir");

    for fixture in manifest.fixtures {
        let path = root.join(&fixture.path);
        let json_out = tmp.path().join("coverage.json");
        let csv_out = tmp.path().join("coverage.csv");

        let status = Command::new(bin)
            .arg("coverage")
            .arg(&path)
            .arg("--export")
            .arg(&json_out)
            .arg("--export-format")
            .arg("json")
            .status()
            .expect("run coverage json");
        assert!(status.success(), "coverage json failed for {:?}", path);
        assert!(json_out.exists());
        let json_data = fs::read_to_string(&json_out).expect("read json export");
        let _: serde_json::Value = serde_json::from_str(&json_data).expect("json export parses");

        let status = Command::new(bin)
            .arg("coverage")
            .arg(&path)
            .arg("--export")
            .arg(&csv_out)
            .arg("--export-format")
            .arg("csv")
            .status()
            .expect("run coverage csv");
        assert!(status.success(), "coverage csv failed for {:?}", path);
        assert!(csv_out.exists());
        let csv_data = fs::read_to_string(&csv_out).expect("read csv export");
        let mut first_line = csv_data.lines().next().unwrap_or("").to_string();
        if let Some(stripped) = first_line.strip_prefix('\u{feff}') {
            first_line = stripped.to_string();
        }
        first_line = first_line.trim_end_matches('\r').to_string();
        assert_eq!(
            first_line,
            "section,code,severity,message,path",
            "unexpected CSV header for {:?} (prefix={:?})",
            path,
            &csv_data.chars().take(64).collect::<String>()
        );
    }
}

#[test]
fn coverage_exports_parts_only() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf();
    let manifest_path = root.join("fixtures/manifest.json");
    let data = fs::read_to_string(&manifest_path).expect("read manifest");
    let manifest: Manifest = serde_json::from_str(&data).expect("parse manifest");

    let bin = env!("CARGO_BIN_EXE_docir");
    let tmp = tempfile::tempdir().expect("tempdir");

    let fixture = manifest.fixtures.get(0).expect("at least one fixture");
    let path = root.join(&fixture.path);
    let json_out = tmp.path().join("coverage_parts.json");
    let csv_out = tmp.path().join("coverage_parts.csv");

    let status = Command::new(bin)
        .arg("coverage")
        .arg(&path)
        .arg("--export")
        .arg(&json_out)
        .arg("--export-format")
        .arg("json")
        .arg("--export-mode")
        .arg("parts")
        .status()
        .expect("run coverage json parts");
    assert!(
        status.success(),
        "coverage json parts failed for {:?}",
        path
    );
    assert!(json_out.exists());
    let json_data = fs::read_to_string(&json_out).expect("read json export");
    let parts: serde_json::Value = serde_json::from_str(&json_data).expect("json export parses");
    assert!(parts.is_array(), "parts export should be a JSON array");

    let status = Command::new(bin)
        .arg("coverage")
        .arg(&path)
        .arg("--export")
        .arg(&csv_out)
        .arg("--export-format")
        .arg("csv")
        .arg("--export-mode")
        .arg("parts")
        .status()
        .expect("run coverage csv parts");
    assert!(status.success(), "coverage csv parts failed for {:?}", path);
    assert!(csv_out.exists());
    let csv_data = fs::read_to_string(&csv_out).expect("read csv export");
    let mut first_line = csv_data.lines().next().unwrap_or("").to_string();
    if let Some(stripped) = first_line.strip_prefix('\u{feff}') {
        first_line = stripped.to_string();
    }
    first_line = first_line.trim_end_matches('\r').to_string();
    assert_eq!(
        first_line,
        "status,path,content_type,parser",
        "unexpected parts CSV header for {:?} (prefix={:?})",
        path,
        &csv_data.chars().take(64).collect::<String>()
    );
}
