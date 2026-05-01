use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

#[derive(serde::Deserialize)]
struct Manifest {
    fixtures: Vec<FixtureEntry>,
}

#[derive(serde::Deserialize)]
struct FixtureEntry {
    path: String,
}

fn fixture_path_by_suffix(root: &Path, manifest: &Manifest, suffix: &str) -> PathBuf {
    manifest
        .fixtures
        .iter()
        .find(|f| f.path.ends_with(suffix))
        .map(|f| root.join(&f.path))
        .unwrap_or_else(|| panic!("missing fixture with suffix {}", suffix))
}

#[test]
fn diff_exports_json() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf();
    let manifest_path = root.join("fixtures/manifest.json");
    let data = fs::read_to_string(&manifest_path).expect("read manifest");
    let manifest: Manifest = serde_json::from_str(&data).expect("parse manifest");

    let left = fixture_path_by_suffix(&root, &manifest, "minimal.docx");
    let right = fixture_path_by_suffix(&root, &manifest, "rich.docx");

    let bin = env!("CARGO_BIN_EXE_docir");
    let tmp = tempfile::tempdir().expect("tempdir");
    let diff_out = tmp.path().join("diff.json");

    let status = Command::new(bin)
        .arg("diff")
        .arg(&left)
        .arg(&right)
        .arg("--pretty")
        .arg("--output")
        .arg(&diff_out)
        .status()
        .expect("run diff");
    assert!(
        status.success(),
        "diff failed for {:?} -> {:?}",
        left,
        right
    );
    assert!(diff_out.exists());
    let diff_data = fs::read_to_string(&diff_out).expect("read diff");
    let value: serde_json::Value = serde_json::from_str(&diff_data).expect("diff json parses");
    assert!(value.get("added").is_some());
    assert!(value.get("removed").is_some());
    assert!(value.get("modified").is_some());
}

#[test]
fn rules_exports_json() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf();
    let manifest_path = root.join("fixtures/manifest.json");
    let data = fs::read_to_string(&manifest_path).expect("read manifest");
    let manifest: Manifest = serde_json::from_str(&data).expect("parse manifest");

    let input = fixture_path_by_suffix(&root, &manifest, "rich.docx");

    let bin = env!("CARGO_BIN_EXE_docir");
    let tmp = tempfile::tempdir().expect("tempdir");
    let rules_out = tmp.path().join("rules.json");

    let status = Command::new(bin)
        .arg("rules")
        .arg(&input)
        .arg("--pretty")
        .arg("--output")
        .arg(&rules_out)
        .status()
        .expect("run rules");
    assert!(status.success(), "rules failed for {:?}", input);
    assert!(rules_out.exists());
    let rules_data = fs::read_to_string(&rules_out).expect("read rules");
    let value: serde_json::Value = serde_json::from_str(&rules_data).expect("rules json parses");
    assert!(value.get("findings").is_some());
}

#[test]
fn rules_accepts_profile_json_file() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf();
    let manifest_path = root.join("fixtures/manifest.json");
    let data = fs::read_to_string(&manifest_path).expect("read manifest");
    let manifest: Manifest = serde_json::from_str(&data).expect("parse manifest");

    let input = fixture_path_by_suffix(&root, &manifest, "rich.docx");
    let bin = env!("CARGO_BIN_EXE_docir");
    let tmp = tempfile::tempdir().expect("tempdir");
    let rules_out = tmp.path().join("rules_with_profile.json");
    let profile_path = tmp.path().join("profile.json");
    fs::write(
        &profile_path,
        r#"{
  "enabled_rules": ["many_external_references"],
  "disabled_rules": [],
  "severity_overrides": {
    "many_external_references": "High"
  },
  "threshold_overrides": {
    "many_external_references": 1
  }
}"#,
    )
    .expect("write profile json");

    let status = Command::new(bin)
        .arg("rules")
        .arg(&input)
        .arg("--profile")
        .arg(&profile_path)
        .arg("--output")
        .arg(&rules_out)
        .status()
        .expect("run rules with profile");
    assert!(
        status.success(),
        "rules with profile failed for {:?}",
        input
    );
    assert!(rules_out.exists());

    let rules_data = fs::read_to_string(&rules_out).expect("read rules output");
    let value: serde_json::Value = serde_json::from_str(&rules_data).expect("rules json parses");
    assert!(value.get("findings").is_some());
}
