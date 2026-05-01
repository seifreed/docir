use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

#[derive(serde::Deserialize)]
struct CoverageReport {
    summary: Option<String>,
    counts: Option<String>,
    parts: Vec<CoverageEntry>,
    part_rows: Vec<CoverageEntry>,
    unknown: Vec<CoverageEntry>,
}

#[derive(serde::Deserialize)]
struct CoverageEntry {
    severity: String,
    code: String,
    message: String,
}

#[derive(serde::Deserialize)]
struct PartRow {
    status: String,
    path: String,
    content_type: String,
    parser: String,
}

#[derive(serde::Deserialize)]
struct Manifest {
    fixtures: Vec<FixtureEntry>,
}

#[derive(serde::Deserialize)]
struct FixtureEntry {
    path: String,
}

fn trim_csv_header(raw: &str) -> String {
    let mut first_line = raw.lines().next().unwrap_or("").to_string();
    if let Some(stripped) = first_line.strip_prefix('\u{feff}') {
        first_line = stripped.to_string();
    }
    first_line.trim_end_matches('\r').to_string()
}

fn is_ooxml_fixture(path: &Path) -> bool {
    matches!(
        path.extension().and_then(|ext| ext.to_str()),
        Some("docx" | "xlsx" | "pptx" | "xlsb")
    )
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
        let report: CoverageReport = serde_json::from_str(&json_data).expect("json export parses");
        if is_ooxml_fixture(&path) {
            assert!(
                report
                    .summary
                    .as_deref()
                    .unwrap_or("")
                    .contains("coverage summary:"),
                "coverage summary should be present for {:?}",
                path
            );
            assert!(
                report.counts.as_deref().unwrap_or("").contains("counts:"),
                "coverage counts should be present for {:?}",
                path
            );
            assert!(
                !report.parts.is_empty(),
                "coverage export should include part diagnostics for {:?}",
                path
            );
            assert!(
                !report.part_rows.is_empty(),
                "coverage export should include tabular part rows for {:?}",
                path
            );
            assert!(
                report
                    .parts
                    .iter()
                    .any(|entry| entry.code == "COVERAGE_PART"),
                "coverage export should include COVERAGE_PART rows for {:?}",
                path
            );
            assert!(
                report.parts.iter().all(|entry| !entry.message.is_empty()),
                "coverage part entries should include messages for {:?}",
                path
            );
            assert!(
                report
                    .parts
                    .iter()
                    .all(|entry| matches!(entry.severity.as_str(), "Info" | "Warning" | "Error")),
                "coverage part severities should be valid enum values for {:?}",
                path
            );
        } else {
            assert!(
                report.summary.is_none() && report.counts.is_none(),
                "non-OOXML fixtures should not emit OOXML coverage summary diagnostics: {:?}",
                path
            );
            assert!(
                report.parts.is_empty() && report.part_rows.is_empty(),
                "non-OOXML fixtures should not emit OOXML coverage part diagnostics: {:?}",
                path
            );
            assert!(
                report.unknown.is_empty(),
                "non-OOXML fixtures should not emit unknown content-type diagnostics: {:?}",
                path
            );
        }
        if is_ooxml_fixture(&path)
            && path
                .file_name()
                .and_then(|name| name.to_str())
                .map(|name| name.starts_with("rich."))
                .unwrap_or(false)
        {
            assert!(
                report
                    .parts
                    .iter()
                    .all(|entry| entry.severity.as_str() == "Info"),
                "rich fixtures should not report pending coverage parts for {:?}",
                path
            );
        }

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
        let first_line = trim_csv_header(&csv_data);
        assert_eq!(
            first_line,
            "section,code,severity,message,path",
            "unexpected CSV header for {:?} (prefix={:?})",
            path,
            &csv_data.chars().take(64).collect::<String>()
        );
        if is_ooxml_fixture(&path) {
            assert!(
                csv_data.contains("summary,COVERAGE_SUMMARY,Info,"),
                "CSV export should include COVERAGE_SUMMARY row for {:?}",
                path
            );
        } else {
            assert!(
                !csv_data.contains("summary,COVERAGE_SUMMARY,Info,"),
                "non-OOXML fixtures should not export COVERAGE_SUMMARY rows: {:?}",
                path
            );
        }
        let data_line_count = csv_data
            .lines()
            .skip(1)
            .filter(|line| !line.is_empty())
            .count();
        if is_ooxml_fixture(&path) {
            assert!(
                data_line_count >= report.parts.len() + report.part_rows.len(),
                "CSV export should include parts and part_rows rows for {:?}",
                path
            );
        } else {
            assert!(
                !csv_data.contains("parts,COVERAGE_PART,"),
                "non-OOXML fixtures should not export COVERAGE_PART rows: {:?}",
                path
            );
            assert!(
                !csv_data.contains("part_rows,COVERAGE_PART_ROW,"),
                "non-OOXML fixtures should not export COVERAGE_PART_ROW rows: {:?}",
                path
            );
        }
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

    let fixture = manifest
        .fixtures
        .iter()
        .find(|entry| {
            matches!(
                PathBuf::from(&entry.path)
                    .extension()
                    .and_then(|ext| ext.to_str()),
                Some("docx" | "xlsx" | "pptx" | "xlsb")
            )
        })
        .expect("manifest should contain at least one OOXML fixture");
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
    let parts: Vec<PartRow> = serde_json::from_str(&json_data).expect("json export parses");
    assert!(!parts.is_empty(), "parts export should include rows");
    assert!(
        parts
            .iter()
            .all(|row| matches!(row.status.as_str(), "complete" | "pending")),
        "parts export should only include complete/pending statuses"
    );
    assert!(
        parts.iter().all(|row| !row.path.is_empty()
            && !row.content_type.is_empty()
            && !row.parser.is_empty()),
        "parts export rows must include path/content_type/parser values"
    );

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
    let first_line = trim_csv_header(&csv_data);
    assert_eq!(
        first_line,
        "status,path,content_type,parser",
        "unexpected parts CSV header for {:?} (prefix={:?})",
        path,
        &csv_data.chars().take(64).collect::<String>()
    );
    let csv_rows = csv_data
        .lines()
        .skip(1)
        .filter(|line| !line.is_empty())
        .count();
    assert_eq!(
        csv_rows,
        parts.len(),
        "parts CSV row count should match JSON parts row count"
    );
}
