//! Extract DDE-style active links into a dedicated report.

use anyhow::Result;
use docir_app::{LinkExtractionReport, ParserConfig};
use std::path::PathBuf;

use crate::commands::util::{
    build_app_and_parse, push_bullet_line, push_labeled_line, run_dual_output,
};

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    json: bool,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let (app, parsed) = build_app_and_parse(&input, parser_config)?;
    let report = app.build_link_extraction_report(&parsed);
    run_dual_output(&report, "report", json, pretty, output, format_report_text)
}

fn format_report_text(report: &LinkExtractionReport) -> String {
    let mut out = String::new();
    push_labeled_line(&mut out, 0, "Format", &report.document_format);
    push_labeled_line(&mut out, 0, "Links", report.link_count);
    if !report.links.is_empty() {
        out.push_str("\nExtracted Links:\n");
        for link in &report.links {
            push_bullet_line(
                &mut out,
                2,
                &link.kind,
                format!("{} [{}]", link.normalized, link.risk),
            );
            push_labeled_line(&mut out, 4, "Raw", &link.raw_text);
            if let Some(location) = link.location.as_deref() {
                push_labeled_line(&mut out, 4, "Location", location);
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::{format_report_text, run};
    use crate::test_support;
    use docir_app::{LinkArtifact, LinkExtractionReport, ParserConfig};
    use docir_core::security::ThreatLevel;
    use std::fs;
    use std::io::Write;
    use zip::write::SimpleFileOptions;
    use zip::CompressionMethod;
    use zip::ZipWriter;

    fn build_test_ods_with_dde() -> Vec<u8> {
        let mut cursor = std::io::Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(&mut cursor);
        let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);

        zip.start_file("mimetype", stored).expect("mimetype");
        zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")
            .expect("mimetype body");

        zip.start_file("META-INF/manifest.xml", stored)
            .expect("manifest");
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>"#,
        )
        .expect("manifest body");

        zip.start_file("content.xml", stored).expect("content");
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell table:formula="of:=DDEAUTO(&quot;cmd&quot;;&quot;/c calc&quot;;&quot;A1&quot;)" office:value-type="string">
            <text:p>dde</text:p>
          </table:table-cell>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#,
        )
        .expect("content body");

        zip.finish().expect("finish zip");
        cursor.into_inner()
    }

    #[test]
    fn format_report_text_includes_normalized_and_raw_fields() {
        let report = LinkExtractionReport {
            document_format: "docx".to_string(),
            link_count: 1,
            links: vec![LinkArtifact {
                kind: "ddeauto".to_string(),
                risk: ThreatLevel::High,
                raw_text: r#"DDEAUTO "cmd" "/c calc" "A1""#.to_string(),
                normalized: "DDEAUTO app=cmd topic=/c calc item=A1".to_string(),
                location: Some("word/document.xml".to_string()),
                application: Some("cmd".to_string()),
                topic: Some("/c calc".to_string()),
                item: Some("A1".to_string()),
            }],
        };

        let text = format_report_text(&report);
        assert!(text.contains("Links: 1"));
        assert!(text.contains("- ddeauto: DDEAUTO app=cmd topic=/c calc item=A1 [HIGH]"));
        assert!(text.contains(r#"Raw: DDEAUTO "cmd" "/c calc" "A1""#));
        assert!(text.contains("Location: word/document.xml"));
    }

    #[test]
    fn extract_links_run_writes_json_for_odf_dde_fixture() {
        let input = test_support::temp_file("dde_fixture", "ods");
        let output = test_support::temp_file("dde", "json");
        fs::write(&input, build_test_ods_with_dde()).expect("fixture");

        run(
            input.clone(),
            true,
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("extract-links json");

        let text = fs::read_to_string(&output).expect("output");
        assert!(text.contains("\"report\""));
        assert!(text.contains("\"kind\": \"ddeauto\""));
        assert!(text.contains("\"normalized\""));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }
}
