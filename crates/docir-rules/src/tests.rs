use super::*;
use docir_parser::DocumentParser;
use std::io::{Cursor, Write};
use zip::write::FileOptions;

fn build_odf_zip(content_xml: &str, manifest_xml: &str, extra_files: &[(&str, &[u8])]) -> Vec<u8> {
    let mut buffer = Vec::new();
    let cursor = Cursor::new(&mut buffer);
    let mut zip = zip::ZipWriter::new(cursor);
    let stored = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(b"application/vnd.oasis.opendocument.text")
        .unwrap();

    zip.start_file("META-INF/manifest.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(manifest_xml.as_bytes()).unwrap();

    zip.start_file("content.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(content_xml.as_bytes()).unwrap();

    zip.start_file("meta.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
  <office:meta>
    <dc:title>Rules</dc:title>
  </office:meta>
</office:document-meta>
"#,
    )
    .unwrap();

    for (path, bytes) in extra_files {
        zip.start_file(*path, FileOptions::<()>::default()).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap();
    buffer
}

#[test]
fn test_rule_engine_basic() {
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p>Safe</text:p>
      <draw:object-ole xlink:href="https://example.com/ole.bin" />
    </office:text>
  </office:body>
  <office:scripts>
    <script:script xlink:href="Scripts/macro.py" />
  </office:scripts>
</office:document-content>
"#;
    let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Scripts/macro.py" manifest:media-type="application/vnd.sun.star.script"/>
  <manifest:file-entry manifest:full-path="Object 1" manifest:media-type="application/vnd.sun.star.oleobject"/>
</manifest:manifest>
"#;
    let zip_data = build_odf_zip(
        content_xml,
        manifest_xml,
        &[("Scripts/macro.py", b"print('hi')"), ("Object 1", b"ole")],
    );

    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    let engine = RuleEngine::with_default_rules();
    let report = engine.run(&parsed.store, parsed.root_id);

    assert!(report.findings.iter().any(|f| f.rule_id == "SEC-001"));
    assert!(report.findings.iter().any(|f| f.rule_id == "SEC-003"));
}

#[test]
fn test_external_reference_rule() {
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p><text:a xlink:href="https://example.com">Link</text:a></text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
    let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#;
    let zip_data = build_odf_zip(content_xml, manifest_xml, &[]);
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let engine = RuleEngine::with_default_rules();
    let report = engine.run(&parsed.store, parsed.root_id);

    assert!(report.findings.iter().any(|f| f.rule_id == "SEC-003"));
}

#[test]
fn test_rule_profile_overrides() {
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <draw:object-ole xlink:href="file:///tmp/ole.bin" />
    </office:text>
  </office:body>
</office:document-content>
"#;
    let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Object 1" manifest:media-type="application/vnd.sun.star.oleobject"/>
</manifest:manifest>
"#;
    let zip_data = build_odf_zip(content_xml, manifest_xml, &[("Object 1", b"ole")]);
    let parser = DocumentParser::new();
    let parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();

    let engine = RuleEngine::with_default_rules();
    let mut profile = RuleProfile::default();
    profile.disabled_rules.push("SEC-002".to_string());
    profile
        .severity_overrides
        .insert("SEC-010".to_string(), Severity::Critical);
    profile.thresholds.max_ole_objects = Some(1);

    let report = engine.run_with_profile(&parsed.store, parsed.root_id, &profile);
    assert!(!report.findings.iter().any(|f| f.rule_id == "SEC-002"));
    assert!(report.findings.iter().any(|f| f.rule_id == "SEC-010"));
    assert!(report
        .findings
        .iter()
        .any(|f| f.rule_id == "SEC-010" && f.severity == Severity::Critical));
}
