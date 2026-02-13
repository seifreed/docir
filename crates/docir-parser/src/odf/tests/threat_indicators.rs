use super::*;

#[test]
fn test_odf_security_threat_indicators() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p>
        <text:a xlink:href="https://example.com">Link</text:a>
      </text:p>
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
    let zip_data = build_odf_zip_custom(
        mimetype,
        content_xml,
        None,
        Some(manifest_xml),
        vec![
            ("Scripts/macro.py", b"print('hi')"),
            ("Object 1", b"oledata"),
        ],
    );
    let parser = DocumentParser::new();
    let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    let doc = parsed.document().unwrap();

    assert!(doc.security.macro_project.is_some());
    assert!(!doc.security.external_refs.is_empty());
    assert!(!doc.security.ole_objects.is_empty());
    assert!(doc
        .security
        .threat_indicators
        .iter()
        .any(|i| i.indicator_type == ThreatIndicatorType::RemoteResource));
    assert!(doc
        .security
        .threat_indicators
        .iter()
        .any(|i| i.indicator_type == ThreatIndicatorType::OleObject));
}
