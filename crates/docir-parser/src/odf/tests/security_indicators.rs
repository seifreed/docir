use super::*;

#[test]
fn test_parse_odf_security_indicators() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let content_xml = odf_security_content_xml();
    let manifest_xml = odf_security_manifest_xml();
    let signatures_xml = odf_security_signatures_xml();
    let zip_data = build_odf_zip_custom(
        mimetype,
        &content_xml,
        None,
        Some(&manifest_xml),
        vec![
            ("META-INF/documentsignatures.xml", &signatures_xml),
            ("Scripts/macro.py", b"print('hi')"),
            ("Object 1", b"oledata"),
        ],
    );
    let parser = DocumentParser::new();
    let mut parsed = parser.parse_reader(Cursor::new(zip_data)).unwrap();
    docir_security::populate_security_indicators(&mut parsed.store, parsed.root_id);
    assert_security_content_present(&parsed);
    assert_signature_and_encryption_diagnostics(&parsed);
}

fn odf_security_content_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?>
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
"#
    .to_string()
}

fn odf_security_manifest_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Scripts/macro.py" manifest:media-type="application/vnd.sun.star.script"/>
  <manifest:file-entry manifest:full-path="Object 1" manifest:media-type="application/vnd.sun.star.oleobject"/>
  <manifest:file-entry manifest:full-path="Encrypted" manifest:media-type="application/vnd.oasis.opendocument.encrypted"/>
</manifest:manifest>
"#
    .to_string()
}

fn odf_security_signatures_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8"?>
<ds:Signatures xmlns:ds="http://www.w3.org/2000/09/xmldsig#">
  <ds:Signature>
    <ds:SignedInfo>
      <ds:SignatureMethod Algorithm="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"/>
      <ds:Reference>
        <ds:DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>
      </ds:Reference>
    </ds:SignedInfo>
    <ds:KeyInfo>
      <ds:X509Data>
        <ds:X509SubjectName>CN=Tester</ds:X509SubjectName>
      </ds:X509Data>
    </ds:KeyInfo>
  </ds:Signature>
</ds:Signatures>
"#
    .to_vec()
}

fn assert_security_content_present(parsed: &ParsedDocument) {
    let doc = parsed.document().unwrap();
    assert!(doc.security.macro_project.is_some());
    assert!(!doc.security.external_refs.is_empty());
    assert!(!doc.security.ole_objects.is_empty());
}

fn assert_signature_and_encryption_diagnostics(parsed: &ParsedDocument) {
    let mut sig_count = 0;
    let mut encryption_diag = false;
    for node in parsed.store.values() {
        if let IRNode::DigitalSignature(_) = node {
            sig_count += 1;
        }
        if let IRNode::Diagnostics(diag) = node {
            if diag.entries.iter().any(|e| e.code == "ODF_ENCRYPTION") {
                encryption_diag = true;
            }
        }
    }
    assert_eq!(sig_count, 1);
    assert!(encryption_diag);
}
