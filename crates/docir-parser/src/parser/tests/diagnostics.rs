use super::super::*;
use super::helpers::create_docx_with_relationships;

#[test]
fn test_content_type_inventory_diagnostics() {
    let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;
    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/extra.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        </Types>"#;
    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;
    let path =
        create_docx_with_relationships(body, rels, content_types, &[("word/extra.xml", "<root/>")]);

    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse docx");
    let doc = parsed.document().expect("doc");
    assert!(!doc.diagnostics.is_empty());
    let mut found = false;
    for diag_id in &doc.diagnostics {
        if let Some(IRNode::Diagnostics(diag)) = parsed.store.get(*diag_id) {
            if diag
                .entries
                .iter()
                .any(|e| e.code == "CONTENT_TYPE_INVENTORY" && e.message.contains("word/extra.xml"))
            {
                found = true;
            }
        }
    }
    assert!(found);
    std::fs::remove_file(path).ok();
}

#[test]
fn test_content_type_unknown_diagnostics() {
    let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;
    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>"#;
    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;
    let path =
        create_docx_with_relationships(body, rels, content_types, &[("word/unknown.bin", "BIN")]);

    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse docx");
    let mut found = false;
    for node in parsed.store.values() {
        if let IRNode::Diagnostics(diag) = node {
            if diag
                .entries
                .iter()
                .any(|e| e.code == "CONTENT_TYPE_UNKNOWN" && e.message.contains("word/unknown.bin"))
            {
                found = true;
            }
        }
    }
    assert!(found);
    std::fs::remove_file(path).ok();
}
