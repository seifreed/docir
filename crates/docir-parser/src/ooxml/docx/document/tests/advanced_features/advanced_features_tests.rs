use super::*;
use crate::ooxml::relationships::TargetMode;
use crate::ooxml::relationships::{rel_type, Relationship};

#[test]
fn test_parse_field_returns_xml_error_for_malformed_input() {
    let xml = r#"
        <w:fldSimple xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:instr="DATE">
          <w:r><w:t>broken</w:r>
        </w:fldSimple>
    "#;

    let mut parser = DocxParser::new();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let err = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:fldSimple" => {
                break parse_field(&mut parser, &mut reader, attr_value(&e, b"w:instr")).err();
            }
            Ok(Event::Eof) => panic!("missing fldSimple"),
            _ => {}
        }
        buf.clear();
    };

    match err.expect("expected xml error") {
        ParseError::Xml { file, .. } => assert_eq!(file, "word/document.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_sdt_inline_collects_run_hyperlink_field_and_revision_in_order() {
    let mut rels = Relationships::default();
    let rel = Relationship {
        id: "rIdLink".to_string(),
        rel_type: rel_type::HYPERLINK.to_string(),
        target: "https://example.test".to_string(),
        target_mode: TargetMode::External,
    };
    rels.by_type
        .entry(rel.rel_type.clone())
        .or_default()
        .push(rel.id.clone());
    rels.by_id.insert(rel.id.clone(), rel);

    let xml = r#"
        <w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:sdtContent>
            <w:r><w:t>A</w:t></w:r>
            <w:hyperlink r:id="rIdLink"><w:r><w:t>B</w:t></w:r></w:hyperlink>
            <w:fldSimple w:instr="DATE"><w:r><w:t>C</w:t></w:r></w:fldSimple>
            <w:ins w:id="5"><w:r><w:t>D</w:t></w:r></w:ins>
          </w:sdtContent>
        </w:sdt>
    "#;

    let mut parser = DocxParser::new();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let sdt_id = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:sdt" => {
                break parse_sdt(&mut parser, &mut reader, &rels, SdtMode::Inline)
                    .expect("parse inline sdt");
            }
            Ok(Event::Eof) => panic!("missing sdt"),
            Err(e) => panic!("xml error: {e}"),
            _ => {}
        }
        buf.clear();
    };

    let store = parser.into_store();
    let control = match store.get(sdt_id) {
        Some(docir_core::ir::IRNode::ContentControl(c)) => c,
        _ => panic!("expected content control"),
    };
    assert_eq!(control.content.len(), 4);
    assert!(matches!(
        store.get(control.content[0]),
        Some(docir_core::ir::IRNode::Run(_))
    ));
    let link = match store.get(control.content[1]) {
        Some(docir_core::ir::IRNode::Hyperlink(h)) => h,
        _ => panic!("expected hyperlink"),
    };
    assert_eq!(link.target, "https://example.test");
    assert!(link.is_external);
    let field = match store.get(control.content[2]) {
        Some(docir_core::ir::IRNode::Field(f)) => f,
        _ => panic!("expected field"),
    };
    assert!(field.instruction_parsed.is_some());
    assert_eq!(field.runs.len(), 1);
    assert!(matches!(
        store.get(control.content[3]),
        Some(docir_core::ir::IRNode::Revision(_))
    ));
}
