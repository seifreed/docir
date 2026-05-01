use super::*;

#[test]
fn test_parse_field_instruction_ref_pageref() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="REF Bookmark1 \\h">
            <w:r><w:t>Ref</w:t></w:r>
          </w:fldSimple>
          <w:fldSimple w:instr="PAGEREF Bookmark1 \\p">
            <w:r><w:t>PageRef</w:t></w:r>
          </w:fldSimple>
        </w:p>
        "#;
    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut para = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                para = Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }
    let para = para.expect("para parsed");
    let store = parser.into_store();
    let para_node = match store.get(para.id) {
        Some(docir_core::ir::IRNode::Paragraph(p)) => p,
        _ => panic!("missing paragraph"),
    };
    let mut kinds = Vec::new();
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
            if let Some(parsed) = f.instruction_parsed.as_ref() {
                kinds.push(parsed.kind.clone());
            }
        }
    }
    assert!(kinds
        .iter()
        .any(|k| matches!(k, docir_core::ir::FieldKind::Ref)));
    assert!(kinds
        .iter()
        .any(|k| matches!(k, docir_core::ir::FieldKind::PageRef)));
}

#[test]
fn test_parse_paragraph_field_char_sequence_creates_field_node() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText>  HYPERLINK "https://example.com"  </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>Visible</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>
        "#;
    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut para = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                para = Some(parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }
    let para = para.expect("paragraph parsed");
    let store = parser.into_store();
    let para_node = match store.get(para.id) {
        Some(docir_core::ir::IRNode::Paragraph(p)) => p,
        _ => panic!("missing paragraph"),
    };
    let field_id = para_node
        .runs
        .iter()
        .copied()
        .find(|id| matches!(store.get(*id), Some(docir_core::ir::IRNode::Field(_))))
        .expect("field node");
    let field = match store.get(field_id) {
        Some(docir_core::ir::IRNode::Field(f)) => f,
        _ => panic!("missing field"),
    };
    assert_eq!(
        field.instruction.as_deref(),
        Some("HYPERLINK \"https://example.com\"")
    );
    assert!(field.runs.len() >= 3);
}
