use super::*;

#[test]
fn test_parse_field_instruction_hyperlink() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="HYPERLINK &quot;https://example.com&quot; \\t &quot;_blank&quot;">
            <w:r><w:t>Link</w:t></w:r>
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
    let mut field = None;
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
            field = Some(f);
            break;
        }
    }
    let field = field.expect("field");
    let parsed = field.instruction_parsed.as_ref().expect("parsed");
    assert!(matches!(parsed.kind, docir_core::ir::FieldKind::Hyperlink));
    assert!(parsed.args.iter().any(|a| a == "https://example.com"));
    assert!(parsed.switches.iter().any(|s| s == "\\t"));
}

#[test]
fn test_parse_field_instruction_includetext() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="INCLUDETEXT &quot;C:\\docs\\file.docx&quot; \\m">
            <w:r><w:t>Include</w:t></w:r>
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
    let mut field = None;
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
            field = Some(f);
            break;
        }
    }
    let field = field.expect("field");
    let parsed = field.instruction_parsed.as_ref().expect("parsed");
    assert!(matches!(
        parsed.kind,
        docir_core::ir::FieldKind::IncludeText
    ));
    assert!(parsed
        .args
        .iter()
        .any(|a| a.contains("C:") && a.contains("docs") && a.contains("file.docx")));
    assert!(parsed.switches.iter().any(|s| s == "\\m"));
}

#[test]
fn test_parse_field_instruction_mergefield() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="MERGEFIELD  CustomerName  \\* MERGEFORMAT">
            <w:r><w:t>Value</w:t></w:r>
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
    let mut field = None;
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
            field = Some(f);
            break;
        }
    }
    let field = field.expect("field");
    let parsed = field.instruction_parsed.as_ref().expect("parsed");
    assert!(matches!(parsed.kind, docir_core::ir::FieldKind::MergeField));
    assert!(parsed.args.iter().any(|a| a == "CustomerName"));
    assert!(parsed.switches.iter().any(|s| s == "\\*"));
}

#[test]
fn test_parse_field_instruction_date() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:fldSimple w:instr="DATE \\@ &quot;MMMM d, yyyy&quot;">
            <w:r><w:t>Today</w:t></w:r>
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
    let mut field = None;
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id) {
            field = Some(f);
            break;
        }
    }
    let field = field.expect("field");
    let parsed = field.instruction_parsed.as_ref().expect("parsed");
    assert!(matches!(parsed.kind, docir_core::ir::FieldKind::Date));
    assert!(parsed.switches.iter().any(|s| s == "\\@"));
}
