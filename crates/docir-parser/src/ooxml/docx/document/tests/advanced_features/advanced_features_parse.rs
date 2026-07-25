use super::*;
use crate::ooxml::docx::document::table::parse_table;
use crate::ooxml::docx::field::parse_field_instruction;

pub(crate) fn parse_single_table(xml: &str) -> (DocxParser, docir_core::types::NodeId) {
    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut table_id = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:tbl" => {
                table_id = Some(parse_table(&mut parser, &mut reader, &rels).expect("table"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }
    (parser, table_id.expect("table parsed"))
}

#[test]
fn test_parse_field_instruction_extended() {
    let parsed = parse_field_instruction("DDE \"cmd\" \"args\"").expect("parsed");
    assert!(matches!(parsed.kind, docir_core::ir::FieldKind::Dde));

    let parsed = parse_field_instruction("DDEAUTO \"cmd\" \"args\"").expect("parsed");
    assert!(matches!(parsed.kind, docir_core::ir::FieldKind::DdeAuto));

    let parsed = parse_field_instruction("ddeauto \"cmd\" \"args\"").expect("parsed");
    assert!(matches!(parsed.kind, docir_core::ir::FieldKind::DdeAuto));

    let parsed = parse_field_instruction("hyperlink \"https://example.test\"").expect("parsed");
    assert!(matches!(parsed.kind, docir_core::ir::FieldKind::Hyperlink));

    let parsed = parse_field_instruction("AUTOTEXT MyEntry").expect("parsed");
    assert!(matches!(parsed.kind, docir_core::ir::FieldKind::AutoText));

    let parsed = parse_field_instruction("AUTOCORRECT MyEntry").expect("parsed");
    assert!(matches!(
        parsed.kind,
        docir_core::ir::FieldKind::AutoCorrect
    ));

    let parsed = parse_field_instruction("INCLUDEPICTURE \"image.png\"").expect("parsed");
    assert!(matches!(
        parsed.kind,
        docir_core::ir::FieldKind::IncludePicture
    ));
}

#[test]
fn test_parse_hyperlink_anchor_tooltip() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:hyperlink w:anchor="BM1" w:tooltip="Go to bookmark">
            <w:r><w:t>Link</w:t></w:r>
          </w:hyperlink>
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
    let mut link = None;
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::Hyperlink(h)) = store.get(*id) {
            link = Some(h);
            break;
        }
    }
    let link = link.expect("hyperlink");
    assert_eq!(link.target, "#BM1");
    assert_eq!(link.tooltip.as_deref(), Some("Go to bookmark"));
    assert!(!link.is_external);
}
#[test]
fn test_parse_content_control_data_binding() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:sdt>
            <w:sdtPr>
              <w:tag w:val="customer"/>
              <w:dataBinding w:xpath="/customer/name" w:storeItemID="{1234}"
                             w:prefixMappings="xmlns:c='urn:customer'"/>
            </w:sdtPr>
            <w:sdtContent>
              <w:r><w:t>Value</w:t></w:r>
            </w:sdtContent>
          </w:sdt>
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
    let mut control = None;
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::ContentControl(c)) = store.get(*id) {
            control = Some(c);
            break;
        }
    }
    let control = control.expect("content control");
    assert_eq!(control.tag.as_deref(), Some("customer"));
    assert_eq!(
        control.data_binding_xpath.as_deref(),
        Some("/customer/name")
    );
    assert_eq!(
        control.data_binding_store_item_id.as_deref(),
        Some("{1234}")
    );
    assert_eq!(
        control.data_binding_prefix_mappings.as_deref(),
        Some("xmlns:c='urn:customer'")
    );
}
#[test]
fn test_parse_run_properties_caps_and_style() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r>
            <w:rPr>
              <w:rStyle w:val="Emphasis"/>
              <w:caps/>
              <w:smallCaps w:val="0"/>
            </w:rPr>
            <w:t>Text</w:t>
          </w:r>
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
    let run = match store.get(para_node.runs[0]) {
        Some(docir_core::ir::IRNode::Run(r)) => r,
        _ => panic!("missing run"),
    };
    assert_eq!(run.properties.style_id.as_deref(), Some("Emphasis"));
    assert_eq!(run.properties.all_caps, Some(true));
    assert_eq!(run.properties.small_caps, Some(false));
}

#[test]
fn test_parse_paragraph_borders_and_flags() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr>
            <w:keepNext/>
            <w:keepLines w:val="1"/>
            <w:pageBreakBefore/>
            <w:widowControl w:val="0"/>
            <w:pBdr>
              <w:top w:val="single" w:sz="4" w:color="00FF00"/>
            </w:pBdr>
          </w:pPr>
          <w:r><w:t>Para</w:t></w:r>
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
    assert_eq!(para_node.properties.keep_next, Some(true));
    assert_eq!(para_node.properties.keep_lines, Some(true));
    assert_eq!(para_node.properties.page_break_before, Some(true));
    assert_eq!(para_node.properties.widow_control, Some(false));
    let border = para_node
        .properties
        .borders
        .as_ref()
        .and_then(|b| b.top.as_ref());
    assert!(border.is_some());
}

#[test]
fn test_parse_note_references_as_fields() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:footnoteReference w:id="1"/></w:r>
          <w:r><w:endnoteReference w:id="2"/></w:r>
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
    let mut args = Vec::new();
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::Field(f)) = store.get(*id)
            && let Some(parsed) = f.instruction_parsed.as_ref()
        {
            kinds.push(parsed.kind.clone());
            args.extend(parsed.args.clone());
        }
    }
    assert!(
        kinds
            .iter()
            .any(|k| matches!(k, docir_core::ir::FieldKind::FootnoteRef))
    );
    assert!(
        kinds
            .iter()
            .any(|k| matches!(k, docir_core::ir::FieldKind::EndnoteRef))
    );
    assert!(args.iter().any(|a| a == "1"));
    assert!(args.iter().any(|a| a == "2"));
}

#[test]
fn test_parse_note_reference_reports_malformed_attributes() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:footnoteReference w:id="1" w:id="2"/></w:r>
        </w:p>
        "#;
    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let err = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                break parse_paragraph(&mut parser, &mut reader, &rels, None).err();
            }
            Ok(Event::Eof) => panic!("missing paragraph"),
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    };

    match err.expect("malformed note reference attributes must fail") {
        ParseError::Xml { file, .. } => assert_eq!(file, "word/document.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_styles_with_effects_and_style_links() {
    let xml = r#"
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="character" w:styleId="CharDefault" w:default="true">
            <w:name>Character Default</w:name>
            <w:basedOn w:val="BaseChar"/>
            <w:next w:val="NextChar"/>
          </w:style>
          <w:style w:type="table" w:styleId="TblStyle" w:default="1">
            <w:name w:val="Table Style"/>
          </w:style>
          <w:style w:type="numbering" w:styleId="NumStyle">
            <w:name w:val="Numbering Style"/>
          </w:style>
          <w:style w:type="custom" w:styleId="OtherStyle">
            <w:name w:val="Other Style"/>
          </w:style>
        </w:styles>
        "#;

    let mut parser = DocxParser::new();
    let styles_id = parser
        .parse_styles_with_effects(xml)
        .expect("styles with effects");
    let store = parser.into_store();
    let styles = match store.get(styles_id) {
        Some(docir_core::ir::IRNode::StyleSet(s)) => s,
        _ => panic!("missing styles"),
    };

    assert!(styles.with_effects);
    assert_eq!(styles.styles.len(), 4);

    assert_eq!(styles.styles[0].name.as_deref(), Some("Character Default"));
    assert_eq!(styles.styles[0].based_on.as_deref(), Some("BaseChar"));
    assert_eq!(styles.styles[0].next.as_deref(), Some("NextChar"));
    assert!(styles.styles[0].is_default);
    assert!(matches!(
        styles.styles[0].style_type,
        docir_core::ir::StyleType::Character
    ));

    assert!(styles.styles[1].is_default);
    assert!(matches!(
        styles.styles[1].style_type,
        docir_core::ir::StyleType::Table
    ));
    assert!(matches!(
        styles.styles[2].style_type,
        docir_core::ir::StyleType::Numbering
    ));
    assert!(matches!(
        styles.styles[3].style_type,
        docir_core::ir::StyleType::Other
    ));
}
