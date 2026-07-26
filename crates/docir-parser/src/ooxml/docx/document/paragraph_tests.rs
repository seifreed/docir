use super::paragraph_props::{alignment_from_val, parse_paragraph_borders};
use super::{
    DocxParser, FieldState, Paragraph, RunParse, handle_field_char, parse_paragraph_properties,
    update_field_from_run,
};
use crate::error::ParseError;
use crate::xml_utils::reader_from_str;
use docir_core::ir::BorderStyle;
use docir_core::ir::{LineSpacingRule, TextAlignment};
use docir_core::types::NodeId;
use quick_xml::events::Event;
use std::collections::HashMap;

#[test]
fn alignment_from_val_maps_known_and_fallback_values() {
    assert_eq!(alignment_from_val("center"), TextAlignment::Center);
    assert_eq!(alignment_from_val("right"), TextAlignment::Right);
    assert_eq!(alignment_from_val("both"), TextAlignment::Justify);
    assert_eq!(alignment_from_val("distribute"), TextAlignment::Distribute);
    assert_eq!(alignment_from_val("unknown"), TextAlignment::Left);
}

#[test]
fn handle_field_char_creates_field_node_on_end() {
    let mut parser = DocxParser::new();
    let mut para = Paragraph::new();
    let mut state = FieldState::new();

    state.start();
    let run_id = NodeId::new();
    state.runs.push(run_id);
    state
        .instr
        .push_str("  HYPERLINK \"https://example.com\"  ");
    handle_field_char(&mut parser, &mut para, &mut state, Some("end"));

    assert_eq!(para.runs.len(), 1);
    let field_id = para.runs[0];
    match parser.store.get(field_id) {
        Some(docir_core::ir::IRNode::Field(field)) => {
            assert_eq!(
                field.instruction.as_deref(),
                Some("HYPERLINK \"https://example.com\"")
            );
            assert_eq!(field.runs, vec![run_id]);
        }
        other => panic!("expected field node, got {other:?}"),
    }
    assert!(!state.active);
    assert!(state.runs.is_empty());
}

#[test]
fn parse_paragraph_borders_returns_none_when_no_valid_entries() {
    let xml = r#"
            <w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:unknown w:foo="bar"/>
            </w:pBdr>
        "#;
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:pBdr" => break,
            Ok(Event::Eof) => panic!("w:pBdr start not found"),
            Ok(_) => {}
            Err(err) => panic!("unexpected xml read error: {err}"),
        }
        buf.clear();
    }

    let borders = parse_paragraph_borders(&mut reader).expect("paragraph borders parse");
    assert!(borders.is_none());
}

#[test]
fn parse_paragraph_properties_reads_flags_spacing_and_section_refs() {
    let xml = r#"
            <w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:pStyle w:val="Heading1"/>
              <w:keepNext w:val="0"/>
              <w:keepLines/>
              <w:pageBreakBefore w:val="false"/>
              <w:widowControl w:val="1"/>
              <w:jc w:val="distribute"/>
              <w:ind w:left="720" w:right="360" w:firstLine="240" w:hanging="120"/>
              <w:spacing w:before="100" w:after="200" w:line="300" w:lineRule="unknown"/>
              <w:outlineLvl w:val="2"/>
              <w:sectPr>
                <w:headerReference w:type="default" r:id="rIdHeader"/>
                <w:footerReference w:type="default" r:id="rIdFooter"/>
              </w:sectPr>
            </w:pPr>
        "#;

    let mut map = HashMap::new();
    let header_id = NodeId::new();
    let footer_id = NodeId::new();
    map.insert("rIdHeader".to_string(), header_id);
    map.insert("rIdFooter".to_string(), footer_id);

    let mut reader = reader_from_str(xml);
    let mut para = Paragraph::new();
    let section_ref = parse_paragraph_properties(&mut reader, &mut para, Some(&map))
        .expect("paragraph properties parse")
        .expect("section refs expected");

    assert_eq!(para.style_id.as_deref(), Some("Heading1"));
    assert_eq!(para.properties.keep_next, Some(false));
    assert_eq!(para.properties.keep_lines, Some(true));
    assert_eq!(para.properties.page_break_before, Some(false));
    assert_eq!(para.properties.widow_control, Some(true));
    assert_eq!(para.properties.alignment, Some(TextAlignment::Distribute));
    let indent = para.properties.indentation.expect("indentation");
    assert_eq!(indent.left, Some(720));
    assert_eq!(indent.right, Some(360));
    assert_eq!(indent.first_line, Some(240));
    assert_eq!(indent.hanging, Some(120));
    let spacing = para.properties.spacing.expect("spacing");
    assert_eq!(spacing.before, Some(100));
    assert_eq!(spacing.after, Some(200));
    assert_eq!(spacing.line, Some(300));
    assert_eq!(spacing.line_rule, None);
    assert_eq!(para.properties.outline_level, Some(2));
    assert_eq!(section_ref.headers, vec![header_id]);
    assert_eq!(section_ref.footers, vec![footer_id]);
}

#[test]
fn parse_paragraph_properties_reports_malformed_attributes() {
    let xml = r#"
            <w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:pStyle w:val="Heading1" w:val="Heading2"/>
            </w:pPr>
        "#;

    let mut reader = reader_from_str(xml);
    let mut para = Paragraph::new();
    let err = parse_paragraph_properties(&mut reader, &mut para, None)
        .expect_err("malformed paragraph property attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "word/document.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn parse_paragraph_properties_reports_malformed_bool_attributes() {
    let xml = r#"
            <w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:keepNext w:val="1" w:val="0"/>
            </w:pPr>
        "#;

    let mut reader = reader_from_str(xml);
    let mut para = Paragraph::new();
    let err = parse_paragraph_properties(&mut reader, &mut para, None)
        .expect_err("malformed paragraph property attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "word/document.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn handle_field_char_end_with_blank_instruction_creates_field_without_instruction() {
    let mut parser = DocxParser::new();
    let mut para = Paragraph::new();
    let mut state = FieldState::new();

    state.start();
    state.runs.push(NodeId::new());
    state.instr.push_str("   ");
    handle_field_char(&mut parser, &mut para, &mut state, Some("end"));

    assert_eq!(para.runs.len(), 1);
    let field_id = para.runs[0];
    match parser.store.get(field_id) {
        Some(docir_core::ir::IRNode::Field(field)) => {
            assert_eq!(field.instruction, None);
            assert_eq!(field.runs.len(), 1);
        }
        other => panic!("expected field node, got {other:?}"),
    }
}

#[test]
fn parse_paragraph_properties_sets_spacing_line_rule_variants_and_paragraph_borders() {
    let xml = r#"
            <w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:spacing w:lineRule="auto"/>
              <w:spacing w:lineRule="exact"/>
              <w:spacing w:lineRule="atLeast"/>
              <w:pBdr>
                <w:top w:val="single" w:sz="4" w:color="00FF00"/>
                <w:right w:val="double" w:sz="6" w:color="auto"/>
              </w:pBdr>
            </w:pPr>
        "#;

    let mut reader = reader_from_str(xml);
    let mut para = Paragraph::new();
    let section_ref =
        parse_paragraph_properties(&mut reader, &mut para, None).expect("paragraph props");
    assert!(section_ref.is_none());

    let spacing = para.properties.spacing.expect("spacing");
    assert_eq!(spacing.line_rule, Some(LineSpacingRule::AtLeast));

    let borders = para.properties.borders.expect("paragraph borders");
    let top = borders.top.expect("top border");
    assert!(matches!(top.style, BorderStyle::Single));
    assert_eq!(top.width, Some(4));
    assert_eq!(top.color.as_deref(), Some("00FF00"));

    let right = borders.right.expect("right border");
    assert!(matches!(right.style, BorderStyle::Double));
    assert_eq!(right.width, Some(6));
    assert_eq!(
        right.color, None,
        "auto border color should normalize to None"
    );
}

#[test]
fn parse_paragraph_properties_reports_malformed_border_width() {
    let xml = r#"
            <w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:pBdr>
                <w:top w:val="single" w:sz="wide" w:color="00FF00"/>
              </w:pBdr>
            </w:pPr>
        "#;

    let mut reader = reader_from_str(xml);
    let mut para = Paragraph::new();
    let err = parse_paragraph_properties(&mut reader, &mut para, None)
        .expect_err("malformed paragraph border width must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "word/document.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn update_field_from_run_stops_collecting_instruction_after_separate() {
    let mut state = FieldState::new();
    state.start();
    let run_a = RunParse {
        run_id: NodeId::new(),
        text: "HYPERLINK ".to_string(),
        has_instr: true,
        field_char: None,
        embedded: Vec::new(),
    };
    update_field_from_run(&run_a, run_a.run_id, &mut state);
    state.separate();
    let run_b = RunParse {
        run_id: NodeId::new(),
        text: "\"https://ignored.example\"".to_string(),
        has_instr: true,
        field_char: None,
        embedded: Vec::new(),
    };
    update_field_from_run(&run_b, run_b.run_id, &mut state);

    assert_eq!(state.instr, "HYPERLINK ");
    assert_eq!(state.runs.len(), 2);
}
