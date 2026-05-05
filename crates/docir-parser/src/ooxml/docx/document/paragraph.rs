use super::inline::{parse_revision_inline, parse_run, parse_sdt, RunParse, SdtMode};
use super::{parse_field, parse_hyperlink, span_from_reader, DocxParser, ParagraphParse};
#[path = "paragraph_props.rs"]
mod paragraph_props;
use super::SectionRef;
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{attr_value, local_name, xml_error};
use docir_core::ir::RevisionType;
use docir_core::ir::{CommentRangeEnd, CommentRangeStart, CommentReference, Field, Paragraph};
use docir_core::types::NodeId;
pub(super) use paragraph_props::parse_paragraph_properties;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::HashMap;

pub(super) fn parse_paragraph(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<ParagraphParse, ParseError> {
    let mut state = ParagraphParseState::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => handle_paragraph_start_event(
                parser,
                reader,
                rels,
                header_footer_map,
                &mut state,
                &e,
            )?,
            Ok(Event::Empty(e)) => handle_paragraph_empty_event(parser, reader, &mut state, &e),
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"p" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    state.para.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = state.para.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::Paragraph(state.para));
    Ok(ParagraphParse {
        id,
        section_ref: state.section_ref,
    })
}

pub(super) fn parse_paragraph_simple(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<NodeId, ParseError> {
    Ok(parse_paragraph(parser, reader, rels, None)?.id)
}

struct FieldState {
    active: bool,
    instr_done: bool,
    instr: String,
    runs: Vec<NodeId>,
}

impl FieldState {
    fn new() -> Self {
        Self {
            active: false,
            instr_done: false,
            instr: String::new(),
            runs: Vec::new(),
        }
    }

    fn start(&mut self) {
        self.active = true;
        self.instr_done = false;
        self.instr.clear();
        self.runs.clear();
    }

    fn separate(&mut self) {
        self.instr_done = true;
    }

    fn finish(&mut self) {
        self.active = false;
        self.instr_done = false;
        self.instr.clear();
        self.runs.clear();
    }
}

struct ParagraphParseState {
    para: Paragraph,
    field_state: FieldState,
    section_ref: Option<SectionRef>,
}

impl ParagraphParseState {
    fn new() -> Self {
        Self {
            para: Paragraph::new(),
            field_state: FieldState::new(),
            section_ref: None,
        }
    }
}

fn update_field_from_run(run: &RunParse, run_id: NodeId, state: &mut FieldState) {
    if state.active {
        state.runs.push(run_id);
        if run.has_instr && !state.instr_done {
            state.instr.push_str(&run.text);
        }
    }
}

fn handle_paragraph_start_event(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> Result<(), ParseError> {
    if local_name(element.name().as_ref()) == b"pPr" {
        state.section_ref = parse_paragraph_properties(reader, &mut state.para, header_footer_map)?;
        return Ok(());
    }

    if handle_inline_start(parser, reader, rels, state, element)? {
        return Ok(());
    }
    if handle_comment_start(parser, reader, state, element) {
        return Ok(());
    }
    if handle_revision_start(parser, reader, rels, state, element)? {
        return Ok(());
    }
    if handle_bookmark_start(parser, state, element) {
        return Ok(());
    }

    Ok(())
}

fn handle_inline_start(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> Result<bool, ParseError> {
    match local_name(element.name().as_ref()) {
        b"r" => {
            let run = parse_run(parser, reader, rels)?;
            let run_id = run.run_id;
            state.para.runs.push(run_id);
            for emb in &run.embedded {
                state.para.runs.push(*emb);
            }
            update_field_from_run(&run, run_id, &mut state.field_state);
            handle_field_char(
                parser,
                &mut state.para,
                &mut state.field_state,
                run.field_char.as_deref(),
            );
            Ok(true)
        }
        b"hyperlink" => {
            let link_id = parse_hyperlink(parser, reader, rels, element)?;
            state.para.runs.push(link_id);
            Ok(true)
        }
        b"sdt" => {
            let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Inline)?;
            state.para.runs.push(sdt_id);
            Ok(true)
        }
        b"fldSimple" => {
            let instr = attr_value(element, b"w:instr");
            let field_id = parse_field(parser, reader, instr)?;
            state.para.runs.push(field_id);
            Ok(true)
        }
        _ => Ok(false),
    }
}

fn handle_comment_start(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> bool {
    let kind = match local_name(element.name().as_ref()) {
        b"commentRangeStart" => CommentNodeKind::RangeStart,
        b"commentRangeEnd" => CommentNodeKind::RangeEnd,
        b"commentReference" => CommentNodeKind::Reference,
        _ => return false,
    };
    insert_comment_node(parser, reader, &mut state.para, element, kind);
    true
}

fn handle_revision_start(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> Result<bool, ParseError> {
    let revision_type = match local_name(element.name().as_ref()) {
        b"ins" => RevisionType::Insert,
        b"del" => RevisionType::Delete,
        b"moveFrom" => RevisionType::MoveFrom,
        b"moveTo" => RevisionType::MoveTo,
        b"pPrChange" | b"rPrChange" => RevisionType::FormatChange,
        _ => return Ok(false),
    };
    push_revision_inline(
        parser,
        reader,
        rels,
        &mut state.para,
        element,
        revision_type,
    )?;
    Ok(true)
}

fn handle_bookmark_start(
    parser: &mut DocxParser,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) -> bool {
    match local_name(element.name().as_ref()) {
        b"bookmarkStart" => {
            insert_bookmark_start(parser, &mut state.para, element);
            true
        }
        b"bookmarkEnd" => {
            insert_bookmark_end(parser, &mut state.para, element);
            true
        }
        _ => false,
    }
}

fn handle_paragraph_empty_event(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    state: &mut ParagraphParseState,
    element: &BytesStart<'_>,
) {
    match local_name(element.name().as_ref()) {
        b"commentRangeStart" => insert_comment_node(
            parser,
            reader,
            &mut state.para,
            element,
            CommentNodeKind::RangeStart,
        ),
        b"commentRangeEnd" => insert_comment_node(
            parser,
            reader,
            &mut state.para,
            element,
            CommentNodeKind::RangeEnd,
        ),
        b"commentReference" => insert_comment_node(
            parser,
            reader,
            &mut state.para,
            element,
            CommentNodeKind::Reference,
        ),
        b"bookmarkStart" => insert_bookmark_start(parser, &mut state.para, element),
        b"bookmarkEnd" => insert_bookmark_end(parser, &mut state.para, element),
        _ => {}
    }
}

fn push_revision_inline(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    para: &mut Paragraph,
    element: &BytesStart<'_>,
    revision_type: RevisionType,
) -> Result<(), ParseError> {
    let rev_id = parse_revision_inline(parser, reader, rels, element, revision_type)?;
    para.runs.push(rev_id);
    Ok(())
}

enum CommentNodeKind {
    RangeStart,
    RangeEnd,
    Reference,
}

fn insert_comment_node(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    para: &mut Paragraph,
    element: &BytesStart<'_>,
    kind: CommentNodeKind,
) {
    if let Some(cid) = attr_value(element, b"w:id") {
        let span = span_from_reader(reader, "word/document.xml");
        let (node, node_id) = match kind {
            CommentNodeKind::RangeStart => {
                let mut node = CommentRangeStart::new(cid);
                node.span = Some(span);
                let node_id = node.id;
                (docir_core::ir::IRNode::CommentRangeStart(node), node_id)
            }
            CommentNodeKind::RangeEnd => {
                let mut node = CommentRangeEnd::new(cid);
                node.span = Some(span);
                let node_id = node.id;
                (docir_core::ir::IRNode::CommentRangeEnd(node), node_id)
            }
            CommentNodeKind::Reference => {
                let mut node = CommentReference::new(cid);
                node.span = Some(span);
                let node_id = node.id;
                (docir_core::ir::IRNode::CommentReference(node), node_id)
            }
        };
        parser.store.insert(node);
        para.runs.push(node_id);
    }
}

fn insert_bookmark_start(parser: &mut DocxParser, para: &mut Paragraph, element: &BytesStart<'_>) {
    if let Some(bm_id) = attr_value(element, b"w:id") {
        let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
        bm.name = attr_value(element, b"w:name");
        bm.col_first = attr_value(element, b"w:colFirst").and_then(|v| v.parse().ok());
        bm.col_last = attr_value(element, b"w:colLast").and_then(|v| v.parse().ok());
        let bm_id = bm.id;
        parser
            .store
            .insert(docir_core::ir::IRNode::BookmarkStart(bm));
        para.runs.push(bm_id);
    }
}

fn insert_bookmark_end(parser: &mut DocxParser, para: &mut Paragraph, element: &BytesStart<'_>) {
    if let Some(bm_id) = attr_value(element, b"w:id") {
        let bm = docir_core::ir::BookmarkEnd::new(bm_id);
        let bm_id = bm.id;
        parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
        para.runs.push(bm_id);
    }
}

fn handle_field_char(
    parser: &mut DocxParser,
    para: &mut Paragraph,
    state: &mut FieldState,
    char_type: Option<&str>,
) {
    match char_type {
        Some("begin") => state.start(),
        Some("separate") => state.separate(),
        Some("end") => {
            if state.active {
                let instr = if state.instr.trim().is_empty() {
                    None
                } else {
                    Some(state.instr.trim().to_string())
                };
                let mut field = Field::new(instr);
                field.runs = state.runs.clone();
                let field_id = field.id;
                parser.store.insert(docir_core::ir::IRNode::Field(field));
                para.runs.push(field_id);
            }
            state.finish();
        }
        _ => {}
    }
}

#[cfg(test)]
mod tests {
    use super::paragraph_props::{alignment_from_val, parse_paragraph_borders};
    use super::{
        handle_field_char, parse_paragraph_properties, update_field_from_run, DocxParser,
        FieldState, Paragraph, RunParse,
    };
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
}
