use super::*;

const DOC_XML_PATH: &str = "word/document.xml";

pub(super) struct RunParse {
    pub(super) run_id: NodeId,
    pub(super) text: String,
    pub(super) has_instr: bool,
    pub(super) field_char: Option<String>,
    pub(super) embedded: Vec<NodeId>,
}

#[derive(Debug, Clone, Copy)]
pub(super) enum SdtMode {
    Block,
    Inline,
}

pub(super) fn parse_run(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<RunParse, ParseError> {
    let mut text = String::new();
    let mut props = RunProperties::default();
    let mut has_instr = false;
    let mut field_char = None;
    let mut embedded = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:rPr" => {
                    parse_run_properties(reader, &mut props)?;
                }
                b"w:drawing" => {
                    if let Some(shape_id) = parse_drawing(parser, reader, rels)? {
                        embedded.push(shape_id);
                    }
                }
                b"w:pict" => {
                    if let Some(shape_id) = parse_vml_pict(parser, reader, rels)? {
                        embedded.push(shape_id);
                    }
                }
                b"w:footnoteReference" => push_note_reference_if_present(
                    parser,
                    reader,
                    &e,
                    docir_core::ir::FieldKind::FootnoteRef,
                    &mut embedded,
                ),
                b"w:endnoteReference" => push_note_reference_if_present(
                    parser,
                    reader,
                    &e,
                    docir_core::ir::FieldKind::EndnoteRef,
                    &mut embedded,
                ),
                b"w:fldChar" => {
                    field_char = attr_value(&e, b"w:fldCharType");
                }
                b"w:t" | b"w:instrText" | b"w:delText" => {
                    let content = reader.read_text(e.name()).unwrap_or_default();
                    if e.name().as_ref() == b"w:instrText" {
                        has_instr = true;
                    }
                    text.push_str(&content);
                }
                b"w:tab" => text.push('\t'),
                _ => {}
            },
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"w:tab" {
                    text.push('\t');
                } else if e.name().as_ref() == b"w:fldChar" {
                    field_char = attr_value(&e, b"w:fldCharType");
                } else if e.name().as_ref() == b"w:footnoteReference" {
                    push_note_reference_if_present(
                        parser,
                        reader,
                        &e,
                        docir_core::ir::FieldKind::FootnoteRef,
                        &mut embedded,
                    );
                } else if e.name().as_ref() == b"w:endnoteReference" {
                    push_note_reference_if_present(
                        parser,
                        reader,
                        &e,
                        docir_core::ir::FieldKind::EndnoteRef,
                        &mut embedded,
                    );
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:r" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }

    let mut run = Run::new(text);
    run.properties = props;
    run.span = Some(span_from_reader(reader, DOC_XML_PATH));
    let run_text = run.text.clone();
    let id = run.id;
    parser.store.insert(docir_core::ir::IRNode::Run(run));
    Ok(RunParse {
        run_id: id,
        text: run_text,
        has_instr,
        field_char,
        embedded,
    })
}

fn push_note_reference_if_present(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    kind: docir_core::ir::FieldKind,
    embedded: &mut Vec<NodeId>,
) {
    if let Some(field_id) = parse_note_reference(parser, reader, start, kind) {
        embedded.push(field_id);
    }
}

fn parse_note_reference(
    parser: &mut DocxParser,
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    kind: docir_core::ir::FieldKind,
) -> Option<NodeId> {
    let id = attr_value(start, b"w:id")?;
    Some(insert_note_reference(parser, reader, kind, id))
}

pub(super) fn parse_revision_inline(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
    change_type: RevisionType,
) -> Result<NodeId, ParseError> {
    parse_revision(
        parser,
        reader,
        rels,
        start,
        change_type,
        RevisionParseMode::Inline,
    )
}

pub(super) fn parse_revision_block(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
    change_type: RevisionType,
) -> Result<NodeId, ParseError> {
    parse_revision(
        parser,
        reader,
        rels,
        start,
        change_type,
        RevisionParseMode::Block,
    )
}

enum RevisionParseMode {
    Inline,
    Block,
}

fn parse_revision(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
    change_type: RevisionType,
    mode: RevisionParseMode,
) -> Result<NodeId, ParseError> {
    let mut revision = Revision::new(change_type);
    revision.revision_id = attr_value(start, b"w:id");
    revision.author = attr_value(start, b"w:author");
    revision.date = attr_value(start, b"w:date");
    revision.span = Some(SourceSpan::new(DOC_XML_PATH));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match mode {
                RevisionParseMode::Inline => {
                    if e.name().as_ref() == b"w:r" {
                        let run = parse_run(parser, reader, rels)?;
                        revision.content.push(run.run_id);
                        revision.content.extend(run.embedded);
                    }
                }
                RevisionParseMode::Block => match e.name().as_ref() {
                    b"w:p" => {
                        let para_id = parse_paragraph_simple(parser, reader, rels)?;
                        revision.content.push(para_id);
                    }
                    b"w:tbl" => {
                        let table_id = parse_table(parser, reader, rels)?;
                        revision.content.push(table_id);
                    }
                    b"w:r" => {
                        let run = parse_run(parser, reader, rels)?;
                        revision.content.push(run.run_id);
                        revision.content.extend(run.embedded);
                    }
                    _ => {}
                },
            },
            Ok(Event::End(e)) => {
                if matches!(
                    e.name().as_ref(),
                    b"w:ins"
                        | b"w:del"
                        | b"w:moveFrom"
                        | b"w:moveTo"
                        | b"w:pPrChange"
                        | b"w:rPrChange"
                ) {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }

    let id = revision.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::Revision(revision));
    Ok(id)
}

pub(super) fn parse_sdt(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    mode: SdtMode,
) -> Result<NodeId, ParseError> {
    let mut control = docir_core::ir::ContentControl::new();
    control.span = Some(span_from_reader(reader, DOC_XML_PATH));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:sdtPr" => {
                    parse_sdt_properties(reader, &mut control)?;
                }
                b"w:sdtContent" => {
                    let content = match mode {
                        SdtMode::Block => parse_sdt_content_block(parser, reader, rels)?,
                        SdtMode::Inline => parse_sdt_content_inline(parser, reader, rels)?,
                    };
                    control.content.extend(content);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:sdt" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }

    let id = control.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::ContentControl(control));
    Ok(id)
}

fn parse_sdt_properties(
    reader: &mut Reader<&[u8]>,
    control: &mut docir_core::ir::ContentControl,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:tag" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        control.tag = Some(val);
                    }
                }
                b"w:alias" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        control.alias = Some(val);
                    }
                }
                b"w:id" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        control.sdt_id = Some(val);
                    }
                }
                b"w:comboBox" => control.control_type = Some("comboBox".to_string()),
                b"w:dropDownList" => control.control_type = Some("dropDownList".to_string()),
                b"w:date" => control.control_type = Some("date".to_string()),
                b"w:checkbox" => control.control_type = Some("checkbox".to_string()),
                b"w:text" => control.control_type = Some("text".to_string()),
                b"w:picture" => control.control_type = Some("picture".to_string()),
                b"w:dataBinding" => {
                    control.data_binding_xpath = attr_value(&e, b"w:xpath");
                    control.data_binding_store_item_id = attr_value(&e, b"w:storeItemID");
                    control.data_binding_prefix_mappings = attr_value(&e, b"w:prefixMappings");
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:sdtPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_sdt_content_block(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Vec<NodeId>, ParseError> {
    let mut content = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:p" => {
                    let para_id = parse_paragraph_simple(parser, reader, rels)?;
                    content.push(para_id);
                }
                b"w:tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    content.push(table_id);
                }
                b"w:sdt" => {
                    let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Block)?;
                    content.push(sdt_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:sdtContent" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(content)
}

fn parse_sdt_content_inline(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<Vec<NodeId>, ParseError> {
    let mut runs = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                handle_sdt_content_inline_start(parser, reader, rels, &e, &mut runs)?
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:sdtContent" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(runs)
}

fn handle_sdt_content_inline_start(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart<'_>,
    runs: &mut Vec<NodeId>,
) -> Result<(), ParseError> {
    match start.name().as_ref() {
        b"w:r" => {
            let run = parse_run(parser, reader, rels)?;
            runs.push(run.run_id);
            runs.extend(run.embedded);
        }
        b"w:hyperlink" => {
            let link_id = parse_hyperlink(parser, reader, rels, start)?;
            runs.push(link_id);
        }
        b"w:fldSimple" => {
            let instr = attr_value(start, b"w:instr");
            let field_id = parse_field(parser, reader, instr)?;
            runs.push(field_id);
        }
        b"w:commentRangeStart" => {
            if let Some(node_id) = insert_comment_range_start(parser, start) {
                runs.push(node_id);
            }
        }
        b"w:commentRangeEnd" => {
            if let Some(node_id) = insert_comment_range_end(parser, start) {
                runs.push(node_id);
            }
        }
        b"w:commentReference" => {
            if let Some(node_id) = insert_comment_reference(parser, start) {
                runs.push(node_id);
            }
        }
        b"w:bookmarkStart" => {
            if let Some(node_id) = insert_bookmark_start(parser, start) {
                runs.push(node_id);
            }
        }
        b"w:bookmarkEnd" => {
            if let Some(node_id) = insert_bookmark_end(parser, start) {
                runs.push(node_id);
            }
        }
        b"w:ins" => {
            let rev_id = parse_revision_inline(parser, reader, rels, start, RevisionType::Insert)?;
            runs.push(rev_id);
        }
        b"w:del" => {
            let rev_id = parse_revision_inline(parser, reader, rels, start, RevisionType::Delete)?;
            runs.push(rev_id);
        }
        _ => {}
    }
    Ok(())
}

fn insert_comment_range_start(parser: &mut DocxParser, start: &BytesStart<'_>) -> Option<NodeId> {
    let cid = attr_value(start, b"w:id")?;
    let mut node = CommentRangeStart::new(cid);
    node.span = Some(SourceSpan::new(DOC_XML_PATH));
    let node_id = node.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::CommentRangeStart(node));
    Some(node_id)
}

fn insert_comment_range_end(parser: &mut DocxParser, start: &BytesStart<'_>) -> Option<NodeId> {
    let cid = attr_value(start, b"w:id")?;
    let mut node = CommentRangeEnd::new(cid);
    node.span = Some(SourceSpan::new(DOC_XML_PATH));
    let node_id = node.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::CommentRangeEnd(node));
    Some(node_id)
}

fn insert_comment_reference(parser: &mut DocxParser, start: &BytesStart<'_>) -> Option<NodeId> {
    let cid = attr_value(start, b"w:id")?;
    let mut node = CommentReference::new(cid);
    node.span = Some(SourceSpan::new(DOC_XML_PATH));
    let node_id = node.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::CommentReference(node));
    Some(node_id)
}

fn insert_bookmark_start(parser: &mut DocxParser, start: &BytesStart<'_>) -> Option<NodeId> {
    let bm_id = attr_value(start, b"w:id")?;
    let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
    bm.name = attr_value(start, b"w:name");
    let node_id = bm.id;
    parser
        .store
        .insert(docir_core::ir::IRNode::BookmarkStart(bm));
    Some(node_id)
}

fn insert_bookmark_end(parser: &mut DocxParser, start: &BytesStart<'_>) -> Option<NodeId> {
    let bm_id = attr_value(start, b"w:id")?;
    let bm = docir_core::ir::BookmarkEnd::new(bm_id);
    let node_id = bm.id;
    parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
    Some(node_id)
}

pub(super) fn parse_hyperlink(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
) -> Result<NodeId, ParseError> {
    let mut link = Hyperlink::new("", false);
    let mut rel_id_opt = None;
    if let Some(tooltip) = attr_value(start, b"w:tooltip") {
        link.tooltip = Some(tooltip);
    }
    if let Some(rel_id) = attr_value(start, b"r:id") {
        if let Some(rel) = rels.get(&rel_id) {
            link.target = rel.target.clone();
            link.is_external = rel.target_mode == TargetMode::External;
            link.relationship_id = Some(rel_id.clone());
            rel_id_opt = Some(rel_id);
        }
    }
    if let Some(anchor) = attr_value(start, b"w:anchor") {
        if link.target.is_empty() {
            link.target = format!("#{}", anchor);
        } else if !link.target.contains('#') {
            link.target = format!("{}#{}", link.target, anchor);
        }
    }
    let mut span = span_from_reader(reader, DOC_XML_PATH);
    span.relationship_id = rel_id_opt.clone();
    link.span = Some(span);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"w:r" {
                    let run = parse_run(parser, reader, rels)?;
                    link.runs.push(run.run_id);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:hyperlink" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }

    let id = link.id;
    parser.store.insert(docir_core::ir::IRNode::Hyperlink(link));
    Ok(id)
}

pub(super) fn parse_field(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    instruction: Option<String>,
) -> Result<NodeId, ParseError> {
    let mut field = Field::new(instruction);
    field.instruction_parsed = field
        .instruction
        .as_deref()
        .and_then(parse_field_instruction);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"w:r" {
                    let run = parse_run(parser, reader, &Relationships::default())?;
                    field.runs.push(run.run_id);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:fldSimple" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }

    let id = field.id;
    parser.store.insert(docir_core::ir::IRNode::Field(field));
    Ok(id)
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::events::Event;

    #[test]
    fn parse_run_records_note_references_and_instruction_metadata() {
        let xml = r#"
            <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:instrText>REF Test</w:instrText>
              <w:tab/>
              <w:fldChar w:fldCharType="begin"/>
              <w:footnoteReference w:id="42"/>
              <w:endnoteReference w:id="99"/>
            </w:r>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut buf = Vec::new();
        let mut run = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                    run = Some(parse_run(&mut parser, &mut reader, &rels).expect("parse run"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {e}"),
                _ => {}
            }
            buf.clear();
        }

        let run = run.expect("run parsed");
        assert_eq!(run.text, "REF Test\t");
        assert!(run.has_instr);
        assert_eq!(run.field_char.as_deref(), Some("begin"));
        assert_eq!(run.embedded.len(), 2);

        let store = parser.into_store();
        let mut has_footnote = false;
        let mut has_endnote = false;
        let mut args = Vec::new();
        for node_id in run.embedded {
            let field = match store.get(node_id) {
                Some(docir_core::ir::IRNode::Field(field)) => field,
                _ => panic!("expected field node"),
            };
            let parsed = field.instruction_parsed.as_ref().expect("parsed field");
            if matches!(parsed.kind, docir_core::ir::FieldKind::FootnoteRef) {
                has_footnote = true;
            }
            if matches!(parsed.kind, docir_core::ir::FieldKind::EndnoteRef) {
                has_endnote = true;
            }
            args.push(parsed.args[0].clone());
        }
        args.sort();
        assert_eq!(args, vec!["42".to_string(), "99".to_string()]);
        assert!(has_footnote);
        assert!(has_endnote);
    }

    #[test]
    fn parse_run_ignores_unresolvable_vml_picture_reference() {
        let xml = r#"
            <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                 xmlns:v="urn:schemas-microsoft-com:vml"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:t>caption</w:t>
              <w:pict>
                <v:shape id="shape1">
                  <v:imagedata r:id="rIdMissing"/>
                </v:shape>
              </w:pict>
            </w:r>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut buf = Vec::new();
        let mut run = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                    run = Some(parse_run(&mut parser, &mut reader, &rels).expect("parse run"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {e}"),
                _ => {}
            }
            buf.clear();
        }

        let run = run.expect("run parsed");
        assert_eq!(run.text, "caption");
        assert!(run.embedded.is_empty());
    }

    #[test]
    fn parse_run_collects_deleted_text_and_ignores_note_without_id() {
        let xml = r#"
            <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:delText>gone</w:delText>
              <w:footnoteReference/>
              <w:endnoteReference w:id="7"/>
            </w:r>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut buf = Vec::new();
        let mut run = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                    run = Some(parse_run(&mut parser, &mut reader, &rels).expect("parse run"));
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {e}"),
                _ => {}
            }
            buf.clear();
        }

        let run = run.expect("run parsed");
        assert_eq!(run.text, "gone");
        assert_eq!(run.embedded.len(), 1);

        let store = parser.into_store();
        let field = match store.get(run.embedded[0]) {
            Some(docir_core::ir::IRNode::Field(field)) => field,
            _ => panic!("expected field node"),
        };
        let parsed = field.instruction_parsed.as_ref().expect("parsed field");
        assert!(matches!(parsed.kind, docir_core::ir::FieldKind::EndnoteRef));
        assert_eq!(parsed.args, vec!["7".to_string()]);
    }

    #[test]
    fn parse_numbering_sets_numbering_only_when_numid_and_level_present() {
        let xml = r#"
            <w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:numId w:val="12"/>
              <w:ilvl w:val="2"/>
            </w:numPr>
        "#;
        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();
        let mut props = ParagraphProperties::default();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:numPr" => {
                    parse_numbering(&mut reader, &mut props).expect("parse numbering");
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {e}"),
                _ => {}
            }
            buf.clear();
        }

        let numbering = props.numbering.expect("numbering should be present");
        assert_eq!(numbering.num_id, 12);
        assert_eq!(numbering.level, 2);

        let xml_missing_level = r#"
            <w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:numId w:val="12"/>
            </w:numPr>
        "#;
        let mut reader = reader_from_str(xml_missing_level);
        let mut buf = Vec::new();
        let mut props = ParagraphProperties::default();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:numPr" => {
                    parse_numbering(&mut reader, &mut props).expect("parse numbering");
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {e}"),
                _ => {}
            }
            buf.clear();
        }
        assert!(props.numbering.is_none());
    }

    #[test]
    fn parse_sdt_inline_collects_supported_content_and_skips_missing_ids() {
        let xml = r#"
            <w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:sdtPr>
                <w:tag w:val="customer"/>
                <w:alias w:val="Customer Field"/>
                <w:id w:val="11"/>
                <w:checkbox/>
                <w:dataBinding w:xpath="/root/customer" w:storeItemID="{store}" w:prefixMappings="xmlns:x='urn:test'"/>
              </w:sdtPr>
              <w:sdtContent>
                <w:commentRangeStart/>
                <w:commentRangeStart w:id="5"/>
                <w:bookmarkStart w:name="missing-id"/>
                <w:bookmarkStart w:id="8" w:name="bm"/>
                <w:fldSimple w:instr="DATE">
                  <w:r><w:t>2024-01-01</w:t></w:r>
                </w:fldSimple>
                <w:ins w:id="1" w:author="Bot">
                  <w:r><w:t>added</w:t></w:r>
                </w:ins>
              </w:sdtContent>
            </w:sdt>
        "#;

        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut buf = Vec::new();
        let mut sdt_id = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:sdt" => {
                    sdt_id =
                        Some(parse_sdt(&mut parser, &mut reader, &rels, SdtMode::Inline).unwrap());
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {e}"),
                _ => {}
            }
            buf.clear();
        }

        let store = parser.into_store();
        let control = match store.get(sdt_id.expect("sdt node should parse")) {
            Some(docir_core::ir::IRNode::ContentControl(control)) => control,
            _ => panic!("expected content control"),
        };
        assert_eq!(control.tag.as_deref(), Some("customer"));
        assert_eq!(control.alias.as_deref(), Some("Customer Field"));
        assert_eq!(control.sdt_id.as_deref(), Some("11"));
        assert_eq!(control.control_type.as_deref(), Some("checkbox"));
        assert_eq!(
            control.data_binding_xpath.as_deref(),
            Some("/root/customer")
        );
        assert_eq!(control.content.len(), 2);

        let mut has_field = false;
        let mut has_revision = false;
        for node_id in &control.content {
            match store.get(*node_id) {
                Some(docir_core::ir::IRNode::Field(field)) => {
                    has_field = true;
                    assert_eq!(field.runs.len(), 1);
                }
                Some(docir_core::ir::IRNode::Revision(rev)) => {
                    has_revision = true;
                    assert!(matches!(rev.change_type, RevisionType::Insert));
                    assert_eq!(rev.content.len(), 1);
                }
                _ => {}
            }
        }
        assert!(has_field);
        assert!(has_revision);
    }

    #[test]
    fn parse_sdt_reports_xml_error_for_malformed_input() {
        let xml = r#"
            <w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:sdtContent>
                <w:r><w:t>unterminated</w:t></w:r>
            </w:sdt>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut buf = Vec::new();
        let mut err = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:sdt" => {
                    err = parse_sdt(&mut parser, &mut reader, &rels, SdtMode::Inline).err();
                    break;
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }

        match err.expect("expected parse error") {
            ParseError::Xml { file, .. } => assert_eq!(file, DOC_XML_PATH),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_run_truncated_run_uses_eof_fallback_without_embeds() {
        let xml = r#"
            <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:t>broken</w:r>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut buf = Vec::new();
        let mut run = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                    run = Some(parse_run(&mut parser, &mut reader, &rels).expect("parse run"));
                    break;
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }

        let run = run.expect("run parsed");
        assert_eq!(run.text, "");
        assert!(run.embedded.is_empty());
        assert!(!run.has_instr);
        assert!(run.field_char.is_none());
    }

    #[test]
    fn parse_revision_inline_records_metadata_and_run_content() {
        let xml = r#"
            <w:ins xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                   w:id="7" w:author="Auditor" w:date="2026-02-28T11:00:00Z">
              <w:customTag/>
              <w:r><w:t>added text</w:t></w:r>
            </w:ins>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut buf = Vec::new();
        let mut revision_id = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:ins" => {
                    revision_id = Some(
                        parse_revision_inline(
                            &mut parser,
                            &mut reader,
                            &rels,
                            &e.into_owned(),
                            RevisionType::Insert,
                        )
                        .expect("parse revision inline"),
                    );
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {e}"),
                _ => {}
            }
            buf.clear();
        }

        let store = parser.into_store();
        let revision = match store.get(revision_id.expect("revision should parse")) {
            Some(docir_core::ir::IRNode::Revision(revision)) => revision,
            _ => panic!("expected revision node"),
        };
        assert_eq!(revision.revision_id.as_deref(), Some("7"));
        assert_eq!(revision.author.as_deref(), Some("Auditor"));
        assert_eq!(revision.date.as_deref(), Some("2026-02-28T11:00:00Z"));
        assert_eq!(revision.content.len(), 1);

        let run = match store.get(revision.content[0]) {
            Some(docir_core::ir::IRNode::Run(run)) => run,
            _ => panic!("expected run node"),
        };
        assert_eq!(run.text, "added text");
    }

    #[test]
    fn parse_revision_inline_reports_xml_error_for_malformed_input() {
        let xml = r#"
            <w:ins xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="9">
              <w:r><w:t>unterminated</w:r>
            </w:ins>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut buf = Vec::new();
        let mut err = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:ins" => {
                    err = parse_revision_inline(
                        &mut parser,
                        &mut reader,
                        &rels,
                        &e.into_owned(),
                        RevisionType::Insert,
                    )
                    .err();
                    break;
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }

        match err.expect("expected xml parse error") {
            ParseError::Xml { file, .. } => assert_eq!(file, DOC_XML_PATH),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_sdt_inline_captures_end_markers_and_delete_revision() {
        let xml = r#"
            <w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:sdtContent>
                <w:commentRangeEnd w:id="12"></w:commentRangeEnd>
                <w:commentReference w:id="12"></w:commentReference>
                <w:bookmarkEnd w:id="20"></w:bookmarkEnd>
                <w:del w:id="3" w:author="Bot">
                  <w:r><w:t>removed</w:t></w:r>
                </w:del>
              </w:sdtContent>
            </w:sdt>
        "#;

        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut buf = Vec::new();
        let mut sdt_id = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:sdt" => {
                    sdt_id =
                        Some(parse_sdt(&mut parser, &mut reader, &rels, SdtMode::Inline).unwrap());
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => panic!("xml error: {e}"),
                _ => {}
            }
            buf.clear();
        }

        let store = parser.into_store();
        let control = match store.get(sdt_id.expect("sdt parsed")) {
            Some(docir_core::ir::IRNode::ContentControl(control)) => control,
            _ => panic!("expected content control"),
        };
        assert_eq!(control.content.len(), 4);

        let mut saw_comment_end = false;
        let mut saw_comment_reference = false;
        let mut saw_bookmark_end = false;
        let mut saw_delete_revision = false;

        for node_id in &control.content {
            match store.get(*node_id) {
                Some(docir_core::ir::IRNode::CommentRangeEnd(node)) => {
                    saw_comment_end = true;
                    assert_eq!(node.comment_id, "12");
                }
                Some(docir_core::ir::IRNode::CommentReference(node)) => {
                    saw_comment_reference = true;
                    assert_eq!(node.comment_id, "12");
                }
                Some(docir_core::ir::IRNode::BookmarkEnd(node)) => {
                    saw_bookmark_end = true;
                    assert_eq!(node.bookmark_id, "20");
                }
                Some(docir_core::ir::IRNode::Revision(revision)) => {
                    saw_delete_revision = true;
                    assert!(matches!(revision.change_type, RevisionType::Delete));
                    assert_eq!(revision.content.len(), 1);
                    let run = match store.get(revision.content[0]) {
                        Some(docir_core::ir::IRNode::Run(run)) => run,
                        _ => panic!("expected run in revision"),
                    };
                    assert_eq!(run.text, "removed");
                }
                _ => {}
            }
        }

        assert!(saw_comment_end);
        assert!(saw_comment_reference);
        assert!(saw_bookmark_end);
        assert!(saw_delete_revision);
    }
}

pub(super) fn parse_numbering(
    reader: &mut Reader<&[u8]>,
    props: &mut ParagraphProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    let mut num_id = None;
    let mut level = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:numId" => {
                    num_id = attr_value(&e, b"w:val").and_then(|v| v.parse().ok());
                }
                b"w:ilvl" => {
                    level = attr_value(&e, b"w:val").and_then(|v| v.parse().ok());
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:numPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }
    if let (Some(num_id), Some(level)) = (num_id, level) {
        props.numbering = Some(NumberingInfo {
            num_id,
            level,
            format: None,
        });
    }
    Ok(())
}

pub(super) fn parse_run_properties(
    reader: &mut Reader<&[u8]>,
    props: &mut RunProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:rStyle" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.style_id = Some(val);
                    }
                }
                b"w:rFonts" => {
                    if let Some(val) = attr_value(&e, b"w:ascii") {
                        props.font_family = Some(val);
                    }
                }
                b"w:sz" => {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        props.font_size = Some(val);
                    }
                }
                b"w:b" => props.bold = Some(true),
                b"w:i" => props.italic = Some(true),
                b"w:u" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.underline = match val.as_str() {
                            "double" => Some(UnderlineStyle::Double),
                            "thick" => Some(UnderlineStyle::Thick),
                            "dotted" => Some(UnderlineStyle::Dotted),
                            "dash" | "dashed" => Some(UnderlineStyle::Dashed),
                            _ => Some(UnderlineStyle::Single),
                        };
                    } else {
                        props.underline = Some(UnderlineStyle::Single);
                    }
                }
                b"w:strike" => props.strike = Some(true),
                b"w:color" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.color = Some(val);
                    }
                }
                b"w:highlight" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.highlight = Some(val);
                    }
                }
                b"w:caps" => {
                    props.all_caps = Some(bool_from_val(&e));
                }
                b"w:smallCaps" => {
                    props.small_caps = Some(bool_from_val(&e));
                }
                b"w:vertAlign" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        props.vertical_align = match val.as_str() {
                            "superscript" => Some(VerticalTextAlignment::Superscript),
                            "subscript" => Some(VerticalTextAlignment::Subscript),
                            _ => None,
                        };
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:rPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(DOC_XML_PATH, e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}
