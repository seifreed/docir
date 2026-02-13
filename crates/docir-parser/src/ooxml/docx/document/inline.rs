use super::*;

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
                b"w:footnoteReference" => {
                    if let Some(field_id) = parse_note_reference(
                        parser,
                        reader,
                        &e,
                        docir_core::ir::FieldKind::FootnoteRef,
                    ) {
                        embedded.push(field_id);
                    }
                }
                b"w:endnoteReference" => {
                    if let Some(field_id) = parse_note_reference(
                        parser,
                        reader,
                        &e,
                        docir_core::ir::FieldKind::EndnoteRef,
                    ) {
                        embedded.push(field_id);
                    }
                }
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
                    if let Some(field_id) = parse_note_reference(
                        parser,
                        reader,
                        &e,
                        docir_core::ir::FieldKind::FootnoteRef,
                    ) {
                        embedded.push(field_id);
                    }
                } else if e.name().as_ref() == b"w:endnoteReference" {
                    if let Some(field_id) = parse_note_reference(
                        parser,
                        reader,
                        &e,
                        docir_core::ir::FieldKind::EndnoteRef,
                    ) {
                        embedded.push(field_id);
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:r" {
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

    let mut run = Run::new(text);
    run.properties = props;
    run.span = Some(span_from_reader(reader, "word/document.xml"));
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
    let mut revision = Revision::new(change_type);
    revision.revision_id = attr_value(start, b"w:id");
    revision.author = attr_value(start, b"w:author");
    revision.date = attr_value(start, b"w:date");
    revision.span = Some(SourceSpan::new("word/document.xml"));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:r" => {
                    let run = parse_run(parser, reader, rels)?;
                    revision.content.push(run.run_id);
                    revision.content.extend(run.embedded);
                }
                _ => {}
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
                return Err(xml_error("word/document.xml", e));
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

pub(super) fn parse_revision_block(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    start: &BytesStart,
    change_type: RevisionType,
) -> Result<NodeId, ParseError> {
    let mut revision = Revision::new(change_type);
    revision.revision_id = attr_value(start, b"w:id");
    revision.author = attr_value(start, b"w:author");
    revision.date = attr_value(start, b"w:date");
    revision.span = Some(SourceSpan::new("word/document.xml"));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
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
                return Err(xml_error("word/document.xml", e));
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
    control.span = Some(span_from_reader(reader, "word/document.xml"));

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
                return Err(xml_error("word/document.xml", e));
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
                return Err(xml_error("word/document.xml", e));
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
                return Err(xml_error("word/document.xml", e));
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
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:r" => {
                    let run = parse_run(parser, reader, rels)?;
                    runs.push(run.run_id);
                    runs.extend(run.embedded);
                }
                b"w:hyperlink" => {
                    let link_id = parse_hyperlink(parser, reader, rels, &e)?;
                    runs.push(link_id);
                }
                b"w:fldSimple" => {
                    let instr = attr_value(&e, b"w:instr");
                    let field_id = parse_field(parser, reader, instr)?;
                    runs.push(field_id);
                }
                b"w:commentRangeStart" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeStart::new(cid);
                        node.span = Some(SourceSpan::new("word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeStart(node));
                        runs.push(node_id);
                    }
                }
                b"w:commentRangeEnd" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeEnd::new(cid);
                        node.span = Some(SourceSpan::new("word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeEnd(node));
                        runs.push(node_id);
                    }
                }
                b"w:commentReference" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentReference::new(cid);
                        node.span = Some(SourceSpan::new("word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentReference(node));
                        runs.push(node_id);
                    }
                }
                b"w:bookmarkStart" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
                        bm.name = attr_value(&e, b"w:name");
                        let bm_id = bm.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::BookmarkStart(bm));
                        runs.push(bm_id);
                    }
                }
                b"w:bookmarkEnd" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let bm = docir_core::ir::BookmarkEnd::new(bm_id);
                        let bm_id = bm.id;
                        parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
                        runs.push(bm_id);
                    }
                }
                b"w:ins" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::Insert)?;
                    runs.push(rev_id);
                }
                b"w:del" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::Delete)?;
                    runs.push(rev_id);
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
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(runs)
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
    let mut span = span_from_reader(reader, "word/document.xml");
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
                return Err(xml_error("word/document.xml", e));
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
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    let id = field.id;
    parser.store.insert(docir_core::ir::IRNode::Field(field));
    Ok(id)
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
                return Err(xml_error("word/document.xml", e));
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
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}
