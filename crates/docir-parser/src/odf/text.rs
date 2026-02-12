//! ODF text parsing helpers.

use super::spreadsheet::parse_content_spreadsheet_fast;
use super::*;

pub(super) fn parse_content_text(
    xml: &[u8],
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<OdfContentResult, ParseError> {
    if limits.fast_mode() {
        return parse_content_spreadsheet_fast(xml, store, limits);
    }
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut section = Section::new();
    section.name = Some("body".to_string());
    let mut in_text = false;
    let mut list_stack: Vec<ListContext> = Vec::new();
    let mut list_id_map: HashMap<String, u32> = HashMap::new();
    let mut next_list_id = 1u32;
    let mut comment_counter = 1u32;
    let mut content_result = OdfContentResult::default();

    let mut pending_inline_nodes: Vec<NodeId> = Vec::new();
    let mut revisions: Vec<NodeId> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"office:text" => in_text = true,
                b"text:list" if in_text => {
                    let style_name = attr_value(&e, b"text:style-name").unwrap_or_default();
                    let num_id = list_id_map.entry(style_name).or_insert_with(|| {
                        let id = next_list_id;
                        next_list_id += 1;
                        id
                    });
                    let level = list_stack.len() as u32;
                    list_stack.push(ListContext {
                        num_id: *num_id,
                        level,
                    });
                }
                b"text:p" | b"text:h" if in_text => {
                    let outline_level =
                        attr_value(&e, b"text:outline-level").and_then(|v| v.parse::<u8>().ok());
                    let numbering = list_stack.last().map(|ctx| NumberingInfo {
                        num_id: ctx.num_id,
                        level: ctx.level,
                        format: None,
                    });
                    let paragraph_id = parse_paragraph(
                        &mut reader,
                        e.name().as_ref(),
                        numbering,
                        outline_level,
                        store,
                        &mut pending_inline_nodes,
                        limits,
                    )?;
                    section.content.extend(pending_inline_nodes.drain(..));
                    section.content.push(paragraph_id);
                }
                b"table:table" if in_text => {
                    let table_id = parse_table(&mut reader, store, limits)?;
                    section.content.extend(pending_inline_nodes.drain(..));
                    section.content.push(table_id);
                }
                b"office:annotation" if in_text => {
                    let comment_id = format!("odf-annotation-{}", comment_counter);
                    comment_counter += 1;
                    let comment_node = parse_annotation(&mut reader, &comment_id, store, limits)?;
                    content_result.comments.push(comment_node);
                    let comment_ref = CommentReference::new(comment_id);
                    let ref_id = comment_ref.id;
                    store.insert(IRNode::CommentReference(comment_ref));
                    section.content.push(ref_id);
                }
                b"text:note" if in_text => {
                    let note_class = attr_value(&e, b"text:note-class")
                        .unwrap_or_else(|| "footnote".to_string());
                    let note_id = format!("odf-note-{}", comment_counter);
                    comment_counter += 1;
                    let note = parse_note(&mut reader, &note_id, &note_class, store, limits)?;
                    match note_class.as_str() {
                        "endnote" => content_result.endnotes.push(note),
                        _ => content_result.footnotes.push(note),
                    }
                }
                b"text:bookmark-start" if in_text => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let mut bookmark = BookmarkStart::new(name.clone());
                        bookmark.name = Some(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkStart(bookmark));
                        section.content.push(bookmark_id);
                    }
                }
                b"text:bookmark-end" if in_text => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let bookmark = BookmarkEnd::new(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkEnd(bookmark));
                        section.content.push(bookmark_id);
                    }
                }
                b"text:reference-mark-start" if in_text => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let mut bookmark = BookmarkStart::new(name.clone());
                        bookmark.name = Some(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkStart(bookmark));
                        section.content.push(bookmark_id);
                    }
                }
                b"text:reference-mark-end" if in_text => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let bookmark = BookmarkEnd::new(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkEnd(bookmark));
                        section.content.push(bookmark_id);
                    }
                }
                b"text:date" if in_text => {
                    let mut field = Field::new(Some("DATE".to_string()));
                    field.instruction_parsed = Some(FieldInstruction {
                        kind: FieldKind::Date,
                        args: Vec::new(),
                        switches: Vec::new(),
                    });
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    section.content.push(field_id);
                }
                b"text:time" if in_text => {
                    let field = Field::new(Some("TIME".to_string()));
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    section.content.push(field_id);
                }
                b"draw:frame" if in_text => {
                    if let Some(shape_id) = parse_draw_frame(&mut reader, &e, store)? {
                        section.content.push(shape_id);
                    }
                }
                b"text:tracked-changes" if in_text => {
                    let mut tracked = parse_tracked_changes(&mut reader, store, limits)?;
                    revisions.append(&mut tracked);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"text:bookmark-start" if in_text => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let mut bookmark = BookmarkStart::new(name.clone());
                        bookmark.name = Some(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkStart(bookmark));
                        section.content.push(bookmark_id);
                    }
                }
                b"text:bookmark-end" if in_text => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let bookmark = BookmarkEnd::new(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkEnd(bookmark));
                        section.content.push(bookmark_id);
                    }
                }
                b"text:reference-mark-start" if in_text => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let mut bookmark = BookmarkStart::new(name.clone());
                        bookmark.name = Some(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkStart(bookmark));
                        section.content.push(bookmark_id);
                    }
                }
                b"text:reference-mark-end" if in_text => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let bookmark = BookmarkEnd::new(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkEnd(bookmark));
                        section.content.push(bookmark_id);
                    }
                }
                b"text:date" if in_text => {
                    let mut field = Field::new(Some("DATE".to_string()));
                    field.instruction_parsed = Some(FieldInstruction {
                        kind: FieldKind::Date,
                        args: Vec::new(),
                        switches: Vec::new(),
                    });
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    section.content.push(field_id);
                }
                b"text:time" if in_text => {
                    let field = Field::new(Some("TIME".to_string()));
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    section.content.push(field_id);
                }
                b"draw:frame" if in_text => {}
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"office:text" => in_text = false,
                b"text:list" => {
                    list_stack.pop();
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "content.xml".to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    section.content.extend(revisions);
    let section_id = section.id;
    store.insert(IRNode::Section(section));
    let mut result = OdfContentResult::default();
    result.content.push(section_id);
    result.comments = content_result.comments;
    result.footnotes = content_result.footnotes;
    result.endnotes = content_result.endnotes;
    Ok(result)
}

pub(super) fn build_paragraph(
    store: &mut IrStore,
    text: &str,
    numbering: Option<NumberingInfo>,
    outline_level: Option<u8>,
) -> NodeId {
    let mut paragraph = Paragraph::new();
    if numbering.is_some() || outline_level.is_some() {
        let mut props = ParagraphProperties::default();
        props.numbering = numbering;
        props.outline_level = outline_level;
        paragraph.properties = props;
    }
    if !text.is_empty() {
        let run = Run::new(text.to_string());
        let run_id = run.id;
        store.insert(IRNode::Run(run));
        paragraph.runs.push(run_id);
    }
    let para_id = paragraph.id;
    store.insert(IRNode::Paragraph(paragraph));
    para_id
}
