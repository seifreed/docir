use super::*;

pub(super) fn parse_paragraph(
    reader: &mut OdfReader<'_>,
    end_name: &[u8],
    numbering: Option<NumberingInfo>,
    outline_level: Option<u8>,
    store: &mut IrStore,
    inline_nodes: &mut Vec<NodeId>,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    limits.bump_paragraphs(1)?;
    let mut buf = Vec::new();
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"text:s" => {
                    let count = attr_value(&e, b"text:c")
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1);
                    text.extend(std::iter::repeat(' ').take(count));
                }
                b"text:tab" => text.push('\t'),
                b"text:line-break" => text.push('\n'),
                b"text:bookmark-start" => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let mut bookmark = BookmarkStart::new(name.clone());
                        bookmark.name = Some(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkStart(bookmark));
                        inline_nodes.push(bookmark_id);
                    }
                }
                b"text:bookmark-end" => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let bookmark = BookmarkEnd::new(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkEnd(bookmark));
                        inline_nodes.push(bookmark_id);
                    }
                }
                b"text:date" => {
                    let mut field = Field::new(Some("DATE".to_string()));
                    field.instruction_parsed = Some(FieldInstruction {
                        kind: FieldKind::Date,
                        args: Vec::new(),
                        switches: Vec::new(),
                    });
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    inline_nodes.push(field_id);
                }
                b"text:time" => {
                    let field = Field::new(Some("TIME".to_string()));
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    inline_nodes.push(field_id);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"text:s" => {
                    let count = attr_value(&e, b"text:c")
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1);
                    text.extend(std::iter::repeat(' ').take(count));
                }
                b"text:tab" => text.push('\t'),
                b"text:line-break" => text.push('\n'),
                b"text:bookmark-start" => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let mut bookmark = BookmarkStart::new(name.clone());
                        bookmark.name = Some(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkStart(bookmark));
                        inline_nodes.push(bookmark_id);
                    }
                }
                b"text:bookmark-end" => {
                    if let Some(name) = attr_value(&e, b"text:name") {
                        let bookmark = BookmarkEnd::new(name);
                        let bookmark_id = bookmark.id;
                        store.insert(IRNode::BookmarkEnd(bookmark));
                        inline_nodes.push(bookmark_id);
                    }
                }
                b"text:date" => {
                    let mut field = Field::new(Some("DATE".to_string()));
                    field.instruction_parsed = Some(FieldInstruction {
                        kind: FieldKind::Date,
                        args: Vec::new(),
                        switches: Vec::new(),
                    });
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    inline_nodes.push(field_id);
                }
                b"text:time" => {
                    let field = Field::new(Some("TIME".to_string()));
                    let field_id = field.id;
                    store.insert(IRNode::Field(field));
                    inline_nodes.push(field_id);
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let chunk = e.unescape().unwrap_or_default();
                text.push_str(&chunk);
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_name {
                    break;
                }
            }
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

    Ok(text::build_paragraph(
        store,
        &text,
        numbering,
        outline_level,
    ))
}
