use super::{
    attr_value, text, BookmarkEnd, BookmarkStart, Field, FieldInstruction, FieldKind, IRNode,
    IrStore, NodeId, NumberingInfo, OdfLimitCounter, OdfReader, ParseError,
};
use crate::xml_utils::xml_error;
use quick_xml::events::{BytesStart, Event};

pub(crate) fn parse_paragraph(
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
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                handle_inline_event(&e, &mut text, store, inline_nodes)
            }
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
            Err(e) => return Err(xml_error("content.xml", e)),
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

fn handle_inline_event(
    event: &BytesStart<'_>,
    text: &mut String,
    store: &mut IrStore,
    inline_nodes: &mut Vec<NodeId>,
) {
    match event.name().as_ref() {
        b"text:s" => {
            let count = attr_value(event, b"text:c")
                .and_then(|v| v.parse::<usize>().ok())
                .unwrap_or(1);
            text.extend(std::iter::repeat_n(' ', count));
        }
        b"text:tab" => text.push('\t'),
        b"text:line-break" => text.push('\n'),
        b"text:bookmark-start" => {
            if let Some(name) = attr_value(event, b"text:name") {
                let mut bookmark = BookmarkStart::new(name.clone());
                bookmark.name = Some(name);
                let bookmark_id = bookmark.id;
                push_inline_node(
                    store,
                    inline_nodes,
                    bookmark_id,
                    IRNode::BookmarkStart(bookmark),
                );
            }
        }
        b"text:bookmark-end" => {
            if let Some(name) = attr_value(event, b"text:name") {
                let bookmark = BookmarkEnd::new(name);
                let bookmark_id = bookmark.id;
                push_inline_node(
                    store,
                    inline_nodes,
                    bookmark_id,
                    IRNode::BookmarkEnd(bookmark),
                );
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
            push_inline_node(store, inline_nodes, field_id, IRNode::Field(field));
        }
        b"text:time" => {
            let field = Field::new(Some("TIME".to_string()));
            let field_id = field.id;
            push_inline_node(store, inline_nodes, field_id, IRNode::Field(field));
        }
        _ => {}
    }
}

fn push_inline_node(store: &mut IrStore, inline_nodes: &mut Vec<NodeId>, id: NodeId, node: IRNode) {
    store.insert(node);
    inline_nodes.push(id);
}
