use super::helpers::append_odf_spaces;
use super::{
    BookmarkEnd, BookmarkStart, Field, FieldInstruction, FieldKind, IRNode, IrStore, NodeId,
    NumberingInfo, OdfLimitCounter, OdfReader, ParseError, text,
};
use crate::xml_utils::{local_name, try_attr_value_by_suffix, xml_error};
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
                handle_inline_event(&e, &mut text, store, inline_nodes)?
            }
            Ok(Event::Text(e)) => {
                let chunk = crate::xml_utils::decoded_text(&e)
                    .map_err(|err| xml_error("content.xml", err))?;
                text.push_str(&chunk);
            }
            Ok(Event::GeneralRef(e)) => {
                let chunk = crate::xml_utils::decoded_general_ref(&e)
                    .map_err(|err| xml_error("content.xml", err))?;
                text.push_str(&chunk);
            }
            Ok(Event::End(e)) if e.name().as_ref() == end_name => {
                break;
            }
            Ok(Event::Eof) => {
                return Err(xml_error("content.xml", "unexpected end of paragraph"));
            }
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
) -> Result<(), ParseError> {
    match local_name(event.name().as_ref()) {
        b"s" => {
            let count = try_attr_value_by_suffix(event, &[b":c"], "content.xml")?
                .map(|value| {
                    value
                        .parse::<usize>()
                        .map_err(|err| xml_error("content.xml", err))
                })
                .transpose()?
                .unwrap_or(1);
            append_odf_spaces(text, count)?;
        }
        b"tab" => text.push('\t'),
        b"line-break" => text.push('\n'),
        b"bookmark-start" => {
            if let Some(name) = try_attr_value_by_suffix(event, &[b":name"], "content.xml")? {
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
        b"bookmark-end" => {
            if let Some(name) = try_attr_value_by_suffix(event, &[b":name"], "content.xml")? {
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
        b"date" => {
            let mut field = Field::new(Some("DATE".to_string()));
            field.instruction_parsed = Some(FieldInstruction {
                kind: FieldKind::Date,
                args: Vec::new(),
                switches: Vec::new(),
            });
            let field_id = field.id;
            push_inline_node(store, inline_nodes, field_id, IRNode::Field(field));
        }
        b"time" => {
            let field = Field::new(Some("TIME".to_string()));
            let field_id = field.id;
            push_inline_node(store, inline_nodes, field_id, IRNode::Field(field));
        }
        _ => {}
    }
    Ok(())
}

fn push_inline_node(store: &mut IrStore, inline_nodes: &mut Vec<NodeId>, id: NodeId, node: IRNode) {
    store.insert(node);
    inline_nodes.push(id);
}

#[cfg(test)]
mod tests {
    use super::super::limits::OdfLimits;
    use super::*;
    use crate::parser::ParserConfig;
    use quick_xml::Reader;
    use std::io::Cursor;

    #[test]
    fn parse_paragraph_rejects_missing_end() {
        let xml = r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">Broken"#;
        let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
        let mut buf = Vec::new();
        assert!(matches!(
            reader.read_event_into(&mut buf),
            Ok(Event::Start(_))
        ));

        let mut store = IrStore::new();
        let mut inline_nodes = Vec::new();
        let limits = OdfLimits::new(&ParserConfig::default(), false);
        let err = match parse_paragraph(
            &mut reader,
            b"text:p",
            None,
            None,
            &mut store,
            &mut inline_nodes,
            &limits,
        ) {
            Ok(_) => panic!("truncated paragraph must fail"),
            Err(err) => err,
        };

        assert!(matches!(err, ParseError::Xml { .. }));
    }

    #[test]
    fn parse_paragraph_rejects_excessive_expanded_spaces() {
        let xml = r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:s text:c="1000001"/></text:p>"#;
        let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
        let mut buf = Vec::new();
        assert!(matches!(
            reader.read_event_into(&mut buf),
            Ok(Event::Start(_))
        ));

        let mut store = IrStore::new();
        let mut inline_nodes = Vec::new();
        let limits = OdfLimits::new(&ParserConfig::default(), false);
        let err = parse_paragraph(
            &mut reader,
            b"text:p",
            None,
            None,
            &mut store,
            &mut inline_nodes,
            &limits,
        )
        .expect_err("excessive ODF spaces must fail");

        assert!(
            matches!(err, ParseError::ResourceLimit(message) if message.contains("expanded spaces"))
        );
    }

    #[test]
    fn parse_paragraph_rejects_cumulative_expanded_text() {
        let controls = r#"<text:s text:c="1000000"/>"#.repeat(11);
        let xml = format!(
            r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{controls}</text:p>"#
        );
        let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
        let mut buf = Vec::new();
        assert!(matches!(
            reader.read_event_into(&mut buf),
            Ok(Event::Start(_))
        ));

        let mut store = IrStore::new();
        let mut inline_nodes = Vec::new();
        let limits = OdfLimits::new(&ParserConfig::default(), false);
        let err = parse_paragraph(
            &mut reader,
            b"text:p",
            None,
            None,
            &mut store,
            &mut inline_nodes,
            &limits,
        )
        .expect_err("cumulative ODF expansion must fail");

        assert!(
            matches!(err, ParseError::ResourceLimit(message) if message.contains("expanded text"))
        );
    }
}
