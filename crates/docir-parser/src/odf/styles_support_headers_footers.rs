use super::super::helpers::parse_table;
use super::super::{
    attr_value, parse_paragraph, text, Footer, Header, IRNode, IrStore, NodeId, NumberingInfo,
    OdfLimitCounter, OdfLimits, OdfReader, ParseError, ParserConfig, SourceSpan,
};
use crate::xml_utils::xml_error;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::HashMap;

pub(crate) fn parse_odf_headers_footers(
    xml: &str,
    store: &mut IrStore,
    config: &ParserConfig,
) -> Result<(Vec<NodeId>, Vec<NodeId>), ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml.as_bytes()));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut headers = Vec::new();
    let mut footers = Vec::new();

    let limits = OdfLimits::new(config, false);

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"style:header" | b"style:header-left" => {
                    let content = parse_odf_header_footer_block(
                        &mut reader,
                        e.name().as_ref(),
                        store,
                        &limits,
                    )?;
                    let mut header = Header::new();
                    header.content = content;
                    header.span = Some(SourceSpan::new("styles.xml"));
                    let id = header.id;
                    store.insert(IRNode::Header(header));
                    headers.push(id);
                }
                b"style:footer" | b"style:footer-left" => {
                    let content = parse_odf_header_footer_block(
                        &mut reader,
                        e.name().as_ref(),
                        store,
                        &limits,
                    )?;
                    let mut footer = Footer::new();
                    footer.content = content;
                    footer.span = Some(SourceSpan::new("styles.xml"));
                    let id = footer.id;
                    store.insert(IRNode::Footer(footer));
                    footers.push(id);
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("styles.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok((headers, footers))
}

#[derive(Debug, Clone, Copy)]
struct ListContext {
    num_id: u32,
    level: u32,
}

struct HeaderFooterParseState<'a> {
    content: &'a mut Vec<NodeId>,
    list_stack: &'a mut Vec<ListContext>,
    list_id_map: &'a mut HashMap<String, u32>,
    next_list_id: &'a mut u32,
    pending_inline_nodes: &'a mut Vec<NodeId>,
}

fn parse_odf_header_footer_block(
    reader: &mut OdfReader<'_>,
    end_name: &[u8],
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<Vec<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut content = Vec::new();
    let mut list_stack: Vec<ListContext> = Vec::new();
    let mut list_id_map: HashMap<String, u32> = HashMap::new();
    let mut next_list_id = 1u32;
    let mut pending_inline_nodes: Vec<NodeId> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => handle_header_footer_start(
                &e,
                reader,
                store,
                limits,
                HeaderFooterParseState {
                    content: &mut content,
                    list_stack: &mut list_stack,
                    list_id_map: &mut list_id_map,
                    next_list_id: &mut next_list_id,
                    pending_inline_nodes: &mut pending_inline_nodes,
                },
            )?,
            Ok(Event::Empty(e)) => handle_header_footer_empty(
                &e,
                store,
                &mut content,
                &list_stack,
                &mut pending_inline_nodes,
            ),
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"text:list" => {
                    list_stack.pop();
                }
                _ if e.name().as_ref() == end_name => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error("styles.xml", e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(content)
}

fn handle_header_footer_start(
    e: &BytesStart<'_>,
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    state: HeaderFooterParseState<'_>,
) -> Result<(), ParseError> {
    let HeaderFooterParseState {
        content,
        list_stack,
        list_id_map,
        next_list_id,
        pending_inline_nodes,
    } = state;
    match e.name().as_ref() {
        b"text:list" => {
            let style_name = attr_value(e, b"text:style-name").unwrap_or_default();
            let num_id = list_id_map.entry(style_name).or_insert_with(|| {
                let id = *next_list_id;
                *next_list_id += 1;
                id
            });
            let level = list_stack.len() as u32;
            list_stack.push(ListContext {
                num_id: *num_id,
                level,
            });
        }
        b"text:p" | b"text:h" => {
            let paragraph_id = parse_header_footer_paragraph(
                e,
                reader,
                store,
                limits,
                list_stack,
                pending_inline_nodes,
            )?;
            content.append(pending_inline_nodes);
            content.push(paragraph_id);
        }
        b"table:table" => {
            let table_id = parse_table(reader, store, limits)?;
            content.append(pending_inline_nodes);
            content.push(table_id);
        }
        _ => {}
    }
    Ok(())
}

fn handle_header_footer_empty(
    e: &BytesStart<'_>,
    store: &mut IrStore,
    content: &mut Vec<NodeId>,
    list_stack: &[ListContext],
    pending_inline_nodes: &mut Vec<NodeId>,
) {
    match e.name().as_ref() {
        b"text:p" | b"text:h" => {
            let outline_level =
                attr_value(e, b"text:outline-level").and_then(|v| v.parse::<u8>().ok());
            let numbering = list_stack.last().map(|ctx| NumberingInfo {
                num_id: ctx.num_id,
                level: ctx.level,
                format: None,
            });
            let paragraph_id = text::build_paragraph(store, "", numbering, outline_level);
            content.append(pending_inline_nodes);
            content.push(paragraph_id);
        }
        _ => {}
    }
}

fn parse_header_footer_paragraph(
    e: &BytesStart<'_>,
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    list_stack: &[ListContext],
    pending_inline_nodes: &mut Vec<NodeId>,
) -> Result<NodeId, ParseError> {
    let outline_level = attr_value(e, b"text:outline-level").and_then(|v| v.parse::<u8>().ok());
    let numbering = list_stack.last().map(|ctx| NumberingInfo {
        num_id: ctx.num_id,
        level: ctx.level,
        format: None,
    });
    parse_paragraph(
        reader,
        e.name().as_ref(),
        numbering,
        outline_level,
        store,
        pending_inline_nodes,
        limits,
    )
}
