use crate::error::ParseError;
#[path = "helpers_parse_events_changes.rs"]
mod helpers_parse_events_changes;
#[path = "helpers_parse_events_tables.rs"]
mod helpers_parse_events_tables;
use crate::odf::{
    limits::OdfLimitCounter, presentation_helpers::classify_media_shape,
    utils::parse_frame_transform, OdfReader,
};
use crate::xml_utils::{
    attr_value, scan_xml_events_until_end, scan_xml_events_with_reader, XmlScanControl,
};
use docir_core::ir::*;
use docir_core::types::*;
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};

const ODF_CONTENT_XML: &str = "content.xml";

#[derive(Debug, Clone)]
pub(crate) struct ValidationDef {
    pub(crate) validation_type: Option<String>,
    pub(crate) operator: Option<String>,
    pub(crate) allow_blank: bool,
    pub(crate) show_input_message: bool,
    pub(crate) show_error_message: bool,
    pub(crate) error_title: Option<String>,
    pub(crate) error: Option<String>,
    pub(crate) prompt_title: Option<String>,
    pub(crate) prompt: Option<String>,
    pub(crate) formula1: Option<String>,
    pub(crate) formula2: Option<String>,
}

pub(crate) fn parse_notes(reader: &mut OdfReader<'_>) -> Result<Option<String>, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    scan_xml_events_with_reader(reader, &mut buf, ODF_CONTENT_XML, |reader, event| {
        match event {
            Event::Start(e) => {
                if e.name().as_ref() == b"text:p" {
                    let para = parse_text_element(reader, b"text:p")?;
                    if !text.is_empty() {
                        text.push('\n');
                    }
                    text.push_str(&para);
                }
            }
            Event::End(e) if e.name().as_ref() == b"presentation:notes" => {
                return Ok(XmlScanControl::Break);
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    if text.is_empty() {
        Ok(None)
    } else {
        Ok(Some(text))
    }
}

pub(crate) fn parse_validation_definition(
    start: &BytesStart<'_>,
) -> Option<(String, ValidationDef)> {
    let name = attr_value(start, b"table:name")?;
    let condition = attr_value(start, b"table:condition");
    let allow_blank = attr_value(start, b"table:allow-empty-cell")
        .map(|v| v == "true")
        .unwrap_or(false);
    let show_input_message = attr_value(start, b"table:display-list")
        .map(|v| v == "true")
        .unwrap_or(false);
    let show_error_message = attr_value(start, b"table:display-list")
        .map(|v| v == "true")
        .unwrap_or(false);
    let def = ValidationDef {
        validation_type: condition.clone(),
        operator: None,
        allow_blank,
        show_input_message,
        show_error_message,
        error_title: None,
        error: None,
        prompt_title: None,
        prompt: None,
        formula1: condition,
        formula2: None,
    };
    Some((name, def))
}

pub(crate) fn parse_ods_conditional_formatting(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<Option<ConditionalFormat>, ParseError> {
    let mut cf = init_conditional_format(start);

    let mut buf = Vec::new();
    let mut depth: usize = 1;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if depth == 1 && e.name().as_ref() == b"table:conditional-format" {
                    cf.rules.push(build_ods_conditional_rule(&e));
                }
                depth = depth.saturating_add(1);
            }
            Ok(Event::Empty(e)) => {
                if depth == 1 && e.name().as_ref() == b"table:conditional-format" {
                    cf.rules.push(build_ods_conditional_rule(&e));
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:conditional-formatting" && depth == 1 {
                    break;
                }
                depth = depth.saturating_sub(1);
            }
            Ok(Event::Eof) => {
                return Err(ParseError::Xml {
                    file: ODF_CONTENT_XML.to_string(),
                    message: "unexpected end-of-file while parsing conditional formatting"
                        .to_string(),
                });
            }
            Err(e) => {
                return Err(ParseError::Xml {
                    file: ODF_CONTENT_XML.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    if cf.rules.is_empty() && cf.ranges.is_empty() {
        Ok(None)
    } else {
        Ok(Some(cf))
    }
}

pub(crate) fn parse_ods_conditional_formatting_empty(
    start: &BytesStart<'_>,
) -> Result<Option<ConditionalFormat>, ParseError> {
    let cf = init_conditional_format(start);
    if cf.rules.is_empty() && cf.ranges.is_empty() {
        Ok(None)
    } else {
        Ok(Some(cf))
    }
}

pub(crate) fn parse_odf_condition_operator(condition: &str) -> Option<String> {
    let lower = condition.to_ascii_lowercase();
    if let Some(idx) = lower.find("cell-content-is-") {
        let rest = &lower[idx + "cell-content-is-".len()..];
        let op = rest.split('(').next().unwrap_or(rest);
        return Some(op.to_string());
    }
    if let Some(idx) = lower.find("is-true-formula") {
        let _ = idx;
        return Some("true-formula".to_string());
    }
    if let Some(idx) = lower.find("formula-is") {
        let _ = idx;
        return Some("formula".to_string());
    }
    None
}

fn init_conditional_format(start: &BytesStart<'_>) -> ConditionalFormat {
    let mut cf = ConditionalFormat {
        id: NodeId::new(),
        ranges: Vec::new(),
        rules: Vec::new(),
        span: Some(SourceSpan::new(ODF_CONTENT_XML)),
    };
    if let Some(ranges) = attr_value(start, b"table:target-range-address")
        .or_else(|| attr_value(start, b"table:cell-range-address"))
    {
        cf.ranges = ranges.split_whitespace().map(|s| s.to_string()).collect();
    }
    cf
}

fn build_ods_conditional_rule(start: &BytesStart<'_>) -> ConditionalRule {
    let mut rule = ConditionalRule {
        rule_type: "odf-condition".to_string(),
        priority: None,
        operator: None,
        formulae: Vec::new(),
    };
    rule.priority = attr_value(start, b"table:priority").and_then(|v| v.parse::<u32>().ok());
    if let Some(condition) = attr_value(start, b"table:condition") {
        rule.operator = parse_odf_condition_operator(&condition);
        rule.formulae.push(condition);
    }
    if let Some(style_name) = attr_value(start, b"table:apply-style-name") {
        rule.formulae.push(format!("apply-style:{}", style_name));
    }
    rule
}

fn append_text_control(text: &mut String, e: &BytesStart<'_>) {
    match e.name().as_ref() {
        b"text:s" => {
            let count = attr_value(e, b"text:c")
                .and_then(|v| v.parse::<usize>().ok())
                .unwrap_or(1);
            text.extend(std::iter::repeat_n(' ', count));
        }
        b"text:tab" => text.push('\t'),
        b"text:line-break" => text.push('\n'),
        _ => {}
    }
}

pub(crate) fn parse_text_element(
    reader: &mut OdfReader<'_>,
    end_name: &[u8],
) -> Result<String, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    scan_xml_events_until_end(
        reader,
        &mut buf,
        "content.xml",
        move |event| matches!(event, Event::End(e) if e.name().as_ref() == end_name),
        |_reader, event| {
            match event {
                Event::Start(e) | Event::Empty(e) => append_text_control(&mut text, e),
                Event::Text(e) => {
                    let chunk = e.unescape().unwrap_or_default();
                    text.push_str(&chunk);
                }
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;
    Ok(text)
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct ListContext {
    pub(crate) num_id: u32,
    pub(crate) level: u32,
}

pub(crate) fn parse_table(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    helpers_parse_events_tables::parse_table(reader, store, limits)
}

pub(crate) fn parse_annotation(
    reader: &mut OdfReader<'_>,
    comment_id: &str,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    helpers_parse_events_changes::parse_annotation(reader, comment_id, store, limits)
}

pub(crate) fn parse_note(
    reader: &mut OdfReader<'_>,
    note_id: &str,
    note_class: &str,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    helpers_parse_events_changes::parse_note(reader, note_id, note_class, store, limits)
}

pub(crate) fn parse_draw_frame(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let mut shape = Shape::new(ShapeType::Picture);
    shape.transform = parse_frame_transform(start);
    shape.name = attr_value(start, b"draw:name");
    let mut buf = Vec::new();
    let mut has_shape = false;

    scan_xml_events_until_end(
        reader,
        &mut buf,
        "content.xml",
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"draw:frame"),
        |_reader, event| {
            match event {
                Event::Start(e) | Event::Empty(e) => match e.name().as_ref() {
                    b"draw:image" => {
                        if let Some(href) = attr_value(e, b"xlink:href") {
                            shape.media_target = Some(href);
                            shape.shape_type = ShapeType::Picture;
                            has_shape = true;
                        }
                    }
                    b"draw:object" | b"draw:object-ole" => {
                        if let Some(href) = attr_value(e, b"xlink:href") {
                            shape.media_target = Some(href);
                        }
                        shape.shape_type = ShapeType::OleObject;
                        has_shape = true;
                    }
                    b"draw:plugin" => {
                        if let Some(href) = attr_value(e, b"xlink:href") {
                            shape.media_target = Some(href.clone());
                            shape.shape_type = classify_media_shape(&href);
                            has_shape = true;
                        }
                    }
                    _ => {}
                },
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    if has_shape {
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
    }
}

pub(crate) fn parse_tracked_changes(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<Vec<NodeId>, ParseError> {
    helpers_parse_events_changes::parse_tracked_changes(reader, store, limits)
}
