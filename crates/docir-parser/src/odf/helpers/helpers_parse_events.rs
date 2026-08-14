use crate::error::ParseError;
#[path = "helpers_parse_events_changes.rs"]
mod helpers_parse_events_changes;
#[path = "helpers_parse_events_tables.rs"]
mod helpers_parse_events_tables;
use crate::odf::{
    OdfReader, limits::OdfLimitCounter, presentation_helpers::classify_media_shape,
    utils::parse_frame_transform,
};
use crate::xml_utils::{
    XmlScanControl, local_name, parse_bool_attr, scan_xml_events_until_end,
    scan_xml_events_with_reader, try_attr_value_by_suffix, xml_error,
};
use docir_core::ir::*;
use docir_core::types::*;
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};

const ODF_CONTENT_XML: &str = "content.xml";
pub(crate) const MAX_ODF_EXPANDED_SPACES: usize = 1_000_000;
pub(crate) const MAX_ODF_EXPANDED_TEXT: usize = 10 * 1024 * 1024;

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
    let mut reached_notes_end = false;
    scan_xml_events_with_reader(reader, &mut buf, ODF_CONTENT_XML, |reader, event| {
        match event {
            Event::Start(e) if local_name(e.name().as_ref()) == b"p" => {
                let para = parse_text_element(reader, e.name().as_ref())?;
                if !text.is_empty() {
                    text.push('\n');
                }
                text.push_str(&para);
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"notes" => {
                reached_notes_end = true;
                return Ok(XmlScanControl::Break);
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;
    if !reached_notes_end {
        return Err(xml_error(
            ODF_CONTENT_XML,
            "unexpected end of XML before closing notes",
        ));
    }

    if text.is_empty() {
        Ok(None)
    } else {
        Ok(Some(text))
    }
}

pub(crate) fn parse_validation_definition(
    start: &BytesStart<'_>,
) -> Result<Option<(String, ValidationDef)>, ParseError> {
    let Some(name) = try_attr_value_by_suffix(start, &[b":name"], "content.xml")? else {
        return Ok(None);
    };
    let condition = try_attr_value_by_suffix(start, &[b":condition"], "content.xml")?;
    let allow_blank = try_attr_value_by_suffix(start, &[b":allow-empty-cell"], "content.xml")?
        .map(|v| parse_bool_attr(v.as_bytes(), "content.xml"))
        .transpose()?
        .unwrap_or(false);
    let show_input_message = try_attr_value_by_suffix(start, &[b":display-list"], "content.xml")?
        .map(|v| parse_bool_attr(v.as_bytes(), "content.xml"))
        .transpose()?
        .unwrap_or(false);
    let show_error_message =
        try_attr_value_by_suffix(start, &[b":display-error-message"], "content.xml")?
            .map(|v| parse_bool_attr(v.as_bytes(), "content.xml"))
            .transpose()?
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
    Ok(Some((name, def)))
}

pub(crate) fn parse_ods_conditional_formatting(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<Option<ConditionalFormat>, ParseError> {
    let mut cf = init_conditional_format(start)?;

    let mut buf = Vec::new();
    let mut depth: usize = 1;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if depth == 1 && local_name(e.name().as_ref()) == b"conditional-format" {
                    cf.rules.push(build_ods_conditional_rule(&e)?);
                }
                depth = depth.saturating_add(1);
            }
            Ok(Event::Empty(e))
                if depth == 1 && local_name(e.name().as_ref()) == b"conditional-format" =>
            {
                cf.rules.push(build_ods_conditional_rule(&e)?);
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"conditional-formatting" && depth == 1 {
                    break;
                }
                depth = depth.saturating_sub(1);
            }
            Ok(Event::Eof) => {
                return Err(xml_error(
                    ODF_CONTENT_XML,
                    "unexpected end-of-file while parsing conditional formatting",
                ));
            }
            Err(e) => {
                return Err(xml_error(ODF_CONTENT_XML, e));
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
    let cf = init_conditional_format(start)?;
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

fn init_conditional_format(start: &BytesStart<'_>) -> Result<ConditionalFormat, ParseError> {
    let mut cf = ConditionalFormat {
        id: NodeId::new(),
        ranges: Vec::new(),
        rules: Vec::new(),
        span: Some(SourceSpan::new(ODF_CONTENT_XML)),
    };
    if let Some(ranges) =
        conditional_attr(start, &[b":target-range-address", b":cell-range-address"])?
    {
        cf.ranges = ranges.split_whitespace().map(|s| s.to_string()).collect();
    }
    Ok(cf)
}

fn build_ods_conditional_rule(start: &BytesStart<'_>) -> Result<ConditionalRule, ParseError> {
    let mut rule = ConditionalRule {
        rule_type: "odf-condition".to_string(),
        priority: None,
        operator: None,
        formulae: Vec::new(),
    };
    rule.priority = conditional_attr(start, &[b":priority"])?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|err| xml_error(ODF_CONTENT_XML, err))
        })
        .transpose()?;
    if let Some(condition) = conditional_attr(start, &[b":condition"])? {
        rule.operator = parse_odf_condition_operator(&condition);
        rule.formulae.push(condition);
    }
    if let Some(style_name) = conditional_attr(start, &[b":apply-style-name"])? {
        rule.formulae.push(format!("apply-style:{}", style_name));
    }
    Ok(rule)
}

fn conditional_attr(
    start: &BytesStart<'_>,
    suffixes: &[&[u8]],
) -> Result<Option<String>, ParseError> {
    for suffix in suffixes {
        if let Some(value) = try_attr_value_by_suffix(start, &[*suffix], ODF_CONTENT_XML)? {
            return Ok(Some(value));
        }
    }
    Ok(None)
}

fn append_text_control(text: &mut String, e: &BytesStart<'_>) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"s" => {
            let count = try_attr_value_by_suffix(e, &[b":c"], ODF_CONTENT_XML)?
                .map(|value| {
                    value
                        .parse::<usize>()
                        .map_err(|err| xml_error(ODF_CONTENT_XML, err))
                })
                .transpose()?
                .unwrap_or(1);
            append_odf_spaces(text, count)?;
        }
        b"tab" => text.push('\t'),
        b"line-break" => text.push('\n'),
        _ => {}
    }
    Ok(())
}

pub(crate) fn append_odf_spaces(text: &mut String, count: usize) -> Result<(), ParseError> {
    if count > MAX_ODF_EXPANDED_SPACES {
        return Err(ParseError::ResourceLimit(format!(
            "ODF expanded spaces exceed limit: {} (max: {})",
            count, MAX_ODF_EXPANDED_SPACES
        )));
    }
    let expanded_len = text.len().checked_add(count).ok_or_else(|| {
        ParseError::ResourceLimit("ODF expanded text length overflow".to_string())
    })?;
    if expanded_len > MAX_ODF_EXPANDED_TEXT {
        return Err(ParseError::ResourceLimit(format!(
            "ODF expanded text exceeds limit: {} (max: {})",
            expanded_len, MAX_ODF_EXPANDED_TEXT
        )));
    }
    text.extend(std::iter::repeat_n(' ', count));
    Ok(())
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
                Event::Start(e) | Event::Empty(e) => append_text_control(&mut text, e)?,
                Event::Text(e) => {
                    let chunk = crate::xml_utils::decoded_text(e)
                        .map_err(|err| xml_error(ODF_CONTENT_XML, err))?;
                    text.push_str(&chunk);
                }
                Event::GeneralRef(e) => {
                    let chunk = crate::xml_utils::decoded_general_ref(e)
                        .map_err(|err| xml_error(ODF_CONTENT_XML, err))?;
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

pub(crate) fn parse_empty_table(store: &mut IrStore) -> NodeId {
    helpers_parse_events_tables::parse_empty_table(store)
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
    shape.transform = parse_frame_transform(start)?;
    shape.name = try_attr_value_by_suffix(start, &[b":name"], ODF_CONTENT_XML)?;
    let mut buf = Vec::new();
    let mut has_shape = false;

    scan_xml_events_until_end(
        reader,
        &mut buf,
        "content.xml",
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"frame"),
        |_reader, event| {
            match event {
                Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"image" => {
                        if let Some(href) =
                            try_attr_value_by_suffix(e, &[b":href"], ODF_CONTENT_XML)?
                        {
                            shape.media_target = Some(href);
                            shape.shape_type = ShapeType::Picture;
                            has_shape = true;
                        }
                    }
                    b"object" | b"object-ole" => {
                        if let Some(href) =
                            try_attr_value_by_suffix(e, &[b":href"], ODF_CONTENT_XML)?
                        {
                            shape.media_target = Some(href);
                        }
                        shape.shape_type = ShapeType::OleObject;
                        has_shape = true;
                    }
                    b"plugin" => {
                        if let Some(href) =
                            try_attr_value_by_suffix(e, &[b":href"], ODF_CONTENT_XML)?
                        {
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
