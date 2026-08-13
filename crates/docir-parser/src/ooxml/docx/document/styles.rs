use super::DocxParser;
use crate::error::ParseError;
use crate::xml_utils::{local_name, track_xml_root_event, try_attr_value, xml_error};
use docir_core::ir::{Paragraph, RunProperties, Style, StyleSet, StyleType};
use docir_core::types::NodeId;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use super::paragraph::parse_paragraph_properties;
use super::table::parse_table_properties;

const STYLES_PATH: &str = "word/styles.xml";

impl DocxParser {
    /// Public API entrypoint: parse_styles.
    pub fn parse_styles(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut styles = StyleSet::new();
        let mut reader = Reader::from_str(xml);
        let config = reader.config_mut();
        config.trim_text(true);
        config.check_end_names = true;
        let mut buf = Vec::new();

        let mut current: Option<Style> = None;
        let mut in_name = false;
        let mut root_name = None;
        let mut root_depth = 0;
        let mut root_closed = false;

        loop {
            let event = reader.read_event_into(&mut buf);
            if let Ok(event) = &event {
                track_xml_root_event(
                    event,
                    &mut root_name,
                    &mut root_depth,
                    &mut root_closed,
                    STYLES_PATH,
                )?;
            }

            if !handle_style_event(event, &mut reader, &mut styles, &mut current, &mut in_name)? {
                break;
            }
            buf.clear();
        }

        let id = styles.id;
        self.store.insert(docir_core::ir::IRNode::StyleSet(styles));
        Ok(id)
    }

    /// Public API entrypoint: parse_styles_with_effects.
    pub fn parse_styles_with_effects(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let id = self.parse_styles(xml)?;
        if let Some(docir_core::ir::IRNode::StyleSet(set)) = self.store.get_mut(id) {
            set.with_effects = true;
        }
        Ok(id)
    }
}

fn handle_style_event(
    event: Result<Event<'_>, quick_xml::Error>,
    reader: &mut Reader<&[u8]>,
    styles: &mut StyleSet,
    current: &mut Option<Style>,
    in_name: &mut bool,
) -> Result<bool, ParseError> {
    match event {
        Ok(Event::Start(e)) => handle_style_start(reader, &e, current, in_name)?,
        Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"style" => {
            styles.styles.push(build_style(&e)?);
        }
        Ok(Event::Empty(e)) => handle_style_empty(&e, current)?,
        Ok(Event::Text(e)) => handle_style_text(&e, current, *in_name)?,
        Ok(Event::GeneralRef(e)) => handle_style_general_ref(&e, current, *in_name)?,
        Ok(Event::End(e)) => finish_style_event(&e, styles, current, in_name),
        Ok(Event::Eof) => return Ok(false),
        Err(e) => return Err(xml_error(STYLES_PATH, e)),
        _ => {}
    }
    Ok(true)
}

fn handle_style_text(
    event: &quick_xml::events::BytesText<'_>,
    current: &mut Option<Style>,
    in_name: bool,
) -> Result<(), ParseError> {
    if in_name && let Some(style) = current.as_mut() {
        style.name =
            Some(crate::xml_utils::decoded_text(event).map_err(|err| xml_error(STYLES_PATH, err))?);
    }
    Ok(())
}

fn handle_style_general_ref(
    event: &quick_xml::events::BytesRef<'_>,
    current: &mut Option<Style>,
    in_name: bool,
) -> Result<(), ParseError> {
    if in_name && let Some(style) = current.as_mut() {
        style.name = Some(
            crate::xml_utils::decoded_general_ref(event)
                .map_err(|err| xml_error(STYLES_PATH, err))?,
        );
    }
    Ok(())
}

fn finish_style_event(
    event: &quick_xml::events::BytesEnd<'_>,
    styles: &mut StyleSet,
    current: &mut Option<Style>,
    in_name: &mut bool,
) {
    if local_name(event.name().as_ref()) == b"name" {
        *in_name = false;
    } else if local_name(event.name().as_ref()) == b"style"
        && let Some(style) = current.take()
    {
        styles.styles.push(style);
    }
}

fn handle_style_start(
    reader: &mut Reader<&[u8]>,
    event: &BytesStart<'_>,
    current: &mut Option<Style>,
    in_name: &mut bool,
) -> Result<(), ParseError> {
    match local_name(event.name().as_ref()) {
        b"style" => {
            *current = Some(build_style(event)?);
        }
        b"name" => {
            *in_name = true;
        }
        b"rPr" => {
            let mut props = RunProperties::default();
            super::parse_run_properties(reader, &mut props)?;
            if let Some(style) = current.as_mut() {
                style.run_props = Some(super::style_run_from_run_props(props));
            }
        }
        b"pPr" => {
            let mut para = Paragraph::new();
            let _ = parse_paragraph_properties(reader, &mut para, None)?;
            if let Some(style) = current.as_mut() {
                style.paragraph_props =
                    Some(super::style_paragraph_from_paragraph_props(para.properties));
            }
        }
        b"tblPr" => {
            if let Some(style) = current.as_mut() {
                let mut props = docir_core::ir::TableProperties::default();
                parse_table_properties(reader, &mut props)?;
                style.table_props = Some(props);
            }
        }
        b"basedOn" => assign_style_attr(event, current.as_mut(), StyleAttr::BasedOn)?,
        b"next" => assign_style_attr(event, current.as_mut(), StyleAttr::Next)?,
        _ => {}
    }
    Ok(())
}

fn handle_style_empty(
    event: &BytesStart<'_>,
    current: &mut Option<Style>,
) -> Result<(), ParseError> {
    match local_name(event.name().as_ref()) {
        b"name" => assign_style_attr(event, current.as_mut(), StyleAttr::Name)?,
        b"basedOn" => assign_style_attr(event, current.as_mut(), StyleAttr::BasedOn)?,
        b"next" => assign_style_attr(event, current.as_mut(), StyleAttr::Next)?,
        _ => {}
    }
    Ok(())
}

fn build_style(event: &BytesStart<'_>) -> Result<Style, ParseError> {
    let style_id = try_attr_value(event, b"w:styleId", STYLES_PATH)?.ok_or_else(|| {
        ParseError::InvalidStructure("word/styles.xml style is missing w:styleId".to_string())
    })?;
    let mut style = Style {
        style_id,
        name: None,
        style_type: StyleType::Other,
        based_on: None,
        next: None,
        is_default: try_attr_value(event, b"w:default", STYLES_PATH)?
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false),
        run_props: None,
        paragraph_props: None,
        table_props: None,
    };
    if let Some(t) = try_attr_value(event, b"w:type", STYLES_PATH)? {
        style.style_type = match t.as_str() {
            "paragraph" => StyleType::Paragraph,
            "character" => StyleType::Character,
            "table" => StyleType::Table,
            "numbering" => StyleType::Numbering,
            _ => StyleType::Other,
        };
    }
    Ok(style)
}

enum StyleAttr {
    Name,
    BasedOn,
    Next,
}

fn assign_style_attr(
    event: &BytesStart<'_>,
    style: Option<&mut Style>,
    attr: StyleAttr,
) -> Result<(), ParseError> {
    let Some(style) = style else {
        return Ok(());
    };
    let Some(val) = try_attr_value(event, b"w:val", STYLES_PATH)? else {
        return Ok(());
    };
    match attr {
        StyleAttr::Name => style.name = Some(val),
        StyleAttr::BasedOn => style.based_on = Some(val),
        StyleAttr::Next => style.next = Some(val),
    }
    Ok(())
}
