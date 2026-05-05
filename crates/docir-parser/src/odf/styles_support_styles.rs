use super::super::{parse_text_alignment, Style, StyleSet, StyleType};
use crate::xml_utils::{attr_value_by_suffix, local_name};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

pub(crate) fn parse_styles(xml: &str) -> Option<StyleSet> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut styles = StyleSet::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"style" => {
                    if let Some(mut style) = build_style_from_start(&e, false) {
                        parse_style_properties(&mut reader, &mut style, e.name().as_ref());
                        styles.styles.push(style);
                    }
                }
                b"default-style" => {
                    if let Some(mut style) = build_style_from_start(&e, true) {
                        parse_style_properties(&mut reader, &mut style, e.name().as_ref());
                        styles.styles.push(style);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"style" => {
                    if let Some(style) = build_style_from_start(&e, false) {
                        styles.styles.push(style);
                    }
                }
                b"default-style" => {
                    if let Some(style) = build_style_from_start(&e, true) {
                        styles.styles.push(style);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }

    if styles.styles.is_empty() {
        None
    } else {
        Some(styles)
    }
}

pub(crate) fn merge_styles(existing: &mut StyleSet, incoming: &mut StyleSet) {
    let mut seen = existing
        .styles
        .iter()
        .map(|s| s.style_id.clone())
        .collect::<std::collections::HashSet<String>>();
    for style in incoming.styles.drain(..) {
        if seen.insert(style.style_id.clone()) {
            existing.styles.push(style);
        }
    }
}

pub(crate) fn parse_master_pages(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == b"master-page" {
                    if let Some(name) = attr_value_by_suffix(&e, &[b":name"]) {
                        out.push(name);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

pub(crate) fn parse_page_layouts(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == b"page-layout" {
                    if let Some(name) = attr_value_by_suffix(&e, &[b":name"]) {
                        out.push(name);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn map_style_family(e: &BytesStart<'_>) -> StyleType {
    match attr_value_by_suffix(e, &[b":family"]).as_deref() {
        Some("paragraph") => StyleType::Paragraph,
        Some("text") => StyleType::Character,
        Some("table") => StyleType::Table,
        Some("list") => StyleType::Numbering,
        _ => StyleType::Other,
    }
}

fn build_style_from_start(start: &BytesStart<'_>, is_default: bool) -> Option<Style> {
    let style_id = attr_value_by_suffix(start, &[b":name"])
        .or_else(|| attr_value_by_suffix(start, &[b":family"]).map(|f| format!("default:{f}")));
    let style_id = style_id?;
    let mut style = Style {
        style_id,
        name: attr_value_by_suffix(start, &[b":display-name"]),
        style_type: map_style_family(start),
        based_on: attr_value_by_suffix(start, &[b":parent-style-name"]),
        next: attr_value_by_suffix(start, &[b":next-style-name"]),
        is_default,
        run_props: None,
        paragraph_props: None,
        table_props: None,
    };
    if let Some(family) = attr_value_by_suffix(start, &[b":family"]) {
        if family == "paragraph" || family == "text" {
            style.is_default = is_default
                || attr_value_by_suffix(start, &[b":default"])
                    .map(|v| v == "true")
                    .unwrap_or(false);
        }
    }
    Some(style)
}

fn parse_style_properties(reader: &mut Reader<&[u8]>, style: &mut Style, end_name: &[u8]) {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"text-properties" => {
                    let mut props = style.run_props.take().unwrap_or_default();
                    if let Some(font) = attr_value_by_suffix(&e, &[b":font-family"])
                        .or_else(|| attr_value_by_suffix(&e, &[b":font-name"]))
                    {
                        props.font_family = Some(font);
                    }
                    if let Some(size) =
                        attr_value_by_suffix(&e, &[b":font-size"]).and_then(|v| parse_font_size(&v))
                    {
                        props.font_size = Some(size);
                    }
                    if let Some(weight) = attr_value_by_suffix(&e, &[b":font-weight"]) {
                        props.bold = Some(weight.eq_ignore_ascii_case("bold"));
                    }
                    if let Some(style_attr) = attr_value_by_suffix(&e, &[b":font-style"]) {
                        props.italic = Some(style_attr.eq_ignore_ascii_case("italic"));
                    }
                    if let Some(color) = attr_value_by_suffix(&e, &[b":color"]) {
                        props.color = Some(color);
                    }
                    style.run_props = Some(props);
                }
                b"paragraph-properties" => {
                    let mut props = style.paragraph_props.take().unwrap_or_default();
                    if let Some(align) = attr_value_by_suffix(&e, &[b":text-align"])
                        .and_then(|v| parse_text_alignment(&v))
                    {
                        props.alignment = Some(align);
                    }
                    style.paragraph_props = Some(props);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_name {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn parse_font_size(value: &str) -> Option<u32> {
    let trimmed = value.trim();
    let num = trimmed
        .trim_end_matches("pt")
        .trim_end_matches("px")
        .trim_end_matches("cm")
        .trim_end_matches("mm");
    num.parse::<f32>().ok().map(|v| v.round() as u32)
}
