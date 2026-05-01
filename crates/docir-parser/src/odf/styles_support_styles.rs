use super::super::{attr_value, parse_text_alignment, Style, StyleSet, StyleType};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

pub(crate) fn parse_styles(xml: &str) -> Option<StyleSet> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut styles = StyleSet::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"style:style" => {
                    if let Some(mut style) = build_style_from_start(&e, false) {
                        parse_style_properties(&mut reader, &mut style, b"style:style");
                        styles.styles.push(style);
                    }
                }
                b"style:default-style" => {
                    if let Some(mut style) = build_style_from_start(&e, true) {
                        parse_style_properties(&mut reader, &mut style, b"style:default-style");
                        styles.styles.push(style);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"style:style" => {
                    if let Some(style) = build_style_from_start(&e, false) {
                        styles.styles.push(style);
                    }
                }
                b"style:default-style" => {
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
                if e.name().as_ref() == b"style:master-page" {
                    if let Some(name) = attr_value(&e, b"style:name") {
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
                if e.name().as_ref() == b"style:page-layout" {
                    if let Some(name) = attr_value(&e, b"style:name") {
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
    match attr_value(e, b"style:family").as_deref() {
        Some("paragraph") => StyleType::Paragraph,
        Some("text") => StyleType::Character,
        Some("table") => StyleType::Table,
        Some("list") => StyleType::Numbering,
        _ => StyleType::Other,
    }
}

fn build_style_from_start(start: &BytesStart<'_>, is_default: bool) -> Option<Style> {
    let style_id = attr_value(start, b"style:name")
        .or_else(|| attr_value(start, b"style:family").map(|f| format!("default:{f}")));
    let style_id = style_id?;
    let mut style = Style {
        style_id,
        name: attr_value(start, b"style:display-name"),
        style_type: map_style_family(start),
        based_on: attr_value(start, b"style:parent-style-name"),
        next: attr_value(start, b"style:next-style-name"),
        is_default,
        run_props: None,
        paragraph_props: None,
        table_props: None,
    };
    if let Some(family) = attr_value(start, b"style:family") {
        if family == "paragraph" || family == "text" {
            style.is_default = is_default
                || attr_value(start, b"style:default")
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
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"style:text-properties" => {
                    let mut props = style.run_props.take().unwrap_or_default();
                    if let Some(font) = attr_value(&e, b"fo:font-family")
                        .or_else(|| attr_value(&e, b"style:font-name"))
                    {
                        props.font_family = Some(font);
                    }
                    if let Some(size) =
                        attr_value(&e, b"fo:font-size").and_then(|v| parse_font_size(&v))
                    {
                        props.font_size = Some(size);
                    }
                    if let Some(weight) = attr_value(&e, b"fo:font-weight") {
                        props.bold = Some(weight.eq_ignore_ascii_case("bold"));
                    }
                    if let Some(style_attr) = attr_value(&e, b"fo:font-style") {
                        props.italic = Some(style_attr.eq_ignore_ascii_case("italic"));
                    }
                    if let Some(color) = attr_value(&e, b"fo:color") {
                        props.color = Some(color);
                    }
                    style.run_props = Some(props);
                }
                b"style:paragraph-properties" => {
                    let mut props = style.paragraph_props.take().unwrap_or_default();
                    if let Some(align) =
                        attr_value(&e, b"fo:text-align").and_then(|v| parse_text_alignment(&v))
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
