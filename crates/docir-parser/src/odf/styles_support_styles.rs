use super::super::{Style, StyleSet, StyleType, parse_text_alignment};
use crate::error::ParseError;
use crate::xml_utils::xml_error;
use crate::xml_utils::{local_name, parse_bool_attr, try_attr_value_by_suffix};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

pub(crate) fn parse_styles(xml: &str, source: &str) -> Result<Option<StyleSet>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut styles = StyleSet::new();
    let mut depth = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"style" => {
                    if let Some(mut style) = build_style_from_start(&e, false, source)? {
                        parse_style_properties(&mut reader, &mut style, e.name().as_ref(), source)?;
                        styles.styles.push(style);
                    } else {
                        skip_element(&mut reader, e.name().as_ref(), source)?;
                    }
                }
                b"default-style" => {
                    if let Some(mut style) = build_style_from_start(&e, true, source)? {
                        parse_style_properties(&mut reader, &mut style, e.name().as_ref(), source)?;
                        styles.styles.push(style);
                    } else {
                        skip_element(&mut reader, e.name().as_ref(), source)?;
                    }
                }
                _ => depth += 1,
            },
            Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"style" => {
                    if let Some(style) = build_style_from_start(&e, false, source)? {
                        styles.styles.push(style);
                    }
                }
                b"default-style" => {
                    if let Some(style) = build_style_from_start(&e, true, source)? {
                        styles.styles.push(style);
                    }
                }
                _ => {}
            },
            Ok(Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(Event::Eof) if depth == 0 => break,
            Ok(Event::Eof) => {
                return Err(xml_error(source, "Unexpected EOF while parsing ODF styles"));
            }
            Err(err) => return Err(xml_error(source, err)),
            _ => {}
        }
        buf.clear();
    }

    if styles.styles.is_empty() {
        Ok(None)
    } else {
        Ok(Some(styles))
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

pub(crate) fn parse_master_pages(xml: &str, source: &str) -> Result<Vec<String>, ParseError> {
    parse_named_elements(xml, source, b"master-page")
}

pub(crate) fn parse_page_layouts(xml: &str, source: &str) -> Result<Vec<String>, ParseError> {
    parse_named_elements(xml, source, b"page-layout")
}

fn parse_named_elements(
    xml: &str,
    source: &str,
    target_name: &[u8],
) -> Result<Vec<String>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                depth += 1;
                if local_name(e.name().as_ref()) == target_name
                    && let Some(name) = try_attr_value_by_suffix(&e, &[b":name"], source)?
                {
                    out.push(name);
                }
            }
            Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == target_name
                    && let Some(name) = try_attr_value_by_suffix(&e, &[b":name"], source)?
                {
                    out.push(name);
                }
            }
            Ok(Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(Event::Eof) if depth == 0 => break,
            Ok(Event::Eof) => {
                return Err(xml_error(source, "Unexpected EOF while parsing ODF styles"));
            }
            Err(err) => return Err(xml_error(source, err)),
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn map_style_family(e: &BytesStart<'_>, source: &str) -> Result<StyleType, ParseError> {
    Ok(
        match try_attr_value_by_suffix(e, &[b":family"], source)?.as_deref() {
            Some("paragraph") => StyleType::Paragraph,
            Some("text") => StyleType::Character,
            Some("table") => StyleType::Table,
            Some("list") => StyleType::Numbering,
            _ => StyleType::Other,
        },
    )
}

fn build_style_from_start(
    start: &BytesStart<'_>,
    is_default: bool,
    source: &str,
) -> Result<Option<Style>, ParseError> {
    let style_id =
        try_attr_value_by_suffix(start, &[b":name"], source)?.or(try_attr_value_by_suffix(
            start,
            &[b":family"],
            source,
        )?
        .map(|f| format!("default:{f}")));
    let Some(style_id) = style_id else {
        return Ok(None);
    };
    let mut style = Style {
        style_id,
        name: try_attr_value_by_suffix(start, &[b":display-name"], source)?,
        style_type: map_style_family(start, source)?,
        based_on: try_attr_value_by_suffix(start, &[b":parent-style-name"], source)?,
        next: try_attr_value_by_suffix(start, &[b":next-style-name"], source)?,
        is_default,
        run_props: None,
        paragraph_props: None,
        table_props: None,
    };
    if let Some(family) = try_attr_value_by_suffix(start, &[b":family"], source)?
        && (family == "paragraph" || family == "text")
        && let Some(value) = try_attr_value_by_suffix(start, &[b":default"], source)?
    {
        style.is_default = is_default || parse_bool_attr(value.as_bytes(), source)?;
    }
    Ok(Some(style))
}

fn parse_style_properties(
    reader: &mut Reader<&[u8]>,
    style: &mut Style,
    end_name: &[u8],
    source: &str,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"text-properties" => {
                    let mut props = style.run_props.take().unwrap_or_default();
                    if let Some(font) = try_attr_value_by_suffix(&e, &[b":font-family"], source)?
                        .or(try_attr_value_by_suffix(&e, &[b":font-name"], source)?)
                    {
                        props.font_family = Some(font);
                    }
                    if let Some(size) = try_attr_value_by_suffix(&e, &[b":font-size"], source)?
                        .and_then(|v| parse_font_size(&v))
                    {
                        props.font_size = Some(size);
                    }
                    if let Some(weight) = try_attr_value_by_suffix(&e, &[b":font-weight"], source)?
                    {
                        props.bold = Some(weight.eq_ignore_ascii_case("bold"));
                    }
                    if let Some(style_attr) =
                        try_attr_value_by_suffix(&e, &[b":font-style"], source)?
                    {
                        props.italic = Some(style_attr.eq_ignore_ascii_case("italic"));
                    }
                    if let Some(color) = try_attr_value_by_suffix(&e, &[b":color"], source)? {
                        props.color = Some(color);
                    }
                    style.run_props = Some(props);
                }
                b"paragraph-properties" => {
                    let mut props = style.paragraph_props.take().unwrap_or_default();
                    if let Some(align) = try_attr_value_by_suffix(&e, &[b":text-align"], source)?
                        .and_then(|v| parse_text_alignment(&v))
                    {
                        props.alignment = Some(align);
                    }
                    style.paragraph_props = Some(props);
                }
                _ => {}
            },
            Ok(Event::End(e)) if e.name().as_ref() == end_name => {
                break;
            }
            Ok(Event::Eof) => {
                return Err(xml_error(source, "Unexpected EOF while parsing ODF style"));
            }
            Err(err) => return Err(xml_error(source, err)),
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn skip_element(
    reader: &mut Reader<&[u8]>,
    end_name: &[u8],
    source: &str,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(_)) => depth += 1,
            Ok(Event::End(e)) if depth == 0 && e.name().as_ref() == end_name => break,
            Ok(Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(Event::Eof) => {
                return Err(xml_error(source, "Unexpected EOF while parsing ODF style"));
            }
            Err(err) => return Err(xml_error(source, err)),
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_font_size(value: &str) -> Option<u32> {
    let trimmed = value.trim();
    for unit in ["pt", "px", "cm", "mm"] {
        if let Some(num) = trimmed.strip_suffix(unit) {
            return parse_finite_font_size(num);
        }
    }
    parse_finite_font_size(trimmed)
}

fn parse_finite_font_size(value: &str) -> Option<u32> {
    let size = value.parse::<f32>().ok()?;
    if size.is_finite() && size >= 0.0 {
        let rounded = f64::from(size.round());
        (rounded <= u32::MAX as f64).then_some(rounded as u32)
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::parse_font_size;

    #[test]
    fn parse_font_size_rejects_non_finite_values() {
        assert_eq!(parse_font_size("NaNpt"), None);
        assert_eq!(parse_font_size("inf"), None);
    }

    #[test]
    fn parse_font_size_rejects_negative_values() {
        assert_eq!(parse_font_size("-1pt"), None);
    }

    #[test]
    fn parse_font_size_rejects_values_that_overflow_u32() {
        assert_eq!(parse_font_size("4294967296pt"), None);
    }
}
