use super::{
    attr_any, local_name, parse_hwpx_paragraph_props, parse_hwpx_table_props,
    run_properties_from_attrs, style_run_props_from_run,
};
use crate::error::ParseError;
use crate::xml_utils::{
    XmlScanControl, parse_bool_attr, reader_from_str, scan_xml_events, track_xml_document_event,
};
use docir_core::ir::{Style, StyleSet, StyleType};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;

pub(super) fn parse_hwpx_styles(xml: &str, source: &str) -> Result<Option<StyleSet>, ParseError> {
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut styles = Vec::new();
    let mut current: Option<Style> = None;
    let mut depth = 0usize;
    let mut root_closed = false;

    scan_xml_events(&mut reader, &mut buf, source, |event| {
        track_xml_document_event(&event, &mut depth, &mut root_closed, source)?;
        match event {
            Event::Start(e) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"style" {
                    current = Some(parse_style_attrs(&e, source)?);
                } else {
                    apply_style_props(local, &e, source, &mut current)?;
                }
            }
            Event::Empty(e) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"style" {
                    styles.push(parse_style_attrs(&e, source)?);
                } else {
                    apply_style_props(local, &e, source, &mut current)?;
                }
            }
            Event::End(e) => {
                let name = e.name().as_ref().to_vec();
                if local_name(&name) == b"style" {
                    styles.extend(current.take());
                }
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(finalize_style_set(styles, source))
}

fn parse_style_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    source: &str,
) -> Result<Style, ParseError> {
    let style_id = attr_any(e, &[b"id", b"styleId", b"style-id"], source)?
        .unwrap_or_else(|| "style".to_string());
    let name = attr_any(e, &[b"name", b"styleName", b"style-name"], source)?;
    let style_type = match attr_any(e, &[b"type", b"styleType"], source)?
        .as_deref()
        .map(|v| v.to_ascii_lowercase())
    {
        Some(t) if t == "paragraph" => StyleType::Paragraph,
        Some(t) if t == "character" => StyleType::Character,
        Some(t) if t == "table" => StyleType::Table,
        _ => StyleType::Other,
    };
    Ok(Style {
        style_id,
        name,
        style_type,
        based_on: attr_any(e, &[b"basedOn", b"based-on"], source)?,
        next: attr_any(e, &[b"next", b"next-style"], source)?,
        is_default: match attr_any(e, &[b"default", b"isDefault"], source)? {
            Some(value) => parse_bool_attr(value.as_bytes(), source)?,
            None => false,
        },
        run_props: None,
        paragraph_props: None,
        table_props: None,
    })
}

fn apply_style_props(
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    source: &str,
    current: &mut Option<Style>,
) -> Result<(), ParseError> {
    if let Some(style) = current.as_mut() {
        if local == b"charPr" || local == b"characterPr" {
            let run_props = run_properties_from_attrs(e, source)?;
            style.run_props = Some(style_run_props_from_run(run_props));
        } else if local == b"paraPr" || local == b"paragraphPr" {
            style.paragraph_props = Some(parse_hwpx_paragraph_props(e, source)?);
        } else if local == b"tblPr" || local == b"tablePr" {
            style.table_props = parse_hwpx_table_props(e, source)?;
        }
    }
    Ok(())
}

fn finalize_style_set(styles: Vec<Style>, source: &str) -> Option<StyleSet> {
    if styles.is_empty() {
        return None;
    }

    let mut set = StyleSet::new();
    set.styles = styles;
    set.span = Some(SourceSpan::new(source));
    Some(set)
}
