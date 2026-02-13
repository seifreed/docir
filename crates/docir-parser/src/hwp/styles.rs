use super::{
    attr_any, local_name, parse_hwpx_paragraph_props, parse_hwpx_table_props,
    run_properties_from_attrs, style_run_props_from_run,
};
use docir_core::ir::{Style, StyleSet, StyleType};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;

pub(super) fn parse_hwpx_styles(xml: &str, source: &str) -> Option<StyleSet> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut styles = Vec::new();
    let mut current: Option<Style> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"style" {
                    current = Some(parse_style_attrs(&e));
                } else {
                    apply_style_props(local, &e, &mut current);
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if local == b"style" {
                    styles.push(parse_style_attrs(&e));
                } else {
                    apply_style_props(local, &e, &mut current);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name().as_ref().to_vec();
                if local_name(&name) == b"style" {
                    styles.extend(current.take());
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    finalize_style_set(styles, source)
}

fn parse_style_attrs(e: &quick_xml::events::BytesStart<'_>) -> Style {
    let style_id =
        attr_any(e, &[b"id", b"styleId", b"style-id"]).unwrap_or_else(|| "style".to_string());
    let name = attr_any(e, &[b"name", b"styleName", b"style-name"]);
    let style_type = match attr_any(e, &[b"type", b"styleType"])
        .as_deref()
        .map(|v| v.to_ascii_lowercase())
    {
        Some(t) if t == "paragraph" => StyleType::Paragraph,
        Some(t) if t == "character" => StyleType::Character,
        Some(t) if t == "table" => StyleType::Table,
        _ => StyleType::Other,
    };
    Style {
        style_id,
        name,
        style_type,
        based_on: attr_any(e, &[b"basedOn", b"based-on"]),
        next: attr_any(e, &[b"next", b"next-style"]),
        is_default: attr_any(e, &[b"default", b"isDefault"])
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false),
        run_props: None,
        paragraph_props: None,
        table_props: None,
    }
}

fn apply_style_props(
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    current: &mut Option<Style>,
) {
    if let Some(style) = current.as_mut() {
        if local == b"charPr" || local == b"characterPr" {
            let run_props = run_properties_from_attrs(e);
            style.run_props = Some(style_run_props_from_run(run_props));
        } else if local == b"paraPr" || local == b"paragraphPr" {
            style.paragraph_props = Some(parse_hwpx_paragraph_props(e));
        } else if local == b"tblPr" || local == b"tablePr" {
            style.table_props = parse_hwpx_table_props(e);
        }
    }
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
