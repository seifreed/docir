use crate::xml_utils::attr_value;
use docir_core::ir::{DefinedName, ShapeTransform};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

pub(super) fn strip_odf_formula_prefix(formula: &str) -> &str {
    if let Some(stripped) = formula.strip_prefix("of:=") {
        stripped
    } else if let Some(stripped) = formula.strip_prefix("of:") {
        stripped
    } else {
        formula
    }
}

pub(super) fn parse_ods_named_ranges(xml: &[u8]) -> Vec<DefinedName> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:named-range" => {
                    if let Some(name) = attr_value(&e, b"table:name") {
                        let value = attr_value(&e, b"table:cell-range-address")
                            .unwrap_or_else(|| String::new());
                        let mut def = DefinedName {
                            id: NodeId::new(),
                            name,
                            value,
                            local_sheet_id: None,
                            hidden: false,
                            comment: attr_value(&e, b"table:comment"),
                            span: Some(SourceSpan::new("content.xml")),
                        };
                        if let Some(hidden) = attr_value(&e, b"table:hidden") {
                            def.hidden = hidden == "true";
                        }
                        out.push(def);
                    }
                }
                b"table:named-expression" => {
                    if let Some(name) = attr_value(&e, b"table:name") {
                        let value =
                            attr_value(&e, b"table:expression").unwrap_or_else(|| String::new());
                        let mut def = DefinedName {
                            id: NodeId::new(),
                            name,
                            value,
                            local_sheet_id: None,
                            hidden: false,
                            comment: attr_value(&e, b"table:comment"),
                            span: Some(SourceSpan::new("content.xml")),
                        };
                        if let Some(hidden) = attr_value(&e, b"table:hidden") {
                            def.hidden = hidden == "true";
                        }
                        out.push(def);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    out
}

pub(super) fn parse_frame_transform(start: &BytesStart<'_>) -> ShapeTransform {
    let mut transform = ShapeTransform::default();
    if let Some(x) = attr_value(start, b"svg:x").and_then(parse_length_emu) {
        transform.x = x;
    }
    if let Some(y) = attr_value(start, b"svg:y").and_then(parse_length_emu) {
        transform.y = y;
    }
    if let Some(width) = attr_value(start, b"svg:width").and_then(parse_length_emu_u64) {
        transform.width = width;
    }
    if let Some(height) = attr_value(start, b"svg:height").and_then(parse_length_emu_u64) {
        transform.height = height;
    }
    transform
}

fn parse_length_emu(value: String) -> Option<i64> {
    parse_length_emu_str(&value).map(|v| v.round() as i64)
}

fn parse_length_emu_u64(value: String) -> Option<u64> {
    parse_length_emu_str(&value).map(|v| v.max(0.0).round() as u64)
}

fn parse_length_emu_str(value: &str) -> Option<f64> {
    let trimmed = value.trim();
    let mut num = String::new();
    let mut unit = String::new();
    for ch in trimmed.chars() {
        if ch.is_ascii_digit() || ch == '.' || ch == '-' {
            num.push(ch);
        } else {
            unit.push(ch);
        }
    }
    let magnitude = num.parse::<f64>().ok()?;
    let emu = match unit.as_str() {
        "cm" => magnitude / 2.54 * 914_400.0,
        "mm" => magnitude / 25.4 * 914_400.0,
        "in" => magnitude * 914_400.0,
        "pt" => magnitude * 12_700.0,
        "pc" => magnitude * 152_400.0,
        "px" => magnitude * 9_525.0,
        "" => magnitude,
        _ => return None,
    };
    Some(emu)
}
