use crate::xml_utils::{attr_value_by_suffix, local_name, scan_xml_events, XmlScanControl};
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

    let _ = scan_xml_events(&mut reader, &mut buf, "content.xml", |event| {
        match event {
            Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"named-range" => {
                    if let Some(def) = parse_ods_named_definition(&e, NamedDefinitionSource::Range)
                    {
                        out.push(def);
                    }
                }
                b"named-expression" => {
                    if let Some(def) =
                        parse_ods_named_definition(&e, NamedDefinitionSource::Expression)
                    {
                        out.push(def);
                    }
                }
                _ => {}
            },
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    });

    out
}

enum NamedDefinitionSource {
    Range,
    Expression,
}

fn parse_ods_named_definition(
    element: &BytesStart<'_>,
    source: NamedDefinitionSource,
) -> Option<DefinedName> {
    let name = attr_value_by_suffix(element, &[b":name"])?;
    let value = match source {
        NamedDefinitionSource::Range => attr_value_by_suffix(element, &[b":cell-range-address"]),
        NamedDefinitionSource::Expression => attr_value_by_suffix(element, &[b":expression"]),
    }
    .unwrap_or_default();

    let mut def = DefinedName {
        id: NodeId::new(),
        name,
        value,
        local_sheet_id: None,
        hidden: false,
        comment: attr_value_by_suffix(element, &[b":comment"]),
        span: Some(SourceSpan::new("content.xml")),
    };

    if let Some(hidden) = attr_value_by_suffix(element, &[b":hidden"]) {
        def.hidden = hidden == "true";
    }
    Some(def)
}

pub(crate) fn parse_frame_transform(start: &BytesStart<'_>) -> ShapeTransform {
    let mut transform = ShapeTransform::default();
    if let Some(x) = parse_length_emu_attr(start, b":x") {
        transform.x = x;
    }
    if let Some(y) = parse_length_emu_attr(start, b":y") {
        transform.y = y;
    }
    if let Some(width) = parse_length_emu_attr_u64(start, b":width") {
        transform.width = width;
    }
    if let Some(height) = parse_length_emu_attr_u64(start, b":height") {
        transform.height = height;
    }
    transform
}

fn parse_length_emu_attr(start: &BytesStart<'_>, key: &[u8]) -> Option<i64> {
    attr_value_by_suffix(start, &[key]).and_then(parse_length_emu)
}

fn parse_length_emu_attr_u64(start: &BytesStart<'_>, key: &[u8]) -> Option<u64> {
    attr_value_by_suffix(start, &[key]).and_then(parse_length_emu_u64)
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strip_odf_formula_prefix_handles_supported_forms() {
        assert_eq!(
            strip_odf_formula_prefix("of:=SUM([.A1:.A2])"),
            "SUM([.A1:.A2])"
        );
        assert_eq!(
            strip_odf_formula_prefix("of:SUM([.A1:.A2])"),
            "SUM([.A1:.A2])"
        );
        assert_eq!(strip_odf_formula_prefix("SUM([.A1:.A2])"), "SUM([.A1:.A2])");
    }

    #[test]
    fn parse_ods_named_ranges_extracts_ranges_expressions_and_hidden_flags() {
        let xml = br#"
            <office:document-content
                xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
              <table:named-expressions>
                <table:named-range table:name="RangeOne" table:cell-range-address="$Sheet1.$A$1:$A$4" table:comment="range comment" table:hidden="true"/>
                <table:named-expression table:name="ExprOne" table:expression="of:=SUM([.A1:.A4])" table:comment="expr comment"/>
                <table:named-range table:cell-range-address="$Sheet1.$B$1:$B$2"/>
              </table:named-expressions>
            </office:document-content>
        "#;

        let parsed = parse_ods_named_ranges(xml);
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].name, "RangeOne");
        assert_eq!(parsed[0].value, "$Sheet1.$A$1:$A$4");
        assert!(parsed[0].hidden);
        assert_eq!(parsed[0].comment.as_deref(), Some("range comment"));
        assert_eq!(parsed[1].name, "ExprOne");
        assert_eq!(parsed[1].value, "of:=SUM([.A1:.A4])");
        assert!(!parsed[1].hidden);
        assert_eq!(parsed[1].comment.as_deref(), Some("expr comment"));
        assert!(parsed[0].span.is_some());
    }

    #[test]
    fn parse_ods_named_ranges_returns_partial_results_on_xml_error() {
        let xml = br#"
            <office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
              <table:named-range table:name="Good" table:cell-range-address="$Sheet1.$C$1"/>
              <table:named-expression table:name="Broken" table:expression="of:=1+1"
        "#;

        let parsed = parse_ods_named_ranges(xml);
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].name, "Good");
    }

    #[test]
    fn parse_frame_transform_converts_supported_units_to_emu() {
        let mut reader = Reader::from_str(
            r#"<draw:frame svg:x="1in" svg:y="2.54cm" svg:width="25.4mm" svg:height="72pt" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"/>"#,
        );
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let start = match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => e.into_owned(),
            other => panic!("unexpected xml event: {other:?}"),
        };

        let transform = parse_frame_transform(&start);
        assert_eq!(transform.x, 914_400);
        assert_eq!(transform.y, 914_400);
        assert_eq!(transform.width, 914_400);
        assert_eq!(transform.height, 914_400);
    }

    #[test]
    fn parse_frame_transform_ignores_unsupported_or_invalid_units() {
        let mut reader = Reader::from_str(
            r#"<draw:frame svg:x="abc" svg:y="3q" svg:width="-2cm" svg:height="10px" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"/>"#,
        );
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let start = match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => e.into_owned(),
            other => panic!("unexpected xml event: {other:?}"),
        };

        let transform = parse_frame_transform(&start);
        assert_eq!(transform.x, 0);
        assert_eq!(transform.y, 0);
        assert_eq!(transform.width, 0);
        assert_eq!(transform.height, 95_250);
    }
}
