use crate::error::ParseError;
use crate::xml_utils::{
    XmlScanControl, local_name, parse_bool_attr, scan_xml_events, track_xml_document_event,
    try_attr_value_by_suffix,
};
use docir_core::ir::{DefinedName, ShapeTransform};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

pub(super) fn strip_odf_formula_prefix(formula: &str) -> &str {
    if let Some(stripped) = formula.strip_prefix("of:=") {
        stripped
    } else if let Some(stripped) = formula.strip_prefix("of:") {
        stripped
    } else {
        formula
    }
}

pub(super) fn parse_ods_named_ranges(xml: &[u8]) -> Result<Vec<DefinedName>, ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;

    scan_xml_events(&mut reader, &mut buf, "content.xml", |event| {
        track_xml_document_event(&event, &mut depth, &mut root_closed, "content.xml")?;
        match event {
            Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"named-range" => {
                    if let Some(def) = parse_ods_named_definition(&e, NamedDefinitionSource::Range)?
                    {
                        out.push(def);
                    }
                }
                b"named-expression" => {
                    if let Some(def) =
                        parse_ods_named_definition(&e, NamedDefinitionSource::Expression)?
                    {
                        out.push(def);
                    }
                }
                _ => {}
            },
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(out)
}

enum NamedDefinitionSource {
    Range,
    Expression,
}

fn parse_ods_named_definition(
    element: &BytesStart<'_>,
    source: NamedDefinitionSource,
) -> Result<Option<DefinedName>, ParseError> {
    let Some(name) = try_attr_value_by_suffix(element, &[b":name"], "content.xml")? else {
        return Ok(None);
    };
    let value = match source {
        NamedDefinitionSource::Range => {
            try_attr_value_by_suffix(element, &[b":cell-range-address"], "content.xml")?
        }
        NamedDefinitionSource::Expression => {
            try_attr_value_by_suffix(element, &[b":expression"], "content.xml")?
        }
    }
    .unwrap_or_default();

    let mut def = DefinedName {
        id: NodeId::new(),
        name,
        value,
        local_sheet_id: None,
        hidden: false,
        comment: try_attr_value_by_suffix(element, &[b":comment"], "content.xml")?,
        span: Some(SourceSpan::new("content.xml")),
    };

    if let Some(hidden) = try_attr_value_by_suffix(element, &[b":hidden"], "content.xml")? {
        def.hidden = parse_bool_attr(hidden.as_bytes(), "content.xml")?;
    }
    Ok(Some(def))
}

pub(crate) fn parse_frame_transform(start: &BytesStart<'_>) -> Result<ShapeTransform, ParseError> {
    let mut transform = ShapeTransform::default();
    if let Some(x) = parse_length_emu_attr(start, b":x")? {
        transform.x = x;
    }
    if let Some(y) = parse_length_emu_attr(start, b":y")? {
        transform.y = y;
    }
    if let Some(width) = parse_length_emu_attr_u64(start, b":width")? {
        transform.width = width;
    }
    if let Some(height) = parse_length_emu_attr_u64(start, b":height")? {
        transform.height = height;
    }
    Ok(transform)
}

fn parse_length_emu_attr(start: &BytesStart<'_>, key: &[u8]) -> Result<Option<i64>, ParseError> {
    Ok(try_attr_value_by_suffix(start, &[key], "content.xml")?.and_then(parse_length_emu))
}

fn parse_length_emu_attr_u64(
    start: &BytesStart<'_>,
    key: &[u8],
) -> Result<Option<u64>, ParseError> {
    Ok(try_attr_value_by_suffix(start, &[key], "content.xml")?.and_then(parse_length_emu_u64))
}

fn parse_length_emu(value: String) -> Option<i64> {
    let rounded = parse_length_emu_str(&value)?.round();
    (rounded >= i64::MIN as f64 && rounded < i64::MAX as f64).then_some(rounded as i64)
}

fn parse_length_emu_u64(value: String) -> Option<u64> {
    let rounded = parse_length_emu_str(&value)?.max(0.0).round();
    (rounded < u64::MAX as f64).then_some(rounded as u64)
}

fn parse_length_emu_str(value: &str) -> Option<f64> {
    let trimmed = value.trim();
    let mut unit_start = trimmed.len();
    for (idx, ch) in trimmed.char_indices() {
        let sign = (ch == '-' || ch == '+') && idx == 0;
        if !(sign || ch.is_ascii_digit() || ch == '.') {
            unit_start = idx;
            break;
        }
    }
    let (num, unit) = trimmed.split_at(unit_start);
    let magnitude = num.parse::<f64>().ok()?;
    if !magnitude.is_finite() {
        return None;
    }
    let emu = match unit {
        "cm" => magnitude / 2.54 * 914_400.0,
        "mm" => magnitude / 25.4 * 914_400.0,
        "in" => magnitude * 914_400.0,
        "pt" => magnitude * 12_700.0,
        "pc" => magnitude * 152_400.0,
        "px" => magnitude * 9_525.0,
        "" => magnitude,
        _ => return None,
    };
    emu.is_finite().then_some(emu)
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

        let parsed = parse_ods_named_ranges(xml).expect("named ranges");
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
    fn parse_ods_named_ranges_reports_xml_error() {
        let xml = br#"
            <office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
              <table:named-range table:name="Good" table:cell-range-address="$Sheet1.$C$1"/>
              <table:named-expression table:name="Broken" table:expression="of:=1+1"
        "#;

        let err = parse_ods_named_ranges(xml).expect_err("malformed named range XML must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("expected content.xml parse error, got {:?}", other),
        }
    }

    #[test]
    fn parse_ods_named_ranges_rejects_multiple_roots() {
        let err = parse_ods_named_ranges(b"<document-content/><document-content/>")
            .expect_err("content XML must have one root");

        assert!(format!("{err}").contains("multiple roots"));
    }

    #[test]
    fn parse_ods_named_ranges_reports_invalid_attribute_entity() {
        let xml = br#"
            <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
              <table:named-range table:name="Broken &" table:cell-range-address="$Sheet1.$A$1"/>
            </office:document-content>
        "#;

        let err = parse_ods_named_ranges(xml).expect_err("invalid attribute entity must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("expected content.xml parse error, got {:?}", other),
        }
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

        let transform = parse_frame_transform(&start).expect("frame transform");
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

        let transform = parse_frame_transform(&start).expect("frame transform");
        assert_eq!(transform.x, 0);
        assert_eq!(transform.y, 0);
        assert_eq!(transform.width, 0);
        assert_eq!(transform.height, 95_250);
    }

    #[test]
    fn parse_length_emu_str_rejects_digits_after_unit() {
        assert_eq!(parse_length_emu_str("1cm2"), None);
        assert_eq!(parse_length_emu_str("2pt3"), None);
    }

    #[test]
    fn parse_length_emu_str_rejects_non_finite_values() {
        assert_eq!(parse_length_emu_str("NaNcm"), None);
        assert_eq!(parse_length_emu_str("infpt"), None);
    }

    #[test]
    fn parse_frame_transform_rejects_lengths_that_overflow_integer_fields() {
        assert_eq!(
            parse_length_emu("999999999999999999999999999999cm".to_string()),
            None
        );
        assert_eq!(
            parse_length_emu_u64("999999999999999999999999999999cm".to_string()),
            None
        );
    }
}
