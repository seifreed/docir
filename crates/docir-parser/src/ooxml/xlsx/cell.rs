use super::XlsxParser;
use crate::error::ParseError;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::visit_attributes;
use crate::xml_utils::{
    XmlScanControl, decoded_text, is_end_event_local, local_name, scan_xml_events_until_end,
    xml_error,
};
use docir_core::ir::{Cell, CellFormula, CellValue};
use docir_core::types::SourceSpan;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

impl XlsxParser {
    pub(super) fn parse_cell(
        &mut self,
        reader: &mut Reader<&[u8]>,
        start: &BytesStart,
        sheet_path: &str,
    ) -> Result<Cell, ParseError> {
        let attrs = CellAttributes::from_start(start, sheet_path)?;
        let coordinates = super::parse_cell_reference(&attrs.reference).ok_or_else(|| {
            ParseError::InvalidStructure(format!("Invalid cell reference: {}", attrs.reference))
        })?;
        let (col, row) = super::validate_cell_coordinates(&attrs.reference, coordinates)?;
        let contents = CellContents::parse(reader, sheet_path)?;

        let mut cell = Cell::new(attrs.reference.clone(), col, row);
        cell.style_id = attrs.style_id;
        if let Some(f) = &contents.formula {
            self.handle_formula_security(&attrs.reference, f, sheet_path);
        }
        cell.formula = contents.formula;
        cell.span = Some(SourceSpan::new(sheet_path));
        cell.value = parse_cell_value(
            &self.shared_strings,
            attrs.cell_type.as_deref(),
            &attrs.reference,
            contents.value_text,
            contents.inline_text,
        )?;

        Ok(cell)
    }
}

struct CellAttributes {
    reference: String,
    cell_type: Option<String>,
    style_id: Option<u32>,
}

impl CellAttributes {
    fn from_start(start: &BytesStart, sheet_path: &str) -> Result<Self, ParseError> {
        let mut reference: Option<String> = None;
        let mut cell_type: Option<String> = None;
        let mut style_raw: Option<String> = None;

        visit_attributes(start, sheet_path, |attr| match attr.key.as_ref() {
            b"r" => reference = Some(lossy_attr_value(attr).to_string()),
            b"t" => cell_type = Some(lossy_attr_value(attr).to_string()),
            b"s" => style_raw = Some(lossy_attr_value(attr).to_string()),
            _ => {}
        })?;

        let reference = reference.ok_or_else(|| {
            ParseError::InvalidStructure("Cell missing reference attribute".to_string())
        })?;
        let style_id = style_raw
            .as_deref()
            .map(|raw| parse_style_id(raw, Some(&reference)))
            .transpose()?;

        Ok(Self {
            reference,
            cell_type,
            style_id,
        })
    }
}

struct CellContents {
    value_text: Option<String>,
    inline_text: Option<String>,
    formula: Option<CellFormula>,
}

impl CellContents {
    fn parse(reader: &mut Reader<&[u8]>, sheet_path: &str) -> Result<Self, ParseError> {
        let mut contents = Self {
            value_text: None,
            inline_text: None,
            formula: None,
        };
        let mut buf = Vec::new();
        scan_xml_events_until_end(
            reader,
            &mut buf,
            sheet_path,
            |event| is_end_event_local(event, b"c"),
            |reader, event| contents.handle_event(reader, event, sheet_path),
        )?;
        Ok(contents)
    }

    fn handle_event(
        &mut self,
        reader: &mut Reader<&[u8]>,
        event: &Event<'_>,
        sheet_path: &str,
    ) -> Result<XmlScanControl, ParseError> {
        match event {
            Event::Start(e) => self.handle_start(reader, e, sheet_path)?,
            Event::Empty(e) if local_name(e.name().as_ref()) == b"f" => {
                self.formula = Some(super::parse_formula_empty(e, sheet_path)?);
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"is" => {
                self.inline_text = Some(String::new());
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    }

    fn handle_start(
        &mut self,
        reader: &mut Reader<&[u8]>,
        start: &BytesStart,
        sheet_path: &str,
    ) -> Result<(), ParseError> {
        match local_name(start.name().as_ref()) {
            b"v" => {
                let text = reader
                    .read_text(start.name())
                    .map_err(|e| xml_error(sheet_path, e))?;
                self.value_text =
                    Some(decoded_text(&text).map_err(|err| xml_error(sheet_path, err))?);
            }
            b"f" => {
                self.formula = Some(super::parse_formula(reader, start, sheet_path)?);
            }
            b"is" => {
                self.inline_text = Some(super::parse_inline_string(reader, sheet_path)?);
            }
            _ => {}
        }
        Ok(())
    }
}

fn parse_style_id(raw: &str, reference: Option<&str>) -> Result<u32, ParseError> {
    raw.parse::<u32>().map_err(|err| {
        let cell_reference = reference.unwrap_or("<unknown>");
        ParseError::InvalidStructure(format!(
            "Invalid style id '{raw}' on cell {cell_reference}: {err}"
        ))
    })
}

fn parse_cell_value(
    shared_strings: &[String],
    cell_type: Option<&str>,
    reference: &str,
    value_text: Option<String>,
    inline_text: Option<String>,
) -> Result<CellValue, ParseError> {
    if let Some(text) = inline_text {
        return Ok(CellValue::InlineString(text));
    }

    let Some(value) = value_text else {
        return Ok(CellValue::Empty);
    };

    match cell_type {
        Some("s") => parse_shared_string_value(shared_strings, reference, value),
        Some("b") => Ok(parse_boolean_value(&value)),
        Some("str") => Ok(CellValue::String(value)),
        Some("e") => Ok(CellValue::Error(super::map_cell_error(&value))),
        Some("d") => Ok(parse_datetime_value(value)),
        _ => Ok(parse_number_or_string(value)),
    }
}

fn parse_shared_string_value(
    shared_strings: &[String],
    reference: &str,
    value: String,
) -> Result<CellValue, ParseError> {
    let idx = value.trim().parse::<u32>().map_err(|err| {
        ParseError::InvalidStructure(format!(
            "Invalid shared-string index '{}' in cell {}: {err}",
            value, reference
        ))
    })?;

    Ok(shared_strings
        .get(idx as usize)
        .map_or(CellValue::SharedString(idx), |s| {
            CellValue::String(s.clone())
        }))
}

fn parse_boolean_value(value: &str) -> CellValue {
    let bool_value = value.trim();
    CellValue::Boolean(bool_value == "1" || bool_value.eq_ignore_ascii_case("true"))
}

fn parse_datetime_value(value: String) -> CellValue {
    match value.trim().parse::<f64>() {
        Ok(v) if v.is_finite() => CellValue::DateTime(v),
        Err(_) => CellValue::String(value),
        Ok(_) => CellValue::String(value),
    }
}

fn parse_number_or_string(value: String) -> CellValue {
    match value.trim().parse::<f64>() {
        Ok(v) if v.is_finite() => CellValue::Number(v),
        Err(_) => CellValue::String(value),
        Ok(_) => CellValue::String(value),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::xlsx::XlsxParser;
    use quick_xml::events::Event;

    fn parse_cell_from_xml(parser: &mut XlsxParser, xml: &str) -> Result<Cell, ParseError> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let start = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"c" => break e.into_owned(),
                Ok(Event::Eof) => panic!("missing <c> start"),
                Ok(_) => {}
                Err(e) => panic!("xml read failed: {e}"),
            }
            buf.clear();
        };
        parser.parse_cell(&mut reader, &start, "xl/worksheets/sheet1.xml")
    }

    #[test]
    fn parse_cell_reports_malformed_attributes() {
        let mut parser = XlsxParser::new();
        let err = parse_cell_from_xml(&mut parser, r#"<c r="A1" r="B1"><v>1</v></c>"#)
            .expect_err("duplicate cell attributes must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "xl/worksheets/sheet1.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_cell_reports_missing_reference() {
        let mut parser = XlsxParser::new();
        let err = parse_cell_from_xml(&mut parser, r#"<c><v>1</v></c>"#).expect_err("must fail");
        match err {
            ParseError::InvalidStructure(msg) => assert!(msg.contains("missing reference")),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_cell_rejects_references_outside_xlsx_limits() {
        let mut parser = XlsxParser::new();
        let err = parse_cell_from_xml(&mut parser, r#"<c r="XFE1"><v>1</v></c>"#)
            .expect_err("column past XFD must fail");

        assert!(
            matches!(err, ParseError::InvalidStructure(message) if message.contains("worksheet limits"))
        );
    }

    #[test]
    fn parse_cell_rejects_truncated_xml() {
        let mut parser = XlsxParser::new();
        let err = parse_cell_from_xml(&mut parser, r#"<c r="A1"><v>1</v>"#)
            .expect_err("truncated cell must fail");

        assert!(matches!(err, ParseError::Xml { .. }));
    }

    #[test]
    fn parse_cell_parses_supported_value_types_and_empty_cell() {
        let mut parser = XlsxParser::new();
        parser.shared_strings = vec!["shared-value".to_string()];

        let shared = parse_cell_from_xml(&mut parser, r#"<c r="A1" t="s"><v>0</v></c>"#)
            .expect("shared string");
        assert!(matches!(shared.value, CellValue::String(ref v) if v == "shared-value"));

        let boolean =
            parse_cell_from_xml(&mut parser, r#"<c r="B1" t="b"><v>true</v></c>"#).expect("bool");
        assert!(matches!(boolean.value, CellValue::Boolean(true)));

        let string =
            parse_cell_from_xml(&mut parser, r#"<c r="C1" t="str"><v>abc</v></c>"#).expect("str");
        assert!(matches!(string.value, CellValue::String(ref v) if v == "abc"));

        let error = parse_cell_from_xml(&mut parser, r#"<c r="D1" t="e"><v>#DIV/0!</v></c>"#)
            .expect("error");
        assert!(matches!(error.value, CellValue::Error(_)));

        let date =
            parse_cell_from_xml(&mut parser, r#"<c r="E1" t="d"><v>123.5</v></c>"#).expect("date");
        assert!(matches!(date.value, CellValue::DateTime(v) if (v - 123.5).abs() < f64::EPSILON));

        let number =
            parse_cell_from_xml(&mut parser, r#"<c r="F1"><v>42</v></c>"#).expect("number");
        assert!(matches!(number.value, CellValue::Number(v) if (v - 42.0).abs() < f64::EPSILON));

        let fallback_string =
            parse_cell_from_xml(&mut parser, r#"<c r="G1"><v>not-a-number</v></c>"#)
                .expect("fallback string");
        assert!(matches!(fallback_string.value, CellValue::String(ref v) if v == "not-a-number"));

        let non_finite_number =
            parse_cell_from_xml(&mut parser, r#"<c r="G2"><v>NaN</v></c>"#).expect("NaN cell");
        assert!(matches!(non_finite_number.value, CellValue::String(ref v) if v == "NaN"));

        let non_finite_date = parse_cell_from_xml(&mut parser, r#"<c r="G3" t="d"><v>NaN</v></c>"#)
            .expect("NaN date cell");
        assert!(matches!(non_finite_date.value, CellValue::String(ref v) if v == "NaN"));

        let empty = parse_cell_from_xml(&mut parser, r#"<c r="H1"></c>"#).expect("empty");
        assert!(matches!(empty.value, CellValue::Empty));
    }

    #[test]
    fn parse_cell_parses_inline_string_and_formula_variants() {
        let mut parser = XlsxParser::new();

        let inline = parse_cell_from_xml(
            &mut parser,
            r#"<c r="I1" t="inlineStr"><is><t>Hello</t></is></c>"#,
        )
        .expect("inline string");
        assert!(matches!(inline.value, CellValue::InlineString(ref v) if v == "Hello"));
        assert!(inline.formula.is_none());
        assert_eq!(inline.style_id, None);
        assert_eq!(
            inline.span.as_ref().map(|s| s.file_path.as_str()),
            Some("xl/worksheets/sheet1.xml")
        );

        let empty_inline = parse_cell_from_xml(&mut parser, r#"<c r="I2" t="inlineStr"><is/></c>"#)
            .expect("empty inline string");
        assert!(matches!(
            empty_inline.value,
            CellValue::InlineString(ref value) if value.is_empty()
        ));

        let formula = parse_cell_from_xml(
            &mut parser,
            r#"<c r="J1" s="5"><f>SUM(A1:A3)</f><v>6</v></c>"#,
        )
        .expect("formula");
        assert_eq!(formula.style_id, Some(5));
        assert!(formula.formula.is_some());
        assert!(matches!(formula.value, CellValue::Number(v) if (v - 6.0).abs() < f64::EPSILON));

        let empty_formula =
            parse_cell_from_xml(&mut parser, r#"<c r="K1"><f/></c>"#).expect("empty formula");
        assert!(empty_formula.formula.is_some());
        assert!(matches!(empty_formula.value, CellValue::Empty));
    }
}
