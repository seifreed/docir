use super::XlsxParser;
use crate::error::ParseError;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{local_name, scan_xml_events_with_reader, xml_error, XmlScanControl};
use docir_core::ir::{Cell, CellFormula, CellValue};
use docir_core::types::SourceSpan;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

impl XlsxParser {
    pub(super) fn parse_cell(
        &mut self,
        reader: &mut Reader<&[u8]>,
        start: &BytesStart,
        sheet_path: &str,
    ) -> Result<Cell, ParseError> {
        let mut cell_ref: Option<String> = None;
        let mut cell_type: Option<String> = None;
        let mut style_id: Option<u32> = None;

        for attr in start.attributes().flatten() {
            match attr.key.as_ref() {
                b"r" => cell_ref = Some(lossy_attr_value(&attr).to_string()),
                b"t" => cell_type = Some(lossy_attr_value(&attr).to_string()),
                b"s" => {
                    let raw = lossy_attr_value(&attr);
                    style_id = Some(raw.parse::<u32>().map_err(|err| {
                        let cell_reference = cell_ref.as_deref().unwrap_or("<unknown>");
                        ParseError::InvalidStructure(format!(
                            "Invalid style id '{raw}' on cell {cell_reference}: {err}"
                        ))
                    })?);
                }
                _ => {}
            }
        }

        let reference = cell_ref.ok_or_else(|| {
            ParseError::InvalidStructure("Cell missing reference attribute".to_string())
        })?;

        let (col, row) = super::parse_cell_reference(&reference).ok_or_else(|| {
            ParseError::InvalidStructure(format!("Invalid cell reference: {reference}"))
        })?;

        let mut value_text: Option<String> = None;
        let mut inline_text: Option<String> = None;
        let mut formula: Option<CellFormula> = None;

        let mut buf = Vec::new();
        scan_xml_events_with_reader(reader, &mut buf, sheet_path, |reader, event| {
            match event {
                Event::Start(e) => match local_name(e.name().as_ref()) {
                    b"v" => {
                        let text = reader
                            .read_text(e.name())
                            .map_err(|e| xml_error(sheet_path, e))?;
                        value_text = Some(text.to_string());
                    }
                    b"f" => {
                        let f = super::parse_formula(reader, &e, sheet_path)?;
                        formula = Some(f);
                    }
                    b"is" => {
                        inline_text = Some(super::parse_inline_string(reader, sheet_path)?);
                    }
                    _ => {}
                },
                Event::Empty(e) => {
                    if local_name(e.name().as_ref()) == b"f" {
                        formula = Some(super::parse_formula_empty(&e));
                    }
                }
                Event::End(e) => {
                    if local_name(e.name().as_ref()) == b"c" {
                        return Ok(XmlScanControl::Break);
                    }
                }
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        })?;

        let mut cell = Cell::new(reference.clone(), col, row);
        cell.style_id = style_id;
        if let Some(f) = &formula {
            self.handle_formula_security(&reference, f, sheet_path);
        }
        cell.formula = formula;
        cell.span = Some(SourceSpan::new(sheet_path));

        cell.value = if let Some(text) = inline_text {
            CellValue::InlineString(text)
        } else if let Some(value) = value_text {
            match cell_type.as_deref() {
                Some("s") => {
                    let idx = value.trim().parse::<u32>().map_err(|err| {
                        ParseError::InvalidStructure(format!(
                            "Invalid shared-string index '{}' in cell {}: {err}",
                            value, reference
                        ))
                    })?;
                    if let Some(s) = self.shared_strings.get(idx as usize) {
                        CellValue::String(s.clone())
                    } else {
                        CellValue::SharedString(idx)
                    }
                }
                Some("b") => {
                    let bool_value = value.trim();
                    CellValue::Boolean(bool_value == "1" || bool_value.eq_ignore_ascii_case("true"))
                }
                Some("str") => CellValue::String(value),
                Some("e") => CellValue::Error(super::map_cell_error(&value)),
                Some("d") => match value.trim().parse::<f64>() {
                    Ok(v) => CellValue::DateTime(v),
                    Err(_) => CellValue::String(value),
                },
                _ => match value.trim().parse::<f64>() {
                    Ok(v) => CellValue::Number(v),
                    Err(_) => CellValue::String(value),
                },
            }
        } else {
            CellValue::Empty
        };

        Ok(cell)
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
    fn parse_cell_reports_missing_reference() {
        let mut parser = XlsxParser::new();
        let err = parse_cell_from_xml(&mut parser, r#"<c><v>1</v></c>"#).expect_err("must fail");
        match err {
            ParseError::InvalidStructure(msg) => assert!(msg.contains("missing reference")),
            other => panic!("unexpected error: {other:?}"),
        }
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
