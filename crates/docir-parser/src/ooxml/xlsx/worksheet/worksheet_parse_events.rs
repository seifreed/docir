use super::WorksheetParseAccum;
use crate::ooxml::relationships::Relationships;
use crate::ooxml::xlsx::{
    parse_color_attr, parse_column, parse_merge_cell, ParseError, Worksheet, XlsxParser,
};
use crate::xml_utils::attr_value;
use crate::xml_utils::{scan_xml_events_until_end, XmlScanControl};
use docir_core::ir::ConditionalFormat;
use docir_core::ir::{DataValidation, SheetPageMargins};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::Event;
use quick_xml::events::{BytesEnd, BytesStart};
use quick_xml::Reader;

pub(crate) fn handle_worksheet_common_tag(
    e: &BytesStart<'_>,
    sheet_path: &str,
    relationships: &Relationships,
    worksheet: &mut Worksheet,
    accum: &mut WorksheetParseAccum,
    parser: &mut XlsxParser,
) -> bool {
    match e.name().as_ref() {
        b"dimension" => {
            if let Some(val) = attr_value(e, b"ref") {
                worksheet.dimension = Some(val);
            }
            true
        }
        b"tabColor" => {
            worksheet.tab_color = parse_color_attr(e);
            true
        }
        b"pageMargins" => {
            worksheet.page_margins = parse_page_margins(e);
            true
        }
        b"col" => {
            parse_column(e, &mut accum.columns);
            true
        }
        b"mergeCell" => {
            if let Some(range) = parse_merge_cell(e) {
                accum.merged_cells.push(range);
            }
            true
        }
        b"hyperlink" => {
            parser.handle_hyperlink(e, relationships, sheet_path);
            true
        }
        _ => false,
    }
}

pub(crate) fn parse_page_margins(start: &BytesStart) -> Option<SheetPageMargins> {
    let mut margins = SheetPageMargins {
        left: None,
        right: None,
        top: None,
        bottom: None,
        header: None,
        footer: None,
    };
    let mut found = false;
    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"left" => {
                margins.left = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            b"right" => {
                margins.right = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            b"top" => {
                margins.top = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            b"bottom" => {
                margins.bottom = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            b"header" => {
                margins.header = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            b"footer" => {
                margins.footer = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            _ => {}
        }
    }
    if found {
        Some(margins)
    } else {
        None
    }
}

pub(crate) fn parse_conditional_formatting_empty(
    start: &BytesStart,
    sheet_path: &str,
) -> ConditionalFormat {
    let mut ranges: Vec<String> = Vec::new();
    for attr in start.attributes().flatten() {
        if attr.key.as_ref() == b"sqref" {
            let val = String::from_utf8_lossy(&attr.value).to_string();
            ranges = val.split_whitespace().map(|s| s.to_string()).collect();
        }
    }
    ConditionalFormat {
        id: NodeId::new(),
        ranges,
        rules: Vec::new(),
        span: Some(SourceSpan::new(sheet_path)),
    }
}

pub(crate) fn parse_data_validations(
    reader: &mut Reader<&[u8]>,
    sheet_path: &str,
) -> Result<Vec<DataValidation>, ParseError> {
    let mut validations: Vec<DataValidation> = Vec::new();
    let mut buf = Vec::new();
    scan_xml_events_until_end(
        reader,
        &mut buf,
        sheet_path,
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"dataValidations"),
        |reader, event| {
            match event {
                Event::Start(e) if e.name().as_ref() == b"dataValidation" => {
                    let val = parse_data_validation(reader, e, sheet_path)?;
                    validations.push(val);
                }
                Event::Empty(e) if e.name().as_ref() == b"dataValidation" => {
                    let val = parse_data_validation_empty(e, sheet_path);
                    validations.push(val);
                }
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    Ok(validations)
}

pub(crate) fn parse_data_validation(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<DataValidation, ParseError> {
    let mut validation = parse_data_validation_empty(start, sheet_path);
    let mut formulas = DataValidationFormulaCapture::default();

    let mut buf = Vec::new();
    scan_xml_events_until_end(
        reader,
        &mut buf,
        sheet_path,
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"dataValidation"),
        |_reader, event| {
            match event {
                Event::Start(e) => {
                    formulas.track_start(e);
                    formulas.track_start_with_context(e, &mut validation);
                }
                Event::Text(e) => {
                    let text = e.unescape().unwrap_or_default().to_string();
                    formulas.append_text(&text);
                }
                Event::End(e) => {
                    formulas.track_end(e, &mut validation);
                }
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    Ok(validation)
}

#[derive(Debug, Default)]
struct DataValidationFormulaCapture {
    in_formula: Option<u8>,
    formula1: String,
    formula2: String,
}

impl DataValidationFormulaCapture {
    fn track_start(&mut self, e: &BytesStart<'_>) {
        match e.name().as_ref() {
            b"formula1" => {
                self.in_formula = Some(1);
                self.formula1.clear();
            }
            b"formula2" => {
                self.in_formula = Some(2);
                self.formula2.clear();
            }
            _ => {}
        }
    }

    fn track_start_with_context(&mut self, e: &BytesStart<'_>, validation: &mut DataValidation) {
        if e.name().as_ref() == b"formula1" {
            if let Some(val) = attr_value(e, b"val") {
                validation.formula1 = Some(val);
                self.in_formula = None;
                self.formula1.clear();
            }
        }
        if e.name().as_ref() == b"formula2" {
            if let Some(val) = attr_value(e, b"val") {
                validation.formula2 = Some(val);
                self.in_formula = None;
                self.formula2.clear();
            }
        }
    }

    fn append_text(&mut self, text: &str) {
        match self.in_formula {
            Some(1) => self.formula1.push_str(text),
            Some(2) => self.formula2.push_str(text),
            _ => {}
        }
    }

    fn track_end(&mut self, e: &BytesEnd<'_>, validation: &mut DataValidation) {
        match (self.in_formula, e.name().as_ref()) {
            (Some(1), b"formula1") => {
                self.in_formula = None;
                if !self.formula1.is_empty() {
                    validation.formula1 = Some(self.formula1.clone());
                }
            }
            (Some(2), b"formula2") => {
                self.in_formula = None;
                if !self.formula2.is_empty() {
                    validation.formula2 = Some(self.formula2.clone());
                }
            }
            _ => {}
        }
    }
}

pub(crate) fn parse_data_validation_empty(start: &BytesStart, sheet_path: &str) -> DataValidation {
    let mut validation = DataValidation {
        id: NodeId::new(),
        validation_type: None,
        operator: None,
        allow_blank: false,
        show_input_message: false,
        show_error_message: false,
        error_title: None,
        error: None,
        prompt_title: None,
        prompt: None,
        ranges: Vec::new(),
        formula1: None,
        formula2: None,
        span: Some(SourceSpan::new(sheet_path)),
    };

    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"type" => {
                validation.validation_type = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"operator" => {
                validation.operator = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"allowBlank" => {
                let v = String::from_utf8_lossy(&attr.value);
                validation.allow_blank = v == "1" || v.eq_ignore_ascii_case("true");
            }
            b"showInputMessage" => {
                let v = String::from_utf8_lossy(&attr.value);
                validation.show_input_message = v == "1" || v.eq_ignore_ascii_case("true");
            }
            b"showErrorMessage" => {
                let v = String::from_utf8_lossy(&attr.value);
                validation.show_error_message = v == "1" || v.eq_ignore_ascii_case("true");
            }
            b"errorTitle" => {
                validation.error_title = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"error" => {
                validation.error = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"promptTitle" => {
                validation.prompt_title = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"prompt" => {
                validation.prompt = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"sqref" => {
                let val = String::from_utf8_lossy(&attr.value).to_string();
                validation.ranges = val.split_whitespace().map(|s| s.to_string()).collect();
            }
            _ => {}
        }
    }

    validation
}

#[cfg(test)]
mod tests {
    use super::{
        parse_conditional_formatting_empty, parse_data_validation, parse_data_validation_empty,
        parse_page_margins,
    };
    use crate::xml_utils::reader_from_str;
    use quick_xml::events::Event;

    #[test]
    fn parse_page_margins_reads_known_attributes() {
        let mut start = quick_xml::events::BytesStart::new("pageMargins");
        start.push_attribute(("left", "0.75"));
        start.push_attribute(("right", "0.5"));
        start.push_attribute(("top", "1.0"));
        let margins = parse_page_margins(&start).expect("margins expected");
        assert_eq!(margins.left, Some(0.75));
        assert_eq!(margins.right, Some(0.5));
        assert_eq!(margins.top, Some(1.0));
        assert_eq!(margins.bottom, None);
    }

    #[test]
    fn parse_data_validation_empty_reads_flags_and_ranges() {
        let mut start = quick_xml::events::BytesStart::new("dataValidation");
        start.push_attribute(("type", "list"));
        start.push_attribute(("operator", "between"));
        start.push_attribute(("allowBlank", "1"));
        start.push_attribute(("showInputMessage", "true"));
        start.push_attribute(("showErrorMessage", "false"));
        start.push_attribute(("sqref", "A1 A2:B2"));

        let validation = parse_data_validation_empty(&start, "xl/worksheets/sheet1.xml");
        assert_eq!(validation.validation_type.as_deref(), Some("list"));
        assert_eq!(validation.operator.as_deref(), Some("between"));
        assert!(validation.allow_blank);
        assert!(validation.show_input_message);
        assert!(!validation.show_error_message);
        assert_eq!(
            validation.ranges,
            vec!["A1".to_string(), "A2:B2".to_string()]
        );
    }

    #[test]
    fn parse_data_validation_reads_formula_nodes_and_attrs() {
        let xml = r#"
            <dataValidation type="whole" sqref="C3">
              <formula1 val="1" />
              <formula2>10</formula2>
            </dataValidation>
        "#;
        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();
        let start = loop {
            match reader.read_event_into(&mut buf).expect("xml") {
                Event::Start(e) if e.name().as_ref() == b"dataValidation" => break e.into_owned(),
                Event::Eof => panic!("missing dataValidation"),
                _ => {}
            }
            buf.clear();
        };
        let validation =
            parse_data_validation(&mut reader, &start, "xl/worksheets/sheet1.xml").expect("ok");
        assert_eq!(validation.formula1, None);
        assert_eq!(validation.formula2.as_deref(), Some("10"));
        assert_eq!(validation.ranges, vec!["C3".to_string()]);
    }

    #[test]
    fn parse_conditional_formatting_empty_splits_sqref_ranges() {
        let mut start = quick_xml::events::BytesStart::new("conditionalFormatting");
        start.push_attribute(("sqref", "A1 B2:C3"));
        let cf = parse_conditional_formatting_empty(&start, "xl/worksheets/sheet1.xml");
        assert_eq!(cf.ranges, vec!["A1".to_string(), "B2:C3".to_string()]);
    }
}
