use super::XlsxParser;
use crate::error::ParseError;
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
                b"r" => cell_ref = Some(String::from_utf8_lossy(&attr.value).to_string()),
                b"t" => cell_type = Some(String::from_utf8_lossy(&attr.value).to_string()),
                b"s" => style_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
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
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"v" => {
                        let text = reader.read_text(e.name()).map_err(|e| ParseError::Xml {
                            file: sheet_path.to_string(),
                            message: e.to_string(),
                        })?;
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
                Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"f" {
                        formula = Some(super::parse_formula_empty(&e));
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"c" {
                        break;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: sheet_path.to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

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
                    let idx = value.trim().parse::<u32>().unwrap_or(0);
                    if let Some(s) = self.shared_strings.get(idx as usize) {
                        CellValue::String(s.clone())
                    } else {
                        CellValue::SharedString(idx)
                    }
                }
                Some("b") => {
                    let v = value.trim();
                    CellValue::Boolean(v == "1" || v.eq_ignore_ascii_case("true"))
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
