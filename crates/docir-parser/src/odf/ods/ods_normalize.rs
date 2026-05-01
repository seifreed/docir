use super::super::helpers::parse_text_element;
use crate::odf::{attr_value, CellFormula, CellValue, OdfReader, ParseError};
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

pub(crate) fn row_repeat_from(start: &BytesStart<'_>) -> u32 {
    attr_value(start, b"table:number-rows-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1)
}

pub(crate) fn read_ods_cell_text(reader: &mut OdfReader<'_>) -> Result<String, ParseError> {
    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"text:p" => {
                let para = parse_text_element(reader, e.name().as_ref())?;
                if !text.is_empty() && !para.is_empty() {
                    text.push('\n');
                }
                text.push_str(&para);
            }
            Ok(Event::Text(e)) => {
                let chunk = e.unescape().unwrap_or_default();
                text.push_str(&chunk);
            }
            Ok(Event::End(e)) if e.name().as_ref() == b"table:table-cell" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::xml_utils::xml_error("content.xml", e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

pub(crate) fn infer_cell_value_type_and_attr(
    value_type: &mut Option<String>,
    value_attr: &mut Option<String>,
    date_value: Option<String>,
    time_value: Option<String>,
    text: &str,
) {
    if value_type.is_none() && !text.is_empty() {
        *value_type = Some("string".to_string());
    }
    if value_attr.is_none() {
        *value_attr = date_value.or(time_value).or_else(|| {
            if text.is_empty() {
                None
            } else {
                Some(text.to_string())
            }
        });
    }
}

pub(crate) fn resolve_style_id(
    start: &BytesStart<'_>,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Option<u32> {
    attr_value(start, b"table:style-name").map(|name| {
        if let Some(id) = style_map.get(&name) {
            *id
        } else {
            let id = *next_style_id;
            *next_style_id += 1;
            style_map.insert(name, id);
            id
        }
    })
}

pub(crate) fn parse_cell_formula(formula_attr: Option<String>) -> Option<CellFormula> {
    formula_attr.map(|formula| CellFormula {
        text: normalize_formula_text(formula),
        formula_type: docir_core::ir::FormulaType::Normal,
        shared_index: None,
        shared_ref: None,
        is_array: false,
        array_ref: None,
    })
}

pub(crate) fn normalize_formula_text(formula: String) -> String {
    if let Some(stripped) = formula.strip_prefix("of:=") {
        stripped.to_string()
    } else if let Some(stripped) = formula.strip_prefix("of:") {
        stripped.to_string()
    } else if let Some(stripped) = formula.strip_prefix('=') {
        stripped.to_string()
    } else {
        formula
    }
}

pub(crate) fn parse_cell_value_with_text(
    value_type: Option<&str>,
    value_attr: Option<&str>,
    text: &str,
) -> CellValue {
    match value_type {
        Some("string") | Some("text") => {
            if !text.is_empty() {
                CellValue::String(text.to_string())
            } else {
                CellValue::String(value_attr.unwrap_or_default().to_string())
            }
        }
        Some("float") | Some("currency") | Some("percentage") => value_attr
            .and_then(|v| v.parse::<f64>().ok())
            .map(CellValue::Number)
            .unwrap_or_else(|| CellValue::String(text.to_string())),
        Some("boolean") => {
            let v = value_attr.unwrap_or(text);
            CellValue::Boolean(v == "true" || v == "1")
        }
        Some("date") | Some("time") => CellValue::String(value_attr.unwrap_or(text).to_string()),
        _ => {
            if !text.is_empty() {
                CellValue::String(text.to_string())
            } else {
                CellValue::Empty
            }
        }
    }
}

pub(crate) fn parse_cell_value_empty(
    value_type: Option<&str>,
    value_attr: Option<&str>,
) -> CellValue {
    match value_type {
        Some("float") | Some("currency") | Some("percentage") => value_attr
            .and_then(|v| v.parse::<f64>().ok())
            .map(CellValue::Number)
            .unwrap_or(CellValue::Empty),
        Some("boolean") => {
            let v = value_attr.unwrap_or_default();
            if v.is_empty() {
                CellValue::Empty
            } else {
                CellValue::Boolean(v == "true" || v == "1")
            }
        }
        Some("string") | Some("text") | Some("date") | Some("time") => value_attr
            .map_or(CellValue::Empty, |value| {
                CellValue::String(value.to_string())
            }),
        _ => CellValue::Empty,
    }
}
