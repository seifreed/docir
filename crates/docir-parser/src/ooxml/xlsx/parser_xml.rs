use crate::error::ParseError;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{local_name, xml_error};
use docir_core::ir::{
    parse_cell_reference, CalcChain, CalcChainEntry, CellError, CellFormula, ColumnDefinition,
    ConditionalFormat, ConditionalRule, FormulaType, MergedCellRange,
};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::HashMap;

pub(super) fn parse_calc_chain(xml: &str, path: &str) -> Result<CalcChain, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut chain = CalcChain::new();
    chain.span = Some(SourceSpan::new(path));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == b"c" {
                    let mut cell_ref = None;
                    let mut sheet_id = None;
                    let mut index = None;
                    let mut level = None;
                    let mut new_value = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"r" => cell_ref = Some(lossy_attr_value(&attr).to_string()),
                            b"i" => index = lossy_attr_value(&attr).parse::<u32>().ok(),
                            b"l" => level = lossy_attr_value(&attr).parse::<u32>().ok(),
                            b"s" => {
                                let value = lossy_attr_value(&attr);
                                new_value =
                                    Some(value == "1" || value.eq_ignore_ascii_case("true"));
                            }
                            b"si" => sheet_id = lossy_attr_value(&attr).parse::<u32>().ok(),
                            _ => {}
                        }
                    }
                    if let Some(cell_ref) = cell_ref {
                        chain.entries.push(CalcChainEntry {
                            cell_ref,
                            sheet_id,
                            index,
                            level,
                            new_value,
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(chain)
}

pub(super) fn map_cell_error(value: &str) -> CellError {
    match value.trim() {
        "#NULL!" => CellError::Null,
        "#DIV/0!" => CellError::DivZero,
        "#VALUE!" => CellError::Value,
        "#REF!" => CellError::Ref,
        "#NAME?" => CellError::Name,
        "#NUM!" => CellError::Num,
        "#N/A" => CellError::NA,
        "#GETTING_DATA" => CellError::GettingData,
        _ => CellError::Value,
    }
}

pub(super) fn parse_conditional_formatting(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<ConditionalFormat, ParseError> {
    let mut ranges: Vec<String> = Vec::new();
    for attr in start.attributes().flatten() {
        if attr.key.as_ref() == b"sqref" {
            let value = lossy_attr_value(&attr).to_string();
            ranges = value.split_whitespace().map(|s| s.to_string()).collect();
        }
    }

    let mut rules: Vec<ConditionalRule> = Vec::new();
    let mut current_rule: Option<ConditionalRule> = None;
    let mut in_formula = false;
    let mut formula_text = String::new();

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"cfRule" => {
                    let mut rule_type = "unknown".to_string();
                    let mut priority = None;
                    let mut operator = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"type" => rule_type = lossy_attr_value(&attr).to_string(),
                            b"priority" => priority = lossy_attr_value(&attr).parse::<u32>().ok(),
                            b"operator" => {
                                operator = Some(lossy_attr_value(&attr).to_string());
                            }
                            _ => {}
                        }
                    }
                    current_rule = Some(ConditionalRule {
                        rule_type,
                        priority,
                        operator,
                        formulae: Vec::new(),
                    });
                }
                b"formula" => {
                    in_formula = true;
                    formula_text.clear();
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_formula {
                    formula_text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => match local_name(e.name().as_ref()) {
                b"formula" => {
                    in_formula = false;
                    if let Some(rule) = current_rule.as_mut() {
                        if !formula_text.is_empty() {
                            rule.formulae.push(formula_text.clone());
                        }
                    }
                }
                b"cfRule" => {
                    if let Some(rule) = current_rule.take() {
                        rules.push(rule);
                    }
                }
                b"conditionalFormatting" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(sheet_path, e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(ConditionalFormat {
        id: NodeId::new(),
        ranges,
        rules,
        span: Some(SourceSpan::new(sheet_path)),
    })
}

pub(super) fn parse_formula(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<CellFormula, ParseError> {
    let (formula_type, shared_index, shared_ref, is_array, array_ref) = parse_formula_attrs(start);

    let text = reader
        .read_text(start.name())
        .map_err(|e| xml_error(sheet_path, e))?;

    Ok(CellFormula {
        text: text.to_string(),
        formula_type,
        shared_index,
        shared_ref,
        is_array,
        array_ref,
    })
}

pub(super) fn extract_formula_function(formula_upper: &str) -> Option<String> {
    let trimmed = formula_upper.trim();
    let trimmed = trimmed.strip_prefix('=').unwrap_or(trimmed);
    let idx = trimmed.find('(')?;
    Some(trimmed[..idx].trim().to_string())
}

pub(super) fn parse_formula_args_text(formula: &str) -> Option<String> {
    let start = formula.find('(')?;
    let end = formula.rfind(')')?;
    if end > start + 1 {
        Some(formula[start + 1..end].to_string())
    } else {
        None
    }
}

pub(super) fn parse_formula_empty(start: &BytesStart) -> CellFormula {
    let (formula_type, shared_index, shared_ref, is_array, array_ref) = parse_formula_attrs(start);

    CellFormula {
        text: String::new(),
        formula_type,
        shared_index,
        shared_ref,
        is_array,
        array_ref,
    }
}

fn parse_formula_attrs(
    start: &BytesStart,
) -> (
    FormulaType,
    Option<u32>,
    Option<String>,
    bool,
    Option<String>,
) {
    let mut formula_type = FormulaType::Normal;
    let mut shared_index = None;
    let mut shared_ref = None;
    let mut array_ref = None;
    let mut is_array = false;

    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"t" => {
                let value = lossy_attr_value(&attr);
                match value.as_ref() {
                    "shared" => formula_type = FormulaType::Shared,
                    "array" => {
                        formula_type = FormulaType::Array;
                        is_array = true;
                    }
                    "dataTable" => formula_type = FormulaType::DataTable,
                    _ => {}
                }
            }
            b"si" => shared_index = lossy_attr_value(&attr).parse::<u32>().ok(),
            b"ref" => {
                let reference = lossy_attr_value(&attr).to_string();
                if formula_type == FormulaType::Shared {
                    shared_ref = Some(reference);
                } else {
                    array_ref = Some(reference);
                }
            }
            _ => {}
        }
    }

    (formula_type, shared_index, shared_ref, is_array, array_ref)
}

pub(super) fn parse_inline_string(
    reader: &mut Reader<&[u8]>,
    sheet_path: &str,
) -> Result<String, ParseError> {
    let mut buf = Vec::new();
    let mut in_t = false;
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if local_name(e.name().as_ref()) == b"t" {
                    in_t = true;
                }
            }
            Ok(Event::Text(e)) => {
                if in_t {
                    let t = e.unescape().map_err(|err| xml_error(sheet_path, err))?;
                    text.push_str(&t);
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"t" {
                    in_t = false;
                } else if local_name(e.name().as_ref()) == b"is" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(sheet_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(text)
}

pub(super) fn parse_column(element: &BytesStart, columns: &mut HashMap<u32, ColumnDefinition>) {
    let mut min = None;
    let mut max = None;
    let mut width = None;
    let mut hidden = false;
    let mut custom_width = false;

    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"min" => min = lossy_attr_value(&attr).parse::<u32>().ok(),
            b"max" => max = lossy_attr_value(&attr).parse::<u32>().ok(),
            b"width" => width = lossy_attr_value(&attr).parse::<f64>().ok(),
            b"hidden" => hidden = attr.value.as_ref() == b"1",
            b"customWidth" => custom_width = attr.value.as_ref() == b"1",
            _ => {}
        }
    }

    let (Some(min), Some(max)) = (min, max) else {
        return;
    };
    for idx in min..=max {
        let col_index = idx.saturating_sub(1);
        columns.insert(
            col_index,
            ColumnDefinition {
                index: col_index,
                width,
                hidden,
                custom_width,
            },
        );
    }
}

pub(super) fn parse_merge_cell(element: &BytesStart) -> Option<MergedCellRange> {
    let mut ref_attr = None;
    for attr in element.attributes().flatten() {
        if attr.key.as_ref() == b"ref" {
            ref_attr = Some(lossy_attr_value(&attr).to_string());
        }
    }

    let ref_attr = ref_attr?;
    let mut parts = ref_attr.split(':');
    let start = parts.next()?;
    let end = parts.next().unwrap_or(start);

    let (start_col, start_row) = parse_cell_reference(start)?;
    let (end_col, end_row) = parse_cell_reference(end)?;

    Some(MergedCellRange {
        start_col,
        start_row,
        end_col,
        end_row,
    })
}
