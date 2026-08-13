use crate::error::ParseError;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{attr_bool_like, decoded_text, local_name, visit_attributes, xml_error};
use docir_core::ir::{
    CalcChain, CalcChainEntry, CellError, CellFormula, ColumnDefinition, ConditionalFormat,
    ConditionalRule, FormulaType, MergedCellRange, parse_cell_reference,
};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

type FormulaAttrs = (
    FormulaType,
    Option<u32>,
    Option<String>,
    bool,
    Option<String>,
);

pub(super) fn parse_calc_chain(xml: &str, path: &str) -> Result<CalcChain, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut chain = CalcChain::new();
    chain.span = Some(SourceSpan::new(path));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"c" => {
                let mut cell_ref = None;
                let mut sheet_id = None;
                let mut index = None;
                let mut level = None;
                let mut new_value = None;
                let mut numeric_error = None;
                visit_attributes(&e, path, |attr| match attr.key.as_ref() {
                    b"r" => cell_ref = Some(lossy_attr_value(attr).to_string()),
                    b"i" => match parse_u32_attr(&lossy_attr_value(attr), path) {
                        Ok(parsed) => index = Some(parsed),
                        Err(err) => numeric_error = Some(err),
                    },
                    b"l" => match parse_u32_attr(&lossy_attr_value(attr), path) {
                        Ok(parsed) => level = Some(parsed),
                        Err(err) => numeric_error = Some(err),
                    },
                    b"s" => {
                        new_value = Some(attr_bool_like(attr.value.as_ref()));
                    }
                    b"si" => match parse_u32_attr(&lossy_attr_value(attr), path) {
                        Ok(parsed) => sheet_id = Some(parsed),
                        Err(err) => numeric_error = Some(err),
                    },
                    _ => {}
                })?;
                if let Some(err) = numeric_error {
                    return Err(err);
                }
                let cell_ref =
                    cell_ref.ok_or_else(|| xml_error(path, "calcChain cell is missing r"))?;
                chain.entries.push(CalcChainEntry {
                    cell_ref,
                    sheet_id,
                    index,
                    level,
                    new_value,
                });
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

fn parse_u32_attr(value: &str, path: &str) -> Result<u32, ParseError> {
    value.parse::<u32>().map_err(|err| xml_error(path, err))
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
    let ranges = conditional_ranges(start, sheet_path)?;
    let mut rules: Vec<ConditionalRule> = Vec::new();
    let mut current_rule: Option<ConditionalRule> = None;
    let mut in_formula = false;
    let mut formula_text = String::new();

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                handle_conditional_start(
                    &e,
                    sheet_path,
                    &mut current_rule,
                    &mut in_formula,
                    &mut formula_text,
                )?;
            }
            Ok(Event::Text(e)) if in_formula => {
                formula_text.push_str(
                    &crate::xml_utils::decoded_text(&e)
                        .map_err(|err| xml_error(sheet_path, err))?,
                );
            }
            Ok(Event::GeneralRef(e)) if in_formula => {
                formula_text.push_str(
                    &crate::xml_utils::decoded_general_ref(&e)
                        .map_err(|err| xml_error(sheet_path, err))?,
                );
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"conditionalFormatting" => {
                break;
            }
            Ok(Event::End(e)) => handle_conditional_end(
                local_name(e.name().as_ref()),
                &mut current_rule,
                &mut rules,
                &mut in_formula,
                &formula_text,
            ),
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

fn conditional_ranges(start: &BytesStart, sheet_path: &str) -> Result<Vec<String>, ParseError> {
    let mut ranges = Vec::new();
    visit_attributes(start, sheet_path, |attr| {
        if attr.key.as_ref() == b"sqref" {
            ranges = lossy_attr_value(attr)
                .split_whitespace()
                .map(|s| s.to_string())
                .collect();
        }
    })?;
    Ok(ranges)
}

fn handle_conditional_start(
    e: &BytesStart<'_>,
    sheet_path: &str,
    current_rule: &mut Option<ConditionalRule>,
    in_formula: &mut bool,
    formula_text: &mut String,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"cfRule" => *current_rule = Some(parse_conditional_rule(e, sheet_path)?),
        b"formula" => {
            *in_formula = true;
            formula_text.clear();
        }
        _ => {}
    }
    Ok(())
}

fn parse_conditional_rule(
    e: &BytesStart<'_>,
    sheet_path: &str,
) -> Result<ConditionalRule, ParseError> {
    let mut rule_type = "unknown".to_string();
    let mut priority = None;
    let mut operator = None;
    let mut numeric_error = None;
    visit_attributes(e, sheet_path, |attr| match attr.key.as_ref() {
        b"type" => rule_type = lossy_attr_value(attr).to_string(),
        b"priority" => match parse_u32_attr(&lossy_attr_value(attr), sheet_path) {
            Ok(parsed) => priority = Some(parsed),
            Err(err) => numeric_error = Some(err),
        },
        b"operator" => operator = Some(lossy_attr_value(attr).to_string()),
        _ => {}
    })?;
    if let Some(err) = numeric_error {
        return Err(err);
    }
    Ok(ConditionalRule {
        rule_type,
        priority,
        operator,
        formulae: Vec::new(),
    })
}

fn handle_conditional_end(
    local: &[u8],
    current_rule: &mut Option<ConditionalRule>,
    rules: &mut Vec<ConditionalRule>,
    in_formula: &mut bool,
    formula_text: &str,
) {
    match local {
        b"formula" => {
            *in_formula = false;
            if let Some(rule) = current_rule.as_mut()
                && !formula_text.is_empty()
            {
                rule.formulae.push(formula_text.to_string());
            }
        }
        b"cfRule" => {
            if let Some(rule) = current_rule.take() {
                rules.push(rule);
            }
        }
        _ => {}
    }
}

pub(super) fn parse_formula(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<CellFormula, ParseError> {
    let (formula_type, shared_index, shared_ref, is_array, array_ref) =
        parse_formula_attrs(start, sheet_path)?;

    let text = reader
        .read_text(start.name())
        .map_err(|e| xml_error(sheet_path, e))?;

    Ok(CellFormula {
        text: decoded_text(&text).map_err(|err| xml_error(sheet_path, err))?,
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

pub(super) fn parse_formula_empty(
    start: &BytesStart,
    sheet_path: &str,
) -> Result<CellFormula, ParseError> {
    let (formula_type, shared_index, shared_ref, is_array, array_ref) =
        parse_formula_attrs(start, sheet_path)?;

    Ok(CellFormula {
        text: String::new(),
        formula_type,
        shared_index,
        shared_ref,
        is_array,
        array_ref,
    })
}

fn parse_formula_attrs(start: &BytesStart, sheet_path: &str) -> Result<FormulaAttrs, ParseError> {
    let mut formula_type = FormulaType::Normal;
    let mut shared_index = None;
    let mut shared_ref = None;
    let mut array_ref = None;
    let mut is_array = false;
    let mut numeric_error = None;

    visit_attributes(start, sheet_path, |attr| match attr.key.as_ref() {
        b"t" => {
            let value = lossy_attr_value(attr);
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
        b"si" => match parse_u32_attr(&lossy_attr_value(attr), sheet_path) {
            Ok(parsed) => shared_index = Some(parsed),
            Err(err) => numeric_error = Some(err),
        },
        b"ref" => {
            let reference = lossy_attr_value(attr).to_string();
            if formula_type == FormulaType::Shared {
                shared_ref = Some(reference);
            } else {
                array_ref = Some(reference);
            }
        }
        _ => {}
    })?;
    if let Some(err) = numeric_error {
        return Err(err);
    }

    Ok((formula_type, shared_index, shared_ref, is_array, array_ref))
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
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"t" => {
                in_t = true;
            }
            Ok(Event::Text(e)) if in_t => {
                let t =
                    crate::xml_utils::decoded_text(&e).map_err(|err| xml_error(sheet_path, err))?;
                text.push_str(&t);
            }
            Ok(Event::GeneralRef(e)) if in_t => {
                let t = crate::xml_utils::decoded_general_ref(&e)
                    .map_err(|err| xml_error(sheet_path, err))?;
                text.push_str(&t);
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

pub(super) fn parse_column(
    element: &BytesStart,
    columns: &mut HashMap<u32, ColumnDefinition>,
    sheet_path: &str,
) -> Result<(), ParseError> {
    let mut min = None;
    let mut max = None;
    let mut width = None;
    let mut hidden = false;
    let mut custom_width = false;
    let mut numeric_error = None;

    visit_attributes(element, sheet_path, |attr| match attr.key.as_ref() {
        b"min" => match parse_u32_attr(&lossy_attr_value(attr), sheet_path) {
            Ok(parsed) => min = Some(parsed),
            Err(err) => numeric_error = Some(err),
        },
        b"max" => match parse_u32_attr(&lossy_attr_value(attr), sheet_path) {
            Ok(parsed) => max = Some(parsed),
            Err(err) => numeric_error = Some(err),
        },
        b"width" => match lossy_attr_value(attr).parse::<f64>() {
            Ok(parsed) => width = Some(parsed),
            Err(err) => numeric_error = Some(xml_error(sheet_path, err)),
        },
        b"hidden" => hidden = attr_bool_like(attr.value.as_ref()),
        b"customWidth" => custom_width = attr_bool_like(attr.value.as_ref()),
        _ => {}
    })?;
    if let Some(err) = numeric_error {
        return Err(err);
    }

    let (Some(min), Some(max)) = (min, max) else {
        return Ok(());
    };
    if min == 0 || max == 0 {
        return Err(xml_error(sheet_path, "column indexes must be 1-based"));
    }
    if max < min {
        return Err(xml_error(sheet_path, "column max is smaller than min"));
    }
    for idx in min..=max {
        let col_index = idx - 1;
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
    Ok(())
}

pub(super) fn parse_merge_cell(
    element: &BytesStart,
    sheet_path: &str,
) -> Result<Option<MergedCellRange>, ParseError> {
    let mut ref_attr = None;
    visit_attributes(element, sheet_path, |attr| {
        if attr.key.as_ref() == b"ref" {
            ref_attr = Some(lossy_attr_value(attr).to_string());
        }
    })?;

    let ref_attr = ref_attr.ok_or_else(|| xml_error(sheet_path, "mergeCell is missing ref"))?;
    let mut parts = ref_attr.split(':');
    let start = parts
        .next()
        .ok_or_else(|| xml_error(sheet_path, "mergeCell ref is empty"))?;
    let end = parts.next().unwrap_or(start);
    if parts.next().is_some() {
        return Err(xml_error(sheet_path, "mergeCell ref has too many ranges"));
    }

    let (start_col, start_row) = parse_cell_reference(start)
        .ok_or_else(|| xml_error(sheet_path, format!("invalid mergeCell start ref: {start}")))?;
    let (end_col, end_row) = parse_cell_reference(end)
        .ok_or_else(|| xml_error(sheet_path, format!("invalid mergeCell end ref: {end}")))?;

    Ok(Some(MergedCellRange {
        start_col,
        start_row,
        end_col,
        end_row,
    }))
}
