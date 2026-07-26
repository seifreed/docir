use super::{CellRange, CellRef, FormulaToken};
use docir_core::ir::CellValue;

pub(super) fn tokenize_formula(formula: &str) -> Vec<FormulaToken> {
    let mut tokens = Vec::new();
    let mut chars = formula.trim().chars().peekable();
    while let Some(&ch) = chars.peek() {
        match ch {
            ' ' | '\t' | '\n' | '\r' => {
                chars.next();
            }
            '+' => {
                chars.next();
                tokens.push(FormulaToken::Plus);
            }
            '-' => {
                chars.next();
                tokens.push(FormulaToken::Minus);
            }
            '*' => {
                chars.next();
                tokens.push(FormulaToken::Star);
            }
            '/' => {
                chars.next();
                tokens.push(FormulaToken::Slash);
            }
            '(' => {
                chars.next();
                tokens.push(FormulaToken::LParen);
            }
            ')' => {
                chars.next();
                tokens.push(FormulaToken::RParen);
            }
            ',' | ';' => {
                chars.next();
                tokens.push(FormulaToken::Comma);
            }
            '[' => {
                chars.next();
                let mut buffer = String::new();
                for c in chars.by_ref() {
                    if c == ']' {
                        break;
                    }
                    buffer.push(c);
                }
                if let Some(token) = parse_bracket_reference(&buffer) {
                    tokens.push(token);
                }
            }
            _ => {
                consume_numeric_or_identifier_token(&mut chars, &mut tokens);
            }
        }
    }
    tokens.push(FormulaToken::End);
    tokens
}

fn consume_numeric_or_identifier_token(
    chars: &mut std::iter::Peekable<std::str::Chars<'_>>,
    tokens: &mut Vec<FormulaToken>,
) {
    let Some(&ch) = chars.peek() else {
        return;
    };
    if ch.is_ascii_digit() || ch == '.' {
        consume_number_token(chars, tokens);
    } else if ch.is_ascii_alphabetic() || ch == '_' {
        consume_identifier_token(chars, tokens);
    } else {
        chars.next();
    }
}

fn consume_number_token(
    chars: &mut std::iter::Peekable<std::str::Chars<'_>>,
    tokens: &mut Vec<FormulaToken>,
) {
    let mut num = String::new();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_digit() || c == '.' {
            num.push(c);
            chars.next();
        } else {
            break;
        }
    }
    match num.parse::<f64>() {
        Ok(value) => tokens.push(FormulaToken::Number(value)),
        Err(_) => tokens.push(FormulaToken::Ident(num)),
    }
}

fn consume_identifier_token(
    chars: &mut std::iter::Peekable<std::str::Chars<'_>>,
    tokens: &mut Vec<FormulaToken>,
) {
    let mut ident = String::new();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_alphanumeric() || c == '_' || c == '.' || c == '$' {
            ident.push(c);
            chars.next();
        } else {
            break;
        }
    }
    if let Some(reference) = parse_simple_reference(&ident) {
        tokens.push(FormulaToken::Ref(reference));
    } else {
        tokens.push(FormulaToken::Ident(ident));
    }
}

fn parse_bracket_reference(input: &str) -> Option<FormulaToken> {
    let trimmed = input.trim();
    if let Some((start, end)) = trimmed.split_once(':') {
        let start_ref = parse_sheeted_cell(start)?;
        let end_ref = parse_sheeted_cell(end)?;
        return Some(FormulaToken::Range(CellRange {
            start: start_ref,
            end: end_ref,
        }));
    }
    parse_sheeted_cell(trimmed).map(FormulaToken::Ref)
}

fn parse_simple_reference(input: &str) -> Option<CellRef> {
    if input.chars().any(|c| c.is_ascii_digit()) && input.chars().any(|c| c.is_ascii_alphabetic()) {
        parse_sheeted_cell(input)
    } else {
        None
    }
}

fn parse_sheeted_cell(input: &str) -> Option<CellRef> {
    let trimmed = input.trim().strip_prefix('.').unwrap_or(input.trim());
    let mut sheet: Option<String> = None;
    let mut cell_part = trimmed;
    if let Some((sheet_part, cell)) = trimmed.rsplit_once('.')
        && !sheet_part.is_empty()
        && !cell.is_empty()
    {
        let sheet_name = sheet_part.trim_matches('\'').replace('$', "");
        if !sheet_name.is_empty() {
            sheet = Some(sheet_name);
        }
        cell_part = cell;
    }
    let cell = parse_cell_ref(cell_part)?;
    Some(CellRef {
        sheet: sheet.or(cell.sheet),
        row: cell.row,
        col: cell.col,
    })
}

fn parse_cell_ref(input: &str) -> Option<CellRef> {
    let mut letters = String::new();
    let mut digits = String::new();
    let mut chars = input.chars().peekable();

    if matches!(chars.peek(), Some('$')) {
        chars.next();
    }

    while let Some(&ch) = chars.peek() {
        if !ch.is_ascii_alphabetic() {
            break;
        }
        letters.push(ch);
        chars.next();
    }

    if matches!(chars.peek(), Some('$')) {
        chars.next();
    }

    while let Some(&ch) = chars.peek() {
        if !ch.is_ascii_digit() {
            break;
        }
        digits.push(ch);
        chars.next();
    }

    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    if chars.next().is_some() {
        return None;
    }

    let col = column_name_to_index(&letters)?;
    let row_number = match digits.parse::<u32>() {
        Ok(value) if value > 0 => value,
        _ => return None,
    };
    Some(CellRef {
        sheet: None,
        row: row_number - 1,
        col,
    })
}

fn column_name_to_index(name: &str) -> Option<u32> {
    let mut index: u32 = 0;
    for ch in name.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        let digit = ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1;
        index = index.checked_mul(26)?.checked_add(digit)?;
    }
    index.checked_sub(1)
}

pub(super) fn cell_value_to_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(num) => Some(*num),
        CellValue::Boolean(v) => Some(if *v { 1.0 } else { 0.0 }),
        CellValue::String(s) => s.parse::<f64>().ok(),
        _ => None,
    }
}

pub(super) fn eval_formula_function(name: &str, values: &[f64]) -> Option<f64> {
    let upper = name.to_ascii_uppercase();
    match upper.as_str() {
        "SUM" => Some(values.iter().sum()),
        "AVERAGE" => {
            if values.is_empty() {
                None
            } else {
                Some(values.iter().sum::<f64>() / values.len() as f64)
            }
        }
        "MIN" => values.iter().copied().reduce(f64::min),
        "MAX" => values.iter().copied().reduce(f64::max),
        "COUNT" => Some(values.len() as f64),
        _ => None,
    }
}
