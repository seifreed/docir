use super::*;

#[derive(Debug, Clone)]
struct CellRef {
    sheet: Option<String>,
    row: u32,
    col: u32,
}

#[derive(Debug, Clone)]
struct CellRange {
    start: CellRef,
    end: CellRef,
}

#[derive(Debug, Clone)]
enum FormulaToken {
    Number(f64),
    Ident(String),
    Ref(CellRef),
    Range(CellRange),
    Plus,
    Minus,
    Star,
    Slash,
    LParen,
    RParen,
    Comma,
    End,
}

struct FormulaEvalContext<'a> {
    sheet_name: &'a str,
    values: HashMap<(u32, u32), CellValue>,
    formulas: &'a HashMap<(u32, u32), String>,
    cache: HashMap<(u32, u32), Option<f64>>,
    stack: Vec<(u32, u32)>,
}

impl<'a> FormulaEvalContext<'a> {
    fn new(
        sheet_name: &'a str,
        values: HashMap<(u32, u32), CellValue>,
        formulas: &'a HashMap<(u32, u32), String>,
    ) -> Self {
        Self {
            sheet_name,
            values,
            formulas,
            cache: HashMap::new(),
            stack: Vec::new(),
        }
    }

    fn eval_formula(&mut self, formula: &str) -> Option<f64> {
        let tokens = tokenize_formula(formula);
        let mut parser = FormulaParser::new(tokens, self);
        parser.parse_expression()
    }

    fn resolve_ref(&mut self, reference: &CellRef) -> Option<f64> {
        if let Some(sheet) = reference.sheet.as_deref() {
            if !sheet.eq_ignore_ascii_case(self.sheet_name) {
                return None;
            }
        }
        let key = (reference.row, reference.col);
        if let Some(value) = self.cache.get(&key) {
            return *value;
        }
        if self.stack.contains(&key) {
            self.cache.insert(key, None);
            return None;
        }
        if let Some(value) = self.values.get(&key) {
            if let Some(number) = cell_value_to_number(value) {
                self.cache.insert(key, Some(number));
                return Some(number);
            }
        }
        let formula_text = self.formulas.get(&key)?.clone();
        self.stack.push(key);
        let result = self.eval_formula(&formula_text);
        self.stack.pop();
        if let Some(number) = result {
            self.values.insert(key, CellValue::Number(number));
        }
        self.cache.insert(key, result);
        result
    }

    fn resolve_range(&mut self, range: &CellRange) -> Option<Vec<f64>> {
        if let Some(sheet) = range.start.sheet.as_deref() {
            if !sheet.eq_ignore_ascii_case(self.sheet_name) {
                return None;
            }
        }
        if let Some(sheet) = range.end.sheet.as_deref() {
            if !sheet.eq_ignore_ascii_case(self.sheet_name) {
                return None;
            }
        }
        let row_start = range.start.row.min(range.end.row);
        let row_end = range.start.row.max(range.end.row);
        let col_start = range.start.col.min(range.end.col);
        let col_end = range.start.col.max(range.end.col);
        let total = (row_end - row_start + 1) as u64 * (col_end - col_start + 1) as u64;
        if total > 1_000_000 {
            return None;
        }
        let mut values = Vec::new();
        for row in row_start..=row_end {
            for col in col_start..=col_end {
                let reference = CellRef {
                    sheet: None,
                    row,
                    col,
                };
                if let Some(number) = self.resolve_ref(&reference) {
                    values.push(number);
                }
            }
        }
        Some(values)
    }
}

struct FormulaParser<'a, 'b> {
    tokens: Vec<FormulaToken>,
    pos: usize,
    ctx: &'a mut FormulaEvalContext<'b>,
}

impl<'a, 'b> FormulaParser<'a, 'b> {
    fn new(tokens: Vec<FormulaToken>, ctx: &'a mut FormulaEvalContext<'b>) -> Self {
        Self {
            tokens,
            pos: 0,
            ctx,
        }
    }

    fn parse_expression(&mut self) -> Option<f64> {
        let mut value = self.parse_term()?;
        loop {
            match self.peek() {
                FormulaToken::Plus => {
                    self.next();
                    value += self.parse_term()?;
                }
                FormulaToken::Minus => {
                    self.next();
                    value -= self.parse_term()?;
                }
                _ => break,
            }
        }
        Some(value)
    }

    fn parse_term(&mut self) -> Option<f64> {
        let mut value = self.parse_factor()?;
        loop {
            match self.peek() {
                FormulaToken::Star => {
                    self.next();
                    value *= self.parse_factor()?;
                }
                FormulaToken::Slash => {
                    self.next();
                    let denom = self.parse_factor()?;
                    if denom == 0.0 {
                        return None;
                    }
                    value /= denom;
                }
                _ => break,
            }
        }
        Some(value)
    }

    fn parse_factor(&mut self) -> Option<f64> {
        let token = self.peek().clone();
        match token {
            FormulaToken::Minus => {
                self.next();
                self.parse_factor().map(|v| -v)
            }
            FormulaToken::Number(value) => {
                self.next();
                Some(value)
            }
            FormulaToken::Ref(reference) => {
                self.next();
                self.ctx.resolve_ref(&reference)
            }
            FormulaToken::Range(range) => {
                self.next();
                let values = self.ctx.resolve_range(&range)?;
                Some(values.iter().sum())
            }
            FormulaToken::Ident(name) => {
                self.next();
                if matches!(self.peek(), FormulaToken::LParen) {
                    self.next();
                    let values = self.parse_function_args()?;
                    if !matches!(self.peek(), FormulaToken::RParen) {
                        return None;
                    }
                    self.next();
                    eval_formula_function(&name, &values)
                } else {
                    None
                }
            }
            FormulaToken::LParen => {
                self.next();
                let value = self.parse_expression()?;
                if !matches!(self.peek(), FormulaToken::RParen) {
                    return None;
                }
                self.next();
                Some(value)
            }
            _ => None,
        }
    }

    fn parse_function_args(&mut self) -> Option<Vec<f64>> {
        let mut values = Vec::new();
        if matches!(self.peek(), FormulaToken::RParen) {
            return Some(values);
        }
        loop {
            if matches!(self.peek(), FormulaToken::Range(_)) {
                if let FormulaToken::Range(range) = self.next().clone() {
                    let range_values = self.ctx.resolve_range(&range)?;
                    values.extend(range_values);
                }
            } else {
                let value = self.parse_expression()?;
                values.push(value);
            }
            match self.peek() {
                FormulaToken::Comma => {
                    self.next();
                }
                FormulaToken::RParen => break,
                _ => return None,
            }
        }
        Some(values)
    }

    fn peek(&self) -> &FormulaToken {
        self.tokens.get(self.pos).unwrap_or(&FormulaToken::End)
    }

    fn next(&mut self) -> &FormulaToken {
        let token = self.tokens.get(self.pos).unwrap_or(&FormulaToken::End);
        self.pos += 1;
        token
    }
}

fn tokenize_formula(formula: &str) -> Vec<FormulaToken> {
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
    if let Ok(value) = num.parse::<f64>() {
        tokens.push(FormulaToken::Number(value));
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
    let trimmed = input.trim().trim_start_matches('.');
    let mut sheet: Option<String> = None;
    let mut cell_part = trimmed;
    if let Some((sheet_part, cell)) = trimmed.rsplit_once('.') {
        if !sheet_part.is_empty() && !cell.is_empty() {
            let sheet_name = sheet_part.trim_matches('\'').replace('$', "");
            if !sheet_name.is_empty() {
                sheet = Some(sheet_name);
            }
            cell_part = cell;
        }
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
    for ch in input.chars() {
        if ch.is_ascii_alphabetic() {
            letters.push(ch);
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else if ch == '$' {
            continue;
        } else {
            break;
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let col = column_name_to_index(&letters)?;
    let row = digits.parse::<u32>().ok()?.saturating_sub(1);
    Some(CellRef {
        sheet: None,
        row,
        col,
    })
}

fn column_name_to_index(name: &str) -> Option<u32> {
    let mut index: u32 = 0;
    for ch in name.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        index = index * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    Some(index.saturating_sub(1))
}

fn cell_value_to_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(num) => Some(*num),
        CellValue::Boolean(v) => Some(if *v { 1.0 } else { 0.0 }),
        CellValue::String(s) => s.parse::<f64>().ok(),
        _ => None,
    }
}

fn eval_formula_function(name: &str, values: &[f64]) -> Option<f64> {
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

pub(super) fn evaluate_ods_formulas(
    sheet_name: &str,
    formula_cells: &[(NodeId, u32, u32, String)],
    store: &mut IrStore,
    cell_values: &mut HashMap<(u32, u32), CellValue>,
    formula_map: &HashMap<(u32, u32), String>,
) {
    let mut ctx = FormulaEvalContext::new(sheet_name, cell_values.clone(), formula_map);
    for (cell_id, row, col, formula) in formula_cells {
        if let Some(IRNode::Cell(cell)) = store.get_mut(*cell_id) {
            if matches!(cell.value, CellValue::Empty) {
                if let Some(value) = ctx.eval_formula(formula) {
                    cell.value = CellValue::Number(value);
                    cell_values.insert((*row, *col), CellValue::Number(value));
                }
            }
        }
    }
}
