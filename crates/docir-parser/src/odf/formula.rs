#[cfg(test)]
use super::Cell;
use super::{CellValue, IRNode, IrStore, NodeId};
use std::collections::HashMap;
#[path = "formula_parse_utils.rs"]
mod formula_parse_utils;
use formula_parse_utils::*;

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

#[cfg(test)]
mod tests {
    use super::*;

    fn number_from_cell(store: &IrStore, id: NodeId) -> Option<f64> {
        let IRNode::Cell(cell) = store.get(id)? else {
            return None;
        };
        match &cell.value {
            CellValue::Number(value) => Some(*value),
            _ => None,
        }
    }

    #[test]
    fn evaluate_ods_formulas_computes_ranges_and_functions() {
        let mut store = IrStore::new();
        let mut cell_values: HashMap<(u32, u32), CellValue> = HashMap::new();
        let mut formula_map: HashMap<(u32, u32), String> = HashMap::new();

        let mut a1 = Cell::new("A1".to_string(), 0, 0);
        a1.value = CellValue::Number(2.0);
        let a1_id = a1.id;
        store.insert(IRNode::Cell(a1));
        cell_values.insert((0, 0), CellValue::Number(2.0));

        let mut b1 = Cell::new("B1".to_string(), 1, 0);
        b1.value = CellValue::Number(3.0);
        let b1_id = b1.id;
        store.insert(IRNode::Cell(b1));
        cell_values.insert((0, 1), CellValue::Number(3.0));

        let c1 = Cell::new("C1".to_string(), 2, 0);
        let c1_id = c1.id;
        store.insert(IRNode::Cell(c1));

        let formula = "SUM([.A1:.B1]) * 2".to_string();
        formula_map.insert((0, 2), formula.clone());
        let formula_cells = vec![(c1_id, 0, 2, formula)];

        evaluate_ods_formulas(
            "Sheet1",
            &formula_cells,
            &mut store,
            &mut cell_values,
            &formula_map,
        );

        assert_eq!(number_from_cell(&store, a1_id), Some(2.0));
        assert_eq!(number_from_cell(&store, b1_id), Some(3.0));
        assert_eq!(number_from_cell(&store, c1_id), Some(10.0));
    }

    #[test]
    fn evaluate_ods_formulas_keeps_cells_empty_on_invalid_expressions() {
        let mut store = IrStore::new();
        let mut cell_values: HashMap<(u32, u32), CellValue> = HashMap::new();
        let mut formula_map: HashMap<(u32, u32), String> = HashMap::new();

        let mut a1 = Cell::new("A1".to_string(), 0, 0);
        a1.value = CellValue::Number(1.0);
        store.insert(IRNode::Cell(a1));
        cell_values.insert((0, 0), CellValue::Number(1.0));

        let b1 = Cell::new("B1".to_string(), 1, 0);
        let b1_id = b1.id;
        store.insert(IRNode::Cell(b1));
        let b1_formula = "A1 / 0".to_string();
        formula_map.insert((0, 1), b1_formula.clone());

        let c1 = Cell::new("C1".to_string(), 2, 0);
        let c1_id = c1.id;
        store.insert(IRNode::Cell(c1));
        let c1_formula = "[Other.A1]".to_string();
        formula_map.insert((0, 2), c1_formula.clone());

        let formula_cells = vec![(b1_id, 0, 1, b1_formula), (c1_id, 0, 2, c1_formula)];
        evaluate_ods_formulas(
            "Sheet1",
            &formula_cells,
            &mut store,
            &mut cell_values,
            &formula_map,
        );

        assert!(number_from_cell(&store, b1_id).is_none());
        assert!(number_from_cell(&store, c1_id).is_none());
    }

    #[test]
    fn evaluate_ods_formulas_handles_circular_references_as_empty() {
        let mut store = IrStore::new();
        let mut cell_values: HashMap<(u32, u32), CellValue> = HashMap::new();
        let mut formula_map: HashMap<(u32, u32), String> = HashMap::new();

        let a1 = Cell::new("A1".to_string(), 0, 0);
        let a1_id = a1.id;
        store.insert(IRNode::Cell(a1));
        formula_map.insert((0, 0), "B1 + 1".to_string());

        let b1 = Cell::new("B1".to_string(), 1, 0);
        let b1_id = b1.id;
        store.insert(IRNode::Cell(b1));
        formula_map.insert((0, 1), "A1 + 1".to_string());

        let formula_cells = vec![
            (a1_id, 0, 0, "B1 + 1".to_string()),
            (b1_id, 0, 1, "A1 + 1".to_string()),
        ];
        evaluate_ods_formulas(
            "Sheet1",
            &formula_cells,
            &mut store,
            &mut cell_values,
            &formula_map,
        );

        assert!(number_from_cell(&store, a1_id).is_none());
        assert!(number_from_cell(&store, b1_id).is_none());
    }

    #[test]
    fn evaluate_ods_formulas_does_not_override_non_empty_cell_values() {
        let mut store = IrStore::new();
        let mut cell_values: HashMap<(u32, u32), CellValue> = HashMap::new();
        let mut formula_map: HashMap<(u32, u32), String> = HashMap::new();

        let mut a1 = Cell::new("A1".to_string(), 0, 0);
        a1.value = CellValue::Number(7.0);
        let a1_id = a1.id;
        store.insert(IRNode::Cell(a1));
        cell_values.insert((0, 0), CellValue::Number(7.0));

        formula_map.insert((0, 0), "1 + 2".to_string());
        let formula_cells = vec![(a1_id, 0, 0, "1 + 2".to_string())];
        evaluate_ods_formulas(
            "Sheet1",
            &formula_cells,
            &mut store,
            &mut cell_values,
            &formula_map,
        );

        assert_eq!(number_from_cell(&store, a1_id), Some(7.0));
    }
}
