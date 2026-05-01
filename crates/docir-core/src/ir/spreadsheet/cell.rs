use crate::ir::new_node_id;
use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// A cell within a worksheet.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Cell {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Cell reference (e.g., "A1", "B2").
    pub reference: String,

    /// Column index (0-based).
    pub column: u32,

    /// Row index (0-based).
    pub row: u32,

    /// Cell value.
    pub value: CellValue,

    /// Formula (if any).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<CellFormula>,

    /// Style ID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<u32>,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Cell {
    /// Creates a new Cell.
    pub fn new(reference: impl Into<String>, column: u32, row: u32) -> Self {
        Self {
            id: new_node_id(),
            reference: reference.into(),
            column,
            row,
            value: CellValue::Empty,
            formula: None,
            style_id: None,
            span: None,
        }
    }

    /// Returns true if this cell has a formula.
    pub fn has_formula(&self) -> bool {
        self.formula.is_some()
    }
}

/// Cell value types.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub enum CellValue {
    /// Empty cell.
    Empty,
    /// String value.
    String(String),
    /// Numeric value.
    Number(f64),
    /// Boolean value.
    Boolean(bool),
    /// Error value.
    Error(CellError),
    /// Date/time (stored as Excel serial date).
    DateTime(f64),
    /// Inline string (rich text).
    InlineString(String),
    /// Shared string reference (index into shared strings table).
    SharedString(u32),
}

/// Cell error types.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellError {
    /// #NULL!
    Null,
    /// #DIV/0!
    DivZero,
    /// #VALUE!
    Value,
    /// #REF!
    Ref,
    /// #NAME?
    Name,
    /// #NUM!
    Num,
    /// #N/A
    NA,
    /// #GETTING_DATA
    GettingData,
}

/// Cell formula information.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct CellFormula {
    /// Formula text (without leading =).
    pub text: String,

    /// Formula type.
    pub formula_type: FormulaType,

    /// Shared formula index (if shared).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shared_index: Option<u32>,

    /// Shared formula reference range (if master).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shared_ref: Option<String>,

    /// Is this an array formula?
    #[serde(default)]
    pub is_array: bool,

    /// Array formula range.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub array_ref: Option<String>,
}

/// Formula types.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FormulaType {
    /// Normal formula.
    #[default]
    Normal,
    /// Array formula.
    Array,
    /// Data table formula.
    DataTable,
    /// Shared formula.
    Shared,
}

/// Converts a 0-based column index to Excel-style letter(s).
pub fn column_to_letter(col: u32) -> String {
    let mut result = String::new();
    let mut col = col;
    loop {
        result.insert(0, (b'A' + (col % 26) as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    result
}

/// Parses an A1-style cell reference into (column, row) 0-based indices.
pub fn parse_cell_reference(reference: &str) -> Option<(u32, u32)> {
    let mut col_str = String::new();
    let mut row_str = String::new();

    for ch in reference.chars() {
        if ch.is_ascii_alphabetic() {
            col_str.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        }
    }

    if col_str.is_empty() || row_str.is_empty() {
        return None;
    }

    // Convert column letters to index
    let mut col: u32 = 0;
    for ch in col_str.chars() {
        col = col * 26 + (ch as u32 - 'A' as u32 + 1);
    }
    col -= 1; // Make 0-based

    // Convert row string to index
    let row: u32 = row_str.parse().ok()?;
    if row == 0 {
        return None;
    }

    Some((col, row - 1))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_to_letter() {
        assert_eq!(column_to_letter(0), "A");
        assert_eq!(column_to_letter(1), "B");
        assert_eq!(column_to_letter(25), "Z");
        assert_eq!(column_to_letter(26), "AA");
        assert_eq!(column_to_letter(27), "AB");
        assert_eq!(column_to_letter(701), "ZZ");
        assert_eq!(column_to_letter(702), "AAA");
    }

    #[test]
    fn test_parse_cell_reference() {
        assert_eq!(parse_cell_reference("A1"), Some((0, 0)));
        assert_eq!(parse_cell_reference("B2"), Some((1, 1)));
        assert_eq!(parse_cell_reference("Z1"), Some((25, 0)));
        assert_eq!(parse_cell_reference("AA1"), Some((26, 0)));
        assert_eq!(parse_cell_reference("AB10"), Some((27, 9)));
    }
}
