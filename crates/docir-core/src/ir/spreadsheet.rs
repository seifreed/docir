//! Spreadsheet (XLSX) IR nodes.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// A worksheet within a spreadsheet.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Worksheet {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Sheet name.
    pub name: String,

    /// Sheet ID (internal).
    pub sheet_id: u32,

    /// Relationship ID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,

    /// Sheet visibility state.
    pub state: SheetState,

    /// Sheet kind (worksheet, chart sheet, macro sheet, etc.).
    pub kind: SheetKind,

    /// Cells containing data.
    pub cells: Vec<NodeId>,

    /// Merged cell ranges.
    pub merged_cells: Vec<MergedCellRange>,

    /// Column definitions.
    pub columns: Vec<ColumnDefinition>,

    /// Drawings/graphics linked to this sheet.
    pub drawings: Vec<NodeId>,

    /// Table definitions linked to this sheet.
    #[serde(default)]
    pub tables: Vec<NodeId>,

    /// Conditional formatting rules linked to this sheet.
    #[serde(default)]
    pub conditional_formats: Vec<NodeId>,

    /// Data validation rules linked to this sheet.
    #[serde(default)]
    pub data_validations: Vec<NodeId>,

    /// Pivot tables linked to this sheet.
    #[serde(default)]
    pub pivot_tables: Vec<NodeId>,

    /// Comments linked to this sheet.
    #[serde(default)]
    pub comments: Vec<NodeId>,

    /// Dimension reference (e.g., A1:C10).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dimension: Option<String>,

    /// Sheet tab color.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,

    /// Page margins settings.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<SheetPageMargins>,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Worksheet {
    /// Creates a new Worksheet.
    pub fn new(name: impl Into<String>, sheet_id: u32) -> Self {
        Self {
            id: NodeId::new(),
            name: name.into(),
            sheet_id,
            relationship_id: None,
            state: SheetState::Visible,
            kind: SheetKind::Worksheet,
            cells: Vec::new(),
            merged_cells: Vec::new(),
            columns: Vec::new(),
            drawings: Vec::new(),
            tables: Vec::new(),
            conditional_formats: Vec::new(),
            data_validations: Vec::new(),
            pivot_tables: Vec::new(),
            comments: Vec::new(),
            dimension: None,
            tab_color: None,
            page_margins: None,
            span: None,
        }
    }

    /// Returns all child node IDs.
    pub fn children(&self) -> Vec<NodeId> {
        let mut out = self.cells.clone();
        out.extend(self.drawings.iter().copied());
        out.extend(self.tables.iter().copied());
        out.extend(self.conditional_formats.iter().copied());
        out.extend(self.data_validations.iter().copied());
        out.extend(self.pivot_tables.iter().copied());
        out.extend(self.comments.iter().copied());
        out
    }
}

/// Sheet visibility state.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetState {
    /// Normal visible sheet.
    Visible,
    /// Hidden (can be unhidden by user).
    Hidden,
    /// Very hidden (can only be unhidden via VBA).
    VeryHidden,
}

/// Sheet kind.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetKind {
    Worksheet,
    ChartSheet,
    DialogSheet,
    MacroSheet,
}

/// Worksheet page margins.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SheetPageMargins {
    pub left: Option<f64>,
    pub right: Option<f64>,
    pub top: Option<f64>,
    pub bottom: Option<f64>,
    pub header: Option<f64>,
    pub footer: Option<f64>,
}

impl Default for SheetKind {
    fn default() -> Self {
        Self::Worksheet
    }
}

impl Default for SheetState {
    fn default() -> Self {
        Self::Visible
    }
}

/// Merged cell range specification.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct MergedCellRange {
    /// Start column (0-based).
    pub start_col: u32,
    /// Start row (0-based).
    pub start_row: u32,
    /// End column (0-based, inclusive).
    pub end_col: u32,
    /// End row (0-based, inclusive).
    pub end_row: u32,
}

impl MergedCellRange {
    /// Returns the A1-style reference for this range.
    pub fn to_a1_notation(&self) -> String {
        format!(
            "{}{}:{}{}",
            column_to_letter(self.start_col),
            self.start_row + 1,
            column_to_letter(self.end_col),
            self.end_row + 1
        )
    }
}

/// Column definition.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ColumnDefinition {
    /// Column index (0-based).
    pub index: u32,
    /// Column width in character units.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    /// Is column hidden?
    #[serde(default)]
    pub hidden: bool,
    /// Custom width?
    #[serde(default)]
    pub custom_width: bool,
}

/// Worksheet drawing container.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct WorksheetDrawing {
    pub id: NodeId,
    pub shapes: Vec<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl WorksheetDrawing {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            shapes: Vec::new(),
            span: None,
        }
    }

    pub fn children(&self) -> Vec<NodeId> {
        self.shapes.clone()
    }
}

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

/// Spreadsheet comment (legacy or threaded).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SheetComment {
    pub id: NodeId,
    pub sheet_name: Option<String>,
    pub cell_ref: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl SheetComment {
    pub fn new(cell_ref: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            sheet_name: None,
            cell_ref: cell_ref.into(),
            author: None,
            text: text.into(),
            span: None,
        }
    }
}

/// Spreadsheet metadata part (xl/metadata.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SheetMetadata {
    pub id: NodeId,
    pub metadata_types: Vec<SheetMetadataType>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cell_metadata_count: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value_metadata_count: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl SheetMetadata {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            metadata_types: Vec::new(),
            cell_metadata_count: None,
            value_metadata_count: None,
            span: None,
        }
    }
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SheetMetadataType {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub min_supported_version: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub copy: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub update: Option<bool>,
}

impl SheetMetadataType {
    pub fn new() -> Self {
        Self {
            name: None,
            min_supported_version: None,
            copy: None,
            update: None,
        }
    }
}

/// Calculation chain entries.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct CalcChain {
    pub id: NodeId,
    pub entries: Vec<CalcChainEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl CalcChain {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            entries: Vec::new(),
            span: None,
        }
    }
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct CalcChainEntry {
    pub cell_ref: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub index: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub level: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub new_value: Option<bool>,
}

impl Cell {
    /// Creates a new Cell.
    pub fn new(reference: impl Into<String>, column: u32, row: u32) -> Self {
        Self {
            id: NodeId::new(),
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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaType {
    /// Normal formula.
    Normal,
    /// Array formula.
    Array,
    /// Data table formula.
    DataTable,
    /// Shared formula.
    Shared,
}

impl Default for FormulaType {
    fn default() -> Self {
        Self::Normal
    }
}

/// Shared strings table.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SharedStringTable {
    pub id: NodeId,
    pub items: Vec<SharedStringItem>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl SharedStringTable {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            items: Vec::new(),
            span: None,
        }
    }
}

/// Shared string item (rich text allowed).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SharedStringItem {
    pub text: String,
    #[serde(default)]
    pub runs: Vec<String>,
}

/// Spreadsheet styles (styles.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SpreadsheetStyles {
    pub id: NodeId,
    pub number_formats: Vec<NumberFormat>,
    pub fonts: Vec<FontDef>,
    pub fills: Vec<FillDef>,
    pub borders: Vec<BorderDef>,
    pub cell_xfs: Vec<CellFormat>,
    pub cell_style_xfs: Vec<CellFormat>,
    pub dxfs: Vec<DxfStyle>,
    pub table_styles: Option<TableStyleInfo>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl SpreadsheetStyles {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            number_formats: Vec::new(),
            fonts: Vec::new(),
            fills: Vec::new(),
            borders: Vec::new(),
            cell_xfs: Vec::new(),
            cell_style_xfs: Vec::new(),
            dxfs: Vec::new(),
            table_styles: None,
            span: None,
        }
    }
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct NumberFormat {
    pub id: u32,
    pub format_code: String,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct FontDef {
    pub name: Option<String>,
    pub size: Option<f64>,
    #[serde(default)]
    pub bold: bool,
    #[serde(default)]
    pub italic: bool,
    #[serde(default)]
    pub underline: bool,
    pub color: Option<String>,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct FillDef {
    pub pattern_type: Option<String>,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct BorderDef {
    pub left: Option<BorderSide>,
    pub right: Option<BorderSide>,
    pub top: Option<BorderSide>,
    pub bottom: Option<BorderSide>,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct BorderSide {
    pub style: Option<String>,
    pub color: Option<String>,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct CellFormat {
    pub num_fmt_id: Option<u32>,
    pub font_id: Option<u32>,
    pub fill_id: Option<u32>,
    pub border_id: Option<u32>,
    pub xf_id: Option<u32>,
    #[serde(default)]
    pub apply_number_format: bool,
    #[serde(default)]
    pub apply_font: bool,
    #[serde(default)]
    pub apply_fill: bool,
    #[serde(default)]
    pub apply_border: bool,
    #[serde(default)]
    pub apply_alignment: bool,
    #[serde(default)]
    pub apply_protection: bool,
    #[serde(default)]
    pub quote_prefix: bool,
    #[serde(default)]
    pub pivot_button: bool,
    pub alignment: Option<CellAlignment>,
    pub protection: Option<CellProtection>,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct CellAlignment {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: bool,
    pub indent: Option<u32>,
    pub text_rotation: Option<i32>,
    pub shrink_to_fit: bool,
    pub reading_order: Option<u32>,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct CellProtection {
    pub locked: Option<bool>,
    pub hidden: Option<bool>,
}

/// Differential format (dxf) style.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct DxfStyle {
    pub num_fmt: Option<NumberFormat>,
    pub font: Option<FontDef>,
    pub fill: Option<FillDef>,
    pub border: Option<BorderDef>,
    pub alignment: Option<CellAlignment>,
    pub protection: Option<CellProtection>,
}

impl DxfStyle {
    pub fn new() -> Self {
        Self {
            num_fmt: None,
            font: None,
            fill: None,
            border: None,
            alignment: None,
            protection: None,
        }
    }
}

/// Table styles info from styles.xml.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct TableStyleInfo {
    pub count: Option<u32>,
    pub default_table_style: Option<String>,
    pub default_pivot_style: Option<String>,
    #[serde(default)]
    pub styles: Vec<TableStyleDef>,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct TableStyleDef {
    pub name: String,
    pub pivot: Option<bool>,
    pub table: Option<bool>,
}

/// Defined name (named range / formula).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct DefinedName {
    pub id: NodeId,
    pub name: String,
    pub value: String,
    pub local_sheet_id: Option<u32>,
    #[serde(default)]
    pub hidden: bool,
    pub comment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

/// Conditional formatting container.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ConditionalFormat {
    pub id: NodeId,
    pub ranges: Vec<String>,
    pub rules: Vec<ConditionalRule>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ConditionalRule {
    pub rule_type: String,
    pub priority: Option<u32>,
    pub operator: Option<String>,
    pub formulae: Vec<String>,
}

/// Data validation rule.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct DataValidation {
    pub id: NodeId,
    pub validation_type: Option<String>,
    pub operator: Option<String>,
    #[serde(default)]
    pub allow_blank: bool,
    #[serde(default)]
    pub show_input_message: bool,
    #[serde(default)]
    pub show_error_message: bool,
    pub error_title: Option<String>,
    pub error: Option<String>,
    pub prompt_title: Option<String>,
    pub prompt: Option<String>,
    pub ranges: Vec<String>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

/// Table definition (ListObject).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct TableDefinition {
    pub id: NodeId,
    pub name: Option<String>,
    pub display_name: Option<String>,
    pub ref_range: Option<String>,
    pub header_row_count: Option<u32>,
    pub totals_row_count: Option<u32>,
    pub columns: Vec<TableColumn>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct TableColumn {
    pub id: u32,
    pub name: Option<String>,
    pub totals_row_label: Option<String>,
    pub totals_row_function: Option<String>,
}

/// Pivot table definition.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct PivotTable {
    pub id: NodeId,
    pub name: Option<String>,
    pub cache_id: Option<u32>,
    pub ref_range: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

/// Pivot cache definition.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct PivotCache {
    pub id: NodeId,
    pub cache_id: u32,
    pub cache_source: Option<String>,
    pub record_count: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub records: Option<NodeId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl PivotCache {
    pub fn new(cache_id: u32) -> Self {
        Self {
            id: NodeId::new(),
            cache_id,
            cache_source: None,
            record_count: None,
            records: None,
            span: None,
        }
    }
}

/// Pivot cache records (pivotCacheRecords*.xml).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct PivotCacheRecords {
    pub id: NodeId,
    pub cache_id: Option<u32>,
    pub record_count: Option<u32>,
    pub field_count: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl PivotCacheRecords {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            cache_id: None,
            record_count: None,
            field_count: None,
            span: None,
        }
    }
}

/// Workbook-level properties.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct WorkbookProperties {
    pub id: NodeId,
    pub date1904: Option<bool>,
    pub calc_mode: Option<String>,
    pub calc_full: Option<bool>,
    pub calc_on_save: Option<bool>,
    pub workbook_protected: bool,
    pub active_tab: Option<u32>,
    pub first_sheet: Option<u32>,
    pub show_horizontal_scroll: Option<bool>,
    pub show_vertical_scroll: Option<bool>,
    pub show_sheet_tabs: Option<bool>,
    pub tab_ratio: Option<u32>,
    pub window_width: Option<u32>,
    pub window_height: Option<u32>,
    pub x_window: Option<i32>,
    pub y_window: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl WorkbookProperties {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            date1904: None,
            calc_mode: None,
            calc_full: None,
            calc_on_save: None,
            workbook_protected: false,
            active_tab: None,
            first_sheet: None,
            show_horizontal_scroll: None,
            show_vertical_scroll: None,
            show_sheet_tabs: None,
            tab_ratio: None,
            window_width: None,
            window_height: None,
            x_window: None,
            y_window: None,
            span: None,
        }
    }
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
