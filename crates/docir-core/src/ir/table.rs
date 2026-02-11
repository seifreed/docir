//! Table IR nodes.

use crate::types::{NodeId, SourceSpan};
use serde::{Deserialize, Serialize};

/// A table structure.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Table {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Table rows.
    pub rows: Vec<NodeId>,

    /// Grid column definitions.
    pub grid: Vec<GridColumn>,

    /// Table properties.
    pub properties: TableProperties,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Table {
    /// Creates a new empty Table.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            rows: Vec::new(),
            grid: Vec::new(),
            properties: TableProperties::default(),
            span: None,
        }
    }

    /// Returns all child node IDs.
    pub fn children(&self) -> Vec<NodeId> {
        self.rows.clone()
    }

    /// Returns the number of columns.
    pub fn column_count(&self) -> usize {
        self.grid.len()
    }

    /// Returns the number of rows.
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }
}

impl Default for Table {
    fn default() -> Self {
        Self::new()
    }
}

/// Grid column definition.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GridColumn {
    /// Column width in twips.
    pub width: u32,
}

/// Table properties.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableProperties {
    /// Table width.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<TableWidth>,

    /// Table alignment.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<TableAlignment>,

    /// Table borders.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<TableBorders>,

    /// Cell margins.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cell_margins: Option<CellMargins>,

    /// Table style ID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
}

/// Table width specification.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableWidth {
    /// Width value.
    pub value: u32,
    /// Width type.
    pub width_type: TableWidthType,
}

/// Table width type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TableWidthType {
    /// Automatic width.
    Auto,
    /// Fixed width in twips.
    Dxa,
    /// Percentage of page width (in fiftieths of a percent).
    Pct,
    /// No specific width.
    Nil,
}

/// Table alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TableAlignment {
    Left,
    Center,
    Right,
}

/// Table borders.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableBorders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inside_h: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inside_v: Option<Border>,
}

/// Border definition.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Border {
    /// Border style.
    pub style: BorderStyle,
    /// Border width in eighths of a point.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<u32>,
    /// Border color (hex without #).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

/// Border styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BorderStyle {
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    Wave,
}

/// Cell margins.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CellMargins {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<u32>,
}

/// A table row.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableRow {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Cells in this row.
    pub cells: Vec<NodeId>,

    /// Row properties.
    pub properties: TableRowProperties,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl TableRow {
    /// Creates a new empty TableRow.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            cells: Vec::new(),
            properties: TableRowProperties::default(),
            span: None,
        }
    }

    /// Returns all child node IDs.
    pub fn children(&self) -> Vec<NodeId> {
        self.cells.clone()
    }
}

impl Default for TableRow {
    fn default() -> Self {
        Self::new()
    }
}

/// Table row properties.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableRowProperties {
    /// Row height.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<RowHeight>,

    /// Is this a header row?
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_header: Option<bool>,

    /// Can this row break across pages?
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cant_split: Option<bool>,
}

/// Row height specification.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RowHeight {
    /// Height value in twips.
    pub value: u32,
    /// Height rule.
    pub rule: RowHeightRule,
}

/// Row height rules.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum RowHeightRule {
    Auto,
    Exact,
    AtLeast,
}

/// A table cell.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Content within this cell (paragraphs, tables, etc.).
    pub content: Vec<NodeId>,

    /// Cell properties.
    pub properties: TableCellProperties,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl TableCell {
    /// Creates a new empty TableCell.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            content: Vec::new(),
            properties: TableCellProperties::default(),
            span: None,
        }
    }

    /// Returns all child node IDs.
    pub fn children(&self) -> Vec<NodeId> {
        self.content.clone()
    }
}

impl Default for TableCell {
    fn default() -> Self {
        Self::new()
    }
}

/// Table cell properties.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableCellProperties {
    /// Cell width.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<TableWidth>,

    /// Vertical merge (for merged cells).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_merge: Option<MergeType>,

    /// Horizontal merge / grid span.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grid_span: Option<u32>,

    /// Vertical alignment.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<CellVerticalAlignment>,

    /// Cell borders.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<TableBorders>,

    /// Cell shading/background.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
}

/// Merge type for table cells.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum MergeType {
    /// Start of a merged region.
    Restart,
    /// Continuation of a merged region.
    Continue,
}

/// Vertical alignment in cells.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum CellVerticalAlignment {
    Top,
    Center,
    Bottom,
}
