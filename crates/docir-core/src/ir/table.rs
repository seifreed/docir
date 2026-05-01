//! Table IR nodes.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// A table structure.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct GridColumn {
    /// Column width in twips.
    pub width: u32,
}

/// Table properties.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct TableWidth {
    /// Width value.
    pub value: u32,
    /// Width type.
    pub width_type: TableWidthType,
}

/// Table width type.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableAlignment {
    Left,
    Center,
    Right,
}

/// Table borders.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct RowHeight {
    /// Height value in twips.
    pub value: u32,
    /// Height rule.
    pub rule: RowHeightRule,
}

/// Row height rules.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RowHeightRule {
    Auto,
    Exact,
    AtLeast,
}

/// A table cell.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MergeType {
    /// Start of a merged region.
    Restart,
    /// Continuation of a merged region.
    Continue,
}

/// Vertical alignment in cells.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellVerticalAlignment {
    Top,
    Center,
    Bottom,
}

#[cfg(test)]
mod tests {
    use super::{
        Border, BorderStyle, CellMargins, CellVerticalAlignment, MergeType, RowHeight,
        RowHeightRule, Table, TableAlignment, TableBorders, TableCell, TableCellProperties,
        TableProperties, TableRow, TableWidth, TableWidthType,
    };
    use crate::types::NodeId;

    #[test]
    fn table_new_and_children_helpers() {
        let mut table = Table::new();
        let row_id = NodeId::new();
        table.rows.push(row_id);
        table.grid.push(super::GridColumn { width: 1800 });
        table.grid.push(super::GridColumn { width: 2200 });

        assert_eq!(table.children(), vec![row_id]);
        assert_eq!(table.column_count(), 2);
        assert_eq!(table.row_count(), 1);
    }

    #[test]
    fn table_row_and_cell_children_helpers() {
        let mut row = TableRow::new();
        let cell_id = NodeId::new();
        row.cells.push(cell_id);
        assert_eq!(row.children(), vec![cell_id]);

        let mut cell = TableCell::new();
        let content_id = NodeId::new();
        cell.content.push(content_id);
        assert_eq!(cell.children(), vec![content_id]);
    }

    #[test]
    fn table_property_structs_can_be_fully_populated() {
        let width = TableWidth {
            value: 5000,
            width_type: TableWidthType::Dxa,
        };
        let border = Border {
            style: BorderStyle::Single,
            width: Some(4),
            color: Some("FF0000".to_string()),
        };
        let borders = TableBorders {
            top: Some(border.clone()),
            bottom: Some(border.clone()),
            left: Some(border.clone()),
            right: Some(border.clone()),
            inside_h: Some(border.clone()),
            inside_v: Some(border),
        };
        let margins = CellMargins {
            top: Some(10),
            bottom: Some(20),
            left: Some(30),
            right: Some(40),
        };
        let props = TableProperties {
            width: Some(width.clone()),
            alignment: Some(TableAlignment::Center),
            borders: Some(borders),
            cell_margins: Some(margins),
            style_id: Some("TableStyle1".to_string()),
        };

        assert_eq!(props.width.as_ref().map(|w| w.value), Some(5000));
        assert_eq!(props.style_id.as_deref(), Some("TableStyle1"));
        assert!(matches!(props.alignment, Some(TableAlignment::Center)));
        assert!(props
            .borders
            .as_ref()
            .and_then(|b| b.inside_h.as_ref())
            .is_some());
        assert_eq!(props.cell_margins.as_ref().and_then(|m| m.left), Some(30));
    }

    #[test]
    fn table_cell_properties_hold_merge_and_alignment_variants() {
        let cell_props = TableCellProperties {
            width: Some(TableWidth {
                value: 2500,
                width_type: TableWidthType::Pct,
            }),
            vertical_merge: Some(MergeType::Restart),
            grid_span: Some(2),
            vertical_align: Some(CellVerticalAlignment::Bottom),
            borders: None,
            shading: Some("CCCCCC".to_string()),
        };
        assert_eq!(cell_props.grid_span, Some(2));
        assert!(matches!(
            cell_props.vertical_merge,
            Some(MergeType::Restart)
        ));
        assert!(matches!(
            cell_props.vertical_align,
            Some(CellVerticalAlignment::Bottom)
        ));
        assert_eq!(cell_props.shading.as_deref(), Some("CCCCCC"));
        assert!(matches!(
            cell_props.width.as_ref().map(|w| w.width_type),
            Some(TableWidthType::Pct)
        ));
    }

    #[test]
    fn row_height_rules_are_assignable() {
        let exact = RowHeight {
            value: 300,
            rule: RowHeightRule::Exact,
        };
        let at_least = RowHeight {
            value: 400,
            rule: RowHeightRule::AtLeast,
        };
        let auto = RowHeight {
            value: 0,
            rule: RowHeightRule::Auto,
        };

        assert!(matches!(exact.rule, RowHeightRule::Exact));
        assert!(matches!(at_least.rule, RowHeightRule::AtLeast));
        assert!(matches!(auto.rule, RowHeightRule::Auto));
    }
}
