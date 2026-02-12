use super::column_to_letter;
use crate::ir::new_node_id;
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
            id: new_node_id(),
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
            id: new_node_id(),
            shapes: Vec::new(),
            span: None,
        }
    }

    pub fn children(&self) -> Vec<NodeId> {
        self.shapes.clone()
    }
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
            id: new_node_id(),
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
            id: new_node_id(),
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
            id: new_node_id(),
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
