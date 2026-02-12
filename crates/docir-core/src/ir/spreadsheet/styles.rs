use crate::ir::new_node_id;
use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

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
            id: new_node_id(),
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
            id: new_node_id(),
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
