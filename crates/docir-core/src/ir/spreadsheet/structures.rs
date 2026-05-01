use crate::ir::new_node_id;
use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

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
    /// Public API entrypoint: new.
    pub fn new(cache_id: u32) -> Self {
        Self {
            id: new_node_id(),
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_node_id(),
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
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: new_node_id(),
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
