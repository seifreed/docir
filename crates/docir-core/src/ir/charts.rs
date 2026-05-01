//! Chart IR nodes (shared for XLSX/PPTX).

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Chart data extracted from chart parts.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ChartData {
    pub id: NodeId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chart_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    /// Legacy series list (names only) for backward compatibility.
    pub series: Vec<String>,
    /// Detailed series data.
    pub series_data: Vec<ChartSeries>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl ChartData {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            chart_type: None,
            title: None,
            series: Vec::new(),
            series_data: Vec::new(),
            span: None,
        }
    }
}

/// Chart series data (name + categories/values).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ChartSeries {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub categories: Vec<String>,
    pub values: Vec<String>,
}

impl ChartSeries {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            name: None,
            categories: Vec::new(),
            values: Vec::new(),
        }
    }
}
