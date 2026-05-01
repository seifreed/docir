use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// External reference (template, link, remote resource).
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ExternalReference {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// Reference type.
    pub ref_type: ExternalRefType,
    /// Target URL or path.
    pub target: String,
    /// Relationship ID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,
    /// Relationship type URI.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationship_type: Option<String>,
    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl ExternalReference {
    /// Creates a new ExternalReference.
    pub fn new(ref_type: ExternalRefType, target: impl Into<String>) -> Self {
        Self {
            id: NodeId::new(),
            ref_type,
            target: target.into(),
            relationship_id: None,
            relationship_type: None,
            span: None,
        }
    }

    /// Returns true if this is a remote reference (http/https/ftp/etc).
    pub fn is_remote(&self) -> bool {
        let target_lower = self.target.to_lowercase();
        target_lower.starts_with("http://")
            || target_lower.starts_with("https://")
            || target_lower.starts_with("ftp://")
            || target_lower.starts_with("smb://")
            || target_lower.starts_with("sftp://")
            || target_lower.starts_with("ldap://")
            || target_lower.starts_with("ldaps://")
            || target_lower.starts_with("file://")
    }
}

/// Types of external references.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExternalRefType {
    /// Attached document template.
    AttachedTemplate,
    /// Hyperlink.
    Hyperlink,
    /// External image.
    Image,
    /// OLE link.
    OleLink,
    /// Frame/subdocument.
    Frame,
    /// External data connection.
    DataConnection,
    /// Slide master/layout.
    SlideMaster,
    /// External stylesheet.
    Stylesheet,
    /// Unknown/other.
    Other,
}

/// DDE (Dynamic Data Exchange) field.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct DdeField {
    /// Field type (DDE or DDEAUTO).
    pub field_type: DdeFieldType,
    /// Application name.
    pub application: String,
    /// Topic/document.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub topic: Option<String>,
    /// Item/command.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub item: Option<String>,
    /// Raw field instruction text.
    pub instruction: String,
    /// Location in document.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub location: Option<SourceSpan>,
}

/// DDE field types.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeFieldType {
    /// Manual update.
    Dde,
    /// Automatic update on document open.
    DdeAuto,
}

/// Excel 4.0 XLM macro.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct XlmMacro {
    /// Sheet name containing the macro.
    pub sheet_name: String,
    /// Sheet visibility.
    pub sheet_state: crate::ir::SheetState,
    /// Dangerous function calls found.
    pub dangerous_functions: Vec<XlmFunction>,
    /// Cell references with macros.
    pub macro_cells: Vec<XlmMacroCell>,
    /// Is there an Auto_Open trigger?
    pub has_auto_open: bool,
}

/// A dangerous XLM function.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct XlmFunction {
    /// Function name (e.g., EXEC, CALL).
    pub name: String,
    /// Arguments (if extractable).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub arguments: Option<String>,
    /// Cell reference.
    pub cell_ref: String,
}

/// A cell containing XLM macro code.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct XlmMacroCell {
    /// Cell reference (e.g., A1).
    pub cell_ref: String,
    /// Formula text.
    pub formula: String,
}
