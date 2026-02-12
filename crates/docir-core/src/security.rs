//! Security-related IR nodes and types.
//!
//! This module defines types for representing security-sensitive
//! elements in Office documents: VBA macros, OLE objects, external
//! references, DDE fields, and more.

use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Aggregate security information for a document.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct SecurityInfo {
    /// VBA macro project (if present).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub macro_project: Option<NodeId>,

    /// Embedded OLE objects.
    pub ole_objects: Vec<NodeId>,

    /// External references (templates, links, etc.).
    pub external_refs: Vec<NodeId>,

    /// ActiveX controls.
    pub activex_controls: Vec<NodeId>,

    /// DDE fields found in document.
    pub dde_fields: Vec<DdeField>,

    /// Excel 4.0 XLM macros (for spreadsheets).
    pub xlm_macros: Vec<XlmMacro>,

    /// Overall threat level assessment.
    pub threat_level: ThreatLevel,

    /// Specific threat indicators found.
    pub threat_indicators: Vec<ThreatIndicator>,
}

impl SecurityInfo {
    /// Creates a new empty SecurityInfo.
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns true if any security-relevant content was found.
    pub fn has_security_content(&self) -> bool {
        self.macro_project.is_some()
            || !self.ole_objects.is_empty()
            || !self.external_refs.is_empty()
            || !self.activex_controls.is_empty()
            || !self.dde_fields.is_empty()
            || !self.xlm_macros.is_empty()
    }

    /// Recalculates the threat level based on indicators.
    pub fn recalculate_threat_level(&mut self) {
        self.threat_level = ThreatLevel::None;

        // Check for various threats in order of severity
        if self.macro_project.is_some() {
            self.threat_level = ThreatLevel::Critical;
            return;
        }

        if !self.xlm_macros.is_empty() {
            self.threat_level = ThreatLevel::Critical;
            return;
        }

        if !self.dde_fields.is_empty() {
            self.threat_level = ThreatLevel::High;
            return;
        }

        if !self.ole_objects.is_empty() || !self.activex_controls.is_empty() {
            self.threat_level = ThreatLevel::High;
            return;
        }

        // Check external refs
        for indicator in &self.threat_indicators {
            match indicator.severity {
                ThreatLevel::Critical => {
                    self.threat_level = ThreatLevel::Critical;
                    return;
                }
                ThreatLevel::High => {
                    if self.threat_level < ThreatLevel::High {
                        self.threat_level = ThreatLevel::High;
                    }
                }
                ThreatLevel::Medium => {
                    if self.threat_level < ThreatLevel::Medium {
                        self.threat_level = ThreatLevel::Medium;
                    }
                }
                ThreatLevel::Low => {
                    if self.threat_level < ThreatLevel::Low {
                        self.threat_level = ThreatLevel::Low;
                    }
                }
                ThreatLevel::None => {}
            }
        }

        if !self.external_refs.is_empty() && self.threat_level < ThreatLevel::Medium {
            self.threat_level = ThreatLevel::Medium;
        }
    }
}

/// Threat level classification.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ThreatLevel {
    /// No security concerns detected.
    None,
    /// Minor concerns (e.g., hyperlinks).
    Low,
    /// Moderate concerns (e.g., external templates).
    Medium,
    /// Significant concerns (e.g., OLE objects, DDE).
    High,
    /// Critical concerns (e.g., VBA macros with auto-exec).
    Critical,
}

impl Default for ThreatLevel {
    fn default() -> Self {
        Self::None
    }
}

impl std::fmt::Display for ThreatLevel {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "NONE"),
            Self::Low => write!(f, "LOW"),
            Self::Medium => write!(f, "MEDIUM"),
            Self::High => write!(f, "HIGH"),
            Self::Critical => write!(f, "CRITICAL"),
        }
    }
}

/// A specific threat indicator.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ThreatIndicator {
    /// Type of threat.
    pub indicator_type: ThreatIndicatorType,

    /// Severity level.
    pub severity: ThreatLevel,

    /// Human-readable description.
    pub description: String,

    /// Location in document (if applicable).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub location: Option<String>,

    /// Related node ID (if applicable).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub node_id: Option<NodeId>,
}

/// Types of threat indicators.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThreatIndicatorType {
    /// VBA macro with auto-execute.
    AutoExecMacro,
    /// Suspicious VBA API call.
    SuspiciousApiCall,
    /// External template reference.
    ExternalTemplate,
    /// Remote image/resource.
    RemoteResource,
    /// DDE command.
    DdeCommand,
    /// OLE object.
    OleObject,
    /// ActiveX control.
    ActiveXControl,
    /// Excel 4.0 XLM macro.
    XlmMacro,
    /// Hidden sheet with macros.
    HiddenMacroSheet,
    /// Suspicious formula.
    SuspiciousFormula,
    /// External hyperlink to suspicious domain.
    SuspiciousLink,
}

/// ActiveX control.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct ActiveXControl {
    /// Unique identifier for this node.
    pub id: NodeId,
    /// Control name (if provided).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// CLSID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub clsid: Option<String>,
    /// ProgID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prog_id: Option<String>,
    /// Control properties (raw key/value).
    pub properties: Vec<(String, String)>,
    /// Source span.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl ActiveXControl {
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            clsid: None,
            prog_id: None,
            properties: Vec::new(),
            span: None,
        }
    }
}

/// VBA macro project.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct MacroProject {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Project name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,

    /// Modules in this project.
    pub modules: Vec<NodeId>,

    /// Does this project have auto-execute macros?
    pub has_auto_exec: bool,

    /// Auto-execute procedure names found.
    pub auto_exec_procedures: Vec<String>,

    /// External references/libraries.
    pub references: Vec<MacroReference>,

    /// Is the project password protected?
    pub is_protected: bool,

    /// Digital signature information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub digital_signature: Option<DigitalSignature>,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl MacroProject {
    /// Creates a new MacroProject.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            name: None,
            modules: Vec::new(),
            has_auto_exec: false,
            auto_exec_procedures: Vec::new(),
            references: Vec::new(),
            is_protected: false,
            digital_signature: None,
            span: None,
        }
    }

    /// Returns all child node IDs.
    pub fn children(&self) -> Vec<NodeId> {
        self.modules.clone()
    }
}

impl Default for MacroProject {
    fn default() -> Self {
        Self::new()
    }
}

/// VBA module within a project.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct MacroModule {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Module name.
    pub name: String,

    /// Module type.
    pub module_type: MacroModuleType,

    /// Procedure names defined in this module.
    pub procedures: Vec<String>,

    /// Suspicious API calls found.
    pub suspicious_calls: Vec<SuspiciousCall>,

    /// Source code (if extracted).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_code: Option<String>,

    /// Source code hash (SHA-256, hex).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_hash: Option<String>,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl MacroModule {
    /// Creates a new MacroModule.
    pub fn new(name: impl Into<String>, module_type: MacroModuleType) -> Self {
        Self {
            id: NodeId::new(),
            name: name.into(),
            module_type,
            procedures: Vec::new(),
            suspicious_calls: Vec::new(),
            source_code: None,
            source_hash: None,
            span: None,
        }
    }
}

/// VBA module types.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MacroModuleType {
    /// Standard code module.
    Standard,
    /// Class module.
    Class,
    /// User form.
    UserForm,
    /// Document module (ThisDocument, ThisWorkbook).
    Document,
}

/// A suspicious API call in VBA code.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct SuspiciousCall {
    /// API/function name.
    pub name: String,
    /// Category of suspicion.
    pub category: SuspiciousCallCategory,
    /// Line number in source (if available).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line: Option<u32>,
}

/// Categories of suspicious API calls.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SuspiciousCallCategory {
    /// Shell execution (Shell, WScript.Shell).
    ShellExecution,
    /// File system operations (FileSystemObject).
    FileSystem,
    /// Network operations (XMLHTTP, WinHTTP).
    Network,
    /// Process manipulation (CreateObject, GetObject).
    ProcessManipulation,
    /// Registry access.
    Registry,
    /// Windows API calls (Declare Function).
    WindowsApi,
    /// PowerShell execution.
    PowerShell,
    /// Encoding/obfuscation (Base64, Chr).
    Obfuscation,
}

/// External library reference in VBA.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct MacroReference {
    /// Reference name.
    pub name: String,
    /// Library GUID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub guid: Option<String>,
    /// Library path.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
    /// Major version.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub major_version: Option<u32>,
    /// Minor version.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minor_version: Option<u32>,
}

/// Digital signature information.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct DigitalSignature {
    /// Is the signature valid?
    pub is_valid: bool,
    /// Signer common name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub signer: Option<String>,
    /// Certificate issuer.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub issuer: Option<String>,
    /// Certificate serial number.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub serial: Option<String>,
    /// Signature timestamp.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub timestamp: Option<String>,
}

/// Embedded OLE object.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct OleObject {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// ProgID (e.g., "Excel.Sheet.12").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prog_id: Option<String>,

    /// Class ID (GUID).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub class_id: Option<String>,

    /// Object name/label.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,

    /// Is this a linked object (vs embedded)?
    pub is_linked: bool,

    /// Link target (if linked).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub link_target: Option<String>,

    /// Size in bytes.
    pub size_bytes: u64,

    /// SHA-256 hash of object data (hex).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_hash: Option<String>,

    /// Relationship ID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl OleObject {
    /// Creates a new OleObject.
    pub fn new() -> Self {
        Self {
            id: NodeId::new(),
            prog_id: None,
            class_id: None,
            name: None,
            is_linked: false,
            link_target: None,
            size_bytes: 0,
            data_hash: None,
            relationship_id: None,
            span: None,
        }
    }
}

impl Default for OleObject {
    fn default() -> Self {
        Self::new()
    }
}

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

    /// Returns true if this is a remote (http/https) reference.
    pub fn is_remote(&self) -> bool {
        let target_lower = self.target.to_lowercase();
        target_lower.starts_with("http://") || target_lower.starts_with("https://")
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
