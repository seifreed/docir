use crate::types::{NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

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
    /// Container path holding the macro project (for example `word/vbaProject.bin`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub container_path: Option<String>,
    /// Root storage/stream path within the container, if known.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub storage_root: Option<String>,
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
            container_path: None,
            storage_root: None,
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
    /// Raw stream name used inside the VBA storage.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stream_name: Option<String>,
    /// Raw stream/storage path inside the container.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stream_path: Option<String>,
    /// Procedure names defined in this module.
    pub procedures: Vec<String>,
    /// Suspicious API calls found.
    pub suspicious_calls: Vec<SuspiciousCall>,
    /// Source code (if extracted).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_code: Option<String>,
    /// Source encoding after normalization.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_encoding: Option<String>,
    /// Source code hash (SHA-256, hex).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_hash: Option<String>,
    /// Compressed VBA stream size in bytes.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub compressed_size: Option<u64>,
    /// Decompressed VBA source size in bytes.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub decompressed_size: Option<u64>,
    /// Best-effort extraction state for this module.
    pub extraction_state: MacroExtractionState,
    /// Partial extraction problems collected for this module.
    #[serde(default)]
    pub extraction_errors: Vec<String>,
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
            stream_name: None,
            stream_path: None,
            procedures: Vec::new(),
            suspicious_calls: Vec::new(),
            source_code: None,
            source_encoding: None,
            source_hash: None,
            compressed_size: None,
            decompressed_size: None,
            extraction_state: MacroExtractionState::NotRequested,
            extraction_errors: Vec::new(),
            span: None,
        }
    }
}

/// Extraction state of a VBA module payload.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MacroExtractionState {
    NotRequested,
    Extracted,
    MissingStream,
    DecodeFailed,
    Partial,
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
