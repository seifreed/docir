//! Security-related IR nodes and types.
//!
//! This module defines types for representing security-sensitive
//! elements in Office documents: VBA macros, OLE objects, external
//! references, DDE fields, and more.

use crate::types::NodeId;
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

mod extraction;
mod security_macros;
mod security_ole;
mod security_references;
mod threat_level;
mod vba;

pub use extraction::{
    ExtractedArtifact, ExtractedArtifactKind, ExtractionManifest, ExtractionWarning,
};
pub use security_macros::{
    DigitalSignature, MacroExtractionState, MacroModule, MacroModuleType, MacroProject,
    MacroReference, SuspiciousCall, SuspiciousCallCategory,
};
pub use security_ole::{ActiveXControl, OleObject};
pub use security_references::{
    DdeField, DdeFieldType, ExternalRefType, ExternalReference, XlmFunction, XlmMacro, XlmMacroCell,
};
use threat_level::max_indicator_threat_level;
pub use threat_level::{ThreatIndicator, ThreatIndicatorType, ThreatLevel};
pub use vba::{
    analyze_vba_source, is_dangerous_xlm_function, VbaAnalysis, AUTO_EXEC_PROCEDURES,
    DANGEROUS_XLM_FUNCTIONS, SUSPICIOUS_VBA_CALLS,
};

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
            || !self.threat_indicators.is_empty()
    }

    /// Returns true if the document has a VBA macro project.
    pub fn has_macro_project(&self) -> bool {
        self.macro_project.is_some()
    }

    /// Returns true if any OLE object was detected.
    pub fn has_ole_objects(&self) -> bool {
        !self.ole_objects.is_empty()
    }

    /// Returns true if any external reference was detected.
    pub fn has_external_references(&self) -> bool {
        !self.external_refs.is_empty()
    }

    /// Returns true if any ActiveX control was detected.
    pub fn has_activex_controls(&self) -> bool {
        !self.activex_controls.is_empty()
    }

    /// Returns true if any DDE field was detected.
    pub fn has_dde_fields(&self) -> bool {
        !self.dde_fields.is_empty()
    }

    /// Returns true if any XLM macro was detected.
    pub fn has_xlm_macros(&self) -> bool {
        !self.xlm_macros.is_empty()
    }

    /// Returns true if any macro-like content exists (VBA project or XLM macros).
    pub fn has_macros(&self) -> bool {
        self.has_macro_project() || self.has_xlm_macros()
    }

    /// Returns the macro project node ID, when available.
    pub fn macro_project_id(&self) -> Option<NodeId> {
        self.macro_project
    }

    /// Returns OLE object node IDs.
    pub fn ole_object_ids(&self) -> &[NodeId] {
        &self.ole_objects
    }

    /// Returns ActiveX control node IDs.
    pub fn activex_control_ids(&self) -> &[NodeId] {
        &self.activex_controls
    }

    /// Returns external reference node IDs.
    pub fn external_ref_ids(&self) -> &[NodeId] {
        &self.external_refs
    }

    /// Returns DDE fields.
    pub fn dde_fields(&self) -> &[DdeField] {
        &self.dde_fields
    }

    /// Returns XLM macro definitions.
    pub fn xlm_macros(&self) -> &[XlmMacro] {
        &self.xlm_macros
    }

    /// Returns threat indicator count.
    pub fn threat_indicator_count(&self) -> usize {
        self.threat_indicators.len()
    }

    /// Returns OLE object count.
    pub fn ole_object_count(&self) -> usize {
        self.ole_objects.len()
    }

    /// Returns ActiveX control count.
    pub fn activex_control_count(&self) -> usize {
        self.activex_controls.len()
    }

    /// Returns external reference count.
    pub fn external_ref_count(&self) -> usize {
        self.external_refs.len()
    }

    /// Returns DDE field count.
    pub fn dde_field_count(&self) -> usize {
        self.dde_fields.len()
    }

    /// Returns XLM macro count.
    pub fn xlm_macro_count(&self) -> usize {
        self.xlm_macros.len()
    }

    /// Replaces derived security state and recalculates threat level.
    pub fn apply_scan_result(
        &mut self,
        scan: SecurityInfo,
        threat_indicators: Vec<ThreatIndicator>,
    ) {
        self.macro_project = scan.macro_project;
        self.ole_objects = scan.ole_objects;
        self.external_refs = scan.external_refs;
        self.activex_controls = scan.activex_controls;
        self.dde_fields = scan.dde_fields;
        self.xlm_macros = scan.xlm_macros;
        self.threat_indicators = threat_indicators;
        self.recalculate_threat_level();
    }

    /// Recalculates the threat level based on indicators.
    pub fn recalculate_threat_level(&mut self) {
        let mut level = ThreatLevel::None;

        if self.macro_project.is_some() || !self.xlm_macros.is_empty() {
            level = ThreatLevel::Critical;
        }

        if !self.dde_fields.is_empty()
            || !self.ole_objects.is_empty()
            || !self.activex_controls.is_empty()
        {
            level = level.max(ThreatLevel::High);
        }

        let indicator_level = max_indicator_threat_level(&self.threat_indicators);
        level = level.max(indicator_level);

        if !self.external_refs.is_empty() && level < ThreatLevel::Medium {
            level = ThreatLevel::Medium;
        }

        self.threat_level = level;
    }
}
