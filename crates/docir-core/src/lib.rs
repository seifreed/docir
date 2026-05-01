//! # docir-core
//!
//! Core IR (Intermediate Representation) definitions for Microsoft Office documents.
//!
//! This crate defines the semantic representation of OOXML documents (DOCX, XLSX, PPTX),
//! providing a unified, typed, and navigable structure for analysis and transformation.

mod defaults;
pub mod equivalence;
pub mod error;
pub mod ir;
pub mod normalize;
pub mod query;
pub mod security;
pub mod types;
pub mod visitor;

pub use equivalence::IrSummary;
pub use error::CoreError;
pub use ir::*;
pub use query::Query;
pub use security::{
    analyze_vba_source, is_dangerous_xlm_function, ActiveXControl, DdeField,
    DigitalSignature as SecurityDigitalSignature, ExternalReference, ExtractedArtifact,
    ExtractedArtifactKind, ExtractionManifest, ExtractionWarning, MacroExtractionState,
    MacroModule, MacroProject, MacroReference, OleObject, SecurityInfo, SuspiciousCall,
    SuspiciousCallCategory, ThreatIndicator, ThreatIndicatorType, ThreatLevel, VbaAnalysis,
    XlmMacro, AUTO_EXEC_PROCEDURES, DANGEROUS_XLM_FUNCTIONS, SUSPICIOUS_VBA_CALLS,
};
pub use types::*;
