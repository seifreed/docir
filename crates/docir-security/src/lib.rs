//! # docir-security
//!
//! Security analysis for docir. Provides utilities for analyzing
//! security-relevant content in Office documents.

/// Core security analyzer implementation.
pub mod analyzer;
/// Helpers to enrich IR stores with derived security indicators.
pub mod enrich;
/// Hashing helpers for artifact fingerprints.
pub mod hash;
/// Atomic indicators used by security and rules layers.
pub mod indicators;
/// VBA and XLM security signature helpers.
pub mod vba;

/// Public analyzer facade used by app and CLI entrypoints.
pub use analyzer::SecurityAnalyzer;
/// Writes computed indicators into the IR store.
pub use enrich::populate_security_indicators;
/// Computes SHA-256 artifact fingerprints.
pub use hash::sha256_hex;
/// Re-exports security indicators.
pub use indicators::*;
/// Re-exports VBA and XLM signature helpers.
pub use vba::{
    analyze_vba_source, contains_dangerous_xlm, is_auto_exec_procedure, is_dangerous_xlm_function,
    scan_vba_source, VbaAnalysis, AUTO_EXEC_PROCEDURES, DANGEROUS_XLM_FUNCTIONS,
    SUSPICIOUS_VBA_CALLS,
};
