//! # docir-security
//!
//! Security analysis for docir. Provides utilities for analyzing
//! security-relevant content in Office documents.

/// Core security analyzer implementation.
pub mod analyzer;
/// Helpers to enrich IR stores with derived security indicators.
pub mod enrich;
/// Atomic indicators used by security and rules layers.
pub mod indicators;

/// Public analyzer facade used by app and CLI entrypoints.
pub use analyzer::SecurityAnalyzer;
/// Writes computed indicators into the IR store.
pub use enrich::populate_security_indicators;
/// Re-exports security indicators.
pub use indicators::*;
