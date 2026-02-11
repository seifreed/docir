//! # docir-security
//!
//! Security analysis for docir. Provides utilities for analyzing
//! security-relevant content in Office documents.

pub mod analyzer;
pub mod enrich;
pub mod indicators;

pub use analyzer::SecurityAnalyzer;
pub use enrich::populate_security_indicators;
pub use indicators::*;
