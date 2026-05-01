//! # docir-parser
//!
//! OOXML parser for docir. Handles secure parsing of DOCX, XLSX, and PPTX files
//! into the docir IR representation.

/// Adapters for parser output sinks and integrations.
pub mod adapters;
/// Configuration types used by parser entrypoints.
pub mod config;
/// Diagnostics helpers for parser reporting.
pub mod diagnostics;
/// Error model shared by parser pipelines.
pub mod error;
/// Document format detection and metadata.
pub mod format;
/// HWP/HWPX parsing module.
pub mod hwp;
mod input;
/// Legacy Office CFB parser module.
pub mod legacy_office;
/// Open Document Format parser module.
pub mod odf;
/// OLE container integration helpers.
pub mod ole;
mod ole_header;
/// OOXML parsing module.
pub mod ooxml;
mod parse_utils;
/// Top-level parser abstractions and dispatch.
pub mod parser;
/// Low-level PowerPoint binary record reader.
pub mod ppt_records;
mod registry_utils;
/// Rich Text Format parser implementation.
pub mod rtf;
mod security_scan;
mod security_utils;
#[cfg(test)]
mod test_support;
mod text_utils;
/// Low-level BIFF/XLS record reader.
pub mod xls_records;
pub(crate) mod xml_utils;
/// ZIP wrapper with security checks.
pub mod zip_handler;

/// Parser configuration values.
pub use config::ParserConfig;
/// Canonical parser error type.
pub use error::ParseError;
/// HWP document parser entrypoints.
pub use hwp::{HwpParser, HwpxParser};
pub use legacy_office::LegacyOfficeParser;
/// ODF parser entrypoint.
pub use odf::OdfParser;
/// Parser facade and OOXML implementation.
pub use parser::{DocumentParser, OoxmlParser};
/// RTF parser entrypoint.
pub use rtf::RtfParser;
/// Default security scanner and scanning trait.
pub use security_scan::{DefaultSecurityScanner, SecurityScanner};

use crate::zip_handler::SecureZipReader;
use docir_core::visitor::IrStore;
use std::io::Cursor;

/// Scans security-relevant artifacts from raw bytes into an existing store.
///
/// The function only acts when the input appears to be a ZIP container and then
/// delegates scanning to the configured default security scanner.
pub fn scan_security_bytes(
    config: &ParserConfig,
    data: &[u8],
    store: &mut IrStore,
) -> Result<(), ParseError> {
    if crate::ole::is_ole_container(data) {
        let cfb = crate::ole::Cfb::parse(data.to_vec())?;
        let scanner = parser::OoxmlSecurityScanner::new(config);
        return scanner.scan_cfb(&cfb, store);
    }
    if !parse_utils::is_zip_container(data) {
        return Ok(());
    }
    let mut zip = SecureZipReader::new(Cursor::new(data), config.zip_config.clone())?;
    let scanner = security_scan::DefaultSecurityScanner;
    scanner.scan_ooxml(config, &mut zip, store)
}
