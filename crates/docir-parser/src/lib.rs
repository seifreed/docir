//! # docir-parser
//!
//! OOXML parser for docir. Handles secure parsing of DOCX, XLSX, and PPTX files
//! into the docir IR representation.

pub mod adapters;
pub mod config;
pub mod diagnostics;
pub mod error;
pub mod format;
pub mod hwp;
mod input;
pub mod odf;
pub mod ole;
pub mod ooxml;
mod parse_utils;
pub mod parser;
mod registry_utils;
pub mod rtf;
mod security_scan;
mod security_utils;
mod text_utils;
pub(crate) mod xml_utils;
pub mod zip_handler;

pub use config::ParserConfig;
pub use error::ParseError;
pub use hwp::{HwpParser, HwpxParser};
pub use odf::OdfParser;
pub use parser::{DocumentParser, OoxmlParser};
pub use rtf::RtfParser;
pub use security_scan::{DefaultSecurityScanner, SecurityScanner};

use crate::zip_handler::SecureZipReader;
use docir_core::visitor::IrStore;
use std::io::Cursor;

/// Scans security-relevant artifacts from raw bytes into an existing store.
pub fn scan_security_bytes(
    config: &ParserConfig,
    data: &[u8],
    store: &mut IrStore,
) -> Result<(), ParseError> {
    if !is_zip_container(data) {
        return Ok(());
    }
    let mut zip = SecureZipReader::new(Cursor::new(data), config.zip_config.clone())?;
    let scanner = security_scan::DefaultSecurityScanner;
    scanner.scan_ooxml(config, &mut zip, store)
}

fn is_zip_container(data: &[u8]) -> bool {
    data.len() >= 4 && data[0] == b'P' && data[1] == b'K'
}
