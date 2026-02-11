//! # docir-parser
//!
//! OOXML parser for docir. Handles secure parsing of DOCX, XLSX, and PPTX files
//! into the docir IR representation.

pub mod diagnostics;
pub mod error;
pub mod hwp;
pub mod odf;
pub mod ole;
pub mod ooxml;
pub mod parser;
pub mod rtf;
mod security_utils;
pub mod zip_handler;

pub use error::ParseError;
pub use hwp::{HwpParser, HwpxParser};
pub use odf::OdfParser;
pub use parser::{DocumentParser, OoxmlParser, ParserConfig};
pub use rtf::RtfParser;
