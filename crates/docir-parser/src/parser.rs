//! Main OOXML parser orchestrator.

pub use crate::config::{ParseMetrics, ParserConfig};
use crate::diagnostics::{push_info, push_warning};
use crate::error::ParseError;
use crate::format::FormatParser;
use crate::hwp::is_hwpx_mimetype;
use crate::input::{enforce_input_size, read_all_with_limit};
use crate::ole::is_ole_container;
use crate::ooxml::content_types::ContentTypes;
use crate::ooxml::docx::DocxParser;
use crate::ooxml::pptx::PptxParser;
use crate::ooxml::relationships::{rel_type, Relationships};
use crate::ooxml::xlsx::XlsxParser;
use crate::rtf::is_rtf_bytes;
use crate::zip_handler::{PackageReader, SecureZipReader};
use docir_core::ir::column_to_letter;
use docir_core::ir::{
    Cell, CellValue, Diagnostics, IRNode, MediaAsset, SheetKind, SheetState, Worksheet,
};
use docir_core::normalize::normalize_store;
use docir_core::security::SecurityInfo;
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use std::collections::HashMap;
use std::io::{Cursor, Read, Seek, SeekFrom};

mod analysis;
mod contracts;
mod coverage;
mod formats;
mod metadata;
mod parser_docx;
mod parser_pptx;
mod parser_xlsx;
mod post_process;
mod security;
mod shared_parts;
mod utils;
mod vba;

mod document;
mod ooxml;
mod types;

use analysis::{hex, map_calamine_error, parse_activex_xml, parse_chart_data, parse_smartart_part};
#[allow(unused_imports)]
pub(crate) use contracts::{
    run_parser_pipeline, NormalizeStage, ParseStage, ParserPipeline, PostprocessStage,
};
pub use document::DocumentParser;
pub use ooxml::OoxmlParser;
pub(crate) use security::SecurityScanner as OoxmlSecurityScanner;
pub use types::ParsedDocument;

#[cfg(test)]
mod tests;
