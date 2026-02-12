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
use crate::ooxml::part_utils::{read_relationships_optional, read_xml_part, read_xml_part_by_rel};
use crate::ooxml::pptx::PptxParser;
use crate::ooxml::relationships::{rel_type, Relationships};
use crate::ooxml::xlsx::XlsxParser;
use crate::rtf::is_rtf_bytes;
use crate::xml_utils::local_name;
use crate::zip_handler::{PackageReader, SecureZipReader};
use docir_core::ir::column_to_letter;
use docir_core::ir::{
    Cell, CellError, CellValue, Diagnostics, Document, ExtensionPart, ExtensionPartKind, IRNode,
    MediaAsset, SheetKind, SheetState, Worksheet,
};
use docir_core::normalize::normalize_store;
use docir_core::security::SecurityInfo;
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use std::collections::HashMap;
use std::io::{Cursor, Read, Seek, SeekFrom};
use std::path::Path;

mod coverage;
mod dispatch;
mod metadata;
mod parser_docx;
mod parser_pptx;
mod parser_xlsx;
mod post_process;
pub(crate) mod security;
mod shared_parts;
mod utils;
mod vba;

mod document;
mod ooxml;
mod types;

pub use document::DocumentParser;
use document::{hex, map_calamine_error, parse_activex_xml, parse_chart_data, parse_smartart_part};
pub use ooxml::OoxmlParser;
pub use types::ParsedDocument;

#[cfg(test)]
mod tests;
