//! ODF (OpenDocument) parsing support.

use crate::diagnostics::{attach_diagnostics_if_any, push_entry};
use crate::error::ParseError;
use crate::format::FormatParser;
use crate::parser::{ParsedDocument, ParserConfig};
use crate::text_utils::parse_text_alignment;
use crate::xml_utils::{attr_value, read_event, scan_xml_events, scan_xml_events_with_reader};
use crate::zip_handler::SecureZipReader;
#[allow(unused_imports)]
use docir_core::ir::{
    BookmarkEnd, BookmarkStart, Cell, CellFormula, CellValue, ChartData, Comment, CommentReference,
    ConditionalFormat, ConditionalRule, DataValidation, DiagnosticSeverity, Diagnostics, Document,
    Endnote, ExtensionPart, ExtensionPartKind, Field, FieldInstruction, FieldKind, Footer,
    Footnote, Header, IRNode, MediaAsset, MediaType, MergedCellRange, NumberingInfo, Paragraph,
    ParagraphProperties, PivotCache, PivotCacheRecords, PivotTable, Revision, RevisionType, Run,
    Section, Shape, ShapeText, ShapeTextParagraph, ShapeTextRun, ShapeType, Slide, SlideAnimation,
    SlideTransition, Style, StyleSet, StyleType, Table, TableCell, TableCellProperties, TableRow,
    Worksheet, WorksheetDrawing,
};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::{Read, Seek};

mod builder;
mod container;
mod formula;
#[allow(clippy::single_match)]
mod helpers;
mod io;
mod limits;
mod manifest;
mod ods;
mod paragraph;
mod presentation;
mod presentation_helpers;
mod sampling;
pub(crate) mod security;
mod security_helpers;
mod spreadsheet;
mod spreadsheet_chunks;
mod styles_support;
mod text;
mod utils;

use self::security::scan_odf_filters;
use self::security_helpers::{build_odf_macro_project, parse_odf_signatures};
use crate::security_scan::DefaultSecurityScanner;

use container::{handle_content_xml, load_meta};
use formula::evaluate_ods_formulas;
use io::{collect_manifest_index, collect_shared_parts};
use limits::{OdfAtomicLimits, OdfLimitCounter, OdfLimits};
use manifest::{
    encrypted_manifest_entries, format_odf_encryption_metadata, is_manifest_entry_encrypted,
    parse_manifest, OdfEncryptionData, OdfManifestEntry,
};
use ods::{parse_ods_table, parse_ods_table_fast};
use paragraph::parse_paragraph;
use presentation_helpers::{build_media_asset, parse_draw_page, parse_odp_transition};
use sampling::parse_ods_row_sample;
use styles_support::{
    merge_styles, parse_master_pages, parse_odf_headers_footers, parse_page_layouts, parse_styles,
};
use utils::{parse_frame_transform, parse_ods_named_ranges, strip_odf_formula_prefix};

type OdfReader<'a> = Reader<std::io::Cursor<&'a [u8]>>;

/// ODF parser (ODT/ODS/ODP).
pub struct OdfParser {
    config: ParserConfig,
}

impl FormatParser for OdfParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
    }
}

#[derive(Debug, Default)]
struct OdfContentResult {
    content: Vec<NodeId>,
    comments: Vec<NodeId>,
    footnotes: Vec<NodeId>,
    endnotes: Vec<NodeId>,
    pivot_caches: Vec<NodeId>,
}

impl Default for OdfParser {
    fn default() -> Self {
        Self::new()
    }
}

impl OdfParser {
    /// Creates a new parser with default configuration.
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    /// Creates a new parser with custom configuration.
    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    crate::impl_parse_entrypoints!();
}

fn parse_content(
    xml: &[u8],
    format: DocumentFormat,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<OdfContentResult, ParseError> {
    match format {
        DocumentFormat::OdfText => text::parse_content_text(xml, store, limits),
        DocumentFormat::OdfSpreadsheet => {
            spreadsheet::parse_content_spreadsheet(xml, store, limits)
        }
        DocumentFormat::OdfPresentation => {
            presentation::parse_content_presentation(xml, store, limits)
        }
        _ => Ok(OdfContentResult::default()),
    }
}

#[cfg(test)]
mod tests;
#[cfg(test)]
mod tests_prelude;
