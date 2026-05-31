//! HWP/HWPX parsing (Hangul Word Processor).

use crate::diagnostics::attach_diagnostics_if_any;
use crate::error::ParseError;
use crate::format::FormatParser;
use crate::parser::{ParsedDocument, ParserConfig};
use crate::xml_utils::{attr_value_by_suffix, local_name};
use docir_core::ir::{IRNode, Shape, ShapeType, Table, TableCell, TableRow};
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::BytesStart;

mod builder;
mod helpers;
mod io;
mod legacy;
mod security;
mod styles;
#[cfg(test)]
mod tests;
use std::collections::HashMap;
use std::io::{Read, Seek};

pub mod part_registry;
mod section;

use helpers::{
    attr_any, parse_hwpx_paragraph_props, parse_hwpx_table_props, run_properties_from_attrs,
    style_run_props_from_run,
};
use legacy::{maybe_decompress_stream, parse_file_header, scan_hwp_external_refs};
use security::scan_hwpx_security;

/// Returns true if the mimetype indicates HWPX.
pub fn is_hwpx_mimetype(value: &str) -> bool {
    let lower = value.trim().to_ascii_lowercase();
    lower.contains("hwp+zip") || lower.contains("hwpx")
}

/// Parser for legacy HWP (OLE/CFB).
pub struct HwpParser {
    pub(crate) config: ParserConfig,
}

impl FormatParser for HwpParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
    }
}

impl Default for HwpParser {
    fn default() -> Self {
        Self::new()
    }
}

impl HwpParser {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    /// Public API entrypoint: with_config.
    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    crate::impl_parse_entrypoints!();
}

/// Parser for HWPX (ZIP + XML).
pub struct HwpxParser {
    config: ParserConfig,
}

impl FormatParser for HwpxParser {
    fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        self.parse_reader(reader)
    }
}

impl Default for HwpxParser {
    fn default() -> Self {
        Self::new()
    }
}

impl HwpxParser {
    /// Public API entrypoint: new.
    pub fn new() -> Self {
        Self {
            config: ParserConfig::default(),
        }
    }

    /// Public API entrypoint: with_config.
    pub fn with_config(config: ParserConfig) -> Self {
        Self { config }
    }

    crate::impl_parse_entrypoints!();
}

fn parse_hwpx_shape(
    e: &BytesStart<'_>,
    local: &[u8],
    source: &str,
    media_lookup: &HashMap<String, NodeId>,
    store: &mut IrStore,
) -> Option<NodeId> {
    let shape_type = match local {
        b"pic" | b"image" | b"img" => ShapeType::Picture,
        b"chart" => ShapeType::Chart,
        b"audio" => ShapeType::Audio,
        b"video" => ShapeType::Video,
        b"ole" | b"object" => ShapeType::OleObject,
        b"shape" | b"draw" => ShapeType::Custom,
        _ => ShapeType::Unknown,
    };
    if matches!(shape_type, ShapeType::Unknown) {
        return None;
    }

    let mut shape = Shape::new(shape_type);
    shape.name = attr_any(e, &[b"name", b"id", b"shapeId", b"shape-id"]);
    shape.alt_text = attr_any(e, &[b"alt", b"altText", b"alt-text"]);
    shape.hyperlink = attr_any(e, &[b"href", b"link", b"xlink:href"]);
    shape.media_target = attr_value_by_suffix(
        e,
        &[
            b"href",
            b"src",
            b"link",
            b"binaryRef",
            b"binData",
            b"binDataId",
        ],
    );
    if let Some(target) = shape.media_target.as_deref()
        && let Some(id) = media_lookup.get(target)
    {
        shape.media_asset = Some(*id);
    }
    if let Some(x) = attr_any(e, &[b"x", b"posX", b"left"]).and_then(|v| v.parse::<i64>().ok()) {
        shape.transform.x = x;
    }
    if let Some(y) = attr_any(e, &[b"y", b"posY", b"top"]).and_then(|v| v.parse::<i64>().ok()) {
        shape.transform.y = y;
    }
    if let Some(width) = attr_any(e, &[b"width", b"w"]).and_then(|v| v.parse::<u64>().ok()) {
        shape.transform.width = width;
    }
    if let Some(height) = attr_any(e, &[b"height", b"h"]).and_then(|v| v.parse::<u64>().ok()) {
        shape.transform.height = height;
    }
    shape.span = Some(SourceSpan::new(source));

    let shape_id = shape.id;
    store.insert(IRNode::Shape(shape));
    Some(shape_id)
}

fn finalize_cell_hwpx(
    current_cell: &mut Option<TableCell>,
    current_row: &mut Option<TableRow>,
    store: &mut IrStore,
) {
    if let Some(cell) = current_cell.take() {
        let cell_id = cell.id;
        store.insert(IRNode::TableCell(cell));
        if let Some(row) = current_row.as_mut() {
            row.cells.push(cell_id);
        }
    }
}

fn finalize_row_hwpx(
    current_row: &mut Option<TableRow>,
    current_table: &mut Option<Table>,
    store: &mut IrStore,
) {
    if let Some(row) = current_row.take() {
        let row_id = row.id;
        store.insert(IRNode::TableRow(row));
        if let Some(table) = current_table.as_mut() {
            table.rows.push(row_id);
        }
    }
}

fn finalize_table_hwpx(
    current_table: &mut Option<Table>,
    current_cell: &mut Option<TableCell>,
    content: &mut Vec<NodeId>,
    store: &mut IrStore,
) {
    if current_cell.is_some() {
        finalize_cell_hwpx(current_cell, &mut None, store);
    }
    if let Some(table) = current_table.take() {
        let table_id = table.id;
        store.insert(IRNode::Table(table));
        content.push(table_id);
    }
}
