//! XLSX workbook and worksheet parsing.

use crate::diagnostics::push_warning;
use crate::error::ParseError;
use crate::ooxml::relationships::{Relationship, Relationships, rel_type};
use docir_core::ir::{
    Cell, ColumnDefinition, ConditionalFormat, Document, IRNode, MergedCellRange, PivotCache,
    Shape, ShapeType, SheetComment, SheetKind, SheetState, Worksheet, WorksheetDrawing,
    parse_cell_reference,
};
use docir_core::security::{ExternalRefType, ExternalReference};
use docir_core::types::{NodeId, SourceSpan};

mod cell;
mod comments;
mod connections;
mod drawing;
mod metadata;
mod parser;
mod relationships;
mod shared_strings;
mod styles;
mod tables;
mod workbook;
mod workbook_parts;
mod worksheet;
mod worksheet_parts;
mod xlm;

const MAX_XLSX_COLUMN_INDEX: u32 = 16_384;
const MAX_XLSX_ROW_INDEX: u32 = 1_048_576;

fn validate_cell_coordinates(
    reference: &str,
    coordinates: (u32, u32),
) -> Result<(u32, u32), ParseError> {
    let (col, row) = coordinates;
    if col >= MAX_XLSX_COLUMN_INDEX || row >= MAX_XLSX_ROW_INDEX {
        return Err(ParseError::InvalidStructure(format!(
            "XLSX cell reference is outside worksheet limits: {reference}"
        )));
    }
    Ok(coordinates)
}

pub(crate) use parser::*;
pub(crate) use shared_strings::parse_shared_strings_table;

pub(crate) use connections::{
    connection_targets, parse_connections_part, parse_external_link_part, parse_query_table_part,
    parse_slicer_part, parse_timeline_part,
};
pub(crate) use metadata::parse_sheet_metadata;
pub(crate) use parser::{
    extract_formula_function, map_cell_error, parse_calc_chain, parse_column,
    parse_conditional_formatting, parse_formula, parse_formula_args_text, parse_formula_empty,
    parse_inline_string, parse_merge_cell, parse_sheet_comments, parse_threaded_comments,
};
pub(crate) use styles::{parse_color_attr, parse_styles};
pub(crate) use tables::{
    parse_pivot_cache_records, parse_pivot_table_definition, parse_table_definition,
};
