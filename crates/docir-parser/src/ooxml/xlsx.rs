//! XLSX workbook and worksheet parsing.

use crate::diagnostics::{attach_diagnostics_if_any, push_warning};
use crate::error::ParseError;
use crate::ooxml::part_utils::parse_xml_part_with_span;
use crate::ooxml::relationships::{rel_type, Relationship, Relationships, TargetMode};
use crate::security_utils::parse_dde_formula;
use crate::xml_utils::attr_value;
use crate::zip_handler::PackageReader;
use docir_core::ir::{
    parse_cell_reference, CalcChain, CalcChainEntry, Cell, CellError, CellFormula,
    ColumnDefinition, ConditionalFormat, ConditionalRule, Diagnostics, Document, FormulaType,
    IRNode, MergedCellRange, PivotCache, Shape, ShapeType, SharedStringItem, SharedStringTable,
    SheetComment, SheetKind, SheetState, Worksheet, WorksheetDrawing,
};
use docir_core::security::{ExternalRefType, ExternalReference, SecurityInfo};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};

mod cell;
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
mod xlm;

pub use parser::*;
pub(crate) use shared_strings::parse_shared_strings_table;

use parser::{
    extract_formula_function, map_cell_error, parse_calc_chain, parse_column,
    parse_conditional_formatting, parse_formula, parse_formula_args_text, parse_formula_empty,
    parse_inline_string, parse_merge_cell, parse_sheet_comments, parse_threaded_comments,
};
use workbook::{auto_open_target_from_defined_name, SheetInfo};

pub(crate) use connections::{
    connection_targets, parse_connections_part, parse_external_link_part, parse_query_table_part,
    parse_slicer_part, parse_timeline_part,
};
pub(crate) use metadata::parse_sheet_metadata;
pub(crate) use styles::{parse_color_attr, parse_styles};
pub(crate) use tables::{
    parse_pivot_cache_records, parse_pivot_table_definition, parse_table_definition,
};
