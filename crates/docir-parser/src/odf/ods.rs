use crate::odf::{CellValue, NodeId};
use std::collections::HashMap;

mod ods_normalize;
mod ods_parse;
mod ods_postprocess;
#[cfg(test)]
mod ods_tests;

pub(crate) use ods_normalize::{
    infer_cell_value_type_and_attr, parse_cell_formula, parse_cell_value_empty,
    parse_cell_value_with_text, read_ods_cell_text, resolve_style_id, row_repeat_from,
};
pub(crate) use ods_parse::{
    parse_ods_cell, parse_ods_cell_empty, parse_ods_table, parse_ods_table_fast,
};
pub(crate) use ods_postprocess::{
    attach_shapes_as_drawing, emit_sampled_row_cells, flush_validation_ranges,
    push_conditional_format,
};

struct RowBuildState<'a> {
    validation_ranges: &'a mut HashMap<String, Vec<String>>,
    cell_values: &'a mut HashMap<(u32, u32), CellValue>,
    formula_cells: &'a mut Vec<(NodeId, u32, u32, String)>,
    formula_map: &'a mut HashMap<(u32, u32), String>,
}

#[cfg(test)]
use crate::error::ParseError;
#[cfg(test)]
use crate::odf::IRNode;
#[cfg(test)]
use crate::odf::OdfLimits;
#[cfg(test)]
use crate::parser::ParserConfig;
#[cfg(test)]
use docir_core::ir::Cell;
#[cfg(test)]
use docir_core::visitor::IrStore;
#[cfg(test)]
use quick_xml::events::BytesStart;
#[cfg(test)]
use quick_xml::events::Event;
#[cfg(test)]
use quick_xml::Reader;
