mod core_normalize;
mod core_parse;
mod core_postprocess;

#[cfg(test)]
mod core_tests;

mod controls;
mod cursor;
mod field_utils;
mod helpers;
mod state;

pub(crate) use core_parse::{
    append_text, apply_border, apply_paragraph_border, color_from_index, ensure_paragraph,
    ensure_section, finalize_cell, finalize_paragraph, finalize_row, flush_text, parse_rtf,
    pending_numbering, RtfParseContext,
};
pub(crate) use cursor::{is_rtf_bytes, RtfCursor};
