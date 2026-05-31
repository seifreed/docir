mod format_helpers;
mod presentation;
mod spreadsheet;
mod summary_normalize;
mod summary_parse;
mod summary_postprocess;
mod summary_primary;
mod summary_secondary;
#[cfg(test)]
mod summary_signature_tests;
mod summary_signatures;
#[cfg(test)]
mod summary_tests;

pub(crate) use self::format_helpers::*;
pub(crate) use self::summary_parse::summarize;
#[cfg(test)]
pub(crate) use self::summary_primary::{
    summarize_cell, summarize_formula, summarize_paragraph, summarize_primary, summarize_shape,
};
#[cfg(test)]
pub(crate) use self::summary_secondary::summarize_secondary;
#[cfg(test)]
pub(crate) use self::summary_signatures::{cell_value_summary, text_from_paragraph};
pub(crate) use self::summary_signatures::{content_signature, style_signature};
#[cfg(test)]
pub(crate) use docir_core::ir::IRNode;
