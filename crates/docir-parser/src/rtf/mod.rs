//! RTF parsing support.

mod core;
mod objects;
mod parser;

pub(crate) use core::{is_rtf_bytes, parse_rtf, RtfCursor, RtfParseContext};
pub use parser::RtfParser;

#[cfg(test)]
mod tests;
