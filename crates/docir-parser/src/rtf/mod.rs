//! RTF parsing support.

mod core;
mod objects;
mod parser;

pub(crate) use core::{RtfCursor, RtfParseContext, is_rtf_bytes, parse_rtf};
pub use parser::RtfParser;

#[cfg(test)]
mod tests;
