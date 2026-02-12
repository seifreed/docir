//! Shared XML parsing helpers.

use crate::error::ParseError;

pub(crate) fn xml_error(file: &str, err: impl std::fmt::Display) -> ParseError {
    crate::xml_utils::xml_error(file, err)
}
