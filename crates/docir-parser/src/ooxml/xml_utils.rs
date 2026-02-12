//! Shared XML parsing helpers.

use crate::error::ParseError;

pub(crate) fn xml_error(file: &str, err: impl std::fmt::Display) -> ParseError {
    ParseError::Xml {
        file: file.to_string(),
        message: err.to_string(),
    }
}
