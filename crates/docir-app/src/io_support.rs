use crate::AppResult;
use docir_parser::ParseError as ParserParseError;
use std::fs;
use std::path::Path;

pub(crate) fn read_bounded_file<P: AsRef<Path>>(
    path: P,
    max_input_size: u64,
) -> AppResult<Vec<u8>> {
    let path = path.as_ref();
    let metadata = fs::metadata(path).map_err(ParserParseError::from)?;
    if metadata.len() > max_input_size {
        return Err(ParserParseError::ResourceLimit(format!(
            "Input exceeds max_input_size ({} > {})",
            metadata.len(),
            max_input_size
        ))
        .into());
    }
    fs::read(path)
        .map_err(ParserParseError::from)
        .map_err(Into::into)
}
