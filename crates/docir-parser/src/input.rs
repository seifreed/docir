//! Shared input helpers for parser entrypoints.

use crate::error::ParseError;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek, SeekFrom};
use std::path::Path;

pub fn open_reader<P: AsRef<Path>>(path: P) -> Result<BufReader<File>, ParseError> {
    let file = File::open(path.as_ref())?;
    Ok(BufReader::new(file))
}

pub fn cursor_from_bytes(data: &[u8]) -> Cursor<&[u8]> {
    Cursor::new(data)
}

pub fn enforce_input_size<R: Seek>(reader: &mut R, max_input_size: u64) -> Result<(), ParseError> {
    let current = reader.stream_position()?;
    let end = reader.seek(SeekFrom::End(0))?;
    reader.seek(SeekFrom::Start(current))?;
    if end > max_input_size {
        return Err(ParseError::ResourceLimit(format!(
            "Input too large: {} bytes (max: {} bytes)",
            end, max_input_size
        )));
    }
    Ok(())
}

pub fn read_all_with_limit<R: Read + Seek>(
    mut reader: R,
    max_input_size: u64,
) -> Result<Vec<u8>, ParseError> {
    enforce_input_size(&mut reader, max_input_size)?;
    let mut data = Vec::new();
    reader.read_to_end(&mut data)?;
    Ok(data)
}
