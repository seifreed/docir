//! Shared input helpers for parser entrypoints.

use crate::error::ParseError;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek, SeekFrom};
use std::path::Path;

/// Public API entrypoint: parse_from_file.
pub fn parse_from_file<P, T, F>(path: P, parse: F) -> Result<T, ParseError>
where
    P: AsRef<Path>,
    F: FnOnce(BufReader<File>) -> Result<T, ParseError>,
{
    let reader = open_reader(path)?;
    parse(reader)
}

/// Public API entrypoint: parse_from_bytes.
pub fn parse_from_bytes<T, F>(data: &[u8], parse: F) -> Result<T, ParseError>
where
    F: FnOnce(Cursor<&[u8]>) -> Result<T, ParseError>,
{
    let reader = cursor_from_bytes(data);
    parse(reader)
}

/// Public API entrypoint: open_reader.
pub fn open_reader<P: AsRef<Path>>(path: P) -> Result<BufReader<File>, ParseError> {
    let file = File::open(path.as_ref())?;
    Ok(BufReader::new(file))
}

/// Public API entrypoint: cursor_from_bytes.
pub fn cursor_from_bytes(data: &[u8]) -> Cursor<&[u8]> {
    Cursor::new(data)
}

/// Public API entrypoint: enforce_input_size.
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

/// Public API entrypoint: read_all_with_limit.
///
/// Note: There is an inherent TOCTOU window between the size check and the
/// read when the reader is backed by a file. For in-memory readers this is
/// not a concern. The risk is mitigated in practice by the `max_input_size`
/// limit which caps the total allocation regardless of file growth.
pub fn read_all_with_limit<R: Read + Seek>(
    mut reader: R,
    max_input_size: u64,
) -> Result<Vec<u8>, ParseError> {
    enforce_input_size(&mut reader, max_input_size)?;
    let read_limit = max_input_size.saturating_add(1);
    let initial_capacity = read_limit.min(64 * 1024).min(usize::MAX as u64) as usize;
    let mut data = Vec::with_capacity(initial_capacity);
    reader.take(read_limit).read_to_end(&mut data)?;
    if data.len() as u64 > max_input_size {
        return Err(ParseError::ResourceLimit(format!(
            "Input too large after read: {} bytes (max: {} bytes)",
            data.len(),
            max_input_size
        )));
    }
    Ok(data)
}

#[macro_export]
macro_rules! impl_parse_entrypoints {
    () => {
        /// Parses a file from the filesystem.
        pub fn parse_file<P: AsRef<std::path::Path>>(
            &self,
            path: P,
        ) -> Result<$crate::parser::ParsedDocument, $crate::error::ParseError> {
            $crate::input::parse_from_file(path, |reader| self.parse_reader(reader))
        }

        /// Parses from a byte slice.
        pub fn parse_bytes(
            &self,
            data: &[u8],
        ) -> Result<$crate::parser::ParsedDocument, $crate::error::ParseError> {
            $crate::input::parse_from_bytes(data, |reader| self.parse_reader(reader))
        }
    };
}

#[cfg(test)]
mod tests {
    use super::read_all_with_limit;
    use crate::error::ParseError;
    use std::cell::Cell;
    use std::io::{Read, Seek, SeekFrom};
    use std::rc::Rc;

    struct GrowingReader {
        reported_len: u64,
        bytes: Vec<u8>,
        pos: usize,
        bytes_read: Rc<Cell<usize>>,
    }

    impl Read for GrowingReader {
        fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
            let remaining = &self.bytes[self.pos..];
            let len = remaining.len().min(buf.len());
            buf[..len].copy_from_slice(&remaining[..len]);
            self.pos += len;
            self.bytes_read.set(self.bytes_read.get() + len);
            Ok(len)
        }
    }

    impl Seek for GrowingReader {
        fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
            match pos {
                SeekFrom::Start(offset) => {
                    self.pos = offset as usize;
                    Ok(offset)
                }
                SeekFrom::End(_) => Ok(self.reported_len),
                SeekFrom::Current(offset) => {
                    self.pos = (self.pos as i64 + offset).max(0) as usize;
                    Ok(self.pos as u64)
                }
            }
        }
    }

    #[test]
    fn read_all_with_limit_stops_after_limit_even_if_reader_grows() {
        let bytes_read = Rc::new(Cell::new(0));
        let reader = GrowingReader {
            reported_len: 4,
            bytes: vec![0; 8],
            pos: 0,
            bytes_read: Rc::clone(&bytes_read),
        };

        let err = read_all_with_limit(reader, 4).unwrap_err();

        assert!(matches!(err, ParseError::ResourceLimit(_)));
        assert_eq!(bytes_read.get(), 5);
    }
}
