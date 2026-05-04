//! Secure ZIP handling for OOXML packages.
//!
//! This module provides secure reading of ZIP archives with protections
//! against zip bombs, path traversal, and other malicious archive attacks.

use crate::error::ParseError;
use std::collections::HashMap;
use std::io::{Read, Seek};
use zip::ZipArchive;

/// Configuration for ZIP extraction limits.
#[derive(Debug, Clone)]
pub struct ZipConfig {
    /// Maximum total uncompressed size (default: 100MB).
    pub max_total_size: u64,
    /// Maximum size per file (default: 50MB).
    pub max_file_size: u64,
    /// Maximum number of files (default: 10000).
    pub max_file_count: usize,
    /// Maximum compression ratio (to detect zip bombs).
    pub max_compression_ratio: f64,
    /// Maximum path depth.
    pub max_path_depth: usize,
}

impl Default for ZipConfig {
    fn default() -> Self {
        Self {
            max_total_size: 100 * 1024 * 1024, // 100MB
            max_file_size: 50 * 1024 * 1024,   // 50MB
            max_file_count: 10000,
            max_compression_ratio: 100.0, // 100:1
            max_path_depth: 20,
        }
    }
}

pub trait PackageReader {
    fn contains(&self, name: &str) -> bool;
    fn read_file(&mut self, name: &str) -> Result<Vec<u8>, ParseError>;
    fn read_file_string(&mut self, name: &str) -> Result<String, ParseError> {
        let bytes = self.read_file(name)?;
        String::from_utf8(bytes)
            .map_err(|e| ParseError::Encoding(format!("Invalid UTF-8 in {name}: {e}")))
    }
    fn file_size(&mut self, name: &str) -> Result<u64, ParseError>;
    fn file_names(&self) -> Vec<String>;
    fn list_prefix(&self, prefix: &str) -> Vec<String>;
    fn list_suffix(&self, suffix: &str) -> Vec<String>;
}

/// A secure wrapper around a ZIP archive.
pub struct SecureZipReader<R: Read + Seek> {
    archive: ZipArchive<R>,
    config: ZipConfig,
    file_index: HashMap<String, usize>,
}

impl<R: Read + Seek> SecureZipReader<R> {
    /// Opens a ZIP archive with security checks.
    pub fn new(reader: R, config: ZipConfig) -> Result<Self, ParseError> {
        let mut archive = ZipArchive::new(reader)?;

        // Check file count
        if archive.len() > config.max_file_count {
            return Err(ParseError::ResourceLimit(format!(
                "Too many files in archive: {} (max: {})",
                archive.len(),
                config.max_file_count
            )));
        }

        // Build index and validate paths
        let mut file_index = HashMap::new();
        let mut total_uncompressed = 0u64;

        for i in 0..archive.len() {
            let file = archive.by_index_raw(i)?;
            let name = file.name().to_string();

            // Check for path traversal
            if Self::is_path_traversal(&name) {
                return Err(ParseError::PathTraversal(name));
            }

            // Check path depth
            let depth = name.matches('/').count();
            if depth > config.max_path_depth {
                return Err(ParseError::ResourceLimit(format!(
                    "Path too deep: {} (max depth: {})",
                    name, config.max_path_depth
                )));
            }

            // Check individual file size
            let uncompressed_size = file.size();
            if uncompressed_size > config.max_file_size {
                return Err(ParseError::ResourceLimit(format!(
                    "File too large: {} ({} bytes, max: {} bytes)",
                    name, uncompressed_size, config.max_file_size
                )));
            }

            // Check compression ratio (zip bomb detection)
            let compressed_size = file.compressed_size();
            if compressed_size > 0 {
                let ratio = uncompressed_size as f64 / compressed_size as f64;
                if ratio > config.max_compression_ratio {
                    return Err(ParseError::ResourceLimit(format!(
                        "Suspicious compression ratio for {}: {:.1}:1 (max: {:.1}:1)",
                        name, ratio, config.max_compression_ratio
                    )));
                }
            }

            total_uncompressed += uncompressed_size;
            file_index.insert(name, i);
        }

        // Check total size
        if total_uncompressed > config.max_total_size {
            return Err(ParseError::ResourceLimit(format!(
                "Total uncompressed size too large: {} bytes (max: {} bytes)",
                total_uncompressed, config.max_total_size
            )));
        }

        Ok(Self {
            archive,
            config,
            file_index,
        })
    }

    /// Checks if a path attempts directory traversal.
    fn is_path_traversal(path: &str) -> bool {
        // Check for parent directory references (including URL-encoded variants)
        if path.contains("..") {
            return true;
        }

        // Check for URL-encoded path traversal attempts
        let lower = path.to_ascii_lowercase();
        if lower.contains("%2e") || lower.contains("%2f") || lower.contains("%5c") {
            return true;
        }

        // Check for null byte injection
        if path.contains('\0') {
            return true;
        }

        // Check for absolute paths
        if path.starts_with('/') || path.starts_with('\\') {
            return true;
        }

        // Check for Windows-style absolute paths (C:\, D:\, etc.)
        let bytes = path.as_bytes();
        if bytes.len() >= 3 {
            let first = bytes[0];
            let second = bytes[1];
            if first.is_ascii_alphabetic() && second == b':' {
                let third = bytes[2];
                if third == b'\\' || third == b'/' {
                    return true;
                }
            }
        }

        // Check for backslash (shouldn't appear in OOXML)
        if path.contains('\\') {
            return true;
        }

        false
    }

    /// Reads a file from the archive by name.
    pub fn read_file(&mut self, name: &str) -> Result<Vec<u8>, ParseError> {
        let index = self
            .file_index
            .get(name)
            .ok_or_else(|| ParseError::MissingPart(name.to_string()))?;

        let mut file = self.archive.by_index(*index)?;

        // Double-check size before reading
        if file.size() > self.config.max_file_size {
            return Err(ParseError::ResourceLimit(format!(
                "File too large: {} ({} bytes)",
                name,
                file.size()
            )));
        }

        let mut contents = Vec::with_capacity(file.size() as usize);
        file.read_to_end(&mut contents)?;

        Ok(contents)
    }

    /// Reads a file as a UTF-8 string.
    pub fn read_file_string(&mut self, name: &str) -> Result<String, ParseError> {
        let bytes = self.read_file(name)?;
        String::from_utf8(bytes)
            .map_err(|e| ParseError::Encoding(format!("Invalid UTF-8 in {}: {}", name, e)))
    }

    /// Returns the uncompressed size for a file.
    pub fn file_size(&mut self, name: &str) -> Result<u64, ParseError> {
        let index = self
            .file_index
            .get(name)
            .ok_or_else(|| ParseError::MissingPart(name.to_string()))?;
        let file = self.archive.by_index(*index)?;
        Ok(file.size())
    }

    /// Checks if a file exists in the archive.
    pub fn contains(&self, name: &str) -> bool {
        self.file_index.contains_key(name)
    }

    /// Returns all file names in the archive.
    pub fn file_names(&self) -> impl Iterator<Item = &str> {
        self.file_index.keys().map(|s| s.as_str())
    }

    /// Returns the number of files in the archive.
    pub fn len(&self) -> usize {
        self.file_index.len()
    }

    /// Returns true if the archive is empty.
    pub fn is_empty(&self) -> bool {
        self.file_index.is_empty()
    }

    /// Lists files matching a prefix.
    pub fn list_prefix(&self, prefix: &str) -> Vec<&str> {
        self.file_index
            .keys()
            .filter(|name| name.starts_with(prefix))
            .map(|s| s.as_str())
            .collect()
    }

    /// Lists files matching a suffix.
    pub fn list_suffix(&self, suffix: &str) -> Vec<&str> {
        self.file_index
            .keys()
            .filter(|name| name.ends_with(suffix))
            .map(|s| s.as_str())
            .collect()
    }
}

impl<R: Read + Seek> PackageReader for SecureZipReader<R> {
    fn contains(&self, name: &str) -> bool {
        SecureZipReader::contains(self, name)
    }

    fn read_file(&mut self, name: &str) -> Result<Vec<u8>, ParseError> {
        SecureZipReader::read_file(self, name)
    }

    fn file_size(&mut self, name: &str) -> Result<u64, ParseError> {
        SecureZipReader::file_size(self, name)
    }

    fn file_names(&self) -> Vec<String> {
        SecureZipReader::file_names(self)
            .map(|name| name.to_string())
            .collect()
    }

    fn list_prefix(&self, prefix: &str) -> Vec<String> {
        SecureZipReader::list_prefix(self, prefix)
            .into_iter()
            .map(|name| name.to_string())
            .collect()
    }

    fn list_suffix(&self, suffix: &str) -> Vec<String> {
        SecureZipReader::list_suffix(self, suffix)
            .into_iter()
            .map(|name| name.to_string())
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;
    use zip::{CompressionMethod, ZipWriter};

    fn make_zip(entries: &[(&str, &[u8])], method: CompressionMethod) -> Vec<u8> {
        let cursor = Cursor::new(Vec::new());
        let mut writer = ZipWriter::new(cursor);
        let options = SimpleFileOptions::default().compression_method(method);

        for (name, contents) in entries {
            writer.start_file(name, options).expect("start file");
            writer.write_all(contents).expect("write file");
        }

        writer.finish().expect("finish zip").into_inner()
    }

    #[test]
    fn test_path_traversal_detection() {
        assert!(SecureZipReader::<Cursor<Vec<u8>>>::is_path_traversal(
            "../etc/passwd"
        ));
        assert!(SecureZipReader::<Cursor<Vec<u8>>>::is_path_traversal(
            "foo/../bar"
        ));
        assert!(SecureZipReader::<Cursor<Vec<u8>>>::is_path_traversal(
            "/absolute/path"
        ));
        assert!(SecureZipReader::<Cursor<Vec<u8>>>::is_path_traversal(
            "C:\\Windows"
        ));
        assert!(SecureZipReader::<Cursor<Vec<u8>>>::is_path_traversal(
            "foo\\bar"
        ));

        assert!(!SecureZipReader::<Cursor<Vec<u8>>>::is_path_traversal(
            "word/document.xml"
        ));
        assert!(!SecureZipReader::<Cursor<Vec<u8>>>::is_path_traversal(
            "[Content_Types].xml"
        ));
        assert!(!SecureZipReader::<Cursor<Vec<u8>>>::is_path_traversal(
            "_rels/.rels"
        ));
    }

    #[test]
    fn secure_zip_reader_reads_and_lists_files() {
        let bytes = make_zip(
            &[
                ("word/document.xml", b"<doc/>"),
                ("word/_rels/document.xml.rels", b"<rels/>"),
            ],
            CompressionMethod::Stored,
        );

        let mut reader =
            SecureZipReader::new(Cursor::new(bytes), ZipConfig::default()).expect("reader");

        assert!(!reader.is_empty());
        assert_eq!(reader.len(), 2);
        assert!(reader.contains("word/document.xml"));
        assert_eq!(
            reader.read_file("word/document.xml").expect("bytes"),
            b"<doc/>".to_vec()
        );
        assert_eq!(
            reader
                .read_file_string("word/_rels/document.xml.rels")
                .expect("string"),
            "<rels/>"
        );
        assert_eq!(reader.file_size("word/document.xml").expect("size"), 6);

        let mut names = reader
            .file_names()
            .map(ToString::to_string)
            .collect::<Vec<_>>();
        names.sort();
        assert_eq!(
            names,
            vec![
                "word/_rels/document.xml.rels".to_string(),
                "word/document.xml".to_string()
            ]
        );

        let mut prefix = reader.list_prefix("word/").to_vec();
        prefix.sort();
        assert_eq!(prefix.len(), 2);
        assert_eq!(
            reader.list_suffix(".rels"),
            vec!["word/_rels/document.xml.rels"]
        );
    }

    #[test]
    fn secure_zip_reader_reports_missing_and_encoding_errors() {
        let bytes = make_zip(
            &[
                ("word/document.xml", b"<doc/>"),
                ("word/binary.bin", &[0xff, 0xfe]),
            ],
            CompressionMethod::Stored,
        );
        let mut reader =
            SecureZipReader::new(Cursor::new(bytes), ZipConfig::default()).expect("reader");

        let err = reader.read_file("word/missing.xml").unwrap_err();
        assert!(matches!(err, ParseError::MissingPart(_)));

        let err = reader.read_file_string("word/binary.bin").unwrap_err();
        assert!(matches!(err, ParseError::Encoding(_)));
    }

    #[test]
    fn secure_zip_reader_rejects_archive_level_limits() {
        let bytes = make_zip(
            &[
                ("a.xml", b"123"),
                ("b.xml", b"456"),
                ("c.xml", b"789"),
                ("deep/path/item.xml", b"x"),
            ],
            CompressionMethod::Stored,
        );

        let count_err = SecureZipReader::new(
            Cursor::new(bytes.clone()),
            ZipConfig {
                max_file_count: 2,
                ..ZipConfig::default()
            },
        )
        .err()
        .expect("file count limit error");
        assert!(matches!(count_err, ParseError::ResourceLimit(_)));

        let depth_err = SecureZipReader::new(
            Cursor::new(bytes.clone()),
            ZipConfig {
                max_path_depth: 1,
                ..ZipConfig::default()
            },
        )
        .err()
        .expect("path depth limit error");
        assert!(matches!(depth_err, ParseError::ResourceLimit(_)));

        let total_err = SecureZipReader::new(
            Cursor::new(bytes),
            ZipConfig {
                max_total_size: 3,
                ..ZipConfig::default()
            },
        )
        .err()
        .expect("total size limit error");
        assert!(matches!(total_err, ParseError::ResourceLimit(_)));
    }

    #[test]
    fn secure_zip_reader_rejects_path_traversal_and_large_files() {
        let traversal = make_zip(&[("../evil.xml", b"x")], CompressionMethod::Stored);
        let err = SecureZipReader::new(Cursor::new(traversal), ZipConfig::default())
            .err()
            .expect("path traversal error");
        assert!(matches!(err, ParseError::PathTraversal(_)));

        let bytes = make_zip(
            &[("word/document.xml", b"12345")],
            CompressionMethod::Stored,
        );
        let err = SecureZipReader::new(
            Cursor::new(bytes),
            ZipConfig {
                max_file_size: 4,
                ..ZipConfig::default()
            },
        )
        .err()
        .expect("file size limit error");
        assert!(matches!(err, ParseError::ResourceLimit(_)));
    }

    #[test]
    fn secure_zip_reader_rejects_suspicious_compression_ratio() {
        let large = vec![b'A'; 3 * 1024 * 1024];
        let bytes = make_zip(
            &[("word/document.xml", &large)],
            CompressionMethod::Deflated,
        );

        let err = SecureZipReader::new(
            Cursor::new(bytes),
            ZipConfig {
                max_file_size: 4 * 1024 * 1024,
                max_total_size: 4 * 1024 * 1024,
                max_compression_ratio: 1.0,
                ..ZipConfig::default()
            },
        )
        .err()
        .expect("compression ratio limit error");
        assert!(matches!(err, ParseError::ResourceLimit(_)));
    }
}
