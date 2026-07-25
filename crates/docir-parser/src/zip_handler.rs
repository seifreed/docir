//! Secure ZIP handling for OOXML packages.
//!
//! This module provides secure reading of ZIP archives with protections
//! against zip bombs, path traversal, and other malicious archive attacks.

use crate::error::ParseError;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek, SeekFrom};
use zip::ZipArchive;

#[cfg(test)]
#[path = "zip_handler_tests.rs"]
mod tests;
#[path = "zip_handler_validation.rs"]
mod zip_handler_validation;
use zip_handler_validation::{validate_archive_entry, validate_total_size};

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
    pub fn new(mut reader: R, config: ZipConfig) -> Result<Self, ParseError> {
        reject_duplicate_central_directory_names(&mut reader, config.max_file_count)?;
        let mut archive = ZipArchive::new(reader)?;

        if archive.len() > config.max_file_count {
            return Err(ParseError::ResourceLimit(format!(
                "Too many files in archive: {} (max: {})",
                archive.len(),
                config.max_file_count
            )));
        }

        let mut file_index = HashMap::new();
        let mut total_uncompressed = 0u64;

        for i in 0..archive.len() {
            let file = archive.by_index_raw(i)?;
            let name = file.name().to_string();
            let uncompressed_size = file.size();
            let compressed_size = file.compressed_size();
            validate_archive_entry(&name, uncompressed_size, compressed_size, &config)?;

            if file_index.contains_key(&name) {
                return Err(ParseError::InvalidZip(format!(
                    "Duplicate file name in archive: {name}"
                )));
            }
            total_uncompressed = total_uncompressed
                .checked_add(uncompressed_size)
                .ok_or_else(|| {
                    ParseError::ResourceLimit("Total uncompressed size overflow".to_string())
                })?;
            file_index.insert(name, i);
        }

        validate_total_size(total_uncompressed, &config)?;

        Ok(Self {
            archive,
            config,
            file_index,
        })
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

const EOCD_SIGNATURE: &[u8; 4] = b"PK\x05\x06";
const ZIP64_EOCD_SIGNATURE: &[u8; 4] = b"PK\x06\x06";
const ZIP64_LOCATOR_SIGNATURE: &[u8; 4] = b"PK\x06\x07";
const CENTRAL_FILE_SIGNATURE: &[u8; 4] = b"PK\x01\x02";
const EOCD_MIN_SIZE: usize = 22;
const EOCD_MAX_SEARCH: u64 = 66_000;
const ZIP64_SENTINEL_U16: u16 = u16::MAX;
const ZIP64_SENTINEL_U32: u32 = u32::MAX;

struct CentralDirectoryInfo {
    total_entries: usize,
    offset: u64,
}

fn read_central_directory_info<R: Read + Seek>(
    reader: &mut R,
) -> Result<Option<CentralDirectoryInfo>, ParseError> {
    let original_pos = reader.stream_position()?;
    let archive_len = reader.seek(SeekFrom::End(0))?;
    let tail_len = archive_len.min(EOCD_MAX_SEARCH) as usize;
    reader.seek(SeekFrom::End(-(tail_len as i64)))?;
    let mut tail = vec![0; tail_len];
    reader.read_exact(&mut tail)?;
    reader.seek(SeekFrom::Start(original_pos))?;

    let Some(eocd) = tail.windows(4).rposition(|window| window == EOCD_SIGNATURE) else {
        return Ok(None);
    };
    if eocd + EOCD_MIN_SIZE > tail.len() {
        return Ok(None);
    }

    let total_entries = u16::from_le_bytes([tail[eocd + 10], tail[eocd + 11]]);
    let central_size = u32::from_le_bytes([
        tail[eocd + 12],
        tail[eocd + 13],
        tail[eocd + 14],
        tail[eocd + 15],
    ]);
    let central_offset = u32::from_le_bytes([
        tail[eocd + 16],
        tail[eocd + 17],
        tail[eocd + 18],
        tail[eocd + 19],
    ]);
    if total_entries == ZIP64_SENTINEL_U16
        || central_size == ZIP64_SENTINEL_U32
        || central_offset == ZIP64_SENTINEL_U32
    {
        let eocd_offset = archive_len - tail_len as u64 + eocd as u64;
        return read_zip64_central_directory_info(reader, eocd_offset, original_pos);
    }

    Ok(Some(CentralDirectoryInfo {
        total_entries: total_entries as usize,
        offset: central_offset as u64,
    }))
}

fn read_zip64_central_directory_info<R: Read + Seek>(
    reader: &mut R,
    eocd_offset: u64,
    original_pos: u64,
) -> Result<Option<CentralDirectoryInfo>, ParseError> {
    const ZIP64_LOCATOR_LEN: u64 = 20;
    if eocd_offset < ZIP64_LOCATOR_LEN {
        return Ok(None);
    }
    reader.seek(SeekFrom::Start(eocd_offset - ZIP64_LOCATOR_LEN))?;
    let mut locator = [0u8; ZIP64_LOCATOR_LEN as usize];
    reader.read_exact(&mut locator)?;
    reader.seek(SeekFrom::Start(original_pos))?;
    if &locator[..4] != ZIP64_LOCATOR_SIGNATURE {
        return Ok(None);
    }

    let zip64_eocd_offset = u64::from_le_bytes([
        locator[8],
        locator[9],
        locator[10],
        locator[11],
        locator[12],
        locator[13],
        locator[14],
        locator[15],
    ]);
    reader.seek(SeekFrom::Start(zip64_eocd_offset))?;
    let mut header = [0u8; 56];
    reader.read_exact(&mut header)?;
    reader.seek(SeekFrom::Start(original_pos))?;
    if &header[..4] != ZIP64_EOCD_SIGNATURE {
        return Ok(None);
    }

    let total_entries = u64::from_le_bytes([
        header[32], header[33], header[34], header[35], header[36], header[37], header[38],
        header[39],
    ]);
    let offset = u64::from_le_bytes([
        header[48], header[49], header[50], header[51], header[52], header[53], header[54],
        header[55],
    ]);
    let total_entries = usize::try_from(total_entries).map_err(|_| {
        ParseError::ResourceLimit("ZIP64 central directory entry count too large".to_string())
    })?;

    Ok(Some(CentralDirectoryInfo {
        total_entries,
        offset,
    }))
}

fn reject_duplicate_central_directory_names<R: Read + Seek>(
    reader: &mut R,
    max_file_count: usize,
) -> Result<(), ParseError> {
    let Some(info) = read_central_directory_info(reader)? else {
        return Ok(());
    };
    let original_pos = reader.stream_position()?;
    let total_entries = info.total_entries;
    if total_entries > max_file_count {
        return Err(ParseError::ResourceLimit(format!(
            "Too many files in archive: {total_entries} (max: {max_file_count})"
        )));
    }

    reader.seek(SeekFrom::Start(info.offset))?;
    let mut seen = HashSet::with_capacity(total_entries);
    for _ in 0..total_entries {
        let mut header = [0u8; 46];
        reader.read_exact(&mut header)?;
        if &header[..4] != CENTRAL_FILE_SIGNATURE {
            return Err(ParseError::InvalidZip(
                "Invalid central directory header".to_string(),
            ));
        }
        let name_len = u16::from_le_bytes([header[28], header[29]]) as usize;
        let extra_len = u16::from_le_bytes([header[30], header[31]]) as i64;
        let comment_len = u16::from_le_bytes([header[32], header[33]]) as i64;
        let mut name = vec![0; name_len];
        reader.read_exact(&mut name)?;
        if !seen.insert(name.clone()) {
            return Err(ParseError::InvalidZip(format!(
                "Duplicate file name in archive: {}",
                String::from_utf8_lossy(&name)
            )));
        }
        reader.seek(SeekFrom::Current(extra_len + comment_len))?;
    }

    reader.seek(SeekFrom::Start(original_pos))?;
    Ok(())
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
