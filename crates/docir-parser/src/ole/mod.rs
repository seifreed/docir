//! Minimal OLE Compound File Binary (CFB) parser for VBA extraction.

mod directory;
mod stream;
mod types;

use std::collections::HashMap;

use crate::error::ParseError;
use crate::ole_header::{
    parse_header, read_difat_chain, read_fat_table, read_mini_fat_table, END_OF_CHAIN, FAT_SECT,
    FREE_SECT, SIGNATURE,
};
use crate::zip_handler::PackageReader;

pub use types::{CfbDirectorySlot, CfbDirectoryState, CfbEntryMetadata, CfbEntryType};

pub struct CfbReader<'a> {
    cfb: &'a Cfb,
}

impl<'a> CfbReader<'a> {
    /// Public API entrypoint: new.
    pub fn new(cfb: &'a Cfb) -> Self {
        Self { cfb }
    }
}

use directory::{
    collect_directory_slots, collect_entry_metadata, read_directory_entries_and_root_stream,
};
use stream::{collect_chain_with_terminal, read_stream_from_mini};
use types::{collect_stream_entries, DirEntry};

/// Parsed CFB file with streams.
pub struct Cfb {
    sector_size: u32,
    mini_sector_size: u32,
    mini_cutoff: u32,
    num_fat_sectors: u32,
    first_dir_sector: u32,
    first_mini_fat: u32,
    num_mini_fat: u32,
    first_difat: u32,
    num_difat: u32,
    difat_entry_count: usize,
    fat: Vec<u32>,
    mini_fat: Vec<u32>,
    root_stream: Vec<u8>,
    streams: HashMap<String, DirEntry>,
    entries: HashMap<String, CfbEntryMetadata>,
    directory_slots: Vec<CfbDirectorySlot>,
    data: Vec<u8>,
}

impl Cfb {
    /// Public API entrypoint: parse.
    pub fn parse(data: Vec<u8>) -> Result<Self, ParseError> {
        let header = parse_header(&data)?;
        let difat = read_difat_chain(&data, &header)?;
        let fat = read_fat_table(&data, header.sector_size, &difat, header.num_fat_sectors)?;
        let (entries, root_stream) = read_directory_entries_and_root_stream(
            &data,
            header.sector_size,
            &fat,
            header.first_dir_sector,
        )?;
        let mini_fat = read_mini_fat_table(
            &data,
            header.sector_size,
            &fat,
            header.first_mini_fat,
            header.num_mini_fat,
        )?;
        let streams = collect_stream_entries(&entries);
        let directory_slots = collect_directory_slots(&entries);
        let entries = collect_entry_metadata(&entries);

        Ok(Self {
            sector_size: header.sector_size,
            mini_sector_size: header.mini_sector_size,
            mini_cutoff: header.mini_cutoff,
            num_fat_sectors: header.num_fat_sectors,
            first_dir_sector: header.first_dir_sector,
            first_mini_fat: header.first_mini_fat,
            num_mini_fat: header.num_mini_fat,
            first_difat: header.first_difat,
            num_difat: header.num_difat,
            difat_entry_count: difat.len(),
            fat,
            mini_fat,
            root_stream,
            streams,
            entries,
            directory_slots,
            data,
        })
    }

    /// Public API entrypoint: read_stream.
    pub fn read_stream(&self, path: &str) -> Option<Vec<u8>> {
        let entry = self.resolve_stream_entry(path)?;
        if entry.object_type != 2 {
            return None;
        }

        if self.should_use_mini_stream(entry) {
            return read_stream_from_mini(
                &self.root_stream,
                self.mini_sector_size,
                &self.mini_fat,
                entry.start_sector,
                entry.size as usize,
            );
        }

        let data = self.read_regular_stream(entry).ok()?;
        let size = usize::try_from(entry.size).unwrap_or(usize::MAX);
        let len = data.len().min(size);
        Some(data[..len].to_vec())
    }

    /// Public API entrypoint: has_stream.
    pub fn has_stream(&self, path: &str) -> bool {
        self.resolve_stream_entry(path).is_some()
    }

    /// Public API entrypoint: list_streams.
    pub fn list_streams(&self) -> Vec<String> {
        let mut keys: Vec<String> = self.streams.keys().cloned().collect();
        keys.sort();
        keys
    }

    /// Public API entrypoint: stream_size.
    pub fn stream_size(&self, path: &str) -> Option<u64> {
        self.resolve_stream_entry(path).map(|entry| entry.size)
    }

    /// Public API entrypoint: list_entries.
    pub fn list_entries(&self) -> Vec<CfbEntryMetadata> {
        let mut entries: Vec<CfbEntryMetadata> = self.entries.values().cloned().collect();
        entries.sort_by(|left, right| left.path.cmp(&right.path));
        entries
    }

    /// Public API entrypoint: list_directory_slots.
    pub fn list_directory_slots(&self) -> Vec<CfbDirectorySlot> {
        self.directory_slots.clone()
    }

    /// Public API entrypoint: entry_metadata.
    pub fn entry_metadata(&self, path: &str) -> Option<&CfbEntryMetadata> {
        self.entries
            .get(path)
            .or_else(|| self.entries.get(&path.replace('\\', "/")))
    }

    /// Public API entrypoint: sector_size.
    pub fn sector_size(&self) -> u32 {
        self.sector_size
    }

    /// Public API entrypoint: mini_sector_size.
    pub fn mini_sector_size(&self) -> u32 {
        self.mini_sector_size
    }

    /// Public API entrypoint: mini_cutoff.
    pub fn mini_cutoff(&self) -> u32 {
        self.mini_cutoff
    }

    /// Public API entrypoint: num_fat_sectors.
    pub fn num_fat_sectors(&self) -> u32 {
        self.num_fat_sectors
    }

    /// Public API entrypoint: first_dir_sector.
    pub fn first_dir_sector(&self) -> u32 {
        self.first_dir_sector
    }

    /// Public API entrypoint: first_mini_fat.
    pub fn first_mini_fat(&self) -> u32 {
        self.first_mini_fat
    }

    /// Public API entrypoint: num_mini_fat.
    pub fn num_mini_fat(&self) -> u32 {
        self.num_mini_fat
    }

    /// Public API entrypoint: first_difat.
    pub fn first_difat(&self) -> u32 {
        self.first_difat
    }

    /// Public API entrypoint: num_difat.
    pub fn num_difat(&self) -> u32 {
        self.num_difat
    }

    /// Public API entrypoint: difat_entry_count.
    pub fn difat_entry_count(&self) -> usize {
        self.difat_entry_count
    }

    /// Public API entrypoint: sector_count.
    pub fn sector_count(&self) -> u32 {
        let count = (self.data.len() / self.sector_size as usize).saturating_sub(1);
        count.try_into().unwrap_or(u32::MAX)
    }

    /// Public API entrypoint: fat_entry_count.
    pub fn fat_entry_count(&self) -> usize {
        self.fat.len()
    }

    /// Public API entrypoint: fat_entry_value.
    pub fn fat_entry_value(&self, sector: u32) -> Option<u32> {
        self.fat.get(sector as usize).copied()
    }

    /// Public API entrypoint: mini_fat_entry_count.
    pub fn mini_fat_entry_count(&self) -> usize {
        self.mini_fat.len()
    }

    /// Public API entrypoint: fat_free_count.
    pub fn fat_free_count(&self) -> usize {
        self.fat.iter().filter(|entry| **entry == FREE_SECT).count()
    }

    /// Public API entrypoint: fat_end_of_chain_count.
    pub fn fat_end_of_chain_count(&self) -> usize {
        self.fat
            .iter()
            .filter(|entry| **entry == END_OF_CHAIN)
            .count()
    }

    /// Public API entrypoint: fat_reserved_count.
    pub fn fat_reserved_count(&self) -> usize {
        self.fat.iter().filter(|entry| **entry == FAT_SECT).count()
    }

    /// Public API entrypoint: stream_sector_chain.
    pub fn stream_sector_chain(&self, path: &str) -> Option<Vec<u32>> {
        let entry = self.resolve_stream_entry(path)?;
        if entry.object_type != 2 {
            return None;
        }
        let (chain, _) = if self.should_use_mini_stream(entry) {
            collect_chain_with_terminal(&self.mini_fat, entry.start_sector)
        } else {
            collect_chain_with_terminal(&self.fat, entry.start_sector)
        };
        Some(chain)
    }

    /// Public API entrypoint: stream_chain_terminal.
    pub fn stream_chain_terminal(&self, path: &str) -> Option<u32> {
        let entry = self.resolve_stream_entry(path)?;
        if entry.object_type != 2 {
            return None;
        }
        let (_, terminal) = if self.should_use_mini_stream(entry) {
            collect_chain_with_terminal(&self.mini_fat, entry.start_sector)
        } else {
            collect_chain_with_terminal(&self.fat, entry.start_sector)
        };
        Some(terminal)
    }

    /// Public API entrypoint: stream_uses_mini_fat.
    pub fn stream_uses_mini_fat(&self, path: &str) -> Option<bool> {
        self.resolve_stream_entry(path)
            .filter(|entry| entry.object_type == 2)
            .map(|entry| self.should_use_mini_stream(entry))
    }

    fn resolve_stream_entry(&self, path: &str) -> Option<&DirEntry> {
        self.streams
            .get(path)
            .or_else(|| self.streams.get(&path.replace('\\', "/")))
    }

    fn should_use_mini_stream(&self, entry: &DirEntry) -> bool {
        entry.size < self.mini_cutoff as u64 && !self.root_stream.is_empty()
    }

    fn read_regular_stream(&self, entry: &DirEntry) -> Result<Vec<u8>, ParseError> {
        read_stream_from_fat(&self.data, self.sector_size, &self.fat, entry.start_sector)
    }
}

impl PackageReader for CfbReader<'_> {
    fn contains(&self, name: &str) -> bool {
        self.cfb.has_stream(name)
    }

    fn read_file(&mut self, name: &str) -> Result<Vec<u8>, ParseError> {
        self.cfb
            .read_stream(name)
            .ok_or_else(|| ParseError::MissingPart(name.to_string()))
    }

    fn file_size(&mut self, name: &str) -> Result<u64, ParseError> {
        self.cfb
            .stream_size(name)
            .ok_or_else(|| ParseError::MissingPart(name.to_string()))
    }

    fn file_names(&self) -> Vec<String> {
        self.cfb.list_streams()
    }

    fn list_prefix(&self, prefix: &str) -> Vec<String> {
        self.cfb
            .list_streams()
            .into_iter()
            .filter(|name| name.starts_with(prefix))
            .collect()
    }

    fn list_suffix(&self, suffix: &str) -> Vec<String> {
        self.cfb
            .list_streams()
            .into_iter()
            .filter(|name| name.ends_with(suffix))
            .collect()
    }
}

/// Public API entrypoint: is_ole_container.
pub fn is_ole_container(data: &[u8]) -> bool {
    data.len() >= SIGNATURE.len() && data[..SIGNATURE.len()] == SIGNATURE
}

pub(crate) use stream::{read_sector, read_stream_from_fat};

#[cfg(test)]
mod tests;
