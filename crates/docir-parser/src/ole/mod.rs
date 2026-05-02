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

    fn read_file_string(&mut self, name: &str) -> Result<String, ParseError> {
        let bytes = self
            .cfb
            .read_stream(name)
            .ok_or_else(|| ParseError::MissingPart(name.to_string()))?;
        String::from_utf8(bytes)
            .map_err(|e| ParseError::Encoding(format!("Invalid UTF-8 in {}: {}", name, e)))
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
mod tests {
    use super::*;
    use crate::ole_header::{END_OF_CHAIN, FREE_SECT, SIGNATURE};

    fn valid_header_template() -> Vec<u8> {
        let mut data = vec![0u8; 512];
        data[..8].copy_from_slice(&SIGNATURE);
        data[0x1E..0x20].copy_from_slice(&(9u16).to_le_bytes());
        data[0x20..0x22].copy_from_slice(&(6u16).to_le_bytes());
        data[0x2C..0x30].copy_from_slice(&(1u32).to_le_bytes());
        data[0x30..0x34].copy_from_slice(&(0u32).to_le_bytes());
        data[0x38..0x3C].copy_from_slice(&(4096u32).to_le_bytes());
        data[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        data[0x40..0x44].copy_from_slice(&(0u32).to_le_bytes());
        data[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        data[0x48..0x4C].copy_from_slice(&(0u32).to_le_bytes());
        for i in 0..109usize {
            let off = 0x4C + i * 4;
            data[off..off + 4].copy_from_slice(&FREE_SECT.to_le_bytes());
        }
        data
    }

    #[test]
    fn ole_signature_detection_and_header_parse() {
        assert!(!is_ole_container(b"not-ole"));
        let header = valid_header_template();
        assert!(is_ole_container(&header));

        let parsed = parse_header(&header).expect("valid header");
        assert_eq!(parsed.sector_size, 512);
        assert_eq!(parsed.mini_sector_size, 64);
        assert_eq!(parsed.num_fat_sectors, 1);
        assert_eq!(parsed.mini_cutoff, 4096);
    }

    #[test]
    fn parse_header_rejects_invalid_signature() {
        let mut bad = vec![0u8; 512];
        bad[..8].copy_from_slice(b"BADHDR!!");
        match parse_header(&bad) {
            Ok(_) => panic!("invalid header should fail"),
            Err(err) => assert!(matches!(err, ParseError::InvalidStructure(_))),
        }
    }

    #[test]
    fn difat_chain_uses_header_entries() {
        let mut header = valid_header_template();
        header[0x4C..0x50].copy_from_slice(&(3u32).to_le_bytes());
        header[0x50..0x54].copy_from_slice(&(7u32).to_le_bytes());

        let parsed = parse_header(&header).expect("header");
        let difat = read_difat_chain(&header, &parsed).expect("difat");
        assert_eq!(difat, vec![3, 7]);
    }

    #[test]
    fn read_sector_and_stream_helpers_handle_bounds_and_chains() {
        let mut data = vec![0u8; 1536];
        data[512..1024].fill(1);
        data[1024..1536].fill(2);

        let s1 = read_sector(&data, 512, 0).expect("sector 0");
        let s2 = read_sector(&data, 512, 1).expect("sector 1");
        assert_eq!(s1[0], 1);
        assert_eq!(s2[0], 2);
        assert!(read_sector(&data, 512, 99).is_err());

        let fat = vec![1, END_OF_CHAIN];
        let stream = read_stream_from_fat(&data, 512, &fat, 0).expect("fat stream");
        assert_eq!(stream.len(), 1024);
        assert_eq!(stream[0], 1);
        assert_eq!(stream[512], 2);

        let mini_stream = b"abcdEFGHijklMNOP".to_vec();
        let mini_fat = vec![1, END_OF_CHAIN];
        let mini = read_stream_from_mini(&mini_stream, 8, &mini_fat, 0, 12).expect("mini stream");
        assert_eq!(mini, b"abcdEFGHijkl".to_vec());
    }

    #[test]
    fn directory_parsing_and_tree_walk_collect_stream_paths() {
        let mut dir = vec![0u8; 128 * 3];

        dir[66] = 5;
        dir[68..72].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[72..76].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[76..80].copy_from_slice(&(1u32).to_le_bytes());
        dir[116..120].copy_from_slice(&(0u32).to_le_bytes());
        dir[120..128].copy_from_slice(&(0u64).to_le_bytes());

        let name1: Vec<u8> = "VBA"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .chain([0, 0])
            .collect();
        dir[128..128 + name1.len()].copy_from_slice(&name1);
        dir[128 + 64..128 + 66].copy_from_slice(&((name1.len()) as u16).to_le_bytes());
        dir[128 + 66] = 2;
        dir[128 + 68..128 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[128 + 72..128 + 76].copy_from_slice(&(2u32).to_le_bytes());
        dir[128 + 76..128 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[128 + 116..128 + 120].copy_from_slice(&(1u32).to_le_bytes());
        dir[128 + 120..128 + 128].copy_from_slice(&(10u64).to_le_bytes());

        let name2: Vec<u8> = "dir"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .chain([0, 0])
            .collect();
        dir[256..256 + name2.len()].copy_from_slice(&name2);
        dir[256 + 64..256 + 66].copy_from_slice(&((name2.len()) as u16).to_le_bytes());
        dir[256 + 66] = 2;
        dir[256 + 68..256 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[256 + 72..256 + 76].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[256 + 76..256 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[256 + 116..256 + 120].copy_from_slice(&(2u32).to_le_bytes());
        dir[256 + 120..256 + 128].copy_from_slice(&(20u64).to_le_bytes());

        let entries = directory::parse_dir_entries(&dir).expect("dir entries");
        assert_eq!(entries.len(), 3);
        assert_eq!(entries[1].name, "VBA");
        assert_eq!(entries[2].name, "dir");

        let mut out = HashMap::new();
        let mut visited = std::collections::HashSet::new();
        types::walk_siblings(entries[0].child, "", &entries, &mut out, 0, &mut visited);
        assert!(out.contains_key("VBA"));
        assert!(out.contains_key("dir"));

        let metadata = collect_entry_metadata(&entries);
        assert!(metadata.contains_key("Root Entry"));
        assert!(metadata.contains_key("VBA"));
        assert_eq!(metadata["VBA"].entry_type, CfbEntryType::Stream);
    }

    #[test]
    fn collect_directory_slots_marks_normal_orphaned_and_free_entries() {
        let mut dir = vec![0u8; 128 * 4];

        dir[66] = 5;
        dir[68..72].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[72..76].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[76..80].copy_from_slice(&(1u32).to_le_bytes());

        let linked_name: Vec<u8> = "WordDocument"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .chain([0, 0])
            .collect();
        dir[128..128 + linked_name.len()].copy_from_slice(&linked_name);
        dir[128 + 64..128 + 66].copy_from_slice(&(linked_name.len() as u16).to_le_bytes());
        dir[128 + 66] = 2;
        dir[128 + 68..128 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[128 + 72..128 + 76].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[128 + 76..128 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[128 + 116..128 + 120].copy_from_slice(&(1u32).to_le_bytes());
        dir[128 + 120..128 + 128].copy_from_slice(&(3u64).to_le_bytes());

        let orphan_name: Vec<u8> = "Ghost"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .chain([0, 0])
            .collect();
        dir[256..256 + orphan_name.len()].copy_from_slice(&orphan_name);
        dir[256 + 64..256 + 66].copy_from_slice(&(orphan_name.len() as u16).to_le_bytes());
        dir[256 + 66] = 2;
        dir[256 + 68..256 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[256 + 72..256 + 76].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[256 + 76..256 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[256 + 116..256 + 120].copy_from_slice(&(2u32).to_le_bytes());
        dir[256 + 120..256 + 128].copy_from_slice(&(5u64).to_le_bytes());

        dir[384 + 68..384 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[384 + 72..384 + 76].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[384 + 76..384 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());

        let entries = directory::parse_dir_entries(&dir).expect("dir entries");
        let slots = collect_directory_slots(&entries);

        assert!(slots
            .iter()
            .any(|slot| slot.path == "WordDocument" && slot.state == CfbDirectoryState::Normal));
        assert!(slots
            .iter()
            .any(|slot| slot.path == "Ghost" && slot.state == CfbDirectoryState::Orphaned));
        assert!(slots
            .iter()
            .any(|slot| slot.entry_index == 3 && slot.state == CfbDirectoryState::Free));
        assert!(slots.iter().any(|slot| slot.path == "WordDocument"
            && slot.name_len_raw as usize == linked_name.len()
            && slot.color_flag_raw == 0
            && slot.left_sibling_raw == FREE_SECT
            && slot.right_sibling_raw == FREE_SECT
            && slot.child_raw == FREE_SECT));
    }

    #[test]
    fn cfb_reader_lists_and_filters_stream_names() {
        let mut streams = HashMap::new();
        streams.insert(
            "VBA/dir".to_string(),
            DirEntry {
                name: "dir".to_string(),
                name_len_raw: 8,
                object_type: 2,
                color_flag: 0,
                left: FREE_SECT,
                right: FREE_SECT,
                child: FREE_SECT,
                start_sector: 0,
                size: 3,
                created_filetime: None,
                modified_filetime: None,
            },
        );
        streams.insert(
            "VBA/Module1".to_string(),
            DirEntry {
                name: "Module1".to_string(),
                name_len_raw: 16,
                object_type: 2,
                color_flag: 0,
                left: FREE_SECT,
                right: FREE_SECT,
                child: FREE_SECT,
                start_sector: 1,
                size: 4,
                created_filetime: None,
                modified_filetime: None,
            },
        );

        let mut entries = HashMap::new();
        entries.insert(
            "Root Entry".to_string(),
            CfbEntryMetadata {
                entry_index: 0,
                path: "Root Entry".to_string(),
                entry_type: CfbEntryType::RootStorage,
                object_type_raw: 5,
                size: 0,
                start_sector: END_OF_CHAIN,
                left_sibling: None,
                right_sibling: None,
                child: Some(1),
                created_filetime: None,
                modified_filetime: None,
            },
        );
        entries.insert(
            "VBA/dir".to_string(),
            CfbEntryMetadata {
                entry_index: 1,
                path: "VBA/dir".to_string(),
                entry_type: CfbEntryType::Stream,
                object_type_raw: 2,
                size: 3,
                start_sector: 0,
                left_sibling: None,
                right_sibling: Some(2),
                child: None,
                created_filetime: None,
                modified_filetime: None,
            },
        );
        entries.insert(
            "VBA/Module1".to_string(),
            CfbEntryMetadata {
                entry_index: 2,
                path: "VBA/Module1".to_string(),
                entry_type: CfbEntryType::Stream,
                object_type_raw: 2,
                size: 4,
                start_sector: 1,
                left_sibling: None,
                right_sibling: None,
                child: None,
                created_filetime: None,
                modified_filetime: None,
            },
        );

        let cfb = Cfb {
            sector_size: 512,
            mini_sector_size: 64,
            mini_cutoff: 4096,
            num_fat_sectors: 1,
            first_dir_sector: 0,
            first_mini_fat: END_OF_CHAIN,
            num_mini_fat: 0,
            first_difat: END_OF_CHAIN,
            num_difat: 0,
            difat_entry_count: 0,
            fat: vec![END_OF_CHAIN, END_OF_CHAIN],
            mini_fat: Vec::new(),
            root_stream: Vec::new(),
            streams,
            entries,
            directory_slots: Vec::new(),
            data: vec![0u8; 1536],
        };
        let mut reader = CfbReader::new(&cfb);

        assert!(reader.contains("VBA/dir"));
        assert_eq!(reader.file_size("VBA/Module1").expect("size"), 4);
        assert!(reader.list_prefix("VBA/").len() >= 2);
        assert_eq!(reader.list_suffix("dir"), vec!["VBA/dir".to_string()]);
        assert!(reader.read_file("missing").is_err());
    }

    #[test]
    fn cfb_path_normalization_and_read_file_string_error_paths() {
        let mut streams = HashMap::new();
        streams.insert(
            "VBA/Module1".to_string(),
            DirEntry {
                name: "Module1".to_string(),
                name_len_raw: 16,
                object_type: 2,
                color_flag: 0,
                left: FREE_SECT,
                right: FREE_SECT,
                child: FREE_SECT,
                start_sector: 0,
                size: 2,
                created_filetime: None,
                modified_filetime: None,
            },
        );

        let mut entries = HashMap::new();
        entries.insert(
            "Root Entry".to_string(),
            CfbEntryMetadata {
                entry_index: 0,
                path: "Root Entry".to_string(),
                entry_type: CfbEntryType::RootStorage,
                object_type_raw: 5,
                size: 0,
                start_sector: END_OF_CHAIN,
                left_sibling: None,
                right_sibling: None,
                child: Some(1),
                created_filetime: None,
                modified_filetime: None,
            },
        );
        entries.insert(
            "VBA/Module1".to_string(),
            CfbEntryMetadata {
                entry_index: 1,
                path: "VBA/Module1".to_string(),
                entry_type: CfbEntryType::Stream,
                object_type_raw: 2,
                size: 2,
                start_sector: 0,
                left_sibling: None,
                right_sibling: None,
                child: None,
                created_filetime: None,
                modified_filetime: None,
            },
        );

        let mut data = vec![0u8; 1024];
        data[512] = 0xFF;
        data[513] = 0xFE;

        let cfb = Cfb {
            sector_size: 512,
            mini_sector_size: 64,
            mini_cutoff: 4096,
            num_fat_sectors: 1,
            first_dir_sector: 0,
            first_mini_fat: END_OF_CHAIN,
            num_mini_fat: 0,
            first_difat: END_OF_CHAIN,
            num_difat: 0,
            difat_entry_count: 0,
            fat: vec![END_OF_CHAIN],
            mini_fat: Vec::new(),
            root_stream: Vec::new(),
            streams,
            entries,
            directory_slots: Vec::new(),
            data,
        };
        let mut reader = CfbReader::new(&cfb);

        assert!(reader.contains("VBA\\Module1"));
        let err = reader
            .read_file_string("VBA\\Module1")
            .expect_err("invalid UTF-8 should fail");
        assert!(matches!(err, ParseError::Encoding(_)));
    }

    #[test]
    fn read_stream_from_mini_handles_out_of_bounds_and_chain_breaks() {
        let mini_stream = b"abcdEFGH".to_vec();
        let mini_fat = vec![END_OF_CHAIN];
        let out = read_stream_from_mini(&mini_stream, 8, &mini_fat, 10, 4).expect("mini stream");
        assert!(out.is_empty());

        let mini_stream = b"abcdefghijklmnop".to_vec();
        let out = read_stream_from_mini(&mini_stream, 8, &mini_fat, 0, 12).expect("mini stream");
        assert_eq!(out, b"abcdefgh".to_vec());
    }

    #[test]
    fn parse_dir_entries_reads_optional_filetime_metadata() {
        let mut dir = vec![0u8; 128 * 2];
        dir[66] = 5;
        dir[76..80].copy_from_slice(&(1u32).to_le_bytes());

        let name1: Vec<u8> = "Storage"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .chain([0, 0])
            .collect();
        dir[128..128 + name1.len()].copy_from_slice(&name1);
        dir[128 + 64..128 + 66].copy_from_slice(&((name1.len()) as u16).to_le_bytes());
        dir[128 + 66] = 1;
        dir[128 + 100..128 + 108].copy_from_slice(&(123456u64).to_le_bytes());
        dir[128 + 108..128 + 116].copy_from_slice(&(654321u64).to_le_bytes());

        let entries = directory::parse_dir_entries(&dir).expect("dir entries");
        assert_eq!(entries[1].created_filetime, Some(123456));
        assert_eq!(entries[1].modified_filetime, Some(654321));
    }
}
