//! Minimal OLE Compound File Binary (CFB) parser for VBA extraction.

use crate::error::ParseError;
use crate::ole_header::{
    parse_header, read_difat_chain, read_fat_table, read_mini_fat_table, read_u16, read_u32,
    END_OF_CHAIN, FAT_SECT, FREE_SECT, SIGNATURE,
};
use crate::zip_handler::PackageReader;
use std::collections::HashMap;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CfbEntryType {
    RootStorage,
    Storage,
    Stream,
}

#[derive(Debug, Clone)]
pub struct CfbEntryMetadata {
    pub entry_index: u32,
    pub path: String,
    pub entry_type: CfbEntryType,
    pub object_type_raw: u8,
    pub size: u64,
    pub start_sector: u32,
    pub left_sibling: Option<u32>,
    pub right_sibling: Option<u32>,
    pub child: Option<u32>,
    pub created_filetime: Option<u64>,
    pub modified_filetime: Option<u64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CfbDirectoryState {
    Normal,
    Free,
    Orphaned,
}

#[derive(Debug, Clone)]
pub struct CfbDirectorySlot {
    pub entry_index: u32,
    pub path: String,
    pub entry_type: Option<CfbEntryType>,
    pub name_len_raw: u16,
    pub object_type_raw: u8,
    pub color_flag_raw: u8,
    pub state: CfbDirectoryState,
    pub size: u64,
    pub start_sector: u32,
    pub left_sibling_raw: u32,
    pub right_sibling_raw: u32,
    pub child_raw: u32,
    pub left_sibling: Option<u32>,
    pub right_sibling: Option<u32>,
    pub child: Option<u32>,
    pub created_filetime: Option<u64>,
    pub modified_filetime: Option<u64>,
}

#[derive(Debug, Clone)]
struct DirEntry {
    name: String,
    name_len_raw: u16,
    object_type: u8,
    color_flag: u8,
    left: u32,
    right: u32,
    child: u32,
    start_sector: u32,
    size: u64,
    created_filetime: Option<u64>,
    modified_filetime: Option<u64>,
}

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

fn read_directory_entries_and_root_stream(
    data: &[u8],
    sector_size: u32,
    fat: &[u32],
    first_dir_sector: u32,
) -> Result<(Vec<DirEntry>, Vec<u8>), ParseError> {
    let dir_stream = read_stream_from_fat(data, sector_size, fat, first_dir_sector)?;
    let entries = parse_dir_entries(&dir_stream)?;
    let root = entries
        .first()
        .ok_or_else(|| ParseError::InvalidStructure("Missing root entry".to_string()))?;
    let root_stream = read_stream_from_fat(data, sector_size, fat, root.start_sector)?;
    Ok((entries, root_stream))
}

fn collect_stream_entries(entries: &[DirEntry]) -> HashMap<String, DirEntry> {
    let mut streams = HashMap::new();
    if let Some(child) = entries.first().map(|e| e.child) {
        walk_siblings(child, "", entries, &mut streams, 0);
    }
    streams
}

pub struct CfbReader<'a> {
    cfb: &'a Cfb,
}

impl<'a> CfbReader<'a> {
    /// Public API entrypoint: new.
    pub fn new(cfb: &'a Cfb) -> Self {
        Self { cfb }
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

fn parse_dir_entries(data: &[u8]) -> Result<Vec<DirEntry>, ParseError> {
    let mut entries = Vec::new();
    for chunk in data.chunks(128) {
        if chunk.len() < 128 {
            break;
        }
        let name_len_raw = read_u16(chunk, 64)?;
        let name_len = name_len_raw as usize;
        let name_raw = &chunk[..64];
        let name = if (2..=64).contains(&name_len) {
            let bytes = &name_raw[..name_len - 2];
            utf16le_to_string(bytes)
        } else {
            String::new()
        };
        let object_type = chunk[66];
        let color_flag = chunk[67];
        let left = read_u32(chunk, 68)?;
        let right = read_u32(chunk, 72)?;
        let child = read_u32(chunk, 76)?;
        let start_sector = read_u32(chunk, 116)?;
        let size = read_u64(chunk, 120)?;
        let created_filetime = normalize_filetime(read_u64(chunk, 100)?);
        let modified_filetime = normalize_filetime(read_u64(chunk, 108)?);
        entries.push(DirEntry {
            name,
            name_len_raw,
            object_type,
            color_flag,
            left,
            right,
            child,
            start_sector,
            size,
            created_filetime,
            modified_filetime,
        });
    }
    Ok(entries)
}

fn normalize_filetime(value: u64) -> Option<u64> {
    if value == 0 {
        None
    } else {
        Some(value)
    }
}

fn entry_type_from_object_type(object_type: u8) -> Option<CfbEntryType> {
    match object_type {
        5 => Some(CfbEntryType::RootStorage),
        1 => Some(CfbEntryType::Storage),
        2 => Some(CfbEntryType::Stream),
        _ => None,
    }
}

fn walk_siblings(
    idx: u32,
    parent: &str,
    entries: &[DirEntry],
    out: &mut HashMap<String, DirEntry>,
    depth: u32,
) {
    if idx == FREE_SECT || idx == END_OF_CHAIN || depth > MAX_RECURSION_DEPTH {
        return;
    }
    let idx_usize = idx as usize;
    if idx_usize >= entries.len() {
        return;
    }
    let entry = &entries[idx_usize];
    walk_siblings(entry.left, parent, entries, out, depth + 1);
    let mut path = String::new();
    if !parent.is_empty() {
        path.push_str(parent);
        path.push('/');
    }
    path.push_str(&entry.name);
    if entry.object_type == 2 {
        out.insert(path.clone(), entry.clone());
    }
    if (entry.object_type == 1 || entry.object_type == 5) && entry.child != FREE_SECT {
        walk_siblings(entry.child, &path, entries, out, depth + 1);
    }
    walk_siblings(entry.right, parent, entries, out, depth + 1);
}

fn collect_entry_metadata(entries: &[DirEntry]) -> HashMap<String, CfbEntryMetadata> {
    let mut out = HashMap::new();
    if let Some(root) = entries.first() {
        let root_path = if root.name.is_empty() {
            "Root Entry".to_string()
        } else {
            root.name.clone()
        };
        out.insert(
            root_path.clone(),
            CfbEntryMetadata {
                entry_index: 0,
                path: root_path,
                entry_type: CfbEntryType::RootStorage,
                object_type_raw: root.object_type,
                size: root.size,
                start_sector: root.start_sector,
                left_sibling: normalize_tree_index(root.left),
                right_sibling: normalize_tree_index(root.right),
                child: normalize_tree_index(root.child),
                created_filetime: root.created_filetime,
                modified_filetime: root.modified_filetime,
            },
        );
        if root.child != FREE_SECT {
            walk_entry_metadata(root.child, "", entries, &mut out, 0);
        }
    }
    out
}

fn collect_directory_slots(entries: &[DirEntry]) -> Vec<CfbDirectorySlot> {
    let linked = collect_linked_indices(entries);
    entries
        .iter()
        .enumerate()
        .map(|(idx, entry)| {
            let path = derive_entry_path(idx as u32, entries).unwrap_or_else(|| {
                if idx == 0 {
                    "Root Entry".to_string()
                } else if entry.name.is_empty() {
                    format!("Entry {idx}")
                } else {
                    entry.name.clone()
                }
            });
            let entry_type = entry_type_from_object_type(entry.object_type);
            let state = if entry.object_type == 0 {
                CfbDirectoryState::Free
            } else if linked.contains(&(idx as u32)) {
                CfbDirectoryState::Normal
            } else {
                CfbDirectoryState::Orphaned
            };
            CfbDirectorySlot {
                entry_index: idx as u32,
                path,
                entry_type,
                name_len_raw: entry.name_len_raw,
                object_type_raw: entry.object_type,
                color_flag_raw: entry.color_flag,
                state,
                size: entry.size,
                start_sector: entry.start_sector,
                left_sibling_raw: entry.left,
                right_sibling_raw: entry.right,
                child_raw: entry.child,
                left_sibling: normalize_tree_index(entry.left),
                right_sibling: normalize_tree_index(entry.right),
                child: normalize_tree_index(entry.child),
                created_filetime: entry.created_filetime,
                modified_filetime: entry.modified_filetime,
            }
        })
        .collect()
}

fn collect_linked_indices(entries: &[DirEntry]) -> std::collections::HashSet<u32> {
    let mut out = std::collections::HashSet::new();
    if entries.is_empty() {
        return out;
    }
    out.insert(0);
    if let Some(root) = entries.first() {
        if root.child != FREE_SECT {
            walk_linked_indices(root.child, entries, &mut out, 0);
        }
    }
    out
}

const MAX_LINKED_DEPTH: u32 = 256;

fn walk_linked_indices(
    idx: u32,
    entries: &[DirEntry],
    out: &mut std::collections::HashSet<u32>,
    depth: u32,
) {
    if idx == FREE_SECT || idx == END_OF_CHAIN || depth > MAX_LINKED_DEPTH {
        return;
    }
    if entries.is_empty() {
        return;
    }
    let idx_usize = idx as usize;
    if idx_usize >= entries.len() || !out.insert(idx) {
        return;
    }
    let entry = &entries[idx_usize];
    walk_linked_indices(entry.left, entries, out, depth + 1);
    if (entry.object_type == 1 || entry.object_type == 5) && entry.child != FREE_SECT {
        walk_linked_indices(entry.child, entries, out, depth + 1);
    }
    walk_linked_indices(entry.right, entries, out, depth + 1);
}

fn derive_entry_path(idx: u32, entries: &[DirEntry]) -> Option<String> {
    if idx as usize >= entries.len() {
        return None;
    }
    if idx == 0 {
        return Some(if entries[0].name.is_empty() {
            "Root Entry".to_string()
        } else {
            entries[0].name.clone()
        });
    }
    let mut out = None;
    if let Some(root) = entries.first() {
        if root.child != FREE_SECT {
            walk_find_path(root.child, "", idx, entries, &mut out, 0);
        }
    }
    out
}

const MAX_RECURSION_DEPTH: u32 = 256;

fn walk_find_path(
    current: u32,
    parent: &str,
    target: u32,
    entries: &[DirEntry],
    out: &mut Option<String>,
    depth: u32,
) {
    if current == FREE_SECT
        || current == END_OF_CHAIN
        || out.is_some()
        || depth > MAX_RECURSION_DEPTH
    {
        return;
    }
    let current_usize = current as usize;
    if current_usize >= entries.len() {
        return;
    }
    let entry = &entries[current_usize];
    walk_find_path(entry.left, parent, target, entries, out, depth + 1);
    if out.is_some() {
        return;
    }
    let mut path = String::new();
    if !parent.is_empty() {
        path.push_str(parent);
        path.push('/');
    }
    path.push_str(&entry.name);
    if current == target {
        *out = Some(path);
        return;
    }
    if (entry.object_type == 1 || entry.object_type == 5) && entry.child != FREE_SECT {
        walk_find_path(entry.child, &path, target, entries, out, depth + 1);
    }
    if out.is_none() {
        walk_find_path(entry.right, parent, target, entries, out, depth + 1);
    }
}

fn walk_entry_metadata(
    idx: u32,
    parent: &str,
    entries: &[DirEntry],
    out: &mut HashMap<String, CfbEntryMetadata>,
    depth: u32,
) {
    if idx == FREE_SECT || idx == END_OF_CHAIN || depth > MAX_RECURSION_DEPTH {
        return;
    }
    let idx_usize = idx as usize;
    if idx_usize >= entries.len() {
        return;
    }
    let entry = &entries[idx_usize];
    walk_entry_metadata(entry.left, parent, entries, out, depth + 1);

    let mut path = String::new();
    if !parent.is_empty() {
        path.push_str(parent);
        path.push('/');
    }
    path.push_str(&entry.name);

    if let Some(entry_type) = entry_type_from_object_type(entry.object_type) {
        out.insert(
            path.clone(),
            CfbEntryMetadata {
                entry_index: idx,
                path: path.clone(),
                entry_type,
                object_type_raw: entry.object_type,
                size: entry.size,
                start_sector: entry.start_sector,
                left_sibling: normalize_tree_index(entry.left),
                right_sibling: normalize_tree_index(entry.right),
                child: normalize_tree_index(entry.child),
                created_filetime: entry.created_filetime,
                modified_filetime: entry.modified_filetime,
            },
        );
    }

    if (entry.object_type == 1 || entry.object_type == 5) && entry.child != FREE_SECT {
        walk_entry_metadata(entry.child, &path, entries, out, depth + 1);
    }
    walk_entry_metadata(entry.right, parent, entries, out, depth + 1);
}

fn normalize_tree_index(value: u32) -> Option<u32> {
    if value == FREE_SECT || value == END_OF_CHAIN {
        None
    } else {
        Some(value)
    }
}

pub(crate) fn read_stream_from_fat(
    data: &[u8],
    sector_size: u32,
    fat: &[u32],
    start_sector: u32,
) -> Result<Vec<u8>, ParseError> {
    let mut out = Vec::new();
    let mut sector = start_sector;
    let mut guard = 0usize;
    while sector != END_OF_CHAIN && sector != FREE_SECT {
        if guard >= fat.len() {
            break;
        }
        let sec = read_sector(data, sector_size, sector)?;
        out.extend_from_slice(&sec);
        let next = *fat.get(sector as usize).unwrap_or(&END_OF_CHAIN);
        sector = next;
        guard += 1;
    }
    Ok(out)
}

fn read_stream_from_mini(
    mini_stream: &[u8],
    mini_sector_size: u32,
    mini_fat: &[u32],
    start_sector: u32,
    size: usize,
) -> Option<Vec<u8>> {
    let mut out = Vec::new();
    let mut sector = start_sector;
    let mut guard = 0usize;
    while sector != END_OF_CHAIN && sector != FREE_SECT && out.len() < size {
        if guard >= mini_fat.len() {
            break;
        }
        let offset = match (sector as usize).checked_mul(mini_sector_size as usize) {
            Some(o) => o,
            None => break,
        };
        let end = offset + mini_sector_size as usize;
        if end > mini_stream.len() {
            break;
        }
        out.extend_from_slice(&mini_stream[offset..end]);
        let next = *mini_fat.get(sector as usize).unwrap_or(&END_OF_CHAIN);
        sector = next;
        guard += 1;
    }
    out.truncate(size);
    Some(out)
}

fn collect_chain_with_terminal(table: &[u32], start_sector: u32) -> (Vec<u32>, u32) {
    let mut out = Vec::new();
    let mut sector = start_sector;
    let mut guard = 0usize;
    while sector != END_OF_CHAIN && sector != FREE_SECT {
        if guard > table.len() {
            return (out, sector);
        }
        out.push(sector);
        sector = *table.get(sector as usize).unwrap_or(&END_OF_CHAIN);
        guard += 1;
    }
    (out, sector)
}

pub(crate) fn read_sector(
    data: &[u8],
    sector_size: u32,
    sector: u32,
) -> Result<Vec<u8>, ParseError> {
    let sector_size_usize = sector_size as usize;
    let sector_usize = sector as usize;
    let offset = sector_usize
        .checked_add(1)
        .and_then(|s| s.checked_mul(sector_size_usize))
        .ok_or_else(|| ParseError::InvalidStructure("OLE sector offset overflow".to_string()))?;
    let end = offset
        .checked_add(sector_size_usize)
        .ok_or_else(|| ParseError::InvalidStructure("OLE sector end overflow".to_string()))?;
    if end > data.len() {
        return Err(ParseError::InvalidStructure(
            "OLE sector out of bounds".to_string(),
        ));
    }
    Ok(data[offset..end].to_vec())
}

fn read_u64(data: &[u8], offset: usize) -> Result<u64, ParseError> {
    if offset + 8 > data.len() {
        return Err(ParseError::InvalidStructure(
            "OLE read_u64 out of bounds".to_string(),
        ));
    }
    Ok(u64::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ]))
}

fn utf16le_to_string(bytes: &[u8]) -> String {
    let mut u16s = Vec::new();
    for chunk in bytes.chunks(2) {
        if chunk.len() == 2 {
            u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
    }
    String::from_utf16_lossy(&u16s)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn valid_header_template() -> Vec<u8> {
        let mut data = vec![0u8; 512];
        data[..8].copy_from_slice(&SIGNATURE);
        data[0x1E..0x20].copy_from_slice(&(9u16).to_le_bytes()); // sector_size=512
        data[0x20..0x22].copy_from_slice(&(6u16).to_le_bytes()); // mini_sector_size=64
        data[0x2C..0x30].copy_from_slice(&(1u32).to_le_bytes()); // num_fat_sectors
        data[0x30..0x34].copy_from_slice(&(0u32).to_le_bytes()); // first_dir_sector
        data[0x38..0x3C].copy_from_slice(&(4096u32).to_le_bytes()); // mini_cutoff
        data[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes()); // first_mini_fat
        data[0x40..0x44].copy_from_slice(&(0u32).to_le_bytes()); // num_mini_fat
        data[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes()); // first_difat
        data[0x48..0x4C].copy_from_slice(&(0u32).to_le_bytes()); // num_difat
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
        let mut data = vec![0u8; 1536]; // 3 sectors of 512 after header
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
        // Build three directory entries of 128 bytes each:
        // root(storage) -> child index 1 ; entry 1 is stream "VBA" with right sibling 2 ("dir").
        let mut dir = vec![0u8; 128 * 3];

        // root entry
        dir[66] = 5; // root
        dir[68..72].copy_from_slice(&FREE_SECT.to_le_bytes()); // left
        dir[72..76].copy_from_slice(&FREE_SECT.to_le_bytes()); // right
        dir[76..80].copy_from_slice(&(1u32).to_le_bytes()); // child=1
        dir[116..120].copy_from_slice(&(0u32).to_le_bytes());
        dir[120..128].copy_from_slice(&(0u64).to_le_bytes());

        // entry 1: stream "VBA"
        let name1: Vec<u8> = "VBA"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .chain([0, 0])
            .collect();
        dir[128..128 + name1.len()].copy_from_slice(&name1);
        dir[128 + 64..128 + 66].copy_from_slice(&((name1.len()) as u16).to_le_bytes());
        dir[128 + 66] = 2; // stream
        dir[128 + 68..128 + 72].copy_from_slice(&FREE_SECT.to_le_bytes()); // left
        dir[128 + 72..128 + 76].copy_from_slice(&(2u32).to_le_bytes()); // right sibling
        dir[128 + 76..128 + 80].copy_from_slice(&FREE_SECT.to_le_bytes()); // child
        dir[128 + 116..128 + 120].copy_from_slice(&(1u32).to_le_bytes()); // start sector
        dir[128 + 120..128 + 128].copy_from_slice(&(10u64).to_le_bytes()); // size

        // entry 2: stream "dir"
        let name2: Vec<u8> = "dir"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .chain([0, 0])
            .collect();
        dir[256..256 + name2.len()].copy_from_slice(&name2);
        dir[256 + 64..256 + 66].copy_from_slice(&((name2.len()) as u16).to_le_bytes());
        dir[256 + 66] = 2; // stream
        dir[256 + 68..256 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[256 + 72..256 + 76].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[256 + 76..256 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());
        dir[256 + 116..256 + 120].copy_from_slice(&(2u32).to_le_bytes());
        dir[256 + 120..256 + 128].copy_from_slice(&(20u64).to_le_bytes());

        let entries = parse_dir_entries(&dir).expect("dir entries");
        assert_eq!(entries.len(), 3);
        assert_eq!(entries[1].name, "VBA");
        assert_eq!(entries[2].name, "dir");

        let mut out = HashMap::new();
        walk_siblings(entries[0].child, "", &entries, &mut out, 0);
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

        let entries = parse_dir_entries(&dir).expect("dir entries");
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

        // Sector 0 data starts at offset 512.
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

        // Backslash path should resolve via normalization.
        assert!(reader.contains("VBA\\Module1"));
        let err = reader
            .read_file_string("VBA\\Module1")
            .expect_err("invalid UTF-8 should fail");
        assert!(matches!(err, ParseError::Encoding(_)));
    }

    #[test]
    fn read_stream_from_mini_handles_out_of_bounds_and_chain_breaks() {
        // start sector points beyond mini stream bounds -> loop breaks and returns empty payload.
        let mini_stream = b"abcdEFGH".to_vec();
        let mini_fat = vec![END_OF_CHAIN];
        let out = read_stream_from_mini(&mini_stream, 8, &mini_fat, 10, 4).expect("mini stream");
        assert!(out.is_empty());

        // Chain points to index without entry in mini FAT -> unwrap_or(END_OF_CHAIN) path.
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

        let entries = parse_dir_entries(&dir).expect("dir entries");
        assert_eq!(entries[1].created_filetime, Some(123456));
        assert_eq!(entries[1].modified_filetime, Some(654321));
    }
}
