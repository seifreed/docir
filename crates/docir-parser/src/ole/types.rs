//! CFB data types and directory entry structures.

use std::collections::HashMap;

use crate::ole_header::{END_OF_CHAIN, FREE_SECT};

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
pub(crate) struct DirEntry {
    pub(crate) name: String,
    pub(crate) name_len_raw: u16,
    pub(crate) object_type: u8,
    pub(crate) color_flag: u8,
    pub(crate) left: u32,
    pub(crate) right: u32,
    pub(crate) child: u32,
    pub(crate) start_sector: u32,
    pub(crate) size: u64,
    pub(crate) created_filetime: Option<u64>,
    pub(crate) modified_filetime: Option<u64>,
}

pub(crate) fn entry_type_from_object_type(object_type: u8) -> Option<CfbEntryType> {
    match object_type {
        5 => Some(CfbEntryType::RootStorage),
        1 => Some(CfbEntryType::Storage),
        2 => Some(CfbEntryType::Stream),
        _ => None,
    }
}

pub(crate) fn normalize_tree_index(value: u32) -> Option<u32> {
    if value == FREE_SECT || value == END_OF_CHAIN {
        None
    } else {
        Some(value)
    }
}

pub(crate) fn normalize_filetime(value: u64) -> Option<u64> {
    if value == 0 {
        None
    } else {
        Some(value)
    }
}

pub(crate) const MAX_STREAM_SIZE: usize = 256 * 1024 * 1024;
pub(crate) const MAX_DIR_ENTRIES: usize = 65_536;
pub(crate) const MAX_RECURSION_DEPTH: u32 = 256;
pub(crate) const MAX_LINKED_DEPTH: u32 = 256;

pub(crate) fn collect_stream_entries(entries: &[DirEntry]) -> HashMap<String, DirEntry> {
    use std::collections::HashSet;
    let mut streams = HashMap::new();
    if let Some(child) = entries.first().map(|e| e.child) {
        let mut visited = HashSet::new();
        walk_siblings(child, "", entries, &mut streams, 0, &mut visited);
    }
    streams
}

pub(crate) fn walk_siblings(
    idx: u32,
    parent: &str,
    entries: &[DirEntry],
    out: &mut HashMap<String, DirEntry>,
    depth: u32,
    visited: &mut std::collections::HashSet<u32>,
) {
    if idx == FREE_SECT || idx == END_OF_CHAIN || depth > MAX_RECURSION_DEPTH {
        return;
    }
    if !visited.insert(idx) {
        return;
    }
    let idx_usize = idx as usize;
    if idx_usize >= entries.len() {
        return;
    }
    let entry = &entries[idx_usize];
    walk_siblings(entry.left, parent, entries, out, depth + 1, visited);
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
        walk_siblings(entry.child, &path, entries, out, depth + 1, visited);
    }
    walk_siblings(entry.right, parent, entries, out, depth + 1, visited);
}
