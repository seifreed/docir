//! CFB directory parsing and metadata collection.

use std::collections::HashMap;
use std::collections::HashSet;

use crate::error::ParseError;
use crate::ole_header::{read_u16, read_u32, END_OF_CHAIN, FREE_SECT};

use super::stream::{read_stream_from_fat, read_u64, utf16le_to_string};
use super::types::{entry_type_from_object_type, normalize_tree_index};
use super::types::{
    CfbDirectorySlot, CfbDirectoryState, CfbEntryMetadata, CfbEntryType, DirEntry, MAX_DIR_ENTRIES,
    MAX_LINKED_DEPTH, MAX_RECURSION_DEPTH,
};

pub(crate) fn read_directory_entries_and_root_stream(
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

pub(crate) fn parse_dir_entries(data: &[u8]) -> Result<Vec<DirEntry>, ParseError> {
    let mut entries = Vec::new();
    for chunk in data.chunks(128) {
        if chunk.len() < 128 {
            break;
        }
        if entries.len() >= MAX_DIR_ENTRIES {
            return Err(ParseError::ResourceLimit(
                "OLE directory entry count exceeds maximum".to_string(),
            ));
        }
        let name_len_raw = read_u16(chunk, 64)?;
        let name_len = name_len_raw as usize;
        let name_raw = &chunk[..64];
        let name = if (2..=64).contains(&name_len) && name_len.is_multiple_of(2) {
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

use super::types::normalize_filetime;

pub(crate) fn collect_entry_metadata(entries: &[DirEntry]) -> HashMap<String, CfbEntryMetadata> {
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

pub(crate) fn collect_directory_slots(entries: &[DirEntry]) -> Vec<CfbDirectorySlot> {
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

fn collect_linked_indices(entries: &[DirEntry]) -> HashSet<u32> {
    let mut out = HashSet::new();
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

fn walk_linked_indices(idx: u32, entries: &[DirEntry], out: &mut HashSet<u32>, depth: u32) {
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
