#[doc(hidden)]
pub enum TestPropertyValue {
    I16(i16),
    U16(u16),
    I32(i32),
    U32(u32),
    I64(i64),
    F64(f64),
    Bool(bool),
    Str(&'static str),
    WStr(&'static str),
    FileTime(u64),
}

#[derive(Debug, Clone, Copy, Default)]
#[doc(hidden)]
pub struct TestCfbDirectoryPatch {
    pub left_sibling_raw: Option<u32>,
    pub right_sibling_raw: Option<u32>,
    pub child_raw: Option<u32>,
    pub start_sector: Option<u32>,
}

#[doc(hidden)]
pub fn build_test_cfb(entries: &[(&str, &[u8])]) -> Vec<u8> {
    build_test_cfb_with_times(entries, &[])
}

#[doc(hidden)]
pub fn build_test_cfb_with_times(entries: &[(&str, &[u8])], times: &[(&str, u64, u64)]) -> Vec<u8> {
    const FREE_SECT: u32 = 0xFFFF_FFFF;
    const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
    let sector_size = 512usize;

    #[derive(Clone)]
    struct DirNode {
        path: String,
        name: String,
        object_type: u8,
        parent: Option<String>,
        start_sector: u32,
        size: u64,
        child: u32,
        right: u32,
        created_filetime: u64,
        modified_filetime: u64,
    }

    fn encode_name(name: &str) -> Vec<u8> {
        name.encode_utf16()
            .flat_map(|unit| unit.to_le_bytes())
            .chain([0, 0])
            .collect()
    }

    fn make_dir_entry(node: &DirNode) -> [u8; 128] {
        let mut entry = [0u8; 128];
        let name_utf16 = encode_name(&node.name);
        entry[..name_utf16.len()].copy_from_slice(&name_utf16);
        entry[64..66].copy_from_slice(&(name_utf16.len() as u16).to_le_bytes());
        entry[66] = node.object_type;
        entry[68..72].copy_from_slice(&FREE_SECT.to_le_bytes());
        entry[72..76].copy_from_slice(&node.right.to_le_bytes());
        entry[76..80].copy_from_slice(&node.child.to_le_bytes());
        entry[100..108].copy_from_slice(&node.created_filetime.to_le_bytes());
        entry[108..116].copy_from_slice(&node.modified_filetime.to_le_bytes());
        entry[116..120].copy_from_slice(&node.start_sector.to_le_bytes());
        entry[120..128].copy_from_slice(&node.size.to_le_bytes());
        entry
    }

    let times: std::collections::HashMap<String, (u64, u64)> = times
        .iter()
        .map(|(path, created, modified)| ((*path).to_string(), (*created, *modified)))
        .collect();

    let mut data_streams = Vec::new();
    let mut stream_meta = Vec::new();
    for (path, bytes) in entries {
        let sectors_needed = bytes.len().max(1).div_ceil(sector_size);
        let start_sector = data_streams.len() as u32 / sector_size as u32;
        for sector_idx in 0..sectors_needed {
            let start = sector_idx * sector_size;
            let end = (start + sector_size).min(bytes.len());
            let mut sector = vec![0u8; sector_size];
            if start < end {
                sector[..end - start].copy_from_slice(&bytes[start..end]);
            }
            data_streams.extend_from_slice(&sector);
        }
        stream_meta.push(((*path).to_string(), start_sector, bytes.len() as u64));
    }

    let mut storage_paths = std::collections::BTreeSet::new();
    for (path, _, _) in &stream_meta {
        let mut current = path.as_str();
        while let Some((parent, _)) = current.rsplit_once('/') {
            storage_paths.insert(parent.to_string());
            current = parent;
        }
    }

    let mut nodes = Vec::new();
    nodes.push(DirNode {
        path: "Root Entry".to_string(),
        name: "Root Entry".to_string(),
        object_type: 5,
        parent: None,
        start_sector: END_OF_CHAIN,
        size: 0,
        child: FREE_SECT,
        right: FREE_SECT,
        created_filetime: 0,
        modified_filetime: 0,
    });

    for path in storage_paths {
        let (created_filetime, modified_filetime) = times.get(&path).copied().unwrap_or((0, 0));
        let parent = path.rsplit_once('/').map(|(parent, _)| parent.to_string());
        nodes.push(DirNode {
            name: path.rsplit('/').next().unwrap_or(&path).to_string(),
            path,
            object_type: 1,
            parent,
            start_sector: END_OF_CHAIN,
            size: 0,
            child: FREE_SECT,
            right: FREE_SECT,
            created_filetime,
            modified_filetime,
        });
    }

    for (path, start_sector, size) in &stream_meta {
        let (created_filetime, modified_filetime) = times.get(path).copied().unwrap_or((0, 0));
        nodes.push(DirNode {
            name: path.rsplit('/').next().unwrap_or(path).to_string(),
            path: path.clone(),
            object_type: 2,
            parent: path.rsplit_once('/').map(|(parent, _)| parent.to_string()),
            start_sector: *start_sector,
            size: *size,
            child: FREE_SECT,
            right: FREE_SECT,
            created_filetime,
            modified_filetime,
        });
    }

    let mut children_by_parent: std::collections::HashMap<Option<String>, Vec<usize>> =
        std::collections::HashMap::new();
    for (idx, node) in nodes.iter().enumerate().skip(1) {
        children_by_parent
            .entry(node.parent.clone())
            .or_default()
            .push(idx);
    }

    for children in children_by_parent.values_mut() {
        children.sort_by(|left, right| nodes[*left].path.cmp(&nodes[*right].path));
        for pair in children.windows(2) {
            nodes[pair[0]].right = pair[1] as u32;
        }
    }

    if let Some(children) = children_by_parent.get(&None) {
        nodes[0].child = children
            .first()
            .copied()
            .map(|idx| idx as u32)
            .unwrap_or(FREE_SECT);
    }
    for node in nodes.iter_mut().skip(1) {
        if node.object_type != 1 {
            continue;
        }
        let key = Some(node.path.clone());
        if let Some(children) = children_by_parent.get(&key) {
            node.child = children
                .first()
                .copied()
                .map(|child| child as u32)
                .unwrap_or(FREE_SECT);
        }
    }

    let dir_entries: Vec<[u8; 128]> = nodes
        .into_iter()
        .map(|node| make_dir_entry(&node))
        .collect();

    let dir_sector_count = (dir_entries.len() * 128).div_ceil(sector_size).max(1);
    let fat_sector_index = data_streams.len() / sector_size + dir_sector_count;
    let dir_sector_start = data_streams.len() / sector_size;

    let mut full = vec![0u8; sector_size];
    full[..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    full[0x1E..0x20].copy_from_slice(&(9u16).to_le_bytes());
    full[0x20..0x22].copy_from_slice(&(6u16).to_le_bytes());
    full[0x2C..0x30].copy_from_slice(&(1u32).to_le_bytes());
    full[0x30..0x34].copy_from_slice(&(dir_sector_start as u32).to_le_bytes());
    full[0x38..0x3C].copy_from_slice(&(4096u32).to_le_bytes());
    full[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    full[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    full[0x4C..0x50].copy_from_slice(&(fat_sector_index as u32).to_le_bytes());
    full.extend_from_slice(&data_streams);

    let mut dir_stream = vec![0u8; dir_sector_count * sector_size];
    for (idx, entry) in dir_entries.into_iter().enumerate() {
        let start = idx * 128;
        dir_stream[start..start + 128].copy_from_slice(&entry);
    }
    full.extend_from_slice(&dir_stream);

    let total_data_sectors = data_streams.len() / sector_size;
    let total_sectors = total_data_sectors + dir_sector_count + 1;
    let mut fat_entries = vec![FREE_SECT; total_sectors];
    let mut stream_sector = 0usize;
    for (_, _, size) in stream_meta {
        let sectors_needed = (size as usize).max(1).div_ceil(sector_size);
        for offset in 0..sectors_needed {
            let sector = stream_sector + offset;
            fat_entries[sector] = if offset + 1 < sectors_needed {
                (sector + 1) as u32
            } else {
                END_OF_CHAIN
            };
        }
        stream_sector += sectors_needed;
    }
    for dir_offset in 0..dir_sector_count {
        let sector = total_data_sectors + dir_offset;
        fat_entries[sector] = if dir_offset + 1 < dir_sector_count {
            (sector + 1) as u32
        } else {
            END_OF_CHAIN
        };
    }
    fat_entries[fat_sector_index] = 0xFFFF_FFFD;

    let mut fat_sector = vec![0u8; sector_size];
    for (idx, entry) in fat_entries.into_iter().enumerate().take(sector_size / 4) {
        fat_sector[idx * 4..idx * 4 + 4].copy_from_slice(&entry.to_le_bytes());
    }
    full.extend_from_slice(&fat_sector);
    full
}

#[doc(hidden)]
pub fn patch_test_cfb_header_u32(bytes: &[u8], header_offset: usize, value: u32) -> Vec<u8> {
    let mut out = bytes.to_vec();
    out[header_offset..header_offset + 4].copy_from_slice(&value.to_le_bytes());
    out
}

#[doc(hidden)]
pub fn patch_test_cfb_fat_entry(bytes: &[u8], fat_index: u32, value: u32) -> Vec<u8> {
    let mut out = bytes.to_vec();
    let sector_size = 1usize << u16::from_le_bytes([out[0x1E], out[0x1F]]);
    let first_fat_sector = u32::from_le_bytes([out[0x4C], out[0x4D], out[0x4E], out[0x4F]]);
    let fat_offset = sector_size + first_fat_sector as usize * sector_size + fat_index as usize * 4;
    out[fat_offset..fat_offset + 4].copy_from_slice(&value.to_le_bytes());
    out
}

#[doc(hidden)]
pub fn patch_test_cfb_directory_entry(
    bytes: &[u8],
    entry_index: u32,
    patch: TestCfbDirectoryPatch,
) -> Vec<u8> {
    let mut out = bytes.to_vec();
    let sector_size = 1usize << u16::from_le_bytes([out[0x1E], out[0x1F]]);
    let dir_sector = u32::from_le_bytes([out[0x30], out[0x31], out[0x32], out[0x33]]);
    let base = sector_size + dir_sector as usize * sector_size + entry_index as usize * 128;
    if let Some(value) = patch.left_sibling_raw {
        out[base + 68..base + 72].copy_from_slice(&value.to_le_bytes());
    }
    if let Some(value) = patch.right_sibling_raw {
        out[base + 72..base + 76].copy_from_slice(&value.to_le_bytes());
    }
    if let Some(value) = patch.child_raw {
        out[base + 76..base + 80].copy_from_slice(&value.to_le_bytes());
    }
    if let Some(value) = patch.start_sector {
        out[base + 116..base + 120].copy_from_slice(&value.to_le_bytes());
    }
    out
}

#[doc(hidden)]
pub fn build_test_property_set_stream(properties: &[(u32, TestPropertyValue)]) -> Vec<u8> {
    const HEADER_SIZE: usize = 0x30;
    let section_offset = 0x30u32;
    let property_count = properties.len() as u32;
    let property_table_size = (property_count as usize) * 8;
    let section_header_size = 8usize;
    let values_base = section_header_size + property_table_size;

    let mut values = Vec::new();
    let mut prop_entries = Vec::new();

    for (property_id, value) in properties {
        let offset = (values_base + values.len()) as u32;
        prop_entries.push((*property_id, offset));
        match value {
            TestPropertyValue::I16(value) => {
                values.extend_from_slice(&2u32.to_le_bytes());
                values.extend_from_slice(&value.to_le_bytes());
                values.extend_from_slice(&0u16.to_le_bytes());
            }
            TestPropertyValue::U16(value) => {
                values.extend_from_slice(&18u32.to_le_bytes());
                values.extend_from_slice(&value.to_le_bytes());
                values.extend_from_slice(&0u16.to_le_bytes());
            }
            TestPropertyValue::I32(value) => {
                values.extend_from_slice(&3u32.to_le_bytes());
                values.extend_from_slice(&value.to_le_bytes());
            }
            TestPropertyValue::F64(value) => {
                values.extend_from_slice(&5u32.to_le_bytes());
                values.extend_from_slice(&value.to_le_bytes());
            }
            TestPropertyValue::Bool(value) => {
                values.extend_from_slice(&11u32.to_le_bytes());
                values.extend_from_slice(&(if *value { 0xFFFFu16 } else { 0u16 }).to_le_bytes());
                values.extend_from_slice(&0u16.to_le_bytes());
            }
            TestPropertyValue::U32(value) => {
                values.extend_from_slice(&19u32.to_le_bytes());
                values.extend_from_slice(&value.to_le_bytes());
            }
            TestPropertyValue::Str(value) => {
                let mut bytes = value.as_bytes().to_vec();
                bytes.push(0);
                values.extend_from_slice(&30u32.to_le_bytes());
                values.extend_from_slice(&(bytes.len() as u32).to_le_bytes());
                values.extend_from_slice(&bytes);
                pad_to_dword(&mut values);
            }
            TestPropertyValue::WStr(value) => {
                let mut utf16: Vec<u8> = value
                    .encode_utf16()
                    .flat_map(|unit| unit.to_le_bytes())
                    .collect();
                utf16.extend_from_slice(&[0, 0]);
                values.extend_from_slice(&31u32.to_le_bytes());
                values.extend_from_slice(&((utf16.len() / 2) as u32).to_le_bytes());
                values.extend_from_slice(&utf16);
                pad_to_dword(&mut values);
            }
            TestPropertyValue::FileTime(value) => {
                values.extend_from_slice(&64u32.to_le_bytes());
                values.extend_from_slice(&value.to_le_bytes());
            }
            TestPropertyValue::I64(value) => {
                values.extend_from_slice(&20u32.to_le_bytes());
                values.extend_from_slice(&value.to_le_bytes());
            }
        }
    }

    let section_size = (section_header_size + property_table_size + values.len()) as u32;
    let mut out = vec![0u8; HEADER_SIZE];
    out[0..2].copy_from_slice(&0xFFFEu16.to_le_bytes());
    out[2..4].copy_from_slice(&0u16.to_le_bytes());
    out[4..6].copy_from_slice(&0u16.to_le_bytes());
    out[24..28].copy_from_slice(&1u32.to_le_bytes());
    out[44..48].copy_from_slice(&section_offset.to_le_bytes());

    out.extend_from_slice(&section_size.to_le_bytes());
    out.extend_from_slice(&property_count.to_le_bytes());
    for (property_id, offset) in prop_entries {
        out.extend_from_slice(&property_id.to_le_bytes());
        out.extend_from_slice(&offset.to_le_bytes());
    }
    out.extend_from_slice(&values);
    out
}

fn pad_to_dword(values: &mut Vec<u8>) {
    while !values.len().is_multiple_of(4) {
        values.push(0);
    }
}
