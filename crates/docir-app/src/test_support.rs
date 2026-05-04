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

const FREE_SECT: u32 = 0xFFFF_FFFF;
const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
const FAT_SECT: u32 = 0xFFFF_FFFD;
const SECTOR_SIZE: usize = 512;

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

struct StreamMeta {
    path: String,
    start_sector: u32,
    size: u64,
}

#[doc(hidden)]
pub fn build_test_cfb(entries: &[(&str, &[u8])]) -> Vec<u8> {
    build_test_cfb_with_times(entries, &[])
}

#[doc(hidden)]
pub fn build_test_cfb_with_times(entries: &[(&str, &[u8])], times: &[(&str, u64, u64)]) -> Vec<u8> {
    let time_map = time_map(times);
    let (data_streams, stream_meta) = build_stream_payload(entries);
    let mut nodes = build_directory_nodes(&stream_meta, &time_map);
    wire_directory_tree(&mut nodes);

    let dir_entries: Vec<[u8; 128]> = nodes.iter().map(make_dir_entry).collect();
    let dir_sector_count = (dir_entries.len() * 128).div_ceil(SECTOR_SIZE).max(1);
    let total_data_sectors = data_streams.len() / SECTOR_SIZE;
    let fat_sector_index = total_data_sectors + dir_sector_count;

    let mut full = cfb_header(total_data_sectors, fat_sector_index);
    full.extend_from_slice(&data_streams);
    full.extend_from_slice(&directory_stream(&dir_entries, dir_sector_count));

    let fat_entries = build_fat_entries(
        &stream_meta,
        total_data_sectors,
        dir_sector_count,
        fat_sector_index,
    );
    let mut fat_sector = vec![0u8; SECTOR_SIZE];
    for (idx, entry) in fat_entries.into_iter().enumerate().take(SECTOR_SIZE / 4) {
        fat_sector[idx * 4..idx * 4 + 4].copy_from_slice(&entry.to_le_bytes());
    }
    full.extend_from_slice(&fat_sector);
    full
}

fn time_map(times: &[(&str, u64, u64)]) -> std::collections::HashMap<String, (u64, u64)> {
    times
        .iter()
        .map(|(path, created, modified)| ((*path).to_string(), (*created, *modified)))
        .collect()
}

fn build_stream_payload(entries: &[(&str, &[u8])]) -> (Vec<u8>, Vec<StreamMeta>) {
    let mut data_streams = Vec::new();
    let mut stream_meta = Vec::new();
    for (path, bytes) in entries {
        let sectors_needed = bytes.len().max(1).div_ceil(SECTOR_SIZE);
        let start_sector = (data_streams.len() / SECTOR_SIZE) as u32;
        append_stream_sectors(&mut data_streams, bytes, sectors_needed);
        stream_meta.push(StreamMeta {
            path: (*path).to_string(),
            start_sector,
            size: bytes.len() as u64,
        });
    }
    (data_streams, stream_meta)
}

fn append_stream_sectors(data_streams: &mut Vec<u8>, bytes: &[u8], sectors_needed: usize) {
    for sector_idx in 0..sectors_needed {
        let start = sector_idx * SECTOR_SIZE;
        let end = (start + SECTOR_SIZE).min(bytes.len());
        let mut sector = vec![0u8; SECTOR_SIZE];
        if start < end {
            sector[..end - start].copy_from_slice(&bytes[start..end]);
        }
        data_streams.extend_from_slice(&sector);
    }
}

fn build_directory_nodes(
    stream_meta: &[StreamMeta],
    times: &std::collections::HashMap<String, (u64, u64)>,
) -> Vec<DirNode> {
    let mut nodes = vec![root_dir_node()];
    for path in storage_paths(stream_meta) {
        nodes.push(storage_dir_node(path, times));
    }
    for meta in stream_meta {
        nodes.push(stream_dir_node(meta, times));
    }
    nodes
}

fn storage_paths(stream_meta: &[StreamMeta]) -> std::collections::BTreeSet<String> {
    let mut paths = std::collections::BTreeSet::new();
    for meta in stream_meta {
        let mut current = meta.path.as_str();
        while let Some((parent, _)) = current.rsplit_once('/') {
            paths.insert(parent.to_string());
            current = parent;
        }
    }
    paths
}

fn root_dir_node() -> DirNode {
    DirNode {
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
    }
}

fn storage_dir_node(
    path: String,
    times: &std::collections::HashMap<String, (u64, u64)>,
) -> DirNode {
    let (created_filetime, modified_filetime) = times.get(&path).copied().unwrap_or((0, 0));
    let parent = path.rsplit_once('/').map(|(parent, _)| parent.to_string());
    DirNode {
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
    }
}

fn stream_dir_node(
    meta: &StreamMeta,
    times: &std::collections::HashMap<String, (u64, u64)>,
) -> DirNode {
    let (created_filetime, modified_filetime) = times.get(&meta.path).copied().unwrap_or((0, 0));
    DirNode {
        name: meta
            .path
            .rsplit('/')
            .next()
            .unwrap_or(&meta.path)
            .to_string(),
        path: meta.path.clone(),
        object_type: 2,
        parent: meta
            .path
            .rsplit_once('/')
            .map(|(parent, _)| parent.to_string()),
        start_sector: meta.start_sector,
        size: meta.size,
        child: FREE_SECT,
        right: FREE_SECT,
        created_filetime,
        modified_filetime,
    }
}

fn wire_directory_tree(nodes: &mut [DirNode]) {
    let mut children_by_parent: std::collections::HashMap<Option<String>, Vec<usize>> =
        std::collections::HashMap::new();
    for (idx, node) in nodes.iter().enumerate().skip(1) {
        children_by_parent
            .entry(node.parent.clone())
            .or_default()
            .push(idx);
    }
    wire_right_siblings(nodes, &mut children_by_parent);
    wire_child_pointers(nodes, &children_by_parent);
}

fn wire_right_siblings(
    nodes: &mut [DirNode],
    children_by_parent: &mut std::collections::HashMap<Option<String>, Vec<usize>>,
) {
    for children in children_by_parent.values_mut() {
        children.sort_by(|left, right| nodes[*left].path.cmp(&nodes[*right].path));
        for pair in children.windows(2) {
            nodes[pair[0]].right = pair[1] as u32;
        }
    }
}

fn wire_child_pointers(
    nodes: &mut [DirNode],
    children_by_parent: &std::collections::HashMap<Option<String>, Vec<usize>>,
) {
    nodes[0].child = first_child(children_by_parent.get(&None));
    for node in nodes.iter_mut().skip(1) {
        if node.object_type == 1 {
            let key = Some(node.path.clone());
            node.child = first_child(children_by_parent.get(&key));
        }
    }
}

fn first_child(children: Option<&Vec<usize>>) -> u32 {
    children
        .and_then(|items| items.first().copied())
        .map(|idx| idx as u32)
        .unwrap_or(FREE_SECT)
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

fn directory_stream(dir_entries: &[[u8; 128]], dir_sector_count: usize) -> Vec<u8> {
    let mut dir_stream = vec![0u8; dir_sector_count * SECTOR_SIZE];
    for (idx, entry) in dir_entries.iter().enumerate() {
        let start = idx * 128;
        dir_stream[start..start + 128].copy_from_slice(entry);
    }
    dir_stream
}

fn cfb_header(dir_sector_start: usize, fat_sector_index: usize) -> Vec<u8> {
    let mut header = vec![0u8; SECTOR_SIZE];
    header[..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    header[0x1E..0x20].copy_from_slice(&(9u16).to_le_bytes());
    header[0x20..0x22].copy_from_slice(&(6u16).to_le_bytes());
    header[0x2C..0x30].copy_from_slice(&(1u32).to_le_bytes());
    header[0x30..0x34].copy_from_slice(&(dir_sector_start as u32).to_le_bytes());
    header[0x38..0x3C].copy_from_slice(&(4096u32).to_le_bytes());
    header[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    header[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    header[0x4C..0x50].copy_from_slice(&(fat_sector_index as u32).to_le_bytes());
    header
}

fn build_fat_entries(
    stream_meta: &[StreamMeta],
    total_data_sectors: usize,
    dir_sector_count: usize,
    fat_sector_index: usize,
) -> Vec<u32> {
    let total_sectors = total_data_sectors + dir_sector_count + 1;
    let mut fat_entries = vec![FREE_SECT; total_sectors];
    wire_stream_fat_entries(&mut fat_entries, stream_meta);
    wire_directory_fat_entries(&mut fat_entries, total_data_sectors, dir_sector_count);
    fat_entries[fat_sector_index] = FAT_SECT;
    fat_entries
}

fn wire_stream_fat_entries(fat_entries: &mut [u32], stream_meta: &[StreamMeta]) {
    let mut stream_sector = 0usize;
    for meta in stream_meta {
        let sectors_needed = (meta.size as usize).max(1).div_ceil(SECTOR_SIZE);
        wire_sector_chain(fat_entries, stream_sector, sectors_needed);
        stream_sector += sectors_needed;
    }
}

fn wire_directory_fat_entries(
    fat_entries: &mut [u32],
    total_data_sectors: usize,
    dir_sector_count: usize,
) {
    wire_sector_chain(fat_entries, total_data_sectors, dir_sector_count);
}

fn wire_sector_chain(fat_entries: &mut [u32], start_sector: usize, sectors_needed: usize) {
    for offset in 0..sectors_needed {
        let sector = start_sector + offset;
        fat_entries[sector] = if offset + 1 < sectors_needed {
            (sector + 1) as u32
        } else {
            END_OF_CHAIN
        };
    }
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
