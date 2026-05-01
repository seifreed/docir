#[cfg(test)]
use crate::ole_header::{END_OF_CHAIN, FREE_SECT};

#[cfg(test)]
fn encode_name(name: &str) -> Vec<u8> {
    name.encode_utf16()
        .flat_map(|unit| unit.to_le_bytes())
        .chain([0, 0])
        .collect()
}

#[cfg(test)]
fn make_dir_entry(
    name: &str,
    object_type: u8,
    left: u32,
    right: u32,
    child: u32,
    start_sector: u32,
    size: u64,
) -> [u8; 128] {
    let mut entry = [0u8; 128];
    let name_utf16 = encode_name(name);
    entry[..name_utf16.len()].copy_from_slice(&name_utf16);
    entry[64..66].copy_from_slice(&(name_utf16.len() as u16).to_le_bytes());
    entry[66] = object_type;
    entry[68..72].copy_from_slice(&left.to_le_bytes());
    entry[72..76].copy_from_slice(&right.to_le_bytes());
    entry[76..80].copy_from_slice(&child.to_le_bytes());
    entry[116..120].copy_from_slice(&start_sector.to_le_bytes());
    entry[120..128].copy_from_slice(&size.to_le_bytes());
    entry
}

#[cfg(test)]
pub(crate) fn build_test_cfb(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let sector_size = 512usize;
    let mut data_streams = Vec::new();
    let mut entry_meta = Vec::new();

    for (path, bytes) in entries {
        let sectors_needed = bytes.len().max(1).div_ceil(sector_size);
        let start_sector = data_streams.len() as u32;
        for sector_idx in 0..sectors_needed {
            let start = sector_idx * sector_size;
            let end = (start + sector_size).min(bytes.len());
            let mut sector = vec![0u8; sector_size];
            if start < end {
                sector[..end - start].copy_from_slice(&bytes[start..end]);
            }
            data_streams.extend_from_slice(&sector);
        }
        entry_meta.push(((*path).to_string(), start_sector, bytes.len() as u64));
    }

    let mut dir_entries = Vec::new();
    dir_entries.push(make_dir_entry(
        "Root Entry",
        5,
        FREE_SECT,
        FREE_SECT,
        if entries.is_empty() { FREE_SECT } else { 1 },
        END_OF_CHAIN,
        0,
    ));

    for (idx, (path, start_sector, size)) in entry_meta.iter().enumerate() {
        let right = if idx + 1 < entry_meta.len() {
            (idx as u32) + 2
        } else {
            FREE_SECT
        };
        dir_entries.push(make_dir_entry(
            path.rsplit('/').next().unwrap_or(path),
            2,
            FREE_SECT,
            right,
            FREE_SECT,
            *start_sector,
            *size,
        ));
    }

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
    for (_, _, size) in entry_meta {
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
