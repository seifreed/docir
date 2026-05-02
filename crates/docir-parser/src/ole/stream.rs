//! CFB stream reading utilities (FAT and mini-FAT).

use crate::error::ParseError;
use crate::ole_header::END_OF_CHAIN;

use super::types::MAX_STREAM_SIZE;

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
        if out.len() >= MAX_STREAM_SIZE {
            return Err(ParseError::ResourceLimit(
                "OLE stream exceeds maximum size".to_string(),
            ));
        }
        let sec = read_sector(data, sector_size, sector)?;
        out.extend_from_slice(&sec);
        let next = *fat.get(sector as usize).unwrap_or(&END_OF_CHAIN);
        sector = next;
        guard += 1;
    }
    Ok(out)
}

pub(crate) fn read_stream_from_mini(
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

pub(crate) fn collect_chain_with_terminal(table: &[u32], start_sector: u32) -> (Vec<u32>, u32) {
    let mut out = Vec::new();
    let mut sector = start_sector;
    let mut guard = 0usize;
    while sector != END_OF_CHAIN && sector != FREE_SECT {
        if guard >= table.len() {
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

use crate::ole_header::FREE_SECT;

pub(crate) fn read_u64(data: &[u8], offset: usize) -> Result<u64, ParseError> {
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

pub(crate) fn utf16le_to_string(bytes: &[u8]) -> String {
    let mut u16s = Vec::new();
    for chunk in bytes.chunks(2) {
        if chunk.len() == 2 {
            u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
    }
    String::from_utf16_lossy(&u16s)
}
