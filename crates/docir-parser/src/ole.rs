//! Minimal OLE Compound File Binary (CFB) parser for VBA extraction.

use crate::error::ParseError;
use crate::zip_handler::PackageReader;
use std::collections::HashMap;

const SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const FREE_SECT: u32 = 0xFFFFFFFF;
const END_OF_CHAIN: u32 = 0xFFFFFFFE;
const FAT_SECT: u32 = 0xFFFFFFFD;
const DIFAT_SECT: u32 = 0xFFFFFFFC;

#[derive(Debug, Clone)]
struct DirEntry {
    name: String,
    object_type: u8,
    left: u32,
    right: u32,
    child: u32,
    start_sector: u32,
    size: u64,
}

/// Parsed CFB file with streams.
pub struct Cfb {
    sector_size: u32,
    mini_sector_size: u32,
    mini_cutoff: u32,
    fat: Vec<u32>,
    mini_fat: Vec<u32>,
    root_stream: Vec<u8>,
    entries: Vec<DirEntry>,
    streams: HashMap<String, DirEntry>,
    data: Vec<u8>,
}

impl Cfb {
    pub fn parse(data: Vec<u8>) -> Result<Self, ParseError> {
        if data.len() < 512 || data[..8] != SIGNATURE {
            return Err(ParseError::InvalidStructure(
                "Invalid OLE header".to_string(),
            ));
        }

        let sector_shift = read_u16(&data, 0x1E)? as u32;
        let mini_sector_shift = read_u16(&data, 0x20)? as u32;
        let sector_size = 1u32 << sector_shift;
        let mini_sector_size = 1u32 << mini_sector_shift;

        let num_fat_sectors = read_u32(&data, 0x2C)?;
        let first_dir_sector = read_u32(&data, 0x30)?;
        let mini_cutoff = read_u32(&data, 0x38)?;
        let first_mini_fat = read_u32(&data, 0x3C)?;
        let num_mini_fat = read_u32(&data, 0x40)?;
        let first_difat = read_u32(&data, 0x44)?;
        let num_difat = read_u32(&data, 0x48)?;

        let mut difat = Vec::new();
        for i in 0..109usize {
            let off = 0x4C + i * 4;
            let v = read_u32(&data, off)?;
            if v != FREE_SECT {
                difat.push(v);
            }
        }

        let mut next_difat = first_difat;
        for _ in 0..num_difat {
            if next_difat == END_OF_CHAIN || next_difat == FREE_SECT {
                break;
            }
            let sector = read_sector(&data, sector_size, next_difat)?;
            let count = (sector_size / 4) as usize - 1;
            for i in 0..count {
                let v = read_u32(&sector, i * 4)?;
                if v != FREE_SECT {
                    difat.push(v);
                }
            }
            next_difat = read_u32(&sector, count * 4)?;
        }

        let mut fat = Vec::new();
        for &fat_sector in difat.iter().take(num_fat_sectors as usize) {
            if fat_sector == FREE_SECT || fat_sector == END_OF_CHAIN || fat_sector == FAT_SECT {
                continue;
            }
            let sector = read_sector(&data, sector_size, fat_sector)?;
            for i in 0..(sector_size / 4) as usize {
                fat.push(read_u32(&sector, i * 4)?);
            }
        }

        let dir_stream = read_stream_from_fat(&data, sector_size, &fat, first_dir_sector)?;
        let entries = parse_dir_entries(&dir_stream)?;

        let root = entries
            .get(0)
            .ok_or_else(|| ParseError::InvalidStructure("Missing root entry".to_string()))?;
        let root_stream = read_stream_from_fat(&data, sector_size, &fat, root.start_sector)?;

        let mut mini_fat = Vec::new();
        if num_mini_fat > 0 && first_mini_fat != END_OF_CHAIN {
            let mini_fat_stream = read_stream_from_fat(&data, sector_size, &fat, first_mini_fat)?;
            for i in 0..(mini_fat_stream.len() / 4) {
                mini_fat.push(read_u32(&mini_fat_stream, i * 4)?);
            }
        }

        let mut streams = HashMap::new();
        if let Some(child) = entries.get(0).map(|e| e.child) {
            walk_siblings(child, "", &entries, &mut streams);
        }

        Ok(Self {
            sector_size,
            mini_sector_size,
            mini_cutoff,
            fat,
            mini_fat,
            root_stream,
            entries,
            streams,
            data,
        })
    }

    pub fn read_stream(&self, path: &str) -> Option<Vec<u8>> {
        let entry = self
            .streams
            .get(path)
            .or_else(|| self.streams.get(&path.replace('\\', "/")))?;
        if entry.object_type != 2 {
            return None;
        }

        if entry.size < self.mini_cutoff as u64 && !self.root_stream.is_empty() {
            return read_stream_from_mini(
                &self.root_stream,
                self.mini_sector_size,
                &self.mini_fat,
                entry.start_sector,
                entry.size as usize,
            );
        }

        let data =
            read_stream_from_fat(&self.data, self.sector_size, &self.fat, entry.start_sector)
                .ok()?;
        Some(data[..entry.size as usize].to_vec())
    }

    pub fn has_stream(&self, path: &str) -> bool {
        self.streams.contains_key(path) || self.streams.contains_key(&path.replace('\\', "/"))
    }

    pub fn list_streams(&self) -> Vec<String> {
        let mut keys: Vec<String> = self.streams.keys().cloned().collect();
        keys.sort();
        keys
    }

    pub fn stream_size(&self, path: &str) -> Option<u64> {
        self.streams
            .get(path)
            .or_else(|| self.streams.get(&path.replace('\\', "/")))
            .map(|entry| entry.size)
    }
}

pub struct CfbReader<'a> {
    cfb: &'a Cfb,
}

impl<'a> CfbReader<'a> {
    pub fn new(cfb: &'a Cfb) -> Self {
        Self { cfb }
    }
}

impl PackageReader for CfbReader<'_> {
    fn contains(&self, name: &str) -> bool {
        self.cfb.has_stream(name)
    }

    fn read_file_string(&mut self, name: &str) -> Result<String, ParseError> {
        let bytes = self
            .cfb
            .read_stream(name)
            .ok_or_else(|| ParseError::MissingPart(name.to_string()))?;
        String::from_utf8(bytes)
            .map_err(|e| ParseError::Encoding(format!("Invalid UTF-8 in {}: {}", name, e)))
    }

    fn file_names(&self) -> Vec<String> {
        self.cfb.list_streams()
    }
}

pub fn is_ole_container(data: &[u8]) -> bool {
    data.len() >= SIGNATURE.len() && data[..SIGNATURE.len()] == SIGNATURE
}

fn parse_dir_entries(data: &[u8]) -> Result<Vec<DirEntry>, ParseError> {
    let mut entries = Vec::new();
    for chunk in data.chunks(128) {
        if chunk.len() < 128 {
            break;
        }
        let name_len = read_u16(chunk, 64)? as usize;
        let name_raw = &chunk[..64];
        let name = if name_len >= 2 {
            let bytes = &name_raw[..name_len - 2];
            utf16le_to_string(bytes)
        } else {
            String::new()
        };
        let object_type = chunk[66];
        let left = read_u32(chunk, 68)?;
        let right = read_u32(chunk, 72)?;
        let child = read_u32(chunk, 76)?;
        let start_sector = read_u32(chunk, 116)?;
        let size = read_u64(chunk, 120)?;
        entries.push(DirEntry {
            name,
            object_type,
            left,
            right,
            child,
            start_sector,
            size,
        });
    }
    Ok(entries)
}

fn walk_siblings(
    idx: u32,
    parent: &str,
    entries: &[DirEntry],
    out: &mut HashMap<String, DirEntry>,
) {
    if idx == FREE_SECT || idx == END_OF_CHAIN {
        return;
    }
    let idx_usize = idx as usize;
    if idx_usize >= entries.len() {
        return;
    }
    let entry = &entries[idx_usize];
    walk_siblings(entry.left, parent, entries, out);
    let mut path = String::new();
    if !parent.is_empty() {
        path.push_str(parent);
        path.push('/');
    }
    path.push_str(&entry.name);
    if entry.object_type == 2 {
        out.insert(path.clone(), entry.clone());
    }
    if entry.object_type == 1 || entry.object_type == 5 {
        if entry.child != FREE_SECT {
            walk_siblings(entry.child, &path, entries, out);
        }
    }
    walk_siblings(entry.right, parent, entries, out);
}

fn read_stream_from_fat(
    data: &[u8],
    sector_size: u32,
    fat: &[u32],
    start_sector: u32,
) -> Result<Vec<u8>, ParseError> {
    let mut out = Vec::new();
    let mut sector = start_sector;
    let mut guard = 0usize;
    while sector != END_OF_CHAIN && sector != FREE_SECT {
        if guard > fat.len() {
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
        if guard > mini_fat.len() {
            break;
        }
        let offset = sector as usize * mini_sector_size as usize;
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

fn read_sector(data: &[u8], sector_size: u32, sector: u32) -> Result<Vec<u8>, ParseError> {
    let offset = (sector as usize + 1) * sector_size as usize;
    let end = offset + sector_size as usize;
    if end > data.len() {
        return Err(ParseError::InvalidStructure(
            "OLE sector out of bounds".to_string(),
        ));
    }
    Ok(data[offset..end].to_vec())
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16, ParseError> {
    if offset + 2 > data.len() {
        return Err(ParseError::InvalidStructure(
            "OLE read_u16 out of bounds".to_string(),
        ));
    }
    Ok(u16::from_le_bytes([data[offset], data[offset + 1]]))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32, ParseError> {
    if offset + 4 > data.len() {
        return Err(ParseError::InvalidStructure(
            "OLE read_u32 out of bounds".to_string(),
        ));
    }
    Ok(u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ]))
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
