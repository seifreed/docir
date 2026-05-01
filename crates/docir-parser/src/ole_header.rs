use crate::error::ParseError;

pub(crate) const SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
pub(crate) const FREE_SECT: u32 = 0xFFFFFFFF;
pub(crate) const END_OF_CHAIN: u32 = 0xFFFFFFFE;
pub(crate) const FAT_SECT: u32 = 0xFFFFFFFD;

pub(crate) struct OleHeader {
    pub(crate) sector_size: u32,
    pub(crate) mini_sector_size: u32,
    pub(crate) num_fat_sectors: u32,
    pub(crate) first_dir_sector: u32,
    pub(crate) mini_cutoff: u32,
    pub(crate) first_mini_fat: u32,
    pub(crate) num_mini_fat: u32,
    pub(crate) first_difat: u32,
    pub(crate) num_difat: u32,
}

pub(crate) fn parse_header(data: &[u8]) -> Result<OleHeader, ParseError> {
    if data.len() < 512 || data[..8] != SIGNATURE {
        return Err(ParseError::InvalidStructure(
            "Invalid OLE header".to_string(),
        ));
    }
    let sector_shift = read_u16(data, 0x1E)? as u32;
    let mini_sector_shift = read_u16(data, 0x20)? as u32;
    if sector_shift >= 32 || mini_sector_shift >= 32 {
        return Err(ParseError::InvalidStructure(
            "OLE header sector shift overflow".to_string(),
        ));
    }
    // CFB spec: sector_shift must be 9 (v3) or 12 (v4); mini_sector_shift must be 6.
    if sector_shift < 9 {
        return Err(ParseError::InvalidStructure(
            "OLE header sector shift too small (minimum 9)".to_string(),
        ));
    }
    if mini_sector_shift < 6 {
        return Err(ParseError::InvalidStructure(
            "OLE header mini sector shift too small (minimum 6)".to_string(),
        ));
    }
    Ok(OleHeader {
        sector_size: 1u32 << sector_shift,
        mini_sector_size: 1u32 << mini_sector_shift,
        num_fat_sectors: read_u32(data, 0x2C)?,
        first_dir_sector: read_u32(data, 0x30)?,
        mini_cutoff: read_u32(data, 0x38)?,
        first_mini_fat: read_u32(data, 0x3C)?,
        num_mini_fat: read_u32(data, 0x40)?,
        first_difat: read_u32(data, 0x44)?,
        num_difat: read_u32(data, 0x48)?,
    })
}

pub(crate) fn read_difat_chain(data: &[u8], header: &OleHeader) -> Result<Vec<u32>, ParseError> {
    let mut difat = Vec::new();
    for i in 0..109usize {
        let off = 0x4C + i * 4;
        let v = read_u32(data, off)?;
        if v != FREE_SECT {
            difat.push(v);
        }
    }

    let mut next_difat = header.first_difat;
    for _ in 0..header.num_difat {
        if next_difat == END_OF_CHAIN || next_difat == FREE_SECT {
            break;
        }
        let sector = crate::ole::read_sector(data, header.sector_size, next_difat)?;
        let count = (header.sector_size / 4) as usize - 1;
        for i in 0..count {
            let v = read_u32(&sector, i * 4)?;
            if v != FREE_SECT {
                difat.push(v);
            }
        }
        next_difat = read_u32(&sector, count * 4)?;
    }
    Ok(difat)
}

pub(crate) fn read_fat_table(
    data: &[u8],
    sector_size: u32,
    difat: &[u32],
    num_fat_sectors: u32,
) -> Result<Vec<u32>, ParseError> {
    let mut fat = Vec::new();
    for &fat_sector in difat.iter().take(num_fat_sectors as usize) {
        if fat_sector == FREE_SECT || fat_sector == END_OF_CHAIN || fat_sector == FAT_SECT {
            continue;
        }
        let sector = crate::ole::read_sector(data, sector_size, fat_sector)?;
        for i in 0..(sector_size / 4) as usize {
            fat.push(read_u32(&sector, i * 4)?);
        }
    }
    Ok(fat)
}

pub(crate) fn read_mini_fat_table(
    data: &[u8],
    sector_size: u32,
    fat: &[u32],
    first_mini_fat: u32,
    num_mini_fat: u32,
) -> Result<Vec<u32>, ParseError> {
    if num_mini_fat == 0 || first_mini_fat == END_OF_CHAIN {
        return Ok(Vec::new());
    }
    let mini_fat_stream = crate::ole::read_stream_from_fat(data, sector_size, fat, first_mini_fat)?;
    let mut mini_fat = Vec::new();
    for i in 0..(mini_fat_stream.len() / 4) {
        mini_fat.push(read_u32(&mini_fat_stream, i * 4)?);
    }
    Ok(mini_fat)
}

pub(crate) fn read_u16(data: &[u8], offset: usize) -> Result<u16, ParseError> {
    if offset + 2 > data.len() {
        return Err(ParseError::InvalidStructure(
            "OLE read_u16 out of bounds".to_string(),
        ));
    }
    Ok(u16::from_le_bytes([data[offset], data[offset + 1]]))
}

pub(crate) fn read_u32(data: &[u8], offset: usize) -> Result<u32, ParseError> {
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
