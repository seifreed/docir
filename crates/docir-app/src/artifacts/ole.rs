use docir_parser::ole::{Cfb, is_ole_container};

use super::helpers::empty_to_none;

#[derive(Debug, Clone)]
pub(super) struct EmbeddedPayload {
    pub stream_name: String,
    pub file_name: Option<String>,
    pub source_path: Option<String>,
    pub temp_path: Option<String>,
    pub data: Vec<u8>,
}

pub(super) fn extract_embedded_payload(data: &[u8]) -> Result<Option<EmbeddedPayload>, String> {
    if !is_ole_container(data) {
        return Ok(None);
    }
    let cfb = Cfb::parse(data.to_vec()).map_err(|err| err.to_string())?;
    for stream_name in ["\u{0001}Ole10Native", "Ole10Native", "Package"] {
        let stream = match cfb.try_read_stream(stream_name) {
            Ok(Some(stream)) => stream,
            Ok(None) => continue,
            Err(err) => return Err(err.to_string()),
        };
        if stream_name.contains("Ole10Native") {
            let payload = parse_ole10_native(&stream)
                .ok_or_else(|| format!("Malformed Ole10Native stream: {stream_name}"))?;
            return Ok(Some(EmbeddedPayload {
                stream_name: stream_name.to_string(),
                file_name: payload.file_name,
                source_path: payload.source_path,
                temp_path: payload.temp_path,
                data: payload.data,
            }));
        } else {
            return Ok(Some(EmbeddedPayload {
                stream_name: stream_name.to_string(),
                file_name: None,
                source_path: None,
                temp_path: None,
                data: stream,
            }));
        }
    }
    Ok(None)
}

pub(super) fn extract_embedded_payload_from_cfb(
    data: &[u8],
) -> Result<Option<EmbeddedPayload>, String> {
    if !is_ole_container(data) {
        return Ok(None);
    }
    extract_embedded_payload(data)
}

#[derive(Debug, Clone)]
pub(super) struct Ole10NativePayload {
    pub file_name: Option<String>,
    pub source_path: Option<String>,
    pub temp_path: Option<String>,
    pub data: Vec<u8>,
}

pub(super) fn parse_ole10_native(data: &[u8]) -> Option<Ole10NativePayload> {
    if data.len() < 6 {
        return None;
    }

    let mut offset = 4usize;
    offset = offset.checked_add(2)?;
    let file_name = read_c_string(data, &mut offset)?;
    let source_path = read_c_string(data, &mut offset)?;
    offset = offset.checked_add(8)?;
    let temp_path = read_c_string(data, &mut offset)?;
    let size_end = offset.checked_add(4)?;
    if size_end > data.len() {
        return None;
    }
    let size = u32::from_le_bytes(data[offset..size_end].try_into().ok()?) as usize;
    offset = size_end;
    let payload_end = offset.checked_add(size)?;
    if payload_end > data.len() {
        return None;
    }

    Some(Ole10NativePayload {
        file_name: empty_to_none(file_name),
        source_path: empty_to_none(source_path),
        temp_path: empty_to_none(temp_path),
        data: data[offset..payload_end].to_vec(),
    })
}

fn read_c_string(data: &[u8], offset: &mut usize) -> Option<String> {
    let start = *offset;
    while *offset < data.len() && data[*offset] != 0 {
        *offset += 1;
    }
    if *offset >= data.len() {
        return None;
    }
    let value = String::from_utf8_lossy(&data[start..*offset]).to_string();
    *offset += 1;
    Some(value)
}
