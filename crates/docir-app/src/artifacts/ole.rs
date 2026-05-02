use docir_parser::ole::{is_ole_container, Cfb};

use super::helpers::empty_to_none;

#[derive(Debug, Clone)]
pub(super) struct EmbeddedPayload {
    pub stream_name: String,
    pub file_name: Option<String>,
    pub source_path: Option<String>,
    pub temp_path: Option<String>,
    pub data: Vec<u8>,
}

pub(super) fn extract_embedded_payload(data: &[u8]) -> Option<EmbeddedPayload> {
    if !is_ole_container(data) {
        return None;
    }
    let cfb = Cfb::parse(data.to_vec()).ok()?;
    for stream_name in ["\u{0001}Ole10Native", "Ole10Native", "Package"] {
        let Some(stream) = cfb.read_stream(stream_name) else {
            continue;
        };
        if stream_name.contains("Ole10Native") {
            if let Some(payload) = parse_ole10_native(&stream) {
                return Some(EmbeddedPayload {
                    stream_name: stream_name.to_string(),
                    file_name: payload.file_name,
                    source_path: payload.source_path,
                    temp_path: payload.temp_path,
                    data: payload.data,
                });
            }
        } else {
            return Some(EmbeddedPayload {
                stream_name: stream_name.to_string(),
                file_name: None,
                source_path: None,
                temp_path: None,
                data: stream,
            });
        }
    }
    None
}

pub(super) fn extract_embedded_payload_from_cfb(data: &[u8]) -> Option<EmbeddedPayload> {
    if !is_ole_container(data) {
        return None;
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
    if offset + 4 > data.len() {
        return None;
    }
    let size = u32::from_le_bytes(data[offset..offset + 4].try_into().ok()?) as usize;
    offset += 4;
    if offset + size > data.len() {
        return None;
    }

    Some(Ole10NativePayload {
        file_name: empty_to_none(file_name),
        source_path: empty_to_none(source_path),
        temp_path: empty_to_none(temp_path),
        data: data[offset..offset + size].to_vec(),
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
