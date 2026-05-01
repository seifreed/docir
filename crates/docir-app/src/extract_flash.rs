use crate::{AppResult, ParserConfig};
use docir_parser::ole::{is_ole_container, Cfb};
use docir_parser::zip_handler::SecureZipReader;
use docir_parser::ParseError as ParserParseError;
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::fs;
use std::io::Cursor;
use std::path::Path;

#[derive(Debug, Clone, Serialize)]
pub struct FlashExtractionReport {
    pub container: String,
    pub object_count: usize,
    pub objects: Vec<FlashObject>,
}

#[derive(Debug, Clone, Serialize)]
pub struct FlashObject {
    pub source_path: String,
    pub offset: usize,
    pub signature: String,
    pub compression: String,
    pub version: u8,
    pub declared_size: u32,
    pub extracted_size: usize,
    pub truncated: bool,
    pub sha256: String,
    #[serde(skip_serializing)]
    pub data: Vec<u8>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub output_path: Option<String>,
}

pub fn extract_flash_path<P: AsRef<Path>>(
    path: P,
    config: &ParserConfig,
) -> AppResult<FlashExtractionReport> {
    let path = path.as_ref();
    let metadata = fs::metadata(path).map_err(ParserParseError::from)?;
    if metadata.len() > config.max_input_size {
        return Err(ParserParseError::ResourceLimit(format!(
            "Input exceeds max_input_size ({} > {})",
            metadata.len(),
            config.max_input_size
        ))
        .into());
    }
    let bytes = fs::read(path).map_err(ParserParseError::from)?;
    extract_flash_bytes(&bytes, config)
}

pub fn extract_flash_bytes(data: &[u8], config: &ParserConfig) -> AppResult<FlashExtractionReport> {
    if is_ole_container(data) {
        return extract_flash_from_cfb(data);
    }
    if looks_like_zip(data) {
        return extract_flash_from_zip(data, config);
    }
    Ok(FlashExtractionReport {
        container: "raw-binary".to_string(),
        object_count: 0,
        objects: find_flash_objects_in_bytes(data, "input"),
    }
    .with_count())
}

impl FlashExtractionReport {
    fn with_count(mut self) -> Self {
        self.object_count = self.objects.len();
        self
    }
}

fn extract_flash_from_cfb(data: &[u8]) -> AppResult<FlashExtractionReport> {
    let cfb = Cfb::parse(data.to_vec())?;
    let mut objects = Vec::new();
    for path in cfb.list_streams() {
        if let Some(bytes) = cfb.read_stream(&path) {
            objects.extend(find_flash_objects_in_bytes(&bytes, &path));
        }
    }
    Ok(FlashExtractionReport {
        container: "cfb".to_string(),
        object_count: objects.len(),
        objects,
    })
}

fn extract_flash_from_zip(data: &[u8], config: &ParserConfig) -> AppResult<FlashExtractionReport> {
    let mut zip = SecureZipReader::new(Cursor::new(data), config.zip_config.clone())?;
    let mut objects = Vec::new();
    let file_names: Vec<String> = zip.file_names().map(|name| name.to_string()).collect();
    for path in file_names {
        let Ok(bytes) = zip.read_file(&path) else {
            continue;
        };
        objects.extend(find_flash_objects_in_bytes(&bytes, &path));
    }
    Ok(FlashExtractionReport {
        container: "zip".to_string(),
        object_count: objects.len(),
        objects,
    })
}

fn looks_like_zip(data: &[u8]) -> bool {
    data.starts_with(b"PK\x03\x04")
        || data.starts_with(b"PK\x05\x06")
        || data.starts_with(b"PK\x07\x08")
}

fn find_flash_objects_in_bytes(data: &[u8], source_path: &str) -> Vec<FlashObject> {
    let mut objects = Vec::new();
    let mut offset = 0usize;
    while offset + 8 <= data.len() {
        if let Some(signature) = swf_signature(&data[offset..]) {
            let version = data[offset + 3];
            let declared_size = u32::from_le_bytes([
                data[offset + 4],
                data[offset + 5],
                data[offset + 6],
                data[offset + 7],
            ]);
            if version == 0 || declared_size < 8 {
                offset += 1;
                continue;
            }
            let available = data.len().saturating_sub(offset);
            let extracted_size = (declared_size as u64).min(available as u64) as usize;
            let payload = &data[offset..offset + extracted_size];
            let truncated = (declared_size as u64) > (available as u64);
            let compression = match signature {
                "FWS" => "none",
                "CWS" => "zlib",
                "ZWS" => "lzma",
                _ => "unknown",
            };
            let sha256 = format!("{:x}", Sha256::digest(payload));
            objects.push(FlashObject {
                source_path: source_path.to_string(),
                offset,
                signature: signature.to_string(),
                compression: compression.to_string(),
                version,
                declared_size,
                extracted_size,
                truncated,
                sha256,
                data: payload.to_vec(),
                output_path: None,
            });
            offset += extracted_size.max(1);
            continue;
        }
        offset += 1;
    }
    objects
}

fn swf_signature(data: &[u8]) -> Option<&'static str> {
    if data.starts_with(b"FWS") {
        Some("FWS")
    } else if data.starts_with(b"CWS") {
        Some("CWS")
    } else if data.starts_with(b"ZWS") {
        Some("ZWS")
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::{extract_flash_bytes, find_flash_objects_in_bytes};
    use crate::test_support::build_test_cfb;
    use crate::ParserConfig;
    use std::io::Write;

    fn swf(signature: &[u8; 3], version: u8, body: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(signature);
        out.push(version);
        out.extend_from_slice(&((body.len() + 8) as u32).to_le_bytes());
        out.extend_from_slice(body);
        out
    }

    #[test]
    fn find_flash_objects_in_bytes_detects_embedded_swf() {
        let mut data = b"prefix".to_vec();
        data.extend_from_slice(&swf(b"FWS", 9, b"payload"));
        let objects = find_flash_objects_in_bytes(&data, "blob");
        assert_eq!(objects.len(), 1);
        assert_eq!(objects[0].signature, "FWS");
        assert_eq!(objects[0].offset, 6);
        assert_eq!(objects[0].compression, "none");
    }

    #[test]
    fn extract_flash_bytes_scans_cfb_streams() {
        let bytes = build_test_cfb(&[("ObjectPool/1/Ole10Native", &swf(b"CWS", 10, b"payload"))]);
        let report = extract_flash_bytes(&bytes, &ParserConfig::default()).expect("report");
        assert_eq!(report.container, "cfb");
        assert_eq!(report.object_count, 1);
        assert_eq!(report.objects[0].signature, "CWS");
    }

    #[test]
    fn extract_flash_bytes_marks_truncated_payload() {
        let mut payload = swf(b"FWS", 9, b"short");
        payload[4..8].copy_from_slice(&100u32.to_le_bytes());
        let report = extract_flash_bytes(&payload, &ParserConfig::default()).expect("report");
        assert_eq!(report.container, "raw-binary");
        assert_eq!(report.object_count, 1);
        assert!(report.objects[0].truncated);
    }

    #[test]
    fn extract_flash_bytes_scans_zip_entries() {
        let mut cursor = std::io::Cursor::new(Vec::<u8>::new());
        {
            let mut zip = zip::ZipWriter::new(&mut cursor);
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("word/media/movie.bin", options)
                .expect("start");
            zip.write_all(&swf(b"FWS", 8, b"payload")).expect("write");
            zip.finish().expect("finish");
        }
        let report =
            extract_flash_bytes(cursor.get_ref(), &ParserConfig::default()).expect("report");
        assert_eq!(report.container, "zip");
        assert_eq!(report.object_count, 1);
        assert_eq!(report.objects[0].source_path, "word/media/movie.bin");
    }

    #[test]
    fn extract_flash_bytes_detects_multiple_signatures() {
        let mut data = swf(b"FWS", 8, b"one");
        data.extend_from_slice(b"pad");
        data.extend_from_slice(&swf(b"ZWS", 13, b"two"));
        let report = extract_flash_bytes(&data, &ParserConfig::default()).expect("report");
        assert_eq!(report.object_count, 2);
        assert_eq!(report.objects[0].signature, "FWS");
        assert_eq!(report.objects[1].signature, "ZWS");
    }

    #[test]
    fn extract_flash_bytes_marks_truncated_cws_inside_zip_entry() {
        let mut payload = swf(b"CWS", 10, b"payload");
        payload[4..8].copy_from_slice(&128u32.to_le_bytes());
        let mut cursor = std::io::Cursor::new(Vec::<u8>::new());
        {
            let mut zip = zip::ZipWriter::new(&mut cursor);
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("ppt/media/flash.bin", options)
                .expect("start");
            zip.write_all(&payload).expect("write");
            zip.finish().expect("finish");
        }

        let report =
            extract_flash_bytes(cursor.get_ref(), &ParserConfig::default()).expect("report");
        assert_eq!(report.container, "zip");
        assert_eq!(report.object_count, 1);
        assert_eq!(report.objects[0].signature, "CWS");
        assert!(report.objects[0].truncated);
        assert_eq!(report.objects[0].source_path, "ppt/media/flash.bin");
    }
}
