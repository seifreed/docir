//! XLSX metadata parsing helpers.

use crate::error::ParseError;
use crate::xml_utils::local_name;
use crate::xml_utils::reader_from_str;
use crate::xml_utils::xml_error;
use docir_core::ir::{SheetMetadata, SheetMetadataType};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;

pub(crate) fn parse_sheet_metadata(xml: &str, path: &str) -> Result<SheetMetadata, ParseError> {
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut metadata = SheetMetadata::new();
    metadata.span = Some(SourceSpan::new(path));

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                match local {
                    b"metadataType" => {
                        let mut mtype = SheetMetadataType::new();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key {
                                b"name" => mtype.name = Some(val),
                                b"minSupportedVersion" => mtype.min_supported_version = Some(val),
                                b"copy" => {
                                    mtype.copy =
                                        Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                }
                                b"update" => {
                                    mtype.update =
                                        Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                }
                                _ => {}
                            }
                        }
                        metadata.metadata_types.push(mtype);
                    }
                    b"cellMetadata" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"count" {
                                metadata.cell_metadata_count =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            }
                        }
                    }
                    b"valueMetadata" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"count" {
                                metadata.value_metadata_count =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(metadata)
}
