//! XLSX metadata parsing helpers.

use crate::error::ParseError;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::reader_from_str;
use crate::xml_utils::{XmlScanControl, local_name, scan_xml_events, visit_attributes, xml_error};
use docir_core::ir::{SheetMetadata, SheetMetadataType};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;

pub(crate) fn parse_sheet_metadata(xml: &str, path: &str) -> Result<SheetMetadata, ParseError> {
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut metadata = SheetMetadata::new();
    metadata.span = Some(SourceSpan::new(path));

    scan_xml_events(&mut reader, &mut buf, path, |event| {
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                match local {
                    b"metadataType" => {
                        let mut mtype = SheetMetadataType::new();
                        visit_attributes(&e, path, |attr| {
                            let key = local_name(attr.key.as_ref());
                            let value = lossy_attr_value(attr).to_string();
                            match key {
                                b"name" => mtype.name = Some(value.clone()),
                                b"minSupportedVersion" => mtype.min_supported_version = Some(value),
                                b"copy" => {
                                    mtype.copy =
                                        Some(value == "1" || value.eq_ignore_ascii_case("true"));
                                }
                                b"update" => {
                                    mtype.update =
                                        Some(value == "1" || value.eq_ignore_ascii_case("true"));
                                }
                                _ => {}
                            }
                        })?;
                        metadata.metadata_types.push(mtype);
                    }
                    b"cellMetadata" => {
                        let mut numeric_error = None;
                        visit_attributes(&e, path, |attr| {
                            if numeric_error.is_some() {
                                return;
                            }
                            let key = local_name(attr.key.as_ref());
                            if key == b"count" {
                                match parse_u32_attr(&lossy_attr_value(attr), path) {
                                    Ok(parsed) => metadata.cell_metadata_count = Some(parsed),
                                    Err(err) => numeric_error = Some(err),
                                }
                            }
                        })?;
                        if let Some(err) = numeric_error {
                            return Err(err);
                        }
                    }
                    b"valueMetadata" => {
                        let mut numeric_error = None;
                        visit_attributes(&e, path, |attr| {
                            if numeric_error.is_some() {
                                return;
                            }
                            let key = local_name(attr.key.as_ref());
                            if key == b"count" {
                                match parse_u32_attr(&lossy_attr_value(attr), path) {
                                    Ok(parsed) => metadata.value_metadata_count = Some(parsed),
                                    Err(err) => numeric_error = Some(err),
                                }
                            }
                        })?;
                        if let Some(err) = numeric_error {
                            return Err(err);
                        }
                    }
                    _ => {}
                }
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(metadata)
}

fn parse_u32_attr(value: &str, path: &str) -> Result<u32, ParseError> {
    value.parse::<u32>().map_err(|err| xml_error(path, err))
}
