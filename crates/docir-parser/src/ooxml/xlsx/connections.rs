//! XLSX connections, external links, and query tables.

use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::ooxml::xml_utils::xml_error;
use crate::xml_utils::local_name;
use docir_core::ir::{
    ConnectionEntry, ConnectionPart, ExternalLinkPart, ExternalLinkSheet, QueryTablePart,
    SlicerPart, TimelinePart,
};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;

pub(crate) fn parse_connections_part(xml: &str, path: &str) -> Result<ConnectionPart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut part = ConnectionPart::new();
    part.span = Some(SourceSpan::new(path));
    let mut current: Option<ConnectionEntry> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"connection" => {
                    let mut entry = ConnectionEntry::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"id" => {
                                entry.connection_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"name" => {
                                entry.name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"description" => {
                                entry.description =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"type" => {
                                entry.connection_type =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"refreshedVersion" => {
                                entry.refreshed_version =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"refreshOnLoad" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.refresh_on_load =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"saveData" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.save_data = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"background" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.background = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"sourceFile" => {
                                entry.source_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            b"odcFile" => {
                                entry.connection_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            _ => {}
                        }
                    }
                    current = Some(entry);
                }
                b"dbPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"connection" => {
                                    entry.connection =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"command" => {
                                    entry.command =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"commandType" => {
                                    entry.command_type =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                                }
                                _ => {}
                            }
                        }
                    }
                }
                b"webPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"url" {
                                entry.url = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                b"textPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"sourceFile" || attr.key.as_ref() == b"file" {
                                entry.source_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"connection" => {
                    let mut entry = ConnectionEntry::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"id" => {
                                entry.connection_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"name" => {
                                entry.name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"description" => {
                                entry.description =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"type" => {
                                entry.connection_type =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"refreshedVersion" => {
                                entry.refreshed_version =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"refreshOnLoad" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.refresh_on_load =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"saveData" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.save_data = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"background" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.background = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"sourceFile" => {
                                entry.source_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            b"odcFile" => {
                                entry.connection_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            _ => {}
                        }
                    }
                    part.entries.push(entry);
                }
                b"dbPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"connection" => {
                                    entry.connection =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"command" => {
                                    entry.command =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"commandType" => {
                                    entry.command_type =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                                }
                                _ => {}
                            }
                        }
                    }
                }
                b"webPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"url" {
                                entry.url = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                b"textPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"sourceFile" || attr.key.as_ref() == b"file" {
                                entry.source_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"connection" {
                    if let Some(entry) = current.take() {
                        part.entries.push(entry);
                    }
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

    Ok(part)
}

pub(crate) fn connection_targets(part: &ConnectionPart) -> Vec<String> {
    let mut targets = Vec::new();
    for entry in &part.entries {
        if let Some(value) = entry.connection.as_ref() {
            targets.push(value.clone());
        }
        if let Some(value) = entry.url.as_ref() {
            targets.push(value.clone());
        }
        if let Some(value) = entry.source_file.as_ref() {
            targets.push(value.clone());
        }
        if let Some(value) = entry.connection_file.as_ref() {
            targets.push(value.clone());
        }
    }
    targets.sort();
    targets.dedup();
    targets
}

pub(crate) fn parse_external_link_part(
    xml: &str,
    path: &str,
    rels: Option<&Relationships>,
) -> Result<ExternalLinkPart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut part = ExternalLinkPart::new();
    part.span = Some(SourceSpan::new(path));

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                match local {
                    b"externalLink" => {
                        // placeholder for type if present
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"linkType" || key == b"type" {
                                part.link_type =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    b"sheetNames" => {}
                    b"sheetName" => {
                        let mut sheet = ExternalLinkSheet {
                            name: None,
                            r_id: None,
                        };
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"val" || key == b"name" {
                                sheet.name = Some(val);
                            }
                        }
                        if let Some(name) = sheet.name {
                            part.sheets.push(ExternalLinkSheet {
                                name: Some(name),
                                r_id: None,
                            });
                        }
                    }
                    b"externalBook" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"id" || key == b"rid" || key == b"rId" {
                                let rel_id = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(rels) = rels {
                                    if let Some(rel) = rels.get(&rel_id) {
                                        part.target = Some(rel.target.clone());
                                        part.link_type = Some(rel.rel_type.clone());
                                    } else {
                                        part.target = Some(rel_id);
                                    }
                                } else {
                                    part.target = Some(rel_id);
                                }
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

    Ok(part)
}

pub(crate) fn parse_slicer_part(xml: &str, path: &str) -> Result<SlicerPart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut slicer = SlicerPart::new();
    slicer.span = Some(SourceSpan::new(path));

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"slicer" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"name" => slicer.name = Some(val),
                            b"caption" => slicer.caption = Some(val),
                            b"cache" | b"cacheId" => slicer.cache_id = Some(val),
                            b"ref" | b"pivotRef" => slicer.target_ref = Some(val),
                            _ => {}
                        }
                    }
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

    Ok(slicer)
}

pub(crate) fn parse_timeline_part(xml: &str, path: &str) -> Result<TimelinePart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut timeline = TimelinePart::new();
    timeline.span = Some(SourceSpan::new(path));

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                if local == b"timeline" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"name" => timeline.name = Some(val),
                            b"cache" | b"cacheId" => timeline.cache_id = Some(val),
                            _ => {}
                        }
                    }
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

    Ok(timeline)
}

pub(crate) fn parse_query_table_part(xml: &str, path: &str) -> Result<QueryTablePart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut query = QueryTablePart::new();
    query.span = Some(SourceSpan::new(path));

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                match local {
                    b"queryTable" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key {
                                b"name" => query.name = Some(val),
                                b"connectionId" | b"connection" => query.connection_id = Some(val),
                                _ => {}
                            }
                        }
                    }
                    b"dbPr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"command" {
                                query.command = Some(val.clone());
                            }
                            if key == b"connection" {
                                query.connection_id = Some(val);
                            }
                        }
                    }
                    b"webPr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"url" {
                                query.url = Some(val.clone());
                                query.source = Some(val);
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

    Ok(query)
}
