//! XLSX connections, external links, and query tables.

use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{attr_bool, attr_u32, attr_value, local_name};
use crate::xml_utils::{scan_xml_events, XmlScanControl};
use docir_core::ir::{
    ConnectionEntry, ConnectionPart, ExternalLinkPart, ExternalLinkSheet, QueryTablePart,
    SlicerPart, TimelinePart,
};
use docir_core::types::SourceSpan;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

fn for_each_local_start_or_empty<F>(
    xml: &str,
    path: &str,
    mut on_event: F,
) -> Result<(), ParseError>
where
    F: FnMut(&[u8], &BytesStart<'_>),
{
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    scan_xml_events(&mut reader, &mut buf, path, |event| {
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                on_event(local, &e);
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(())
}

pub(crate) fn parse_connections_part(xml: &str, path: &str) -> Result<ConnectionPart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut part = ConnectionPart::new();
    part.span = Some(SourceSpan::new(path));
    let mut current: Option<ConnectionEntry> = None;

    scan_xml_events(&mut reader, &mut buf, path, |event| {
        match event {
            Event::Start(e) => {
                if e.name().as_ref() == b"connection" {
                    current = Some(connection_entry_from_attrs(&e));
                } else {
                    apply_connection_child_attrs(&mut current, &e);
                }
            }
            Event::Empty(e) => {
                if e.name().as_ref() == b"connection" {
                    part.entries.push(connection_entry_from_attrs(&e));
                } else {
                    apply_connection_child_attrs(&mut current, &e);
                }
            }
            Event::End(e) => {
                if e.name().as_ref() == b"connection" {
                    if let Some(entry) = current.take() {
                        part.entries.push(entry);
                    }
                }
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(part)
}

fn connection_entry_from_attrs(e: &BytesStart<'_>) -> ConnectionEntry {
    let mut entry = ConnectionEntry::new();
    apply_connection_attrs(&mut entry, e);
    entry
}

fn apply_connection_attrs(entry: &mut ConnectionEntry, e: &BytesStart<'_>) {
    entry.connection_id = attr_u32(e, b"id");
    entry.name = attr_value(e, b"name");
    entry.description = attr_value(e, b"description");
    entry.connection_type = attr_u32(e, b"type");
    entry.refreshed_version = attr_u32(e, b"refreshedVersion");
    entry.refresh_on_load = attr_bool(e, b"refreshOnLoad");
    entry.save_data = attr_bool(e, b"saveData");
    entry.background = attr_bool(e, b"background");
    entry.source_file = attr_value(e, b"sourceFile");
    entry.connection_file = attr_value(e, b"odcFile");
}

fn apply_dbpr_attrs(entry: &mut ConnectionEntry, e: &BytesStart<'_>) {
    entry.connection = attr_value(e, b"connection");
    entry.command = attr_value(e, b"command");
    entry.command_type = attr_u32(e, b"commandType");
}

fn apply_textpr_attrs(entry: &mut ConnectionEntry, e: &BytesStart<'_>) {
    entry.source_file = attr_value(e, b"sourceFile").or_else(|| attr_value(e, b"file"));
}

fn apply_connection_child_attrs(current: &mut Option<ConnectionEntry>, e: &BytesStart<'_>) {
    let Some(entry) = current.as_mut() else {
        return;
    };
    match e.name().as_ref() {
        b"dbPr" => apply_dbpr_attrs(entry, e),
        b"webPr" => {
            if let Some(url) = attr_value(e, b"url") {
                entry.url = Some(url);
            }
        }
        b"textPr" => apply_textpr_attrs(entry, e),
        _ => {}
    }
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
    let mut part = ExternalLinkPart::new();
    part.span = Some(SourceSpan::new(path));

    for_each_local_start_or_empty(xml, path, |local, e| match local {
        b"externalLink" => {
            // placeholder for type if present
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                if key == b"linkType" || key == b"type" {
                    part.link_type = Some(lossy_attr_value(&attr).to_string());
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
                let value = lossy_attr_value(&attr).to_string();
                if key == b"val" || key == b"name" {
                    sheet.name = Some(value);
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
                    let rel_id = lossy_attr_value(&attr).to_string();
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
    })?;

    Ok(part)
}

pub(crate) fn parse_slicer_part(xml: &str, path: &str) -> Result<SlicerPart, ParseError> {
    let mut slicer = SlicerPart::new();
    slicer.span = Some(SourceSpan::new(path));

    for_each_local_start_or_empty(xml, path, |local, e| {
        if local == b"slicer" {
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                let value = lossy_attr_value(&attr).to_string();
                match key {
                    b"name" => slicer.name = Some(value),
                    b"caption" => slicer.caption = Some(value),
                    b"cache" | b"cacheId" => slicer.cache_id = Some(value),
                    b"ref" | b"pivotRef" => slicer.target_ref = Some(value),
                    _ => {}
                }
            }
        }
    })?;

    Ok(slicer)
}

pub(crate) fn parse_timeline_part(xml: &str, path: &str) -> Result<TimelinePart, ParseError> {
    let mut timeline = TimelinePart::new();
    timeline.span = Some(SourceSpan::new(path));

    for_each_local_start_or_empty(xml, path, |local, e| {
        if local == b"timeline" {
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                let value = lossy_attr_value(&attr).to_string();
                match key {
                    b"name" => timeline.name = Some(value),
                    b"cache" | b"cacheId" => timeline.cache_id = Some(value),
                    _ => {}
                }
            }
        }
    })?;

    Ok(timeline)
}

pub(crate) fn parse_query_table_part(xml: &str, path: &str) -> Result<QueryTablePart, ParseError> {
    let mut query = QueryTablePart::new();
    query.span = Some(SourceSpan::new(path));

    for_each_local_start_or_empty(xml, path, |local, e| match local {
        b"queryTable" => {
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                let value = lossy_attr_value(&attr).to_string();
                match key {
                    b"name" => query.name = Some(value),
                    b"connectionId" | b"connection" => query.connection_id = Some(value),
                    _ => {}
                }
            }
        }
        b"dbPr" => {
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                let value = lossy_attr_value(&attr).to_string();
                if key == b"command" {
                    query.command = Some(value.clone());
                }
                if key == b"connection" {
                    query.connection_id = Some(value);
                }
            }
        }
        b"webPr" => {
            for attr in e.attributes().flatten() {
                let key = local_name(attr.key.as_ref());
                let value = lossy_attr_value(&attr).to_string();
                if key == b"url" {
                    query.url = Some(value.clone());
                    query.source = Some(value);
                }
            }
        }
        _ => {}
    })?;

    Ok(query)
}
