//! XLSX connections, external links, and query tables.

use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{XmlScanControl, scan_xml_events, visit_attributes};
use crate::xml_utils::{attr_bool_like, local_name};
use docir_core::ir::{
    ConnectionEntry, ConnectionPart, ExternalLinkPart, ExternalLinkSheet, QueryTablePart,
    SlicerPart, TimelinePart,
};
use docir_core::types::SourceSpan;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

fn for_each_local_start_or_empty<F>(
    xml: &str,
    path: &str,
    mut on_event: F,
) -> Result<(), ParseError>
where
    F: FnMut(&[u8], &BytesStart<'_>) -> Result<(), ParseError>,
{
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    scan_xml_events(&mut reader, &mut buf, path, |event| {
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = local_name(&name_buf);
                on_event(local, &e)?;
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
                if local_name(e.name().as_ref()) == b"connection" {
                    current = Some(connection_entry_from_attrs(&e, path)?);
                } else {
                    apply_connection_child_attrs(&mut current, &e, path)?;
                }
            }
            Event::Empty(e) => {
                if local_name(e.name().as_ref()) == b"connection" {
                    part.entries.push(connection_entry_from_attrs(&e, path)?);
                } else {
                    apply_connection_child_attrs(&mut current, &e, path)?;
                }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == b"connection"
                    && let Some(entry) = current.take()
                {
                    part.entries.push(entry);
                }
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    Ok(part)
}

fn connection_entry_from_attrs(
    e: &BytesStart<'_>,
    path: &str,
) -> Result<ConnectionEntry, ParseError> {
    let mut entry = ConnectionEntry::new();
    apply_connection_attrs(&mut entry, e, path)?;
    Ok(entry)
}

fn apply_connection_attrs(
    entry: &mut ConnectionEntry,
    e: &BytesStart<'_>,
    path: &str,
) -> Result<(), ParseError> {
    visit_attributes(e, path, |attr| {
        let key = local_name(attr.key.as_ref());
        let value = lossy_attr_value(attr);
        match key {
            b"id" => entry.connection_id = value.parse::<u32>().ok(),
            b"name" => entry.name = Some(value.into_owned()),
            b"description" => entry.description = Some(value.into_owned()),
            b"type" => entry.connection_type = value.parse::<u32>().ok(),
            b"refreshedVersion" => entry.refreshed_version = value.parse::<u32>().ok(),
            b"refreshOnLoad" => entry.refresh_on_load = Some(attr_bool_like(attr.value.as_ref())),
            b"saveData" => entry.save_data = Some(attr_bool_like(attr.value.as_ref())),
            b"background" => entry.background = Some(attr_bool_like(attr.value.as_ref())),
            b"sourceFile" => entry.source_file = Some(value.into_owned()),
            b"odcFile" => entry.connection_file = Some(value.into_owned()),
            _ => {}
        }
    })
}

fn apply_dbpr_attrs(
    entry: &mut ConnectionEntry,
    e: &BytesStart<'_>,
    path: &str,
) -> Result<(), ParseError> {
    visit_attributes(e, path, |attr| {
        let key = local_name(attr.key.as_ref());
        let value = lossy_attr_value(attr);
        match key {
            b"connection" => entry.connection = Some(value.into_owned()),
            b"command" => entry.command = Some(value.into_owned()),
            b"commandType" => entry.command_type = value.parse::<u32>().ok(),
            _ => {}
        }
    })
}

fn apply_textpr_attrs(
    entry: &mut ConnectionEntry,
    e: &BytesStart<'_>,
    path: &str,
) -> Result<(), ParseError> {
    let mut source_file = None;
    let mut file = None;
    visit_attributes(e, path, |attr| {
        let key = local_name(attr.key.as_ref());
        let value = lossy_attr_value(attr);
        match key {
            b"sourceFile" => source_file = Some(value.into_owned()),
            b"file" => file = Some(value.into_owned()),
            _ => {}
        }
    })?;
    entry.source_file = source_file.or(file);
    Ok(())
}

fn apply_connection_child_attrs(
    current: &mut Option<ConnectionEntry>,
    e: &BytesStart<'_>,
    path: &str,
) -> Result<(), ParseError> {
    let Some(entry) = current.as_mut() else {
        return Ok(());
    };
    match local_name(e.name().as_ref()) {
        b"dbPr" => apply_dbpr_attrs(entry, e, path)?,
        b"webPr" => {
            visit_attributes(e, path, |attr| {
                if local_name(attr.key.as_ref()) == b"url" {
                    entry.url = Some(lossy_attr_value(attr).into_owned());
                }
            })?;
        }
        b"textPr" => apply_textpr_attrs(entry, e, path)?,
        _ => {}
    }
    Ok(())
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

    for_each_local_start_or_empty(xml, path, |local, e| {
        match local {
            b"externalLink" => {
                // placeholder for type if present
                visit_attributes(e, path, |attr| {
                    let key = local_name(attr.key.as_ref());
                    if key == b"linkType" || key == b"type" {
                        part.link_type = Some(lossy_attr_value(attr).to_string());
                    }
                })?;
            }
            b"sheetNames" => {}
            b"sheetName" => {
                let mut sheet = ExternalLinkSheet {
                    name: None,
                    r_id: None,
                };
                visit_attributes(e, path, |attr| {
                    let key = local_name(attr.key.as_ref());
                    let value = lossy_attr_value(attr).to_string();
                    if key == b"val" || key == b"name" {
                        sheet.name = Some(value);
                    }
                })?;
                if let Some(name) = sheet.name {
                    part.sheets.push(ExternalLinkSheet {
                        name: Some(name),
                        r_id: None,
                    });
                }
            }
            b"externalBook" => {
                visit_attributes(e, path, |attr| {
                    let key = local_name(attr.key.as_ref());
                    if key == b"id" || key == b"rid" || key == b"rId" {
                        let rel_id = lossy_attr_value(attr).to_string();
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
                })?;
            }
            _ => {}
        }
        Ok(())
    })?;

    Ok(part)
}

pub(crate) fn parse_slicer_part(xml: &str, path: &str) -> Result<SlicerPart, ParseError> {
    let mut slicer = SlicerPart::new();
    slicer.span = Some(SourceSpan::new(path));

    for_each_local_start_or_empty(xml, path, |local, e| {
        if local == b"slicer" {
            visit_attributes(e, path, |attr| {
                let key = local_name(attr.key.as_ref());
                let value = lossy_attr_value(attr).to_string();
                match key {
                    b"name" => slicer.name = Some(value),
                    b"caption" => slicer.caption = Some(value),
                    b"cache" | b"cacheId" => slicer.cache_id = Some(value),
                    b"ref" | b"pivotRef" => slicer.target_ref = Some(value),
                    _ => {}
                }
            })?;
        }
        Ok(())
    })?;

    Ok(slicer)
}

pub(crate) fn parse_timeline_part(xml: &str, path: &str) -> Result<TimelinePart, ParseError> {
    let mut timeline = TimelinePart::new();
    timeline.span = Some(SourceSpan::new(path));

    for_each_local_start_or_empty(xml, path, |local, e| {
        if local == b"timeline" {
            visit_attributes(e, path, |attr| {
                let key = local_name(attr.key.as_ref());
                let value = lossy_attr_value(attr).to_string();
                match key {
                    b"name" => timeline.name = Some(value),
                    b"cache" | b"cacheId" => timeline.cache_id = Some(value),
                    _ => {}
                }
            })?;
        }
        Ok(())
    })?;

    Ok(timeline)
}

pub(crate) fn parse_query_table_part(xml: &str, path: &str) -> Result<QueryTablePart, ParseError> {
    let mut query = QueryTablePart::new();
    query.span = Some(SourceSpan::new(path));

    for_each_local_start_or_empty(xml, path, |local, e| {
        match local {
            b"queryTable" => {
                visit_attributes(e, path, |attr| {
                    let key = local_name(attr.key.as_ref());
                    let value = lossy_attr_value(attr).to_string();
                    match key {
                        b"name" => query.name = Some(value),
                        b"connectionId" | b"connection" => query.connection_id = Some(value),
                        _ => {}
                    }
                })?;
            }
            b"dbPr" => {
                visit_attributes(e, path, |attr| {
                    let key = local_name(attr.key.as_ref());
                    let value = lossy_attr_value(attr).to_string();
                    if key == b"command" {
                        query.command = Some(value.clone());
                    }
                    if key == b"connection" {
                        query.connection_id = Some(value);
                    }
                })?;
            }
            b"webPr" => {
                visit_attributes(e, path, |attr| {
                    let key = local_name(attr.key.as_ref());
                    let value = lossy_attr_value(attr).to_string();
                    if key == b"url" {
                        query.url = Some(value.clone());
                        query.source = Some(value);
                    }
                })?;
            }
            _ => {}
        }
        Ok(())
    })?;

    Ok(query)
}

#[cfg(test)]
mod tests {
    use super::{
        connection_targets, parse_connections_part, parse_external_link_part,
        parse_query_table_part, parse_slicer_part, parse_timeline_part,
    };
    use crate::error::ParseError;

    #[test]
    fn parse_connections_part_accepts_prefixed_connection_tags() {
        let xml = r#"
        <x:connections xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <x:connection id="1" name="Db" type="5" refreshOnLoad="1">
            <x:dbPr connection="Server=example" command="select * from t"/>
          </x:connection>
          <x:connection id="2" name="Web" type="4">
            <x:webPr url="https://example.test/feed"/>
          </x:connection>
        </x:connections>
        "#;

        let part = parse_connections_part(xml, "xl/connections.xml").expect("connections");

        assert_eq!(part.entries.len(), 2);
        assert_eq!(part.entries[0].connection_id, Some(1));
        assert_eq!(
            part.entries[0].connection.as_deref(),
            Some("Server=example")
        );
        assert_eq!(
            part.entries[1].url.as_deref(),
            Some("https://example.test/feed")
        );
        assert_eq!(
            connection_targets(&part),
            vec![
                "Server=example".to_string(),
                "https://example.test/feed".to_string()
            ]
        );
    }

    #[test]
    fn connection_parsers_report_malformed_attributes() {
        let connection_cases: [(&str, &str); 4] = [
            (
                r#"<connections><connection id="1" id="2"/></connections>"#,
                "xl/connections-connection-broken.xml",
            ),
            (
                r#"<connections><connection id="1"><dbPr connection="a" connection="b"/></connection></connections>"#,
                "xl/connections-dbpr-broken.xml",
            ),
            (
                r#"<connections><connection id="1"><webPr url="a" url="b"/></connection></connections>"#,
                "xl/connections-webpr-broken.xml",
            ),
            (
                r#"<connections><connection id="1"><textPr sourceFile="a" sourceFile="b"/></connection></connections>"#,
                "xl/connections-textpr-broken.xml",
            ),
        ];

        for (xml, path) in connection_cases {
            assert_xml_file(parse_connections_part(xml, path), path);
        }

        let external_link_cases: [(&str, &str); 3] = [
            (
                r#"<externalLink linkType="a" linkType="b"/>"#,
                "xl/external-link-type-broken.xml",
            ),
            (
                r#"<externalLink><sheetName val="A" val="B"/></externalLink>"#,
                "xl/external-link-sheet-broken.xml",
            ),
            (
                r#"<externalLink><externalBook r:id="rId1" r:id="rId2"/></externalLink>"#,
                "xl/external-link-book-broken.xml",
            ),
        ];

        for (xml, path) in external_link_cases {
            assert_xml_file(parse_external_link_part(xml, path, None), path);
        }

        assert_xml_file(
            parse_slicer_part(r#"<slicer name="a" name="b"/>"#, "xl/slicer-broken.xml"),
            "xl/slicer-broken.xml",
        );
        assert_xml_file(
            parse_timeline_part(r#"<timeline name="a" name="b"/>"#, "xl/timeline-broken.xml"),
            "xl/timeline-broken.xml",
        );

        let query_cases: [(&str, &str); 3] = [
            (
                r#"<queryTable name="a" name="b"/>"#,
                "xl/query-table-broken.xml",
            ),
            (
                r#"<queryTable><dbPr command="a" command="b"/></queryTable>"#,
                "xl/query-dbpr-broken.xml",
            ),
            (
                r#"<queryTable><webPr url="a" url="b"/></queryTable>"#,
                "xl/query-webpr-broken.xml",
            ),
        ];

        for (xml, path) in query_cases {
            assert_xml_file(parse_query_table_part(xml, path), path);
        }
    }

    fn assert_xml_file<T>(result: Result<T, ParseError>, path: &str) {
        match result {
            Ok(_) => panic!("malformed attributes must fail"),
            Err(ParseError::Xml { file, .. }) => assert_eq!(file, path),
            Err(other) => panic!("expected XML error, got {other:?}"),
        }
    }
}
