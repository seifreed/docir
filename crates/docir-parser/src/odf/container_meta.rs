use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::{Read, Seek};

use docir_core::ir::DocumentMetadata;

use super::{Document, IRNode, IrStore, ParseError, SecureZipReader};
use crate::xml_utils::{
    XmlScanControl, decoded_general_ref, decoded_text, local_name, scan_xml_events, xml_error,
};

#[derive(Clone, Copy)]
enum MetaField {
    Title,
    Subject,
    Creator,
    Keywords,
    Description,
    Created,
    Modified,
}

pub(super) fn load_meta<R: Read + Seek>(
    zip: &mut SecureZipReader<R>,
    store: &mut IrStore,
    doc: &mut Document,
) -> Result<(), ParseError> {
    if zip.contains("meta.xml") {
        let meta_xml = zip.read_file_string("meta.xml")?;
        if let Some(meta) = parse_meta(&meta_xml)? {
            let meta_id = meta.id;
            store.insert(IRNode::Metadata(meta));
            doc.metadata = Some(meta_id);
        }
    }
    Ok(())
}

pub(super) fn parse_meta(xml: &str) -> Result<Option<DocumentMetadata>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut meta = DocumentMetadata::new();
    let mut current = None;

    scan_xml_events(&mut reader, &mut buf, "meta.xml", |event| {
        match event {
            Event::Start(e) => {
                current = meta_field_for_name(local_name(e.name().as_ref()));
            }
            Event::Text(e) => {
                if let Some(field) = current {
                    let value = decoded_text(&e).map_err(|err| xml_error("meta.xml", err))?;
                    set_meta_field(&mut meta, field, value);
                }
            }
            Event::GeneralRef(e) => {
                if let Some(field) = current {
                    let value =
                        decoded_general_ref(&e).map_err(|err| xml_error("meta.xml", err))?;
                    set_meta_field(&mut meta, field, value);
                }
            }
            Event::End(_) => {
                current = None;
            }
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    if meta_has_any_field(&meta) {
        Ok(Some(meta))
    } else {
        Ok(None)
    }
}

fn meta_field_for_name(name: &[u8]) -> Option<MetaField> {
    match name {
        b"title" => Some(MetaField::Title),
        b"subject" => Some(MetaField::Subject),
        b"creator" => Some(MetaField::Creator),
        b"keyword" => Some(MetaField::Keywords),
        b"description" => Some(MetaField::Description),
        b"creation-date" => Some(MetaField::Created),
        b"date" => Some(MetaField::Modified),
        _ => None,
    }
}

fn set_meta_field(meta: &mut DocumentMetadata, field: MetaField, value: String) {
    match field {
        MetaField::Title => meta.title = Some(value),
        MetaField::Subject => meta.subject = Some(value),
        MetaField::Creator => meta.creator = Some(value),
        MetaField::Keywords => meta.keywords = Some(value),
        MetaField::Description => meta.description = Some(value),
        MetaField::Created => meta.created = Some(value),
        MetaField::Modified => meta.modified = Some(value),
    }
}

fn meta_has_any_field(meta: &DocumentMetadata) -> bool {
    meta.title.is_some()
        || meta.subject.is_some()
        || meta.creator.is_some()
        || meta.keywords.is_some()
        || meta.description.is_some()
        || meta.created.is_some()
        || meta.modified.is_some()
}
