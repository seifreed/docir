//! Shared OOXML part utilities.

use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::zip_handler::PackageReader;
use docir_core::ir::{Document, IRNode};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;

/// Get the relationships file path for a given part.
pub(crate) fn get_rels_path(part_path: &str) -> String {
    if let Some(idx) = part_path.rfind('/') {
        let dir = &part_path[..idx + 1];
        let file = &part_path[idx + 1..];
        format!("{}_rels/{}.rels", dir, file)
    } else {
        format!("_rels/{}.rels", part_path)
    }
}

/// Read relationships for a part, returning an error if the relationships file is invalid.
pub(crate) fn read_relationships(
    zip: &mut impl PackageReader,
    part_path: &str,
) -> Result<Relationships, ParseError> {
    let rels_path = get_rels_path(part_path);
    if !zip.contains(&rels_path) {
        return Ok(Relationships::default());
    }
    let rels_xml = zip.read_file_string(&rels_path)?;
    Relationships::parse(&rels_xml)
}

/// Read relationships for a part, returning a default value on any failure.
pub(crate) fn read_relationships_optional(
    zip: &mut impl PackageReader,
    part_path: &str,
) -> Relationships {
    let rels_path = get_rels_path(part_path);
    if zip.contains(&rels_path) {
        if let Ok(rels_xml) = zip.read_file_string(&rels_path) {
            if let Ok(rels) = Relationships::parse(&rels_xml) {
                return rels;
            }
        }
    }
    Relationships::default()
}

pub(crate) fn read_xml_part(
    zip: &mut impl PackageReader,
    part_path: &str,
) -> Result<Option<String>, ParseError> {
    if !zip.contains(part_path) {
        return Ok(None);
    }
    Ok(Some(zip.read_file_string(part_path)?))
}

pub(crate) fn read_xml_part_by_rel(
    zip: &mut impl PackageReader,
    main_part_path: &str,
    rels: &Relationships,
    rel_type: &str,
) -> Result<Option<(String, String)>, ParseError> {
    let Some(rel) = rels.get_first_by_type(rel_type) else {
        return Ok(None);
    };
    let part_path = Relationships::resolve_target(main_part_path, &rel.target);
    let Some(xml) = read_xml_part(zip, &part_path)? else {
        return Ok(None);
    };
    Ok(Some((part_path, xml)))
}

pub(crate) fn parse_xml_part_with_span<T, F, S>(
    zip: &mut impl PackageReader,
    part_path: &str,
    parse: F,
    set_span: S,
) -> Result<Option<T>, ParseError>
where
    F: FnOnce(&str, &str) -> Result<T, ParseError>,
    S: FnOnce(&mut T, &str),
{
    let Some(xml) = read_xml_part(zip, part_path)? else {
        return Ok(None);
    };
    let mut part = parse(&xml, part_path)?;
    set_span(&mut part, part_path);
    Ok(Some(part))
}

pub(crate) fn insert_shared_part(
    store: &mut IrStore,
    document: &mut Document,
    node: IRNode,
    id: NodeId,
) {
    store.insert(node);
    document.shared_parts.push(id);
}
