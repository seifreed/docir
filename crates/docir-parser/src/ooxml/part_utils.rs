//! Shared OOXML part utilities.

use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::zip_handler::PackageReader;

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
