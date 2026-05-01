use crate::ooxml::part_utils::read_relationships;
use crate::ooxml::relationships::Relationships;
use crate::ooxml::xlsx::{
    connection_targets, parse_connections_part, parse_external_link_part, rel_type,
    ExternalRefType, ExternalReference, IRNode, ParseError, XlsxParser,
};
use crate::zip_handler::PackageReader;
use docir_core::types::SourceSpan;

pub(super) fn parse_external_links_and_connections_impl(
    parser: &mut XlsxParser,
    zip: &mut impl PackageReader,
    workbook_path: &str,
    workbook_rels: &Relationships,
) -> Result<(), ParseError> {
    parse_external_link_parts(parser, zip, workbook_path, workbook_rels)?;
    parse_connections(parser, zip)?;
    Ok(())
}

fn parse_external_link_parts(
    parser: &mut XlsxParser,
    zip: &mut impl PackageReader,
    workbook_path: &str,
    workbook_rels: &Relationships,
) -> Result<(), ParseError> {
    for rel in workbook_rels.get_by_type(rel_type::EXTERNAL_LINK) {
        let external_path = Relationships::resolve_target(workbook_path, &rel.target);
        if !zip.contains(&external_path) {
            continue;
        }

        let rels = read_relationships(zip, &external_path)?;
        if let Ok(xml) = zip.read_file_string(&external_path) {
            if let Ok(mut part) = parse_external_link_part(&xml, &external_path, Some(&rels)) {
                part.span = Some(SourceSpan::new(&external_path));
                let part_id = part.id;
                parser.store.insert(IRNode::ExternalLinkPart(part));
                push_shared_part(parser, part_id);
            }
        }

        for ext in rels.by_id.values() {
            let target = &ext.target;
            let ext_ref = ExternalReference::new(ExternalRefType::DataConnection, target);
            let ext_ref = ExternalReference {
                relationship_id: Some(ext.id.clone()),
                relationship_type: Some(ext.rel_type.clone()),
                ..ext_ref
            };
            push_external_reference(parser, ext_ref);
        }
    }

    Ok(())
}

fn parse_connections(
    parser: &mut XlsxParser,
    zip: &mut impl PackageReader,
) -> Result<(), ParseError> {
    if !zip.contains("xl/connections.xml") {
        return Ok(());
    }

    let xml = zip.read_file_string("xl/connections.xml")?;
    if let Ok(mut part) = parse_connections_part(&xml, "xl/connections.xml") {
        part.span = Some(SourceSpan::new("xl/connections.xml"));
        let part_id = part.id;
        let targets = connection_targets(&part);
        parser.store.insert(IRNode::ConnectionPart(part));
        push_shared_part(parser, part_id);
        for target in targets {
            let ext_ref = ExternalReference::new(ExternalRefType::DataConnection, target);
            push_external_reference(parser, ext_ref);
        }
    }
    Ok(())
}

fn push_shared_part(parser: &mut XlsxParser, part_id: docir_core::types::NodeId) {
    if let Some(IRNode::Document(doc)) = parser.store.get_mut(parser.root_id) {
        doc.shared_parts.push(part_id);
    }
}

fn push_external_reference(parser: &mut XlsxParser, ext_ref: ExternalReference) {
    let id = ext_ref.id;
    parser.store.insert(IRNode::ExternalReference(ext_ref));
    parser.security_info.external_refs.push(id);
}
