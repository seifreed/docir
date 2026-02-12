use super::*;

impl PptxParser {
    pub(super) fn process_external_relationships(&mut self, rels: &Relationships, file_path: &str) {
        for rel in rels.external_relationships() {
            let ref_type = classify_relationship(&rel.rel_type);
            self.add_external_reference(rel, ref_type, file_path);
        }
    }

    pub(super) fn add_external_reference(
        &mut self,
        rel: &Relationship,
        ref_type: ExternalRefType,
        file_path: &str,
    ) {
        let key = format!("{file_path}::{id}", id = rel.id);
        if !self.external_rel_ids.insert(key) {
            return;
        }

        let mut ext_ref = ExternalReference::new(ref_type, &rel.target);
        ext_ref.relationship_id = Some(rel.id.clone());
        ext_ref.relationship_type = Some(rel.rel_type.clone());
        ext_ref.span = Some(SourceSpan::new(file_path).with_relationship(rel.id.clone()));

        let ext_id = ext_ref.id;
        self.store.insert(IRNode::ExternalReference(ext_ref));
        self.security_info.external_refs.push(ext_id);
    }
}

pub(super) fn classify_relationship(rel_type_uri: &str) -> ExternalRefType {
    if rel_type_uri.contains("hyperlink") {
        ExternalRefType::Hyperlink
    } else if rel_type_uri.contains("image") {
        ExternalRefType::Image
    } else if rel_type_uri.contains("slideMaster") || rel_type_uri.contains("slideLayout") {
        ExternalRefType::SlideMaster
    } else if rel_type_uri.contains("oleObject") {
        ExternalRefType::OleLink
    } else if rel_type_uri.contains("external") {
        ExternalRefType::DataConnection
    } else {
        ExternalRefType::Other
    }
}
