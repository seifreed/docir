//! Deterministic normalization for IR trees.

use crate::ir::{IRNode, IrNode as IrNodeTrait};
use crate::types::{NodeId, NodeType};
use crate::visitor::IrStore;

/// Normalize node ordering for deterministic output.
pub fn normalize_store(store: &mut IrStore, root: NodeId) {
    normalize_document(store, root);
    normalize_worksheets(store);
}

fn normalize_document(store: &mut IrStore, root: NodeId) {
    let Some(IRNode::Document(doc)) = store.get(root) else {
        return;
    };

    let mut content = doc.content.clone();
    let mut shared_parts = doc.shared_parts.clone();
    let mut defined_names = doc.defined_names.clone();
    let mut pivot_caches = doc.pivot_caches.clone();
    let mut diagnostics = doc.diagnostics.clone();

    let _ = doc;

    content.sort_by_key(|id| node_sort_key(store, *id));
    shared_parts.sort_by_key(|id| node_sort_key(store, *id));
    defined_names.sort_by_key(|id| defined_name_key(store, *id));
    pivot_caches.sort_by_key(|id| pivot_cache_key(store, *id));
    diagnostics.sort_by_key(|id| node_sort_key(store, *id));

    if let Some(IRNode::Document(doc)) = store.get_mut(root) {
        doc.content = content;
        doc.shared_parts = shared_parts;
        doc.defined_names = defined_names;
        doc.pivot_caches = pivot_caches;
        doc.diagnostics = diagnostics;
    }
}

fn normalize_worksheets(store: &mut IrStore) {
    let worksheet_ids: Vec<NodeId> = store.iter_ids_by_type(NodeType::Worksheet).collect();

    for ws_id in worksheet_ids {
        let Some(IRNode::Worksheet(ws)) = store.get(ws_id) else {
            continue;
        };
        let mut cells = ws.cells.clone();
        let mut drawings = ws.drawings.clone();
        let mut tables = ws.tables.clone();
        let mut conds = ws.conditional_formats.clone();
        let mut validations = ws.data_validations.clone();
        let mut pivots = ws.pivot_tables.clone();
        let mut comments = ws.comments.clone();
        let _ = ws;

        cells.sort_by_key(|id| cell_key(store, *id));
        drawings.sort_by_key(|id| node_sort_key(store, *id));
        tables.sort_by_key(|id| node_sort_key(store, *id));
        conds.sort_by_key(|id| node_sort_key(store, *id));
        validations.sort_by_key(|id| node_sort_key(store, *id));
        pivots.sort_by_key(|id| node_sort_key(store, *id));
        comments.sort_by_key(|id| node_sort_key(store, *id));

        if let Some(IRNode::Worksheet(ws)) = store.get_mut(ws_id) {
            ws.cells = cells;
            ws.drawings = drawings;
            ws.tables = tables;
            ws.conditional_formats = conds;
            ws.data_validations = validations;
            ws.pivot_tables = pivots;
            ws.comments = comments;
        }
    }
}

fn node_sort_key(store: &IrStore, id: NodeId) -> (u32, String) {
    let Some(node) = store.get(id) else {
        return (u32::MAX, id.as_u64().to_string());
    };
    let rank = match node.node_type() {
        NodeType::Section => 10,
        NodeType::Worksheet => 20,
        NodeType::Slide => 30,
        NodeType::Theme => 40,
        NodeType::VmlDrawing => 45,
        NodeType::DrawingPart => 47,
        NodeType::ConnectionPart => 48,
        NodeType::RelationshipGraph => 50,
        NodeType::CustomXmlPart => 60,
        NodeType::MediaAsset => 70,
        NodeType::DigitalSignature => 80,
        NodeType::ExtensionPart => 90,
        NodeType::PivotCacheRecords => 95,
        _ => 500,
    };
    let key = match node {
        IRNode::Worksheet(ws) => ws.name.clone(),
        IRNode::Slide(slide) => format!("{:08}", slide.number),
        IRNode::Section(section) => section.name.clone().unwrap_or_default(),
        IRNode::Theme(theme) => theme.name.clone().unwrap_or_default(),
        IRNode::VmlDrawing(drawing) => drawing.path.clone(),
        IRNode::DrawingPart(part) => part.path.clone(),
        IRNode::ConnectionPart(part) => part
            .span
            .as_ref()
            .map(|s| s.file_path.clone())
            .unwrap_or_default(),
        IRNode::CustomXmlPart(part) => part.path.clone(),
        IRNode::MediaAsset(media) => media.path.clone(),
        IRNode::RelationshipGraph(graph) => graph.source.clone(),
        IRNode::DigitalSignature(sig) => sig.signature_id.clone().unwrap_or_default(),
        IRNode::ExtensionPart(part) => part.path.clone(),
        IRNode::PresentationTag(tag) => tag.name.clone(),
        _ => format!("{:?}", node.node_type()),
    };
    (rank, key)
}

fn defined_name_key(store: &IrStore, id: NodeId) -> (String, u32) {
    let Some(node) = store.get(id) else {
        return (String::new(), 0);
    };
    if let IRNode::DefinedName(name) = node {
        (name.name.clone(), name.local_sheet_id.unwrap_or(0))
    } else {
        (format!("{:?}", node.node_type()), 0)
    }
}

fn pivot_cache_key(store: &IrStore, id: NodeId) -> u32 {
    let Some(node) = store.get(id) else {
        return 0;
    };
    if let IRNode::PivotCache(cache) = node {
        cache.cache_id
    } else {
        0
    }
}

fn cell_key(store: &IrStore, id: NodeId) -> (u32, u32) {
    let Some(node) = store.get(id) else {
        return (u32::MAX, u32::MAX);
    };
    if let IRNode::Cell(cell) = node {
        (cell.row, cell.column)
    } else {
        (u32::MAX, u32::MAX)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{
        Cell, DefinedName, Document, ExtensionPart, ExtensionPartKind, IRNode, MediaAsset,
        MediaType, PivotCache, Section, Worksheet,
    };
    use crate::types::DocumentFormat;

    #[test]
    fn normalize_store_sorts_document_collections_deterministically() {
        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::Spreadsheet);

        let mut section_b = Section::new();
        section_b.name = Some("B".to_string());
        let section_b_id = section_b.id;
        store.insert(IRNode::Section(section_b));

        let mut section_a = Section::new();
        section_a.name = Some("A".to_string());
        let section_a_id = section_a.id;
        store.insert(IRNode::Section(section_a));

        let media = MediaAsset::new("z.bin", MediaType::Other, 1);
        let media_id = media.id;
        store.insert(IRNode::MediaAsset(media));
        let part = ExtensionPart::new("a.ext", 1, ExtensionPartKind::Unknown);
        let part_id = part.id;
        store.insert(IRNode::ExtensionPart(part));

        let dn_b = DefinedName {
            id: NodeId::new(),
            name: "Zeta".to_string(),
            value: "Sheet1!$A$1".to_string(),
            local_sheet_id: Some(2),
            hidden: false,
            comment: None,
            span: None,
        };
        let dn_b_id = dn_b.id;
        store.insert(IRNode::DefinedName(dn_b));
        let dn_a = DefinedName {
            id: NodeId::new(),
            name: "Alpha".to_string(),
            value: "Sheet1!$B$1".to_string(),
            local_sheet_id: Some(1),
            hidden: false,
            comment: None,
            span: None,
        };
        let dn_a_id = dn_a.id;
        store.insert(IRNode::DefinedName(dn_a));

        let cache_2 = PivotCache::new(2);
        let cache_2_id = cache_2.id;
        store.insert(IRNode::PivotCache(cache_2));
        let cache_1 = PivotCache::new(1);
        let cache_1_id = cache_1.id;
        store.insert(IRNode::PivotCache(cache_1));

        doc.content = vec![section_b_id, section_a_id];
        doc.shared_parts = vec![part_id, media_id];
        doc.defined_names = vec![dn_b_id, dn_a_id];
        doc.pivot_caches = vec![cache_2_id, cache_1_id];
        doc.diagnostics = vec![part_id, media_id];
        let root = doc.id;
        store.insert(IRNode::Document(doc));

        normalize_store(&mut store, root);

        let IRNode::Document(doc) = store.get(root).expect("document present") else {
            panic!("expected document node");
        };
        assert_eq!(doc.content, vec![section_a_id, section_b_id]);
        assert_eq!(doc.shared_parts, vec![media_id, part_id]);
        assert_eq!(doc.defined_names, vec![dn_a_id, dn_b_id]);
        assert_eq!(doc.pivot_caches, vec![cache_1_id, cache_2_id]);
        assert_eq!(doc.diagnostics, vec![media_id, part_id]);
    }

    #[test]
    fn normalize_store_sorts_worksheet_children_and_handles_non_document_root() {
        let mut store = IrStore::new();

        let mut worksheet = Worksheet::new("Sheet1", 1);
        let c_b2 = Cell::new("B2", 1, 1);
        let c_b2_id = c_b2.id;
        store.insert(IRNode::Cell(c_b2));
        let c_a1 = Cell::new("A1", 0, 0);
        let c_a1_id = c_a1.id;
        store.insert(IRNode::Cell(c_a1));

        let part = ExtensionPart::new("z.ext", 1, ExtensionPartKind::Unknown);
        let part_id = part.id;
        store.insert(IRNode::ExtensionPart(part));
        let media = MediaAsset::new("a.bin", MediaType::Other, 1);
        let media_id = media.id;
        store.insert(IRNode::MediaAsset(media));

        worksheet.cells = vec![c_b2_id, c_a1_id];
        worksheet.drawings = vec![part_id, media_id];
        worksheet.tables = vec![part_id, media_id];
        worksheet.conditional_formats = vec![part_id, media_id];
        worksheet.data_validations = vec![part_id, media_id];
        worksheet.pivot_tables = vec![part_id, media_id];
        worksheet.comments = vec![part_id, media_id];
        let ws_id = worksheet.id;
        store.insert(IRNode::Worksheet(worksheet));

        // Non-document root: normalization should still process worksheets.
        normalize_store(&mut store, ws_id);

        let IRNode::Worksheet(ws) = store.get(ws_id).expect("worksheet present") else {
            panic!("expected worksheet node");
        };
        assert_eq!(ws.cells, vec![c_a1_id, c_b2_id]);
        assert_eq!(ws.drawings, vec![media_id, part_id]);
        assert_eq!(ws.tables, vec![media_id, part_id]);
        assert_eq!(ws.conditional_formats, vec![media_id, part_id]);
        assert_eq!(ws.data_validations, vec![media_id, part_id]);
        assert_eq!(ws.pivot_tables, vec![media_id, part_id]);
        assert_eq!(ws.comments, vec![media_id, part_id]);
    }
}
