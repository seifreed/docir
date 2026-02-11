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
