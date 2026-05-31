use docir_core::ir::{
    DrawingPart, GlossaryEntry, IRNode, IrNode as IrNodeTrait, VmlDrawing, VmlShape, WebExtension,
    WebExtensionTaskpane,
};

use super::format_helpers::{opt_bool, opt_str};

pub(crate) fn summarize_secondary(node: &IRNode) -> String {
    match node {
        IRNode::CommentExtensionSet(set) => format!("entries={}", set.entries.len()),
        IRNode::CommentIdMap(map) => format!("mappings={}", map.mappings.len()),
        IRNode::CommentRangeStart(start) => format!("comment_id={}", start.comment_id),
        IRNode::CommentRangeEnd(end) => format!("comment_id={}", end.comment_id),
        IRNode::CommentReference(reference) => format!("comment_id={}", reference.comment_id),
        IRNode::PeoplePart(people) => format!("people={}", people.people.len()),
        IRNode::SmartArtPart(part) => format!("kind={} path={}", part.kind, part.path),
        IRNode::WebExtension(ext) => summarize_web_extension(ext),
        IRNode::WebExtensionTaskpane(pane) => summarize_web_extension_taskpane(pane),
        IRNode::GlossaryDocument(doc) => format!("entries={}", doc.entries.len()),
        IRNode::GlossaryEntry(entry) => summarize_glossary_entry(entry),
        IRNode::VmlDrawing(drawing) => summarize_vml_drawing(drawing),
        IRNode::VmlShape(shape) => summarize_vml_shape(shape),
        IRNode::DrawingPart(part) => summarize_drawing_part(part),
        IRNode::Diagnostics(diag) => format!("entries={}", diag.entries.len()),
        _ => format!("unsupported={:?}", node.node_type()),
    }
}

fn summarize_web_extension(ext: &WebExtension) -> String {
    format!(
        "id={} store={} version={} properties={}",
        opt_str(&ext.extension_id),
        opt_str(&ext.store),
        opt_str(&ext.version),
        ext.properties.len()
    )
}

fn summarize_web_extension_taskpane(pane: &WebExtensionTaskpane) -> String {
    format!(
        "ref={} dock_state={} visible={}",
        opt_str(&pane.web_extension_ref),
        opt_str(&pane.dock_state),
        opt_bool(pane.visibility)
    )
}

fn summarize_glossary_entry(entry: &GlossaryEntry) -> String {
    format!(
        "name={} gallery={} content_nodes={}",
        opt_str(&entry.name),
        opt_str(&entry.gallery),
        entry.content.len()
    )
}

fn summarize_vml_drawing(drawing: &VmlDrawing) -> String {
    format!("path={} shapes={}", drawing.path, drawing.shapes.len())
}

fn summarize_vml_shape(shape: &VmlShape) -> String {
    format!(
        "name={} rel_id={} image_target={}",
        opt_str(&shape.name),
        opt_str(&shape.rel_id),
        opt_str(&shape.image_target)
    )
}

fn summarize_drawing_part(part: &DrawingPart) -> String {
    format!("path={} shapes={}", part.path, part.shapes.len())
}
