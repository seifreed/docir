use crate::ooxml::relationships::{Relationships, TargetMode};
use docir_core::ir::{RelationshipEntry, RelationshipGraph, RelationshipTargetMode};
use docir_core::types::SourceSpan;

/// Public API entrypoint: build_relationship_graph.
pub fn build_relationship_graph(
    source: &str,
    rels_path: &str,
    rels: &Relationships,
) -> RelationshipGraph {
    let mut graph = RelationshipGraph::new(source);
    // Coverage wants to mark the `.rels` part itself as "seen".
    // We still keep `graph.source` to point at the owning part (e.g. word/document.xml).
    graph.span = Some(SourceSpan::new(rels_path));

    for rel in rels.by_id.values() {
        let target_mode = match rel.target_mode {
            TargetMode::Internal => RelationshipTargetMode::Internal,
            TargetMode::External => RelationshipTargetMode::External,
        };
        graph.relationships.push(RelationshipEntry {
            id: rel.id.clone(),
            rel_type: rel.rel_type.clone(),
            target: rel.target.clone(),
            target_mode,
        });
    }

    graph
}
