use crate::summary::{content_signature, short_hash, style_signature, summarize};
use docir_core::ir::IRNode;
use docir_core::ir::IrNode as IrNodeTrait;
use docir_core::types::{NodeId, NodeType};
use docir_core::visitor::IrStore;
use std::collections::{BTreeMap, HashMap, HashSet};
#[path = "index_intrinsic.rs"]
mod index_intrinsic;

#[derive(Debug, Clone)]
pub(crate) struct NodeSnapshot {
    pub(crate) node_type: NodeType,
    pub(crate) summary: String,
    pub(crate) content_sig: Option<String>,
    pub(crate) style_sig: Option<String>,
}

pub(crate) fn build_index(store: &IrStore, root: NodeId) -> BTreeMap<String, NodeSnapshot> {
    let mut index = BTreeMap::new();
    let mut used = HashSet::new();
    let root_key = match store.get(root) {
        Some(node) => {
            intrinsic_key(node, store).unwrap_or_else(|| format!("{:?}", node.node_type()))
        }
        None => "Document".to_string(),
    };
    walk_with_key(store, root, root_key, &mut index, &mut used);
    index
}

fn walk_with_key(
    store: &IrStore,
    node_id: NodeId,
    key: String,
    index: &mut BTreeMap<String, NodeSnapshot>,
    used_keys: &mut HashSet<String>,
) {
    let Some(node) = store.get(node_id) else {
        return;
    };
    let node_type = node.node_type();

    let summary = summarize(node, store);
    let content_sig = content_signature(node, store);
    let style_sig = style_signature(node, store);
    index.insert(
        key.clone(),
        NodeSnapshot {
            node_type,
            summary,
            content_sig,
            style_sig,
        },
    );

    let mut counters: HashMap<NodeType, usize> = HashMap::new();
    let mut used_local: HashSet<String> = HashSet::new();

    for child_id in node.children() {
        let Some(child) = store.get(child_id) else {
            continue;
        };
        let child_type = child.node_type();
        let entry = counters.entry(child_type).or_insert(0);
        *entry += 1;
        let ordinal = *entry;

        let child_local = local_key_with_index(child, store, ordinal, &mut used_local);
        let child_key = format!("{key}/{child_local}");
        if !used_keys.insert(child_key.clone()) {
            continue;
        }
        walk_with_key(store, child_id, child_key, index, used_keys);
    }
}

fn local_key_with_index(
    node: &IRNode,
    store: &IrStore,
    ordinal: usize,
    used: &mut HashSet<String>,
) -> String {
    let base = if let Some(key) = intrinsic_key(node, store) {
        key
    } else if matches!(node.node_type(), NodeType::Paragraph | NodeType::Run) {
        format!("{:?}[{}]", node.node_type(), ordinal)
    } else if let Some(sig) = content_signature(node, store) {
        format!("{:?}[{}]", node.node_type(), short_hash(&sig))
    } else {
        format!("{:?}[{}]", node.node_type(), ordinal)
    };
    let mut candidate = base.clone();
    let mut counter = 1usize;
    while used.contains(&candidate) {
        counter += 1;
        candidate = format!("{base}#{counter}");
    }
    used.insert(candidate.clone());
    if candidate == base {
        base
    } else {
        candidate
    }
}

fn intrinsic_key(node: &IRNode, _store: &IrStore) -> Option<String> {
    index_intrinsic::intrinsic_key(node)
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::{
        BookmarkEnd, BookmarkStart, CalcChain, Cell, Comment, ConnectionPart, DataValidation,
        DefinedName, Diagnostics, Document, DrawingPart, Endnote, ExternalLinkPart, Footnote,
        GlossaryEntry, IRNode, Paragraph, PivotCache, PivotCacheRecords, PptxCommentAuthor,
        PresentationTag, QueryTablePart, Revision, RevisionType, Run, Shape, ShapeType,
        SheetComment, SheetMetadata, SlicerPart, Slide, SmartArtPart, SpreadsheetStyles,
        TableDefinition, TimelinePart, VmlDrawing, VmlShape, WebExtension, WebExtensionTaskpane,
        Worksheet, WorksheetDrawing,
    };
    use docir_core::security::{MacroProject, OleObject};
    use docir_core::types::DocumentFormat;
    use docir_core::visitor::IrStore;

    fn assert_intrinsic_key(node: IRNode, expected: &str, store: &IrStore) {
        assert_eq!(
            intrinsic_key(&node, store),
            Some(expected.to_string()),
            "wrong intrinsic key for {node:?}"
        );
    }

    #[test]
    fn build_index_assigns_stable_paths_and_snapshots() {
        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::Spreadsheet);

        let mut ws = Worksheet::new("Sheet1", 1);
        let mut cell = Cell::new("A1", 0, 0);
        cell.value = docir_core::ir::CellValue::String("alpha".to_string());

        let ws_id = ws.id;
        let cell_id = cell.id;
        ws.cells.push(cell_id);
        doc.content.push(ws_id);

        store.insert(IRNode::Document(doc.clone()));
        store.insert(IRNode::Worksheet(ws.clone()));
        store.insert(IRNode::Cell(cell.clone()));

        let index = build_index(&store, doc.id);
        assert!(index.keys().any(|k| k.contains("Worksheet[Sheet1]")));
        assert!(index.keys().any(|k| k.contains("Cell[A1]")));

        let cell_key = index
            .keys()
            .find(|k| k.contains("Cell[A1]"))
            .expect("cell key must exist");
        let snapshot = index.get(cell_key).expect("snapshot for cell");
        assert_eq!(snapshot.node_type, docir_core::types::NodeType::Cell);
        assert!(snapshot.summary.contains("ref=A1"));
    }

    #[test]
    fn local_key_with_index_disambiguates_duplicate_nodes() {
        let mut store = IrStore::new();
        let run_a = Run::new("same");
        let run_b = Run::new("same");
        let mut para = Paragraph::new();
        para.runs = vec![run_a.id, run_b.id];

        store.insert(IRNode::Paragraph(para.clone()));
        store.insert(IRNode::Run(run_a));
        store.insert(IRNode::Run(run_b));

        let index = build_index(&store, para.id);
        let run_keys: Vec<_> = index
            .keys()
            .filter(|k| k.contains("Run["))
            .cloned()
            .collect();
        assert_eq!(run_keys.len(), 2);
        assert_ne!(run_keys[0], run_keys[1]);
    }

    #[test]
    fn intrinsic_key_covers_word_variants() {
        let store = IrStore::new();
        assert_intrinsic_key(IRNode::Comment(Comment::new("c1")), "Comment[c1]", &store);
        assert_intrinsic_key(
            IRNode::Footnote(Footnote::new("f1")),
            "Footnote[f1]",
            &store,
        );
        assert_intrinsic_key(IRNode::Endnote(Endnote::new("e1")), "Endnote[e1]", &store);
        assert_intrinsic_key(
            IRNode::BookmarkStart(BookmarkStart::new("b1")),
            "BookmarkStart[b1]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::BookmarkEnd(BookmarkEnd::new("b1")),
            "BookmarkEnd[b1]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::Revision(Revision::new(RevisionType::Insert)),
            "Revision[Insert]",
            &store,
        );
    }

    #[test]
    fn intrinsic_key_covers_annotation_variants() {
        let store = IrStore::new();
        let mut glossary = GlossaryEntry::new();
        glossary.name = Some("entry".to_string());
        assert_intrinsic_key(
            IRNode::GlossaryEntry(glossary),
            "GlossaryEntry[entry]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::VmlDrawing(VmlDrawing::new("word/vml.vml")),
            "VmlDrawing[word/vml.vml]",
            &store,
        );
        let mut vml_shape = VmlShape::new();
        vml_shape.name = Some("shape1".to_string());
        assert_intrinsic_key(IRNode::VmlShape(vml_shape), "VmlShape[shape1]", &store);
        assert_intrinsic_key(
            IRNode::DrawingPart(DrawingPart::new("word/drawing1.xml")),
            "DrawingPart[word/drawing1.xml]",
            &store,
        );
    }

    #[test]
    fn intrinsic_key_covers_presentation_variants() {
        let store = IrStore::new();
        assert_intrinsic_key(IRNode::Slide(Slide::new(2)), "Slide[2]", &store);
        let mut shape = Shape::new(ShapeType::TextBox);
        shape.name = Some("Title".to_string());
        assert_intrinsic_key(IRNode::Shape(shape), "Shape[Title]", &store);
        assert_intrinsic_key(
            IRNode::PptxCommentAuthor(PptxCommentAuthor {
                id: docir_core::types::NodeId::new(),
                author_id: 7,
                name: None,
                initials: None,
                span: None,
            }),
            "PptxCommentAuthor[7]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::PresentationTag(PresentationTag {
                id: docir_core::types::NodeId::new(),
                name: "tag".to_string(),
                value: None,
                span: None,
            }),
            "PresentationTag[tag]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::SmartArtPart(SmartArtPart {
                id: docir_core::types::NodeId::new(),
                kind: "diagramData".to_string(),
                path: "ppt/diagrams/data1.xml".to_string(),
                root_element: None,
                point_count: None,
                connection_count: None,
                rel_ids: Vec::new(),
                span: None,
            }),
            "SmartArtPart[diagramData]",
            &store,
        );
    }

    #[test]
    fn intrinsic_key_covers_security_variants() {
        let store = IrStore::new();
        let mut project = MacroProject::new();
        project.name = Some("VBAProject".to_string());
        assert_intrinsic_key(
            IRNode::MacroProject(project),
            "MacroProject[VBAProject]",
            &store,
        );
        let mut ole = OleObject::new();
        ole.name = Some("Obj".to_string());
        assert_intrinsic_key(IRNode::OleObject(ole), "OleObject[Obj]", &store);
        let mut ext = WebExtension::new();
        ext.extension_id = Some("ext-1".to_string());
        assert_intrinsic_key(IRNode::WebExtension(ext), "WebExtension[ext-1]", &store);
        let mut taskpane = WebExtensionTaskpane::new();
        taskpane.web_extension_ref = Some("ext-1".to_string());
        assert_intrinsic_key(
            IRNode::WebExtensionTaskpane(taskpane),
            "WebExtensionTaskpane[ext-1]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::Diagnostics(Diagnostics::new()),
            "Diagnostics",
            &store,
        );
    }

    #[test]
    fn intrinsic_key_covers_spreadsheet_core_variants() {
        let store = IrStore::new();

        assert_intrinsic_key(
            IRNode::Worksheet(Worksheet::new("SheetX", 1)),
            "Worksheet[SheetX]",
            &store,
        );
        assert_intrinsic_key(IRNode::Cell(Cell::new("C3", 2, 2)), "Cell[C3]", &store);
        let defined_name = DefinedName {
            id: docir_core::types::NodeId::new(),
            name: "NamedRange".to_string(),
            value: "SheetX!$A$1".to_string(),
            local_sheet_id: None,
            hidden: false,
            comment: None,
            span: None,
        };
        assert_intrinsic_key(
            IRNode::DefinedName(defined_name),
            "DefinedName[NamedRange]",
            &store,
        );

        let mut table = TableDefinition {
            id: docir_core::types::NodeId::new(),
            name: None,
            display_name: None,
            ref_range: None,
            header_row_count: None,
            totals_row_count: None,
            columns: Vec::new(),
            span: None,
        };
        table.display_name = Some("Orders".to_string());
        assert_intrinsic_key(
            IRNode::TableDefinition(table),
            "TableDefinition[Orders]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::WorksheetDrawing(WorksheetDrawing::new()),
            "WorksheetDrawing",
            &store,
        );
        assert_intrinsic_key(
            IRNode::SheetComment(SheetComment::new("A1", "note")),
            "SheetComment[A1]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::SheetMetadata(SheetMetadata::new()),
            "SheetMetadata",
            &store,
        );
    }

    #[test]
    fn intrinsic_key_covers_spreadsheet_other_variants() {
        let store = IrStore::new();
        assert_intrinsic_key(IRNode::CalcChain(CalcChain::new()), "CalcChain", &store);
        let mut external_link = ExternalLinkPart::new();
        external_link.target = Some("xl/externalLinks/ext.xml".to_string());
        assert_intrinsic_key(
            IRNode::ExternalLinkPart(external_link),
            "ExternalLinkPart[xl/externalLinks/ext.xml]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::ConnectionPart(ConnectionPart::new()),
            "ConnectionPart[0]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::SlicerPart(SlicerPart::new()),
            "SlicerPart[-]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::TimelinePart(TimelinePart::new()),
            "TimelinePart[-]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::QueryTablePart(QueryTablePart::new()),
            "QueryTablePart[-]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::SharedStringTable(docir_core::ir::SharedStringTable::new()),
            "SharedStringTable",
            &store,
        );
        assert_intrinsic_key(
            IRNode::SpreadsheetStyles(SpreadsheetStyles::new()),
            "SpreadsheetStyles",
            &store,
        );
        let conditional = docir_core::ir::ConditionalFormat {
            id: docir_core::types::NodeId::new(),
            ranges: Vec::new(),
            rules: Vec::new(),
            span: None,
        };
        assert_intrinsic_key(
            IRNode::ConditionalFormat(conditional),
            "ConditionalFormat",
            &store,
        );
        let validation = DataValidation {
            id: docir_core::types::NodeId::new(),
            validation_type: None,
            operator: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            ranges: Vec::new(),
            formula1: None,
            formula2: None,
            span: None,
        };
        assert_intrinsic_key(IRNode::DataValidation(validation), "DataValidation", &store);
        assert_intrinsic_key(
            IRNode::PivotCache(PivotCache::new(4)),
            "PivotCache[4]",
            &store,
        );
        assert_intrinsic_key(
            IRNode::PivotCacheRecords(PivotCacheRecords::new()),
            "PivotCacheRecords[0]",
            &store,
        );
    }

    #[test]
    fn build_index_disambiguates_colliding_intrinsic_child_keys() {
        let mut store = IrStore::new();
        let mut doc = Document::new(DocumentFormat::Spreadsheet);

        let d1 = WorksheetDrawing::new();
        let d2 = WorksheetDrawing::new();
        let d1_id = d1.id;
        let d2_id = d2.id;

        doc.content.push(d1_id);
        doc.content.push(d2_id);
        store.insert(IRNode::Document(doc.clone()));
        store.insert(IRNode::WorksheetDrawing(d1));
        store.insert(IRNode::WorksheetDrawing(d2));

        let index = build_index(&store, doc.id);
        let keys: Vec<String> = index
            .keys()
            .filter(|k| k.contains("WorksheetDrawing"))
            .cloned()
            .collect();
        assert_eq!(keys.len(), 2);
        assert!(keys.iter().any(|k| k.ends_with("WorksheetDrawing")));
        assert!(keys.iter().any(|k| k.ends_with("WorksheetDrawing#2")));
    }
}
