use crate::summary::{content_signature, short_hash, style_signature, summarize};
use docir_core::ir::IRNode;
use docir_core::ir::IrNode as IrNodeTrait;
use docir_core::types::{NodeId, NodeType};
use docir_core::visitor::IrStore;
use std::collections::{BTreeMap, HashMap, HashSet};

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
    intrinsic_key_spreadsheet(node)
        .or_else(|| intrinsic_key_word(node))
        .or_else(|| intrinsic_key_presentation(node))
        .or_else(|| intrinsic_key_security(node))
}

fn intrinsic_key_spreadsheet(node: &IRNode) -> Option<String> {
    match node {
        IRNode::Worksheet(ws) => Some(format!("Worksheet[{}]", ws.name)),
        IRNode::Cell(cell) => Some(format!("Cell[{}]", cell.reference)),
        IRNode::DefinedName(def) => Some(format!("DefinedName[{}]", def.name)),
        IRNode::TableDefinition(table) => table
            .display_name
            .as_ref()
            .map(|name| format!("TableDefinition[{name}]")),
        IRNode::PivotTable(pivot) => pivot
            .name
            .as_ref()
            .map(|name| format!("PivotTable[{name}]")),
        IRNode::WorksheetDrawing(_) => Some("WorksheetDrawing".to_string()),
        IRNode::CalcChain(_) => Some("CalcChain".to_string()),
        IRNode::SheetComment(comment) => Some(format!("SheetComment[{}]", comment.cell_ref)),
        IRNode::SheetMetadata(_) => Some("SheetMetadata".to_string()),
        IRNode::ExternalLinkPart(part) => Some(format!(
            "ExternalLinkPart[{}]",
            part.target.as_deref().unwrap_or("-")
        )),
        IRNode::ConnectionPart(part) => Some(format!("ConnectionPart[{}]", part.entries.len())),
        IRNode::SlicerPart(part) => Some(format!(
            "SlicerPart[{}]",
            part.name.as_deref().unwrap_or("-")
        )),
        IRNode::TimelinePart(part) => Some(format!(
            "TimelinePart[{}]",
            part.name.as_deref().unwrap_or("-")
        )),
        IRNode::QueryTablePart(part) => Some(format!(
            "QueryTablePart[{}]",
            part.name.as_deref().unwrap_or("-")
        )),
        IRNode::SharedStringTable(_) => Some("SharedStringTable".to_string()),
        IRNode::SpreadsheetStyles(_) => Some("SpreadsheetStyles".to_string()),
        IRNode::ConditionalFormat(_) => Some("ConditionalFormat".to_string()),
        IRNode::DataValidation(_) => Some("DataValidation".to_string()),
        IRNode::PivotCache(cache) => Some(format!("PivotCache[{}]", cache.cache_id)),
        IRNode::PivotCacheRecords(records) => Some(format!(
            "PivotCacheRecords[{}]",
            records.record_count.unwrap_or(0)
        )),
        IRNode::WorkbookProperties(_) => Some("WorkbookProperties".to_string()),
        _ => None,
    }
}

fn intrinsic_key_word(node: &IRNode) -> Option<String> {
    match node {
        IRNode::Hyperlink(link) => Some(format!("Hyperlink[{}]", link.target)),
        IRNode::StyleSet(styles) => Some(format!("StyleSet[{}]", styles.styles.len())),
        IRNode::NumberingSet(nums) => Some(format!("NumberingSet[{}]", nums.abstract_nums.len())),
        IRNode::Comment(comment) => Some(format!("Comment[{}]", comment.comment_id)),
        IRNode::Footnote(note) => Some(format!("Footnote[{}]", note.footnote_id)),
        IRNode::Endnote(note) => Some(format!("Endnote[{}]", note.endnote_id)),
        IRNode::Header(_) => Some("Header".to_string()),
        IRNode::Footer(_) => Some("Footer".to_string()),
        IRNode::WordSettings(_) => Some("WordSettings".to_string()),
        IRNode::WebSettings(_) => Some("WebSettings".to_string()),
        IRNode::FontTable(_) => Some("FontTable".to_string()),
        IRNode::ContentControl(_) => Some("ContentControl".to_string()),
        IRNode::BookmarkStart(start) => Some(format!("BookmarkStart[{}]", start.bookmark_id)),
        IRNode::BookmarkEnd(end) => Some(format!("BookmarkEnd[{}]", end.bookmark_id)),
        IRNode::Field(field) => Some(format!(
            "Field[{}]",
            field.instruction.as_deref().unwrap_or("-")
        )),
        IRNode::Revision(rev) => Some(format!("Revision[{:?}]", rev.change_type)),
        IRNode::CommentExtensionSet(_) => Some("CommentExtensionSet".to_string()),
        IRNode::CommentIdMap(_) => Some("CommentIdMap".to_string()),
        IRNode::CommentRangeStart(start) => {
            Some(format!("CommentRangeStart[{}]", start.comment_id))
        }
        IRNode::CommentRangeEnd(end) => Some(format!("CommentRangeEnd[{}]", end.comment_id)),
        IRNode::CommentReference(reference) => {
            Some(format!("CommentReference[{}]", reference.comment_id))
        }
        IRNode::GlossaryDocument(_) => Some("GlossaryDocument".to_string()),
        IRNode::GlossaryEntry(entry) => Some(format!(
            "GlossaryEntry[{}]",
            entry.name.as_deref().unwrap_or("-")
        )),
        IRNode::VmlDrawing(drawing) => Some(format!("VmlDrawing[{}]", drawing.path)),
        IRNode::VmlShape(shape) => Some(format!(
            "VmlShape[{}]",
            shape.name.as_deref().unwrap_or("-")
        )),
        IRNode::DrawingPart(part) => Some(format!("DrawingPart[{}]", part.path)),
        _ => None,
    }
}

fn intrinsic_key_presentation(node: &IRNode) -> Option<String> {
    match node {
        IRNode::Slide(slide) => Some(format!("Slide[{}]", slide.number)),
        IRNode::Shape(shape) => shape.name.as_ref().map(|name| format!("Shape[{name}]")),
        IRNode::SlideMaster(_) => Some("SlideMaster".to_string()),
        IRNode::SlideLayout(_) => Some("SlideLayout".to_string()),
        IRNode::NotesMaster(_) => Some("NotesMaster".to_string()),
        IRNode::HandoutMaster(_) => Some("HandoutMaster".to_string()),
        IRNode::NotesSlide(_) => Some("NotesSlide".to_string()),
        IRNode::ChartData(_) => Some("ChartData".to_string()),
        IRNode::PresentationProperties(_) => Some("PresentationProperties".to_string()),
        IRNode::ViewProperties(_) => Some("ViewProperties".to_string()),
        IRNode::TableStyleSet(_) => Some("TableStyleSet".to_string()),
        IRNode::PptxCommentAuthor(author) => {
            Some(format!("PptxCommentAuthor[{}]", author.author_id))
        }
        IRNode::PptxComment(_) => Some("PptxComment".to_string()),
        IRNode::PresentationTag(tag) => Some(format!("PresentationTag[{}]", tag.name)),
        IRNode::PresentationInfo(_) => Some("PresentationInfo".to_string()),
        IRNode::PeoplePart(_) => Some("PeoplePart".to_string()),
        IRNode::SmartArtPart(part) => Some(format!("SmartArtPart[{}]", part.kind)),
        _ => None,
    }
}

fn intrinsic_key_security(node: &IRNode) -> Option<String> {
    match node {
        IRNode::ExternalReference(ext) => Some(format!("ExternalReference[{}]", ext.target)),
        IRNode::ActiveXControl(ctrl) => ctrl
            .prog_id
            .as_ref()
            .map(|p| format!("ActiveXControl[{p}]")),
        IRNode::MacroModule(module) => Some(format!("MacroModule[{}]", module.name)),
        IRNode::MacroProject(project) => project
            .name
            .as_ref()
            .map(|name| format!("MacroProject[{name}]")),
        IRNode::OleObject(ole) => {
            if let Some(name) = &ole.name {
                Some(format!("OleObject[{}]", name))
            } else {
                ole.prog_id
                    .as_ref()
                    .map(|prog_id| format!("OleObject[{prog_id}]"))
            }
        }
        IRNode::WebExtension(ext) => Some(format!(
            "WebExtension[{}]",
            ext.extension_id.as_deref().unwrap_or("-")
        )),
        IRNode::WebExtensionTaskpane(pane) => Some(format!(
            "WebExtensionTaskpane[{}]",
            pane.web_extension_ref.as_deref().unwrap_or("-")
        )),
        IRNode::Diagnostics(_) => Some("Diagnostics".to_string()),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::{
        BookmarkEnd, BookmarkStart, Cell, Comment, Diagnostics, Document, DrawingPart, Endnote,
        Footnote, GlossaryEntry, IRNode, Paragraph, PptxCommentAuthor, PresentationTag, Revision,
        RevisionType, Run, Shape, ShapeType, Slide, SmartArtPart, VmlDrawing, VmlShape,
        WebExtension, WebExtensionTaskpane, Worksheet,
    };
    use docir_core::security::{MacroProject, OleObject};
    use docir_core::types::DocumentFormat;
    use docir_core::visitor::IrStore;

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
    fn intrinsic_key_covers_word_presentation_and_security_variants() {
        let store = IrStore::new();

        assert_eq!(
            intrinsic_key(&IRNode::Comment(Comment::new("c1")), &store),
            Some("Comment[c1]".to_string())
        );
        assert_eq!(
            intrinsic_key(&IRNode::Footnote(Footnote::new("f1")), &store),
            Some("Footnote[f1]".to_string())
        );
        assert_eq!(
            intrinsic_key(&IRNode::Endnote(Endnote::new("e1")), &store),
            Some("Endnote[e1]".to_string())
        );
        assert_eq!(
            intrinsic_key(&IRNode::BookmarkStart(BookmarkStart::new("b1")), &store),
            Some("BookmarkStart[b1]".to_string())
        );
        assert_eq!(
            intrinsic_key(&IRNode::BookmarkEnd(BookmarkEnd::new("b1")), &store),
            Some("BookmarkEnd[b1]".to_string())
        );

        let mut glossary = GlossaryEntry::new();
        glossary.name = Some("entry".to_string());
        assert_eq!(
            intrinsic_key(&IRNode::GlossaryEntry(glossary), &store),
            Some("GlossaryEntry[entry]".to_string())
        );
        assert_eq!(
            intrinsic_key(&IRNode::VmlDrawing(VmlDrawing::new("word/vml.vml")), &store),
            Some("VmlDrawing[word/vml.vml]".to_string())
        );
        let mut vml_shape = VmlShape::new();
        vml_shape.name = Some("shape1".to_string());
        assert_eq!(
            intrinsic_key(&IRNode::VmlShape(vml_shape), &store),
            Some("VmlShape[shape1]".to_string())
        );
        assert_eq!(
            intrinsic_key(&IRNode::DrawingPart(DrawingPart::new("word/drawing1.xml")), &store),
            Some("DrawingPart[word/drawing1.xml]".to_string())
        );

        assert_eq!(
            intrinsic_key(&IRNode::Slide(Slide::new(2)), &store),
            Some("Slide[2]".to_string())
        );
        let mut shape = Shape::new(ShapeType::TextBox);
        shape.name = Some("Title".to_string());
        assert_eq!(
            intrinsic_key(&IRNode::Shape(shape), &store),
            Some("Shape[Title]".to_string())
        );
        assert_eq!(
            intrinsic_key(
                &IRNode::PptxCommentAuthor(PptxCommentAuthor {
                    id: docir_core::types::NodeId::new(),
                    author_id: 7,
                    name: None,
                    initials: None,
                    span: None,
                }),
                &store
            ),
            Some("PptxCommentAuthor[7]".to_string())
        );
        assert_eq!(
            intrinsic_key(
                &IRNode::PresentationTag(PresentationTag {
                    id: docir_core::types::NodeId::new(),
                    name: "tag".to_string(),
                    value: None,
                    span: None,
                }),
                &store
            ),
            Some("PresentationTag[tag]".to_string())
        );
        assert_eq!(
            intrinsic_key(
                &IRNode::SmartArtPart(SmartArtPart {
                    id: docir_core::types::NodeId::new(),
                    kind: "diagramData".to_string(),
                    path: "ppt/diagrams/data1.xml".to_string(),
                    root_element: None,
                    point_count: None,
                    connection_count: None,
                    rel_ids: Vec::new(),
                    span: None,
                }),
                &store
            ),
            Some("SmartArtPart[diagramData]".to_string())
        );

        let mut project = MacroProject::new();
        project.name = Some("VBAProject".to_string());
        assert_eq!(
            intrinsic_key(&IRNode::MacroProject(project), &store),
            Some("MacroProject[VBAProject]".to_string())
        );
        let mut ole = OleObject::new();
        ole.name = Some("Obj".to_string());
        assert_eq!(
            intrinsic_key(&IRNode::OleObject(ole), &store),
            Some("OleObject[Obj]".to_string())
        );
        let mut ext = WebExtension::new();
        ext.extension_id = Some("ext-1".to_string());
        assert_eq!(
            intrinsic_key(&IRNode::WebExtension(ext), &store),
            Some("WebExtension[ext-1]".to_string())
        );
        let mut taskpane = WebExtensionTaskpane::new();
        taskpane.web_extension_ref = Some("ext-1".to_string());
        assert_eq!(
            intrinsic_key(&IRNode::WebExtensionTaskpane(taskpane), &store),
            Some("WebExtensionTaskpane[ext-1]".to_string())
        );
        assert_eq!(
            intrinsic_key(&IRNode::Diagnostics(Diagnostics::new()), &store),
            Some("Diagnostics".to_string())
        );

        let revision = Revision::new(RevisionType::Insert);
        assert_eq!(
            intrinsic_key(&IRNode::Revision(revision), &store),
            Some("Revision[Insert]".to_string())
        );
    }
}
