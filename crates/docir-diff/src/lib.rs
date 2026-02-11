//! # docir-diff
//!
//! Structural and semantic diffing for docir IR.

use docir_core::ir::IrNode as IrNodeTrait;
use docir_core::ir::{
    Cell, CellFormula, CellValue, Document, Hyperlink, IRNode, Paragraph, Run, Section, Shape,
    Slide, Worksheet,
};
use docir_core::security::{ExternalReference, MacroModule, MacroProject, OleObject};
use docir_core::types::{NodeId, NodeType};
use docir_core::visitor::IrStore;
use serde::{Deserialize, Serialize};
use sha2::Digest;
use std::collections::{BTreeMap, HashMap, HashSet};

/// Diff result between two IR trees.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiffResult {
    /// Nodes only present in the right-hand document.
    pub added: Vec<NodeChange>,
    /// Nodes only present in the left-hand document.
    pub removed: Vec<NodeChange>,
    /// Nodes present in both but with differing summaries.
    pub modified: Vec<NodeModification>,
}

impl DiffResult {
    /// Returns true if there are no differences.
    pub fn is_empty(&self) -> bool {
        self.added.is_empty() && self.removed.is_empty() && self.modified.is_empty()
    }
}

/// A node that was added or removed.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NodeChange {
    pub key: String,
    pub node_type: NodeType,
    pub summary: String,
}

/// A node that was modified.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NodeModification {
    pub key: String,
    pub node_type: NodeType,
    pub before: String,
    pub after: String,
    pub change_kind: ChangeKind,
}

/// Kind of change detected between two matched nodes.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum ChangeKind {
    Content,
    Style,
    Both,
    Metadata,
}

/// Diff engine for docir IR.
pub struct DiffEngine;

impl DiffEngine {
    /// Computes a diff between two IR trees.
    pub fn diff(
        left: &IrStore,
        left_root: NodeId,
        right: &IrStore,
        right_root: NodeId,
    ) -> DiffResult {
        let left_index = build_index(left, left_root);
        let right_index = build_index(right, right_root);

        let mut added = Vec::new();
        let mut removed = Vec::new();
        let mut modified = Vec::new();

        let mut keys: BTreeMap<&String, ()> = BTreeMap::new();
        for key in left_index.keys() {
            keys.insert(key, ());
        }
        for key in right_index.keys() {
            keys.insert(key, ());
        }

        for key in keys.keys() {
            match (left_index.get(*key), right_index.get(*key)) {
                (Some(left_snap), Some(right_snap)) => {
                    if left_snap.summary != right_snap.summary {
                        let content_same = left_snap.content_sig == right_snap.content_sig;
                        let style_same = left_snap.style_sig == right_snap.style_sig;
                        let change_kind = if content_same && !style_same {
                            ChangeKind::Style
                        } else if !content_same && style_same {
                            ChangeKind::Content
                        } else if !content_same && !style_same {
                            ChangeKind::Both
                        } else {
                            ChangeKind::Metadata
                        };
                        modified.push(NodeModification {
                            key: (*key).clone(),
                            node_type: left_snap.node_type,
                            before: left_snap.summary.clone(),
                            after: right_snap.summary.clone(),
                            change_kind,
                        });
                    }
                }
                (Some(left_snap), None) => {
                    removed.push(NodeChange {
                        key: (*key).clone(),
                        node_type: left_snap.node_type,
                        summary: left_snap.summary.clone(),
                    });
                }
                (None, Some(right_snap)) => {
                    added.push(NodeChange {
                        key: (*key).clone(),
                        node_type: right_snap.node_type,
                        summary: right_snap.summary.clone(),
                    });
                }
                (None, None) => {}
            }
        }

        DiffResult {
            added,
            removed,
            modified,
        }
    }
}

#[derive(Debug, Clone)]
struct NodeSnapshot {
    node_type: NodeType,
    summary: String,
    content_sig: Option<String>,
    style_sig: Option<String>,
}

fn build_index(store: &IrStore, root: NodeId) -> BTreeMap<String, NodeSnapshot> {
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
        IRNode::Slide(slide) => Some(format!("Slide[{}]", slide.number)),
        IRNode::Shape(shape) => {
            if let Some(name) = &shape.name {
                Some(format!("Shape[{}]", name))
            } else {
                None
            }
        }
        IRNode::Hyperlink(link) => Some(format!("Hyperlink[{}]", link.target)),
        IRNode::ExternalReference(ext) => Some(format!("ExternalReference[{}]", ext.target)),
        IRNode::ActiveXControl(ctrl) => ctrl
            .prog_id
            .as_ref()
            .map(|p| format!("ActiveXControl[{p}]")),
        IRNode::MacroModule(module) => Some(format!("MacroModule[{}]", module.name)),
        IRNode::MacroProject(project) => project
            .name
            .as_ref()
            .map(|name| format!("MacroProject[{}]", name)),
        IRNode::OleObject(ole) => {
            if let Some(name) = &ole.name {
                Some(format!("OleObject[{}]", name))
            } else if let Some(prog_id) = &ole.prog_id {
                Some(format!("OleObject[{}]", prog_id))
            } else {
                None
            }
        }
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
        IRNode::SlideMaster(_) => Some("SlideMaster".to_string()),
        IRNode::SlideLayout(_) => Some("SlideLayout".to_string()),
        IRNode::NotesMaster(_) => Some("NotesMaster".to_string()),
        IRNode::HandoutMaster(_) => Some("HandoutMaster".to_string()),
        IRNode::NotesSlide(_) => Some("NotesSlide".to_string()),
        IRNode::WorksheetDrawing(_) => Some("WorksheetDrawing".to_string()),
        IRNode::ChartData(_) => Some("ChartData".to_string()),
        IRNode::CalcChain(_) => Some("CalcChain".to_string()),
        IRNode::SheetComment(comment) => Some(format!("SheetComment[{}]", comment.cell_ref)),
        IRNode::SheetMetadata(_) => Some("SheetMetadata".to_string()),
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
        IRNode::WebExtension(ext) => Some(format!(
            "WebExtension[{}]",
            ext.extension_id.as_deref().unwrap_or("-")
        )),
        IRNode::WebExtensionTaskpane(pane) => Some(format!(
            "WebExtensionTaskpane[{}]",
            pane.web_extension_ref.as_deref().unwrap_or("-")
        )),
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
        IRNode::Diagnostics(_) => Some("Diagnostics".to_string()),
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
        IRNode::Paragraph(_para) => None,
        _ => None,
    }
}

fn summarize(node: &IRNode, store: &IrStore) -> String {
    match node {
        IRNode::Document(doc) => summarize_document(doc),
        IRNode::Section(section) => summarize_section(section),
        IRNode::Paragraph(para) => summarize_paragraph(para, store),
        IRNode::Run(run) => summarize_run(run),
        IRNode::Hyperlink(link) => summarize_hyperlink(link, store),
        IRNode::Table(table) => format!(
            "rows={} cols={} style={}",
            table.rows.len(),
            table.grid.len(),
            opt_str(&table.properties.style_id)
        ),
        IRNode::TableRow(row) => format!("cells={}", row.cells.len()),
        IRNode::TableCell(cell) => format!(
            "content_nodes={} span={}",
            cell.content.len(),
            cell.properties.grid_span.unwrap_or(1)
        ),
        IRNode::Worksheet(ws) => summarize_worksheet(ws),
        IRNode::Cell(cell) => summarize_cell(cell),
        IRNode::SharedStringTable(table) => format!("items={}", table.items.len()),
        IRNode::ConnectionPart(part) => format!("connections={}", part.entries.len()),
        IRNode::SpreadsheetStyles(styles) => {
            format!(
                "numfmts={} fonts={} fills={} borders={} xfs={} style_xfs={} dxfs={} table_styles={}",
                styles.number_formats.len(),
                styles.fonts.len(),
                styles.fills.len(),
                styles.borders.len(),
                styles.cell_xfs.len(),
                styles.cell_style_xfs.len(),
                styles.dxfs.len(),
                styles.table_styles
                    .as_ref()
                    .and_then(|t| t.default_table_style.as_ref())
                    .map_or("-", |v| v.as_str())
            )
        }
        IRNode::DefinedName(name) => format!(
            "name={} scope={}",
            name.name,
            name.local_sheet_id
                .map_or("global".to_string(), |id| format!("sheet:{id}"))
        ),
        IRNode::ConditionalFormat(fmt) => {
            format!("ranges={} rules={}", fmt.ranges.len(), fmt.rules.len())
        }
        IRNode::DataValidation(val) => format!(
            "ranges={} type={}",
            val.ranges.len(),
            opt_str(&val.validation_type)
        ),
        IRNode::TableDefinition(table) => {
            format!(
                "name={} cols={} ref={}",
                opt_str(&table.display_name),
                table.columns.len(),
                opt_str(&table.ref_range)
            )
        }
        IRNode::PivotTable(pivot) => format!(
            "name={} cache_id={}",
            opt_str(&pivot.name),
            opt_u32(pivot.cache_id)
        ),
        IRNode::PivotCache(cache) => format!(
            "cache_id={} source={}",
            cache.cache_id,
            opt_str(&cache.cache_source)
        ),
        IRNode::PivotCacheRecords(records) => format!(
            "records={} fields={}",
            records.record_count.unwrap_or(0),
            records.field_count.unwrap_or(0)
        ),
        IRNode::WorkbookProperties(props) => format!(
            "calc_mode={} protected={}",
            opt_str(&props.calc_mode),
            props.workbook_protected
        ),
        IRNode::Slide(slide) => summarize_slide(slide),
        IRNode::Shape(shape) => summarize_shape(shape),
        IRNode::MacroProject(project) => summarize_macro_project(project),
        IRNode::MacroModule(module) => summarize_macro_module(module),
        IRNode::OleObject(ole) => summarize_ole(ole),
        IRNode::ExternalReference(ext) => summarize_external_ref(ext),
        IRNode::ActiveXControl(ctrl) => format!(
            "name={} clsid={} prog_id={}",
            opt_str(&ctrl.name),
            opt_str(&ctrl.clsid),
            opt_str(&ctrl.prog_id)
        ),
        IRNode::Metadata(meta) => format!(
            "title={} author={}",
            opt_str(&meta.title),
            opt_str(&meta.creator)
        ),
        IRNode::Theme(theme) => format!(
            "name={} colors={} fonts={}",
            opt_str(&theme.name),
            theme.colors.len(),
            theme.fonts.major.as_deref().unwrap_or("-")
        ),
        IRNode::MediaAsset(media) => format!(
            "path={} type={:?} size={}",
            media.path, media.media_type, media.size_bytes
        ),
        IRNode::CustomXmlPart(part) => {
            format!("path={} root={}", part.path, opt_str(&part.root_element))
        }
        IRNode::RelationshipGraph(graph) => {
            format!("source={} rels={}", graph.source, graph.relationships.len())
        }
        IRNode::DigitalSignature(sig) => format!(
            "id={} method={}",
            opt_str(&sig.signature_id),
            opt_str(&sig.signature_method)
        ),
        IRNode::ExtensionPart(part) => format!(
            "path={} kind={:?} size={}",
            part.path, part.kind, part.size_bytes
        ),
        IRNode::StyleSet(styles) => format!("styles={}", styles.styles.len()),
        IRNode::NumberingSet(nums) => format!(
            "abstracts={} nums={}",
            nums.abstract_nums.len(),
            nums.nums.len()
        ),
        IRNode::Comment(comment) => format!(
            "id={} author={} content_nodes={}",
            comment.comment_id,
            opt_str(&comment.author),
            comment.content.len()
        ),
        IRNode::Footnote(note) => format!(
            "id={} content_nodes={}",
            note.footnote_id,
            note.content.len()
        ),
        IRNode::Endnote(note) => format!(
            "id={} content_nodes={}",
            note.endnote_id,
            note.content.len()
        ),
        IRNode::Header(header) => format!("content_nodes={}", header.content.len()),
        IRNode::Footer(footer) => format!("content_nodes={}", footer.content.len()),
        IRNode::WordSettings(settings) => format!("entries={}", settings.entries.len()),
        IRNode::WebSettings(settings) => format!("entries={}", settings.entries.len()),
        IRNode::FontTable(table) => format!("fonts={}", table.fonts.len()),
        IRNode::ContentControl(control) => format!(
            "content_nodes={} tag={}",
            control.content.len(),
            opt_str(&control.tag)
        ),
        IRNode::BookmarkStart(start) => {
            format!("id={} name={}", start.bookmark_id, opt_str(&start.name))
        }
        IRNode::BookmarkEnd(end) => format!("id={}", end.bookmark_id),
        IRNode::Field(field) => format!(
            "runs={} instr={}",
            field.runs.len(),
            opt_str(&field.instruction)
        ),
        IRNode::Revision(rev) => format!(
            "type={:?} content_nodes={}",
            rev.change_type,
            rev.content.len()
        ),
        IRNode::CommentExtensionSet(set) => format!("entries={}", set.entries.len()),
        IRNode::CommentIdMap(map) => format!("mappings={}", map.mappings.len()),
        IRNode::CommentRangeStart(start) => format!("comment_id={}", start.comment_id),
        IRNode::CommentRangeEnd(end) => format!("comment_id={}", end.comment_id),
        IRNode::CommentReference(reference) => format!("comment_id={}", reference.comment_id),
        IRNode::SlideMaster(master) => format!(
            "shapes={} layouts={}",
            master.shapes.len(),
            master.layouts.len()
        ),
        IRNode::SlideLayout(layout) => format!("shapes={}", layout.shapes.len()),
        IRNode::NotesMaster(master) => format!("shapes={}", master.shapes.len()),
        IRNode::HandoutMaster(master) => format!("shapes={}", master.shapes.len()),
        IRNode::NotesSlide(slide) => format!(
            "shapes={} text={}",
            slide.shapes.len(),
            opt_str(&slide.text)
        ),
        IRNode::WorksheetDrawing(d) => format!("shapes={}", d.shapes.len()),
        IRNode::ChartData(c) => format!(
            "type={} series={} series_data={}",
            opt_str(&c.chart_type),
            c.series.len(),
            c.series_data.len()
        ),
        IRNode::CalcChain(chain) => format!("entries={}", chain.entries.len()),
        IRNode::SheetComment(comment) => format!(
            "cell_ref={} author={} text={}",
            comment.cell_ref,
            opt_str(&comment.author),
            abbreviate(&comment.text, 80)
        ),
        IRNode::SheetMetadata(meta) => format!(
            "types={} cell_count={} value_count={}",
            meta.metadata_types.len(),
            opt_u32(meta.cell_metadata_count),
            opt_u32(meta.value_metadata_count)
        ),
        IRNode::PresentationProperties(props) => format!(
            "auto_compress={} compat={} rtl={}",
            opt_bool(props.auto_compress_pictures),
            opt_str(&props.compat_mode),
            opt_bool(props.rtl)
        ),
        IRNode::ViewProperties(props) => format!(
            "last_view={} zoom={}",
            opt_str(&props.last_view),
            opt_u32(props.zoom)
        ),
        IRNode::TableStyleSet(styles) => format!(
            "default={} styles={}",
            opt_str(&styles.default_style_id),
            styles.styles.len()
        ),
        IRNode::PptxCommentAuthor(author) => format!(
            "author_id={} name={}",
            author.author_id,
            opt_str(&author.name)
        ),
        IRNode::PptxComment(comment) => format!(
            "author_id={} text={}",
            comment
                .author_id
                .map(|v| v.to_string())
                .unwrap_or_else(|| "-".to_string()),
            abbreviate(&comment.text, 80)
        ),
        IRNode::PresentationTag(tag) => format!("name={} value={}", tag.name, opt_str(&tag.value)),
        IRNode::PresentationInfo(info) => format!(
            "slide_size={} notes_size={} show_type={}",
            info.slide_size
                .as_ref()
                .map(|s| format!("{}x{}", s.cx, s.cy))
                .unwrap_or_else(|| "-".to_string()),
            info.notes_size
                .as_ref()
                .map(|s| format!("{}x{}", s.cx, s.cy))
                .unwrap_or_else(|| "-".to_string()),
            opt_str(&info.show_type)
        ),
        IRNode::PeoplePart(people) => format!("people={}", people.people.len()),
        IRNode::SmartArtPart(part) => format!("kind={} path={}", part.kind, part.path),
        IRNode::WebExtension(ext) => format!(
            "id={} store={} version={} properties={}",
            opt_str(&ext.extension_id),
            opt_str(&ext.store),
            opt_str(&ext.version),
            ext.properties.len()
        ),
        IRNode::WebExtensionTaskpane(pane) => format!(
            "ref={} dock_state={} visible={}",
            opt_str(&pane.web_extension_ref),
            opt_str(&pane.dock_state),
            opt_bool(pane.visibility)
        ),
        IRNode::GlossaryDocument(doc) => format!("entries={}", doc.entries.len()),
        IRNode::GlossaryEntry(entry) => format!(
            "name={} gallery={} content_nodes={}",
            opt_str(&entry.name),
            opt_str(&entry.gallery),
            entry.content.len()
        ),
        IRNode::VmlDrawing(drawing) => {
            format!("path={} shapes={}", drawing.path, drawing.shapes.len())
        }
        IRNode::VmlShape(shape) => format!(
            "name={} rel_id={} image_target={}",
            opt_str(&shape.name),
            opt_str(&shape.rel_id),
            opt_str(&shape.image_target)
        ),
        IRNode::DrawingPart(part) => format!("path={} shapes={}", part.path, part.shapes.len()),
        IRNode::ExternalLinkPart(part) => format!(
            "type={} target={} sheets={}",
            opt_str(&part.link_type),
            opt_str(&part.target),
            part.sheets.len()
        ),
        IRNode::SlicerPart(part) => format!(
            "name={} caption={} cache_id={}",
            opt_str(&part.name),
            opt_str(&part.caption),
            opt_str(&part.cache_id)
        ),
        IRNode::TimelinePart(part) => format!(
            "name={} cache_id={}",
            opt_str(&part.name),
            opt_str(&part.cache_id)
        ),
        IRNode::QueryTablePart(part) => format!(
            "name={} connection_id={} url={}",
            opt_str(&part.name),
            opt_str(&part.connection_id),
            opt_str(&part.url)
        ),
        IRNode::Diagnostics(diag) => format!("entries={}", diag.entries.len()),
    }
}

fn content_signature(node: &IRNode, store: &IrStore) -> Option<String> {
    match node {
        IRNode::Paragraph(para) => Some(text_from_paragraph(para, store)),
        IRNode::Run(run) => Some(run.text.clone()),
        IRNode::Hyperlink(link) => Some(link.target.clone()),
        IRNode::Cell(cell) => Some(cell_content_signature(cell)),
        IRNode::Worksheet(ws) => Some(worksheet_content_signature(ws, store)),
        IRNode::Shape(shape) => shape.text.as_ref().map(shape_text),
        IRNode::MacroModule(module) => Some(module.name.clone()),
        IRNode::MacroProject(project) => project.name.clone(),
        IRNode::ExternalReference(ext) => Some(ext.target.clone()),
        IRNode::OleObject(ole) => ole.prog_id.clone().or_else(|| ole.name.clone()),
        IRNode::ActiveXControl(ctrl) => ctrl.prog_id.clone().or_else(|| ctrl.name.clone()),
        IRNode::DefinedName(def) => Some(def.name.clone()),
        IRNode::TableDefinition(table) => table.display_name.clone().or_else(|| table.name.clone()),
        _ => None,
    }
}

fn style_signature(node: &IRNode, _store: &IrStore) -> Option<String> {
    match node {
        IRNode::Paragraph(para) => {
            Some(serde_json::to_string(&para.properties).unwrap_or_default())
        }
        IRNode::Run(run) => Some(serde_json::to_string(&run.properties).unwrap_or_default()),
        IRNode::Table(table) => Some(serde_json::to_string(&table.properties).unwrap_or_default()),
        IRNode::Cell(cell) => Some(format!(
            "style={}",
            cell.style_id.map_or("-".to_string(), |id| id.to_string())
        )),
        IRNode::Worksheet(ws) => Some(format!("state={:?} kind={:?}", ws.state, ws.kind)),
        IRNode::Shape(shape) => Some(format!(
            "type={:?} has_text={}",
            shape.shape_type,
            shape.text.is_some()
        )),
        IRNode::Slide(slide) => Some(format!(
            "layout_id={} master_id={}",
            opt_str(&slide.layout_id),
            opt_str(&slide.master_id)
        )),
        _ => None,
    }
}

fn text_from_paragraph(para: &Paragraph, store: &IrStore) -> String {
    let mut out = String::new();
    for run_id in &para.runs {
        if let Some(IRNode::Run(run)) = store.get(*run_id) {
            if !run.text.is_empty() {
                out.push_str(&run.text);
            }
        }
    }
    out
}

fn cell_content_signature(cell: &Cell) -> String {
    let mut out = String::new();
    out.push_str(&cell.reference);
    out.push('=');
    out.push_str(&cell_value_summary(&cell.value));
    if let Some(formula) = &cell.formula {
        out.push_str(";");
        out.push_str(&cell_formula_summary(formula));
    }
    out
}

fn worksheet_content_signature(ws: &Worksheet, store: &IrStore) -> String {
    let mut entries: Vec<String> = ws
        .cells
        .iter()
        .filter_map(|id| store.get(*id))
        .filter_map(|node| {
            if let IRNode::Cell(cell) = node {
                Some(cell_content_signature(cell))
            } else {
                None
            }
        })
        .collect();
    entries.sort();
    let joined = entries.join("|");
    short_hash(&joined)
}

fn cell_value_summary(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "empty".to_string(),
        CellValue::Number(n) => format!("n:{n}"),
        CellValue::Boolean(b) => format!("b:{b}"),
        CellValue::String(s) => format!("s:{s}"),
        CellValue::InlineString(s) => format!("is:{s}"),
        CellValue::SharedString(idx) => format!("ss:{idx}"),
        CellValue::Error(err) => format!("e:{err:?}"),
        CellValue::DateTime(dt) => format!("dt:{dt}"),
    }
}

fn cell_formula_summary(formula: &CellFormula) -> String {
    formula.text.clone()
}

fn short_hash(input: &str) -> String {
    let mut hasher = sha2::Sha256::new();
    hasher.update(input.as_bytes());
    let hash = hasher.finalize();
    to_hex(&hash[..8])
}

fn to_hex(data: &[u8]) -> String {
    data.iter().map(|b| format!("{:02x}", b)).collect()
}

fn summarize_document(doc: &Document) -> String {
    format!(
        "format={:?} content_nodes={} macros={} ole={} external_refs={} threat={:?}",
        doc.format,
        doc.content.len(),
        doc.security.macro_project.is_some(),
        doc.security.ole_objects.len(),
        doc.security.external_refs.len(),
        doc.security.threat_level,
    )
}

fn summarize_section(section: &Section) -> String {
    format!(
        "name={} content_nodes={} columns={} orientation={:?}",
        opt_str(&section.name),
        section.content.len(),
        section.properties.columns.unwrap_or(1),
        section.properties.orientation,
    )
}

fn summarize_paragraph(para: &Paragraph, store: &IrStore) -> String {
    let text = paragraph_text(para, store);
    format!(
        "style={} runs={} text=\"{}\"",
        opt_str(&para.style_id),
        para.runs.len(),
        abbreviate(&text, 80)
    )
}

fn summarize_run(run: &Run) -> String {
    format!(
        "text=\"{}\" bold={} italic={} size={}",
        abbreviate(&run.text, 80),
        opt_bool(run.properties.bold),
        opt_bool(run.properties.italic),
        run.properties
            .font_size
            .map(|v| v.to_string())
            .unwrap_or_else(|| "-".to_string()),
    )
}

fn summarize_hyperlink(link: &Hyperlink, store: &IrStore) -> String {
    let text = runs_text(&link.runs, store);
    format!(
        "target={} external={} runs={} text=\"{}\"",
        link.target,
        link.is_external,
        link.runs.len(),
        abbreviate(&text, 80)
    )
}

fn summarize_worksheet(ws: &Worksheet) -> String {
    format!(
        "name={} sheet_id={} state={:?} cells={} merged={}",
        ws.name,
        ws.sheet_id,
        ws.state,
        ws.cells.len(),
        ws.merged_cells.len(),
    )
}

fn summarize_cell(cell: &Cell) -> String {
    let value = match &cell.value {
        CellValue::Empty => "empty".to_string(),
        CellValue::String(v) => format!("str:{}", abbreviate(v, 60)),
        CellValue::Number(v) => format!("num:{}", format_float(*v)),
        CellValue::Boolean(v) => format!("bool:{}", v),
        CellValue::Error(e) => format!("error:{:?}", e),
        CellValue::DateTime(v) => format!("date:{}", format_float(*v)),
        CellValue::InlineString(v) => format!("inline:{}", abbreviate(v, 60)),
        CellValue::SharedString(i) => format!("shared:{}", i),
    };
    let formula = cell
        .formula
        .as_ref()
        .map(summarize_formula)
        .unwrap_or_else(|| "-".to_string());
    format!("ref={} value={} formula={}", cell.reference, value, formula)
}

fn summarize_formula(formula: &CellFormula) -> String {
    format!(
        "{} type={:?}",
        abbreviate(&formula.text, 80),
        formula.formula_type,
    )
}

fn summarize_slide(slide: &Slide) -> String {
    format!(
        "number={} name={} shapes={} hidden={}",
        slide.number,
        opt_str(&slide.name),
        slide.shapes.len(),
        slide.hidden,
    )
}

fn summarize_shape(shape: &Shape) -> String {
    let text = shape.text.as_ref().map(shape_text).unwrap_or_default();
    format!(
        "type={:?} name={} text=\"{}\" x={} y={} w={} h={} link={}",
        shape.shape_type,
        opt_str(&shape.name),
        abbreviate(&text, 80),
        shape.transform.x,
        shape.transform.y,
        shape.transform.width,
        shape.transform.height,
        opt_str(&shape.hyperlink),
    )
}

fn summarize_macro_project(project: &MacroProject) -> String {
    format!(
        "name={} modules={} auto_exec={} protected={}",
        opt_str(&project.name),
        project.modules.len(),
        project.has_auto_exec,
        project.is_protected,
    )
}

fn summarize_macro_module(module: &MacroModule) -> String {
    format!(
        "name={} module_type={:?} suspicious_calls={}",
        module.name,
        module.module_type,
        module.suspicious_calls.len(),
    )
}

fn summarize_ole(ole: &OleObject) -> String {
    format!(
        "name={} prog_id={} linked={} size={} hash={}",
        opt_str(&ole.name),
        opt_str(&ole.prog_id),
        ole.is_linked,
        ole.size_bytes,
        opt_str(&ole.data_hash),
    )
}

fn summarize_external_ref(ext: &ExternalReference) -> String {
    format!("type={:?} target={}", ext.ref_type, ext.target,)
}

fn paragraph_text(para: &Paragraph, store: &IrStore) -> String {
    let mut out = String::new();
    for run_id in &para.runs {
        if let Some(node) = store.get(*run_id) {
            match node {
                IRNode::Run(run) => out.push_str(&run.text),
                IRNode::Hyperlink(link) => out.push_str(&runs_text(&link.runs, store)),
                _ => {}
            }
        }
    }
    out
}

fn runs_text(run_ids: &[NodeId], store: &IrStore) -> String {
    let mut out = String::new();
    for run_id in run_ids {
        if let Some(IRNode::Run(run)) = store.get(*run_id) {
            out.push_str(&run.text);
        }
    }
    out
}

fn shape_text(text: &docir_core::ir::ShapeText) -> String {
    let mut out = String::new();
    for (p_idx, para) in text.paragraphs.iter().enumerate() {
        if p_idx > 0 {
            out.push('\n');
        }
        for run in &para.runs {
            out.push_str(&run.text);
        }
    }
    out
}

fn opt_str(value: &Option<String>) -> String {
    value.clone().unwrap_or_else(|| "-".to_string())
}

fn opt_bool(value: Option<bool>) -> String {
    value
        .map(|v| if v { "true" } else { "false" }.to_string())
        .unwrap_or_else(|| "-".to_string())
}

fn opt_u32(value: Option<u32>) -> String {
    value
        .map(|v| v.to_string())
        .unwrap_or_else(|| "-".to_string())
}

fn abbreviate(value: &str, max: usize) -> String {
    if value.len() <= max {
        return value.to_string();
    }
    let mut out = value.chars().take(max).collect::<String>();
    out.push_str("...");
    out
}

fn format_float(value: f64) -> String {
    if value.fract() == 0.0 {
        format!("{:.0}", value)
    } else {
        format!("{:.4}", value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_parser::DocumentParser;
    use std::io::{Cursor, Write};
    use zip::write::FileOptions;

    fn build_odf_zip(content_xml: &str) -> Vec<u8> {
        let mut buffer = Vec::new();
        let cursor = Cursor::new(&mut buffer);
        let mut zip = zip::ZipWriter::new(cursor);
        let stored =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(b"application/vnd.oasis.opendocument.text")
            .unwrap();

        zip.start_file("META-INF/manifest.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"#,
        )
        .unwrap();

        zip.start_file("content.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(content_xml.as_bytes()).unwrap();

        zip.start_file("meta.xml", FileOptions::<()>::default())
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
  <office:meta>
    <dc:title>Diff</dc:title>
  </office:meta>
</office:document-meta>
"#,
        )
        .unwrap();

        zip.finish().unwrap();
        buffer
    }

    #[test]
    fn test_diff_odt_modified_text() {
        let content_left = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Hello</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
        let content_right = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:text>
      <text:p>Hello world</text:p>
    </office:text>
  </office:body>
</office:document-content>
"#;
        let left_zip = build_odf_zip(content_left);
        let right_zip = build_odf_zip(content_right);

        let parser = DocumentParser::new();
        let left_doc = parser.parse_reader(Cursor::new(left_zip)).unwrap();
        let right_doc = parser.parse_reader(Cursor::new(right_zip)).unwrap();

        let diff = DiffEngine::diff(
            &left_doc.store,
            left_doc.root_id,
            &right_doc.store,
            right_doc.root_id,
        );

        assert!(diff
            .modified
            .iter()
            .any(|m| m.change_kind == ChangeKind::Content));
    }
}
