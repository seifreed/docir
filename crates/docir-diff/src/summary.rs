use docir_core::ir::{
    BookmarkEnd, BookmarkStart, Cell, CellFormula, CellValue, Comment, ContentControl, Document,
    DrawingPart, Endnote, Field, Footnote, GlossaryEntry, Hyperlink, IRNode, Paragraph, Revision,
    Run, Section, Shape, Slide, Table, TableCell, TableRow, VmlDrawing, VmlShape, WebExtension,
    WebExtensionTaskpane, Worksheet,
};
use docir_core::security::{ExternalReference, MacroModule, MacroProject, OleObject};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use sha2::Digest;

mod presentation;
mod spreadsheet;

pub(crate) fn summarize(node: &IRNode, store: &IrStore) -> String {
    if let Some(summary) = spreadsheet::summarize(node, store) {
        return summary;
    }
    if let Some(summary) = presentation::summarize(node, store) {
        return summary;
    }
    summarize_with_fallback(node, store)
}

fn summarize_with_fallback(node: &IRNode, store: &IrStore) -> String {
    summarize_primary(node, store).unwrap_or_else(|| summarize_secondary(node))
}

fn summarize_primary(node: &IRNode, store: &IrStore) -> Option<String> {
    match node {
        IRNode::Document(doc) => Some(summarize_document(doc)),
        IRNode::Section(section) => Some(summarize_section(section)),
        IRNode::Paragraph(para) => Some(summarize_paragraph(para, store)),
        IRNode::Run(run) => Some(summarize_run(run)),
        IRNode::Hyperlink(link) => Some(summarize_hyperlink(link, store)),
        IRNode::Table(table) => Some(summarize_table(table)),
        IRNode::TableRow(row) => Some(summarize_table_row(row)),
        IRNode::TableCell(cell) => Some(summarize_table_cell(cell)),
        IRNode::MacroProject(project) => Some(summarize_macro_project(project)),
        IRNode::MacroModule(module) => Some(summarize_macro_module(module)),
        IRNode::OleObject(ole) => Some(summarize_ole(ole)),
        IRNode::ExternalReference(ext) => Some(summarize_external_ref(ext)),
        IRNode::ActiveXControl(ctrl) => Some(format!(
            "name={} clsid={} prog_id={}",
            opt_str(&ctrl.name),
            opt_str(&ctrl.clsid),
            opt_str(&ctrl.prog_id)
        )),
        IRNode::Metadata(meta) => Some(format!(
            "title={} author={}",
            opt_str(&meta.title),
            opt_str(&meta.creator)
        )),
        IRNode::Theme(theme) => Some(format!(
            "name={} colors={} fonts={}",
            opt_str(&theme.name),
            theme.colors.len(),
            theme.fonts.major.as_deref().unwrap_or("-")
        )),
        IRNode::MediaAsset(media) => Some(format!(
            "path={} type={:?} size={}",
            media.path, media.media_type, media.size_bytes
        )),
        IRNode::CustomXmlPart(part) => Some(format!(
            "path={} root={}",
            part.path,
            opt_str(&part.root_element)
        )),
        IRNode::RelationshipGraph(graph) => Some(format!(
            "source={} rels={}",
            graph.source,
            graph.relationships.len()
        )),
        IRNode::DigitalSignature(sig) => Some(format!(
            "id={} method={}",
            opt_str(&sig.signature_id),
            opt_str(&sig.signature_method)
        )),
        IRNode::ExtensionPart(part) => Some(format!(
            "path={} kind={:?} size={}",
            part.path, part.kind, part.size_bytes
        )),
        IRNode::StyleSet(styles) => Some(format!("styles={}", styles.styles.len())),
        IRNode::NumberingSet(nums) => Some(format!(
            "abstracts={} nums={}",
            nums.abstract_nums.len(),
            nums.nums.len()
        )),
        IRNode::Comment(comment) => Some(summarize_comment(comment)),
        IRNode::Footnote(note) => Some(summarize_footnote(note)),
        IRNode::Endnote(note) => Some(summarize_endnote(note)),
        IRNode::Header(header) => Some(summarize_header_footer(header.content.len())),
        IRNode::Footer(footer) => Some(summarize_header_footer(footer.content.len())),
        IRNode::WordSettings(settings) => Some(format!("entries={}", settings.entries.len())),
        IRNode::WebSettings(settings) => Some(format!("entries={}", settings.entries.len())),
        IRNode::FontTable(table) => Some(format!("fonts={}", table.fonts.len())),
        IRNode::ContentControl(control) => Some(summarize_content_control(control)),
        IRNode::BookmarkStart(start) => Some(summarize_bookmark_start(start)),
        IRNode::BookmarkEnd(end) => Some(summarize_bookmark_end(end)),
        IRNode::Field(field) => Some(summarize_field(field)),
        IRNode::Revision(rev) => Some(summarize_revision(rev)),
        _ => None,
    }
}

fn summarize_secondary(node: &IRNode) -> String {
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
        _ => unreachable!("summary missing for node"),
    }
}

fn summarize_table(table: &Table) -> String {
    format!(
        "rows={} cols={} style={}",
        table.rows.len(),
        table.grid.len(),
        opt_str(&table.properties.style_id)
    )
}

fn summarize_table_row(row: &TableRow) -> String {
    format!("cells={}", row.cells.len())
}

fn summarize_table_cell(cell: &TableCell) -> String {
    format!(
        "content_nodes={} span={}",
        cell.content.len(),
        cell.properties.grid_span.unwrap_or(1)
    )
}

fn summarize_comment(comment: &Comment) -> String {
    format!(
        "id={} author={} content_nodes={}",
        comment.comment_id,
        opt_str(&comment.author),
        comment.content.len()
    )
}

fn summarize_footnote(note: &Footnote) -> String {
    format!(
        "id={} content_nodes={}",
        note.footnote_id,
        note.content.len()
    )
}

fn summarize_endnote(note: &Endnote) -> String {
    format!(
        "id={} content_nodes={}",
        note.endnote_id,
        note.content.len()
    )
}

fn summarize_header_footer(content_len: usize) -> String {
    format!("content_nodes={}", content_len)
}

fn summarize_content_control(control: &ContentControl) -> String {
    format!(
        "content_nodes={} tag={}",
        control.content.len(),
        opt_str(&control.tag)
    )
}

fn summarize_bookmark_start(start: &BookmarkStart) -> String {
    format!("id={} name={}", start.bookmark_id, opt_str(&start.name))
}

fn summarize_bookmark_end(end: &BookmarkEnd) -> String {
    format!("id={}", end.bookmark_id)
}

fn summarize_field(field: &Field) -> String {
    format!(
        "runs={} instr={}",
        field.runs.len(),
        opt_str(&field.instruction)
    )
}

fn summarize_revision(rev: &Revision) -> String {
    format!(
        "type={:?} content_nodes={}",
        rev.change_type,
        rev.content.len()
    )
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

pub(crate) fn content_signature(node: &IRNode, store: &IrStore) -> Option<String> {
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

pub(crate) fn style_signature(node: &IRNode, _store: &IrStore) -> Option<String> {
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

pub(crate) fn short_hash(input: &str) -> String {
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
    use docir_core::ir::{
        BookmarkEnd, BookmarkStart, Cell, CellFormula, CellValue, Comment, ContentControl,
        CustomXmlPart, DiagnosticEntry, DiagnosticSeverity, Diagnostics, DigitalSignature,
        DocumentMetadata, DrawingPart, Endnote, ExtensionPart, ExtensionPartKind, Field, Footer,
        FormulaType, GlossaryDocument, GlossaryEntry, Header, Hyperlink, NumberingSet, Paragraph,
        PeoplePart, RelationshipGraph, Revision, RevisionType, Run, Shape, ShapeText,
        ShapeTextParagraph, ShapeTextRun, ShapeType, SmartArtPart, StyleSet, Table, TableCell,
        TableRow, Theme, ThemeColor, VmlDrawing, VmlShape, WebExtension, WebExtensionTaskpane,
        Worksheet,
    };
    use docir_core::ir::{CommentExtensionSet, CommentIdMap, CommentRangeEnd, CommentRangeStart};
    use docir_core::ir::{CommentReference, MediaAsset, MediaType};
    use docir_core::security::ExternalRefType;
    use docir_core::types::{DocumentFormat, NodeId};
    use docir_core::visitor::IrStore;

    #[test]
    fn helper_functions_are_deterministic() {
        assert_eq!(opt_bool(Some(true)), "true");
        assert_eq!(opt_bool(Some(false)), "false");
        assert_eq!(opt_bool(None), "-");
        assert_eq!(opt_u32(Some(42)), "42");
        assert_eq!(opt_u32(None), "-");
        assert_eq!(abbreviate("abc", 3), "abc");
        assert_eq!(abbreviate("abcdef", 3), "abc...");
        assert_eq!(format_float(2.0), "2");
        assert_eq!(format_float(2.5), "2.5000");
        assert_eq!(short_hash("same"), short_hash("same"));
    }

    #[test]
    fn summarizes_primary_word_and_security_nodes() {
        let store = IrStore::new();

        let doc = IRNode::Document(docir_core::ir::Document::new(DocumentFormat::WordProcessing));
        assert!(summarize_primary(&doc, &store)
            .unwrap()
            .contains("format=WordProcessing"));

        let mut section = docir_core::ir::Section::new();
        section.name = Some("Body".to_string());
        assert!(summarize_primary(&IRNode::Section(section), &store)
            .unwrap()
            .contains("name=Body"));

        let para_id = NodeId::new();
        let mut with_store = IrStore::new();
        let run = Run::new("Hello");
        with_store.insert(IRNode::Run(run.clone()));
        let mut para = Paragraph::new();
        para.id = para_id;
        para.runs.push(run.id);
        assert!(summarize_primary(&IRNode::Paragraph(para), &with_store)
            .unwrap()
            .contains("text=\"Hello\""));

        let run = Run::new("Run text");
        assert!(summarize_primary(&IRNode::Run(run.clone()), &store)
            .unwrap()
            .contains("bold=-"));

        let mut link = Hyperlink::new("https://example.test", true);
        link.runs.push(run.id);
        assert!(summarize_primary(&IRNode::Hyperlink(link), &with_store)
            .unwrap()
            .contains("external=true"));

        let mut macro_project = docir_core::security::MacroProject::new();
        macro_project.name = Some("VBA".to_string());
        macro_project.has_auto_exec = true;
        assert!(summarize_primary(&IRNode::MacroProject(macro_project), &store)
            .unwrap()
            .contains("auto_exec=true"));

        let mut ext =
            docir_core::security::ExternalReference::new(ExternalRefType::Hyperlink, "http://example.test");
        ext.ref_type = ExternalRefType::Hyperlink;
        assert!(summarize_primary(&IRNode::ExternalReference(ext), &store)
            .unwrap()
            .contains("target=http://example.test"));
    }

    #[test]
    fn secondary_summary_and_signatures_cover_key_branches() {
        let mut store = IrStore::new();
        let run_a = Run::new("A");
        let run_b = Run::new("B");
        store.insert(IRNode::Run(run_a.clone()));
        store.insert(IRNode::Run(run_b.clone()));

        let mut para = Paragraph::new();
        para.runs = vec![run_a.id, run_b.id];
        let para_sig = content_signature(&IRNode::Paragraph(para), &store).unwrap();
        assert_eq!(para_sig, "AB");

        let mut cell = Cell::new("C3", 2, 2);
        cell.value = CellValue::Number(9.0);
        cell.formula = Some(CellFormula {
            text: "SUM(A1:A2)".to_string(),
            formula_type: FormulaType::Normal,
            shared_index: None,
            shared_ref: None,
            is_array: false,
            array_ref: None,
        });
        let cell_sig = content_signature(&IRNode::Cell(cell.clone()), &store).unwrap();
        assert!(cell_sig.contains("C3=n:9;SUM(A1:A2)"));

        let mut worksheet = Worksheet::new("Data", 1);
        let cell_id = NodeId::new();
        store.insert(IRNode::Cell(cell));
        worksheet.cells.push(cell_id);
        let ws_sig = content_signature(&IRNode::Worksheet(worksheet), &store).unwrap();
        assert_eq!(ws_sig.len(), 16);

        let shape = Shape {
            text: Some(ShapeText {
                paragraphs: vec![ShapeTextParagraph {
                    runs: vec![ShapeTextRun {
                        text: "Title".to_string(),
                        bold: None,
                        italic: None,
                        font_size: None,
                        font_family: None,
                    }],
                    alignment: None,
                }],
            }),
            ..Shape::new(ShapeType::TextBox)
        };
        assert_eq!(
            content_signature(&IRNode::Shape(shape.clone()), &store).unwrap(),
            "Title"
        );
        assert!(style_signature(&IRNode::Shape(shape), &store)
            .unwrap()
            .contains("has_text=true"));

        let secondary = summarize_secondary(&IRNode::CommentExtensionSet(
            docir_core::ir::CommentExtensionSet::new(),
        ));
        assert_eq!(secondary, "entries=0");
    }

    #[test]
    fn summarize_covers_primary_word_and_package_nodes() {
        let store = IrStore::new();

        let mut table = Table::new();
        table.properties.style_id = Some("Grid".to_string());
        assert!(summarize(&IRNode::Table(table), &store).contains("style=Grid"));

        let mut row = TableRow::new();
        row.cells.push(NodeId::new());
        assert_eq!(summarize(&IRNode::TableRow(row), &store), "cells=1");

        let mut cell = TableCell::new();
        cell.properties.grid_span = Some(3);
        assert_eq!(summarize(&IRNode::TableCell(cell), &store), "content_nodes=0 span=3");

        let mut comment = Comment::new("c-1");
        comment.author = Some("alice".to_string());
        assert!(summarize(&IRNode::Comment(comment), &store).contains("author=alice"));

        let mut footnote = docir_core::ir::Footnote::new("f-1");
        footnote.content.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::Footnote(footnote), &store),
            "id=f-1 content_nodes=1"
        );

        let mut endnote = Endnote::new("e-1");
        endnote.content.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::Endnote(endnote), &store),
            "id=e-1 content_nodes=1"
        );

        let mut header = Header::new();
        header.content.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::Header(header), &store),
            "content_nodes=1"
        );

        let mut footer = Footer::new();
        footer.content.push(NodeId::new());
        assert_eq!(
            summarize(&IRNode::Footer(footer), &store),
            "content_nodes=1"
        );

        let mut control = ContentControl::new();
        control.tag = Some("tag-1".to_string());
        assert!(summarize(&IRNode::ContentControl(control), &store).contains("tag=tag-1"));

        let mut start = BookmarkStart::new("b-1");
        start.name = Some("chapter".to_string());
        assert_eq!(
            summarize(&IRNode::BookmarkStart(start), &store),
            "id=b-1 name=chapter"
        );
        assert_eq!(
            summarize(&IRNode::BookmarkEnd(BookmarkEnd::new("b-1")), &store),
            "id=b-1"
        );

        let mut field = Field::new(Some("HYPERLINK".to_string()));
        field.runs.push(NodeId::new());
        assert!(summarize(&IRNode::Field(field), &store).contains("runs=1"));

        let mut rev = Revision::new(RevisionType::Insert);
        rev.content.push(NodeId::new());
        assert!(summarize(&IRNode::Revision(rev), &store).contains("content_nodes=1"));

        let mut theme = Theme::new();
        theme.name = Some("Office".to_string());
        theme.colors.push(ThemeColor {
            name: "accent1".to_string(),
            value: Some("FF0000".to_string()),
        });
        theme.fonts.major = Some("Calibri".to_string());
        assert!(summarize(&IRNode::Theme(theme), &store).contains("colors=1"));

        let media = MediaAsset::new("word/media/image1.png", MediaType::Image, 42);
        assert!(summarize(&IRNode::MediaAsset(media), &store).contains("size=42"));

        let mut custom = CustomXmlPart::new("customXml/item1.xml", 7);
        custom.root_element = Some("root".to_string());
        assert!(summarize(&IRNode::CustomXmlPart(custom), &store).contains("root=root"));

        let rel = RelationshipGraph::new("word/document.xml");
        assert!(summarize(&IRNode::RelationshipGraph(rel), &store).contains("rels=0"));

        let mut sig = DigitalSignature::new();
        sig.signature_id = Some("sig1".to_string());
        sig.signature_method = Some("rsa-sha256".to_string());
        assert!(summarize(&IRNode::DigitalSignature(sig), &store).contains("method=rsa-sha256"));

        let ext = ExtensionPart::new("word/unknown.bin", 9, ExtensionPartKind::Unknown);
        assert!(summarize(&IRNode::ExtensionPart(ext), &store).contains("size=9"));

        assert_eq!(summarize(&IRNode::StyleSet(StyleSet::new()), &store), "styles=0");
        assert_eq!(
            summarize(&IRNode::NumberingSet(NumberingSet::new()), &store),
            "abstracts=0 nums=0"
        );

        let mut meta = DocumentMetadata::new();
        meta.title = Some("Doc".to_string());
        meta.creator = Some("Bob".to_string());
        assert_eq!(
            summarize(&IRNode::Metadata(meta), &store),
            "title=Doc author=Bob"
        );
    }

    #[test]
    fn summarize_covers_secondary_nodes_and_slide_style_signature() {
        let store = IrStore::new();

        assert_eq!(
            summarize(&IRNode::CommentExtensionSet(CommentExtensionSet::new()), &store),
            "entries=0"
        );
        assert_eq!(
            summarize(&IRNode::CommentIdMap(CommentIdMap::new()), &store),
            "mappings=0"
        );
        assert_eq!(
            summarize(&IRNode::CommentRangeStart(CommentRangeStart::new("2")), &store),
            "comment_id=2"
        );
        assert_eq!(
            summarize(&IRNode::CommentRangeEnd(CommentRangeEnd::new("2")), &store),
            "comment_id=2"
        );
        assert_eq!(
            summarize(&IRNode::CommentReference(CommentReference::new("2")), &store),
            "comment_id=2"
        );
        assert_eq!(
            summarize(&IRNode::PeoplePart(PeoplePart::new()), &store),
            "people=0"
        );

        let smart = SmartArtPart {
            id: NodeId::new(),
            kind: "diagramData".to_string(),
            path: "ppt/diagrams/data1.xml".to_string(),
            root_element: None,
            point_count: None,
            connection_count: None,
            rel_ids: Vec::new(),
            span: None,
        };
        assert_eq!(
            summarize(&IRNode::SmartArtPart(smart), &store),
            "kind=diagramData path=ppt/diagrams/data1.xml"
        );

        let mut web = WebExtension::new();
        web.extension_id = Some("ext-id".to_string());
        web.store = Some("store".to_string());
        web.version = Some("1.0".to_string());
        assert!(summarize(&IRNode::WebExtension(web), &store).contains("properties=0"));

        let mut pane = WebExtensionTaskpane::new();
        pane.web_extension_ref = Some("ext-id".to_string());
        pane.dock_state = Some("right".to_string());
        pane.visibility = Some(true);
        assert_eq!(
            summarize(&IRNode::WebExtensionTaskpane(pane), &store),
            "ref=ext-id dock_state=right visible=true"
        );

        assert_eq!(
            summarize(&IRNode::GlossaryDocument(GlossaryDocument::new()), &store),
            "entries=0"
        );
        let mut glossary = GlossaryEntry::new();
        glossary.name = Some("quick".to_string());
        glossary.gallery = Some("auto".to_string());
        assert_eq!(
            summarize(&IRNode::GlossaryEntry(glossary), &store),
            "name=quick gallery=auto content_nodes=0"
        );

        let drawing = VmlDrawing::new("word/vmlDrawing1.vml");
        assert_eq!(
            summarize(&IRNode::VmlDrawing(drawing), &store),
            "path=word/vmlDrawing1.vml shapes=0"
        );
        let mut vml_shape = VmlShape::new();
        vml_shape.name = Some("shape1".to_string());
        vml_shape.rel_id = Some("rId1".to_string());
        vml_shape.image_target = Some("media/image1.png".to_string());
        assert_eq!(
            summarize(&IRNode::VmlShape(vml_shape), &store),
            "name=shape1 rel_id=rId1 image_target=media/image1.png"
        );

        let part = DrawingPart::new("word/drawings/drawing1.xml");
        assert_eq!(
            summarize(&IRNode::DrawingPart(part), &store),
            "path=word/drawings/drawing1.xml shapes=0"
        );

        let mut diag = Diagnostics::new();
        diag.entries.push(DiagnosticEntry {
            severity: DiagnosticSeverity::Warning,
            code: "W001".to_string(),
            message: "warn".to_string(),
            path: None,
        });
        assert_eq!(summarize(&IRNode::Diagnostics(diag), &store), "entries=1");

        let mut slide = docir_core::ir::Slide::new(1);
        slide.layout_id = Some("layout-1".to_string());
        slide.master_id = Some("master-1".to_string());
        assert_eq!(
            style_signature(&IRNode::Slide(slide), &store).unwrap(),
            "layout_id=layout-1 master_id=master-1"
        );
    }

    #[test]
    fn summary_and_signature_cover_macro_and_hyperlink_fallback_paths() {
        let mut store = IrStore::new();

        let run_a = Run::new("PartA");
        let run_b = Run::new("PartB");
        store.insert(IRNode::Run(run_a.clone()));
        store.insert(IRNode::Run(run_b.clone()));

        let mut link = Hyperlink::new("https://example.test", true);
        link.runs = vec![run_a.id, run_b.id];
        let link_id = link.id;
        store.insert(IRNode::Hyperlink(link.clone()));

        let mut para = Paragraph::new();
        para.runs.push(link_id);
        let para_text = summarize_paragraph(&para, &store);
        assert!(para_text.contains("PartAPartB"));

        assert_eq!(
            content_signature(&IRNode::Hyperlink(link), &store).as_deref(),
            Some("https://example.test")
        );

        let mut module = docir_core::security::MacroModule::new(
            "AutoOpen",
            docir_core::security::MacroModuleType::Standard,
        );
        module
            .suspicious_calls
            .push(docir_core::security::SuspiciousCall {
                name: "Shell".to_string(),
                category: docir_core::security::SuspiciousCallCategory::ShellExecution,
                line: Some(1),
            });
        let module_summary = summarize(&IRNode::MacroModule(module.clone()), &store);
        assert!(module_summary.contains("suspicious_calls=1"));
        assert_eq!(
            content_signature(&IRNode::MacroModule(module), &store).as_deref(),
            Some("AutoOpen")
        );

        let mut project = docir_core::security::MacroProject::new();
        project.name = Some("VBAProject".to_string());
        assert_eq!(
            content_signature(&IRNode::MacroProject(project), &store).as_deref(),
            Some("VBAProject")
        );

        let mut activex = docir_core::ActiveXControl::new();
        activex.name = Some("Btn".to_string());
        assert_eq!(
            content_signature(&IRNode::ActiveXControl(activex), &store).as_deref(),
            Some("Btn")
        );

        let mut ole = docir_core::security::OleObject::new();
        ole.name = Some("Object1".to_string());
        assert_eq!(
            content_signature(&IRNode::OleObject(ole), &store).as_deref(),
            Some("Object1")
        );

        let defined = docir_core::ir::DefinedName {
            id: NodeId::new(),
            name: "MyName".to_string(),
            value: "Sheet1!$A$1".to_string(),
            local_sheet_id: None,
            hidden: false,
            comment: None,
            span: None,
        };
        assert_eq!(
            content_signature(&IRNode::DefinedName(defined), &store).as_deref(),
            Some("MyName")
        );

        let table = docir_core::ir::TableDefinition {
            id: NodeId::new(),
            name: Some("TableFallback".to_string()),
            display_name: None,
            ref_range: None,
            header_row_count: None,
            totals_row_count: None,
            columns: vec![],
            span: None,
        };
        assert_eq!(
            content_signature(&IRNode::TableDefinition(table), &store).as_deref(),
            Some("TableFallback")
        );
    }

    #[test]
    fn style_signature_and_shape_text_cover_remaining_branches() {
        let store = IrStore::new();

        let para = Paragraph::new();
        let para_sig = style_signature(&IRNode::Paragraph(para), &store).unwrap();
        assert!(para_sig.starts_with('{'));

        let run_sig = style_signature(&IRNode::Run(Run::new("r")), &store).unwrap();
        assert!(run_sig.starts_with('{'));

        let table_sig = style_signature(&IRNode::Table(Table::new()), &store).unwrap();
        assert!(table_sig.starts_with('{'));

        let shape = Shape {
            text: Some(ShapeText {
                paragraphs: vec![
                    ShapeTextParagraph {
                        runs: vec![ShapeTextRun {
                            text: "L1".to_string(),
                            bold: None,
                            italic: None,
                            font_size: None,
                            font_family: None,
                        }],
                        alignment: None,
                    },
                    ShapeTextParagraph {
                        runs: vec![ShapeTextRun {
                            text: "L2".to_string(),
                            bold: None,
                            italic: None,
                            font_size: None,
                            font_family: None,
                        }],
                        alignment: None,
                    },
                ],
            }),
            ..Shape::new(ShapeType::TextBox)
        };
        let summary = summarize_shape(&shape);
        assert!(summary.contains("L1\nL2"));
    }
}
