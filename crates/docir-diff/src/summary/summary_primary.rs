use docir_core::ir::{
    BookmarkEnd, BookmarkStart, Cell, CellFormula, CellValue, Comment, ContentControl, Document,
    Endnote, Field, Footnote, Hyperlink, IRNode, Paragraph, Revision, Run, Section, Shape, Slide,
    Table, TableCell, TableRow, Worksheet,
};
use docir_core::security::{ExternalReference, MacroModule, MacroProject, OleObject};
use docir_core::visitor::IrStore;

use super::format_helpers::{
    abbreviate, format_float, opt_bool, opt_str, paragraph_text, runs_text, shape_text,
};

pub(crate) fn summarize_primary(node: &IRNode, store: &IrStore) -> Option<String> {
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

fn summarize_document(doc: &Document) -> String {
    format!(
        "format={:?} content_nodes={} macros={} ole={} external_refs={} threat={:?}",
        doc.format,
        doc.content.len(),
        doc.security.has_macro_project(),
        doc.security.ole_object_count(),
        doc.security.external_ref_count(),
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

pub(crate) fn summarize_paragraph(para: &Paragraph, store: &IrStore) -> String {
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

pub(crate) fn summarize_worksheet(ws: &Worksheet) -> String {
    format!(
        "name={} sheet_id={} state={:?} cells={} merged={}",
        ws.name,
        ws.sheet_id,
        ws.state,
        ws.cells.len(),
        ws.merged_cells.len(),
    )
}

pub(crate) fn summarize_cell(cell: &Cell) -> String {
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

pub(crate) fn summarize_formula(formula: &CellFormula) -> String {
    format!(
        "{} type={:?}",
        abbreviate(&formula.text, 80),
        formula.formula_type,
    )
}

pub(crate) fn summarize_slide(slide: &Slide) -> String {
    format!(
        "number={} name={} shapes={} hidden={}",
        slide.number,
        opt_str(&slide.name),
        slide.shapes.len(),
        slide.hidden,
    )
}

pub(crate) fn summarize_shape(shape: &Shape) -> String {
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
