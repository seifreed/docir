use docir_core::ir::{
    Cell, CellFormula, CellValue, Document, Hyperlink, IRNode, Paragraph, Run, Section, Shape,
    Slide, Worksheet,
};
use docir_core::security::{ExternalReference, MacroModule, MacroProject, OleObject};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use sha2::Digest;

mod presentation;
mod spreadsheet;

pub(crate) fn summarize(node: &IRNode, store: &IrStore) -> String {
    match node {
        IRNode::Worksheet(_)
        | IRNode::Cell(_)
        | IRNode::SharedStringTable(_)
        | IRNode::ConnectionPart(_)
        | IRNode::SpreadsheetStyles(_)
        | IRNode::DefinedName(_)
        | IRNode::ConditionalFormat(_)
        | IRNode::DataValidation(_)
        | IRNode::TableDefinition(_)
        | IRNode::PivotTable(_)
        | IRNode::PivotCache(_)
        | IRNode::PivotCacheRecords(_)
        | IRNode::WorkbookProperties(_)
        | IRNode::CalcChain(_)
        | IRNode::SheetComment(_)
        | IRNode::SheetMetadata(_)
        | IRNode::WorksheetDrawing(_)
        | IRNode::ChartData(_)
        | IRNode::ExternalLinkPart(_)
        | IRNode::SlicerPart(_)
        | IRNode::TimelinePart(_)
        | IRNode::QueryTablePart(_) => {
            spreadsheet::summarize(node, store).expect("spreadsheet summary missing for node")
        }
        IRNode::Slide(_)
        | IRNode::Shape(_)
        | IRNode::SlideMaster(_)
        | IRNode::SlideLayout(_)
        | IRNode::NotesMaster(_)
        | IRNode::HandoutMaster(_)
        | IRNode::NotesSlide(_)
        | IRNode::PresentationProperties(_)
        | IRNode::ViewProperties(_)
        | IRNode::TableStyleSet(_)
        | IRNode::PptxCommentAuthor(_)
        | IRNode::PptxComment(_)
        | IRNode::PresentationTag(_)
        | IRNode::PresentationInfo(_) => {
            presentation::summarize(node, store).expect("presentation summary missing for node")
        }
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
        IRNode::Diagnostics(diag) => format!("entries={}", diag.entries.len()),
    }
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
