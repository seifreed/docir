use docir_core::ir::{Cell, CellFormula, CellValue, IRNode, Paragraph, Worksheet};
use docir_core::visitor::IrStore;

use super::format_helpers::{opt_str, paragraph_text, shape_text, short_hash};

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
        IRNode::Paragraph(para) => serde_json::to_string(&para.properties)
            .ok()
            .filter(|s| !s.is_empty()),
        IRNode::Run(run) => serde_json::to_string(&run.properties)
            .ok()
            .filter(|s| !s.is_empty()),
        IRNode::Table(table) => serde_json::to_string(&table.properties)
            .ok()
            .filter(|s| !s.is_empty()),
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

pub(crate) fn text_from_paragraph(para: &Paragraph, store: &IrStore) -> String {
    paragraph_text(para, store)
}

fn cell_content_signature(cell: &Cell) -> String {
    let mut out = String::new();
    out.push_str(&cell.reference);
    out.push('=');
    out.push_str(&cell_value_summary(&cell.value));
    if let Some(formula) = &cell.formula {
        out.push(';');
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

pub(crate) fn cell_value_summary(value: &CellValue) -> String {
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
