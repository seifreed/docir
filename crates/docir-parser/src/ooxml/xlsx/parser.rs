//! XLSX workbook and worksheet parsing.

use crate::diagnostics::{attach_diagnostics_if_any, push_warning};
use crate::error::ParseError;
use crate::ooxml::part_utils::parse_xml_part_with_span;
use crate::ooxml::relationships::{rel_type, Relationship, Relationships, TargetMode};
use crate::security_utils::parse_dde_formula;
use crate::zip_handler::PackageReader;
use docir_core::ir::{
    parse_cell_reference, CalcChain, CalcChainEntry, Cell, CellError, CellFormula,
    ColumnDefinition, ConditionalFormat, ConditionalRule, Diagnostics, Document, FormulaType,
    IRNode, MergedCellRange, PivotCache, Shape, ShapeType, SheetComment, SheetKind, SheetState,
    Worksheet, WorksheetDrawing,
};
use docir_core::security::{ExternalRefType, ExternalReference, SecurityInfo};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};

use super::connections::{
    connection_targets, parse_connections_part, parse_external_link_part, parse_query_table_part,
    parse_slicer_part, parse_timeline_part,
};
use super::metadata::parse_sheet_metadata;
use super::relationships::classify_relationship;
use super::styles::{parse_color_attr, parse_styles};
use super::tables::{
    parse_pivot_cache_records, parse_pivot_table_definition, parse_table_definition,
};
use super::workbook::{
    auto_open_target_from_defined_name, parse_workbook_info, SheetInfo, WorkbookInfo,
};

/// XLSX parser for workbook.xml and worksheets.
pub struct XlsxParser {
    pub(super) store: IrStore,
    pub(super) security_info: SecurityInfo,
    pub(super) shared_strings: Vec<String>,
    pub(super) external_rel_ids: HashSet<String>,
    pub(super) chart_nodes: Vec<NodeId>,
    pub(super) root_id: NodeId,
    pub(super) current_sheet_kind: Option<SheetKind>,
    pub(super) current_sheet_name: Option<String>,
    pub(super) current_sheet_state: Option<SheetState>,
    pub(super) current_xlm_index: Option<usize>,
    pub(super) diagnostics: Diagnostics,
}

impl XlsxParser {
    /// Creates a new XLSX parser.
    pub fn new() -> Self {
        Self {
            store: IrStore::new(),
            security_info: SecurityInfo::default(),
            shared_strings: Vec::new(),
            external_rel_ids: HashSet::new(),
            chart_nodes: Vec::new(),
            root_id: NodeId::new(),
            current_sheet_kind: None,
            current_sheet_name: None,
            current_sheet_state: None,
            current_xlm_index: None,
            diagnostics: Diagnostics::new(),
        }
    }

    /// Parses the workbook and all worksheets.
    pub fn parse_workbook(
        &mut self,
        zip: &mut impl PackageReader,
        workbook_xml: &str,
        workbook_rels: &Relationships,
        workbook_path: &str,
    ) -> Result<NodeId, ParseError> {
        let (mut document, workbook_info) =
            self.pipeline_discover(workbook_xml, workbook_path, workbook_rels)?;
        let (auto_open_targets, pivot_cache_refs) = self.pipeline_parse_core(
            &mut document,
            zip,
            workbook_path,
            workbook_rels,
            workbook_info,
        )?;
        self.pipeline_enrich(
            &mut document,
            zip,
            workbook_path,
            workbook_rels,
            &auto_open_targets,
            pivot_cache_refs,
        )?;
        Ok(self.pipeline_finalize(document, workbook_path))
    }

    /// Returns the IR store.
    pub fn into_store(self) -> IrStore {
        self.store
    }

    fn pipeline_discover(
        &mut self,
        workbook_xml: &str,
        workbook_path: &str,
        workbook_rels: &Relationships,
    ) -> Result<(Document, WorkbookInfo), ParseError> {
        let mut document = Document::new(DocumentFormat::Spreadsheet);
        self.root_id = document.id;
        document.span = Some(SourceSpan::new(workbook_path));
        self.process_external_relationships(workbook_rels, workbook_path);
        let workbook_info = parse_workbook_info(workbook_xml)?;
        Ok((document, workbook_info))
    }

    fn pipeline_parse_core(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
        workbook_path: &str,
        workbook_rels: &Relationships,
        workbook_info: WorkbookInfo,
    ) -> Result<(Vec<Option<String>>, Vec<super::workbook::PivotCacheRef>), ParseError> {
        let WorkbookInfo {
            sheets,
            defined_names,
            workbook_properties,
            pivot_cache_refs,
        } = workbook_info;

        self.load_workbook_properties(document, workbook_properties, workbook_path);
        self.load_shared_strings(document, zip)?;
        self.load_styles(document, zip)?;
        self.load_calc_chain(document, zip)?;
        self.load_people_part(document, zip)?;
        let auto_open_targets = self.load_defined_names(document, defined_names);
        self.load_sheets(document, zip, workbook_path, workbook_rels, sheets)?;

        Ok((auto_open_targets, pivot_cache_refs))
    }

    fn pipeline_enrich(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
        workbook_path: &str,
        workbook_rels: &Relationships,
        auto_open_targets: &[Option<String>],
        pivot_cache_refs: Vec<super::workbook::PivotCacheRef>,
    ) -> Result<(), ParseError> {
        self.finalize_auto_open_targets(auto_open_targets);
        self.load_pivot_caches(
            document,
            zip,
            workbook_path,
            workbook_rels,
            pivot_cache_refs,
        )?;
        self.parse_external_links_and_connections(zip, workbook_path, workbook_rels)?;
        self.load_sheet_metadata(document, zip)?;
        self.load_slicer_parts(document, zip)?;
        self.load_timeline_parts(document, zip)?;
        self.load_query_tables(document, zip)?;
        Ok(())
    }

    fn pipeline_finalize(&mut self, mut document: Document, workbook_path: &str) -> NodeId {
        document.shared_parts.extend(self.chart_nodes.drain(..));
        document.security = std::mem::take(&mut self.security_info);

        let mut diagnostics = std::mem::replace(&mut self.diagnostics, Diagnostics::new());
        if !diagnostics.entries.is_empty() {
            diagnostics.span = Some(SourceSpan::new(workbook_path));
            attach_diagnostics_if_any(&mut self.store, &mut document, diagnostics);
        }

        let doc_id = document.id;
        self.store.insert(IRNode::Document(document));
        doc_id
    }

    pub(super) fn parse_chart(&mut self, xml: &str, chart_path: &str) -> Option<NodeId> {
        crate::ooxml::shared::parse_chart_data(xml, chart_path, &mut self.store)
    }

    pub(super) fn parse_empty_cell(
        &self,
        start: &BytesStart,
        sheet_path: &str,
    ) -> Result<Cell, ParseError> {
        let mut cell_ref: Option<String> = None;
        for attr in start.attributes().flatten() {
            if attr.key.as_ref() == b"r" {
                cell_ref = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
        }

        let reference = cell_ref.ok_or_else(|| {
            ParseError::InvalidStructure("Cell missing reference attribute".to_string())
        })?;
        let (col, row) = parse_cell_reference(&reference).ok_or_else(|| {
            ParseError::InvalidStructure(format!("Invalid cell reference: {reference}"))
        })?;

        let mut cell = Cell::new(reference, col, row);
        cell.span = Some(SourceSpan::new(sheet_path));
        Ok(cell)
    }

    pub(super) fn handle_hyperlink(
        &mut self,
        element: &BytesStart,
        relationships: &Relationships,
        sheet_path: &str,
    ) {
        let mut rel_id: Option<String> = None;
        for attr in element.attributes().flatten() {
            if attr.key.as_ref() == b"r:id" {
                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
        }

        let Some(rel_id) = rel_id else {
            return;
        };
        let Some(rel) = relationships.get(&rel_id) else {
            return;
        };
        if rel.target_mode != TargetMode::External {
            return;
        }

        let ref_type = classify_relationship(&rel.rel_type);
        self.add_external_reference(rel, ref_type, sheet_path);
    }

    pub(super) fn handle_formula_security(
        &mut self,
        cell_ref: &str,
        formula: &CellFormula,
        sheet_path: &str,
    ) {
        let text = formula.text.trim();
        let upper = text.to_ascii_uppercase();

        // DDE detection in Excel formulas
        if upper.starts_with("DDEAUTO") || upper.starts_with("DDE") {
            if let Some(dde) = parse_dde_formula(
                text,
                SourceSpan::new(sheet_path).with_xml_path(cell_ref),
                true,
            ) {
                self.security_info.dde_fields.push(dde);
            }
        }

        self.record_xlm_formula(cell_ref, text, &upper, sheet_path);
    }
}

impl Default for XlsxParser {
    fn default() -> Self {
        Self::new()
    }
}

pub(super) fn parse_calc_chain(xml: &str, path: &str) -> Result<CalcChain, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut chain = CalcChain::new();
    chain.span = Some(SourceSpan::new(path));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"c" {
                    let mut cell_ref = None;
                    let mut sheet_id = None;
                    let mut index = None;
                    let mut level = None;
                    let mut new_value = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"r" => {
                                cell_ref = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"i" => {
                                index = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"l" => {
                                level = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"s" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                new_value = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"si" => {
                                sheet_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            _ => {}
                        }
                    }
                    if let Some(cell_ref) = cell_ref {
                        chain.entries.push(CalcChainEntry {
                            cell_ref,
                            sheet_id,
                            index,
                            level,
                            new_value,
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(chain)
}

pub(super) fn parse_sheet_comments(
    xml: &str,
    path: &str,
    sheet_name: Option<&str>,
) -> Result<Vec<SheetComment>, ParseError> {
    super::comments::parse_sheet_comments_impl(
        xml,
        path,
        sheet_name,
        super::comments::CommentFlavor::Legacy,
    )
}

pub(super) fn parse_threaded_comments(
    xml: &str,
    path: &str,
    sheet_name: Option<&str>,
) -> Result<Vec<SheetComment>, ParseError> {
    super::comments::parse_sheet_comments_impl(
        xml,
        path,
        sheet_name,
        super::comments::CommentFlavor::Threaded,
    )
}

pub(super) fn map_cell_error(value: &str) -> CellError {
    match value.trim() {
        "#NULL!" => CellError::Null,
        "#DIV/0!" => CellError::DivZero,
        "#VALUE!" => CellError::Value,
        "#REF!" => CellError::Ref,
        "#NAME?" => CellError::Name,
        "#NUM!" => CellError::Num,
        "#N/A" => CellError::NA,
        "#GETTING_DATA" => CellError::GettingData,
        _ => CellError::Value,
    }
}

pub(super) fn parse_conditional_formatting(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<ConditionalFormat, ParseError> {
    let mut ranges: Vec<String> = Vec::new();
    for attr in start.attributes().flatten() {
        if attr.key.as_ref() == b"sqref" {
            let val = String::from_utf8_lossy(&attr.value).to_string();
            ranges = val.split_whitespace().map(|s| s.to_string()).collect();
        }
    }

    let mut rules: Vec<ConditionalRule> = Vec::new();
    let mut current_rule: Option<ConditionalRule> = None;
    let mut in_formula = false;
    let mut formula_text = String::new();

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"cfRule" => {
                    let mut rule_type = "unknown".to_string();
                    let mut priority = None;
                    let mut operator = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"type" => rule_type = String::from_utf8_lossy(&attr.value).to_string(),
                            b"priority" => {
                                priority = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"operator" => {
                                operator = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            _ => {}
                        }
                    }
                    current_rule = Some(ConditionalRule {
                        rule_type,
                        priority,
                        operator,
                        formulae: Vec::new(),
                    });
                }
                b"formula" => {
                    in_formula = true;
                    formula_text.clear();
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_formula {
                    formula_text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"formula" => {
                    in_formula = false;
                    if let Some(rule) = current_rule.as_mut() {
                        if !formula_text.is_empty() {
                            rule.formulae.push(formula_text.clone());
                        }
                    }
                }
                b"cfRule" => {
                    if let Some(rule) = current_rule.take() {
                        rules.push(rule);
                    }
                }
                b"conditionalFormatting" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: sheet_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(ConditionalFormat {
        id: NodeId::new(),
        ranges,
        rules,
        span: Some(SourceSpan::new(sheet_path)),
    })
}

pub(super) fn parse_formula(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<CellFormula, ParseError> {
    let mut formula_type = FormulaType::Normal;
    let mut shared_index = None;
    let mut shared_ref = None;
    let mut array_ref = None;
    let mut is_array = false;

    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"t" => {
                let v = String::from_utf8_lossy(&attr.value);
                match v.as_ref() {
                    "shared" => formula_type = FormulaType::Shared,
                    "array" => {
                        formula_type = FormulaType::Array;
                        is_array = true;
                    }
                    "dataTable" => formula_type = FormulaType::DataTable,
                    _ => {}
                }
            }
            b"si" => shared_index = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"ref" => {
                let r = String::from_utf8_lossy(&attr.value).to_string();
                if formula_type == FormulaType::Shared {
                    shared_ref = Some(r);
                } else {
                    array_ref = Some(r);
                }
            }
            _ => {}
        }
    }

    let text = reader
        .read_text(start.name())
        .map_err(|e| ParseError::Xml {
            file: sheet_path.to_string(),
            message: e.to_string(),
        })?;

    Ok(CellFormula {
        text: text.to_string(),
        formula_type,
        shared_index,
        shared_ref,
        is_array,
        array_ref,
    })
}

pub(super) fn extract_formula_function(formula_upper: &str) -> Option<String> {
    let trimmed = formula_upper.trim();
    let trimmed = trimmed.strip_prefix('=').unwrap_or(trimmed);
    let idx = trimmed.find('(')?;
    Some(trimmed[..idx].trim().to_string())
}

pub(super) fn parse_formula_args_text(formula: &str) -> Option<String> {
    let start = formula.find('(')?;
    let end = formula.rfind(')')?;
    if end > start + 1 {
        Some(formula[start + 1..end].to_string())
    } else {
        None
    }
}

pub(super) fn parse_formula_empty(start: &BytesStart) -> CellFormula {
    let mut formula_type = FormulaType::Normal;
    let mut shared_index = None;
    let mut shared_ref = None;
    let mut array_ref = None;
    let mut is_array = false;

    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"t" => {
                let v = String::from_utf8_lossy(&attr.value);
                match v.as_ref() {
                    "shared" => formula_type = FormulaType::Shared,
                    "array" => {
                        formula_type = FormulaType::Array;
                        is_array = true;
                    }
                    "dataTable" => formula_type = FormulaType::DataTable,
                    _ => {}
                }
            }
            b"si" => shared_index = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"ref" => {
                let r = String::from_utf8_lossy(&attr.value).to_string();
                if formula_type == FormulaType::Shared {
                    shared_ref = Some(r);
                } else {
                    array_ref = Some(r);
                }
            }
            _ => {}
        }
    }

    CellFormula {
        text: String::new(),
        formula_type,
        shared_index,
        shared_ref,
        is_array,
        array_ref,
    }
}

pub(super) fn parse_inline_string(
    reader: &mut Reader<&[u8]>,
    sheet_path: &str,
) -> Result<String, ParseError> {
    let mut buf = Vec::new();
    let mut in_t = false;
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"t" {
                    in_t = true;
                }
            }
            Ok(Event::Text(e)) => {
                if in_t {
                    let t = e.unescape().map_err(|err| ParseError::Xml {
                        file: sheet_path.to_string(),
                        message: err.to_string(),
                    })?;
                    text.push_str(&t);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"t" {
                    in_t = false;
                } else if e.name().as_ref() == b"is" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: sheet_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(text)
}

pub(super) fn parse_column(element: &BytesStart, columns: &mut HashMap<u32, ColumnDefinition>) {
    let mut min = None;
    let mut max = None;
    let mut width = None;
    let mut hidden = false;
    let mut custom_width = false;

    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"min" => min = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"max" => max = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"width" => width = String::from_utf8_lossy(&attr.value).parse::<f64>().ok(),
            b"hidden" => hidden = attr.value.as_ref() == b"1",
            b"customWidth" => custom_width = attr.value.as_ref() == b"1",
            _ => {}
        }
    }

    let (Some(min), Some(max)) = (min, max) else {
        return;
    };
    for idx in min..=max {
        let col_index = idx.saturating_sub(1);
        columns.insert(
            col_index,
            ColumnDefinition {
                index: col_index,
                width,
                hidden,
                custom_width,
            },
        );
    }
}

pub(super) fn parse_merge_cell(element: &BytesStart) -> Option<MergedCellRange> {
    let mut ref_attr = None;
    for attr in element.attributes().flatten() {
        if attr.key.as_ref() == b"ref" {
            ref_attr = Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }

    let ref_attr = ref_attr?;
    let mut parts = ref_attr.split(':');
    let start = parts.next()?;
    let end = parts.next().unwrap_or(start);

    let (start_col, start_row) = parse_cell_reference(start)?;
    let (end_col, end_row) = parse_cell_reference(end)?;

    Some(MergedCellRange {
        start_col,
        start_row,
        end_col,
        end_row,
    })
}

#[cfg(test)]
mod tests;
