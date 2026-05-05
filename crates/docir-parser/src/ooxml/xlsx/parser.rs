//! XLSX workbook and worksheet parsing.

use crate::diagnostics::attach_diagnostics_if_any;
use crate::error::ParseError;
use crate::ooxml::relationships::{Relationships, TargetMode};
use crate::security_utils::parse_dde_formula;
use crate::xml_utils::{attr_value, attr_value_by_suffix};
use crate::zip_handler::PackageReader;
use docir_core::ir::{
    parse_cell_reference, CalcChain, Cell, CellError, CellFormula, ColumnDefinition,
    ConditionalFormat, Diagnostics, Document, IRNode, MergedCellRange, SheetComment, SheetKind,
    SheetState,
};
use docir_core::security::SecurityInfo;
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::BytesStart;
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};

use super::relationships::classify_relationship;
use super::workbook::{parse_workbook_info, WorkbookInfo};
#[path = "parser_xml.rs"]
mod parser_xml;

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
        document.shared_parts.append(&mut self.chart_nodes);
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
        let reference = attr_value(start, b"r").ok_or_else(|| {
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
        let Some(rel_id) = attr_value_by_suffix(element, &[b":id"]) else {
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

pub(crate) fn parse_calc_chain(xml: &str, path: &str) -> Result<CalcChain, ParseError> {
    parser_xml::parse_calc_chain(xml, path)
}

pub(crate) fn parse_sheet_comments(
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

pub(crate) fn parse_threaded_comments(
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

pub(crate) fn map_cell_error(value: &str) -> CellError {
    parser_xml::map_cell_error(value)
}

pub(crate) fn parse_conditional_formatting(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<ConditionalFormat, ParseError> {
    parser_xml::parse_conditional_formatting(reader, start, sheet_path)
}

pub(crate) fn parse_formula(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<CellFormula, ParseError> {
    parser_xml::parse_formula(reader, start, sheet_path)
}

pub(crate) fn extract_formula_function(formula_upper: &str) -> Option<String> {
    parser_xml::extract_formula_function(formula_upper)
}

pub(crate) fn parse_formula_args_text(formula: &str) -> Option<String> {
    parser_xml::parse_formula_args_text(formula)
}

pub(crate) fn parse_formula_empty(start: &BytesStart) -> CellFormula {
    parser_xml::parse_formula_empty(start)
}

pub(crate) fn parse_inline_string(
    reader: &mut Reader<&[u8]>,
    sheet_path: &str,
) -> Result<String, ParseError> {
    parser_xml::parse_inline_string(reader, sheet_path)
}

pub(crate) fn parse_column(element: &BytesStart, columns: &mut HashMap<u32, ColumnDefinition>) {
    parser_xml::parse_column(element, columns);
}

pub(crate) fn parse_merge_cell(element: &BytesStart) -> Option<MergedCellRange> {
    parser_xml::parse_merge_cell(element)
}

#[cfg(test)]
mod tests;
