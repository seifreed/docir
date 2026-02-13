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
mod tests {
    use super::*;
    use docir_core::ir::IRNode;
    use docir_core::CellValue;

    #[test]
    fn test_parse_workbook_info_sheets() {
        let xml = r#"
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <bookViews>
            <workbookView activeTab="1" firstSheet="0" showHorizontalScroll="1"
                          showVerticalScroll="0" showSheetTabs="1" tabRatio="400"
                          windowWidth="12000" windowHeight="8000" xWindow="120" yWindow="240"/>
          </bookViews>
          <sheets>
            <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
            <sheet name="Hidden" sheetId="2" r:id="rId2" state="hidden"/>
            <sheet name="VeryHidden" sheetId="3" r:id="rId3" state="veryHidden"/>
          </sheets>
        </workbook>
        "#;

        let info = parse_workbook_info(xml).expect("parse workbook info");
        assert_eq!(info.sheets.len(), 3);
        assert_eq!(info.sheets[0].name, "Sheet1");
        assert_eq!(info.sheets[1].state, SheetState::Hidden);
        assert_eq!(info.sheets[2].state, SheetState::VeryHidden);
        let props = info.workbook_properties.expect("workbook props");
        assert_eq!(props.active_tab, Some(1));
        assert_eq!(props.first_sheet, Some(0));
        assert_eq!(props.show_horizontal_scroll, Some(true));
        assert_eq!(props.show_vertical_scroll, Some(false));
        assert_eq!(props.show_sheet_tabs, Some(true));
        assert_eq!(props.tab_ratio, Some(400));
        assert_eq!(props.window_width, Some(12000));
        assert_eq!(props.window_height, Some(8000));
        assert_eq!(props.x_window, Some(120));
        assert_eq!(props.y_window, Some(240));
    }

    #[test]
    fn test_parse_worksheet_cells() {
        let xml = r#"
        <worksheet>
          <cols>
            <col min="1" max="2" width="10" customWidth="1"/>
          </cols>
          <sheetData>
            <row r="1">
              <c r="A1" t="s"><v>0</v></c>
              <c r="B1"><v>42</v></c>
              <c r="C1" t="b"><v>1</v></c>
              <c r="D1" t="e"><v>#REF!</v></c>
              <c r="E1"><f>SUM(A1:B1)</f><v>3</v></c>
              <c r="F1"><is><t>Inline</t></is></c>
            </row>
          </sheetData>
          <mergeCells>
            <mergeCell ref="A1:B1"/>
          </mergeCells>
        </worksheet>
        "#;

        let mut parser = XlsxParser::new();
        parser.shared_strings = vec!["Hello".to_string()];

        let sheet = SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };

        let mut zip = build_empty_zip();
        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                xml,
                &sheet,
                "xl/worksheets/sheet1.xml",
                &Relationships::default(),
                SheetKind::Worksheet,
            )
            .expect("parse worksheet");
        let store = parser.into_store();

        let worksheet = match store.get(ws_id) {
            Some(IRNode::Worksheet(ws)) => ws,
            _ => panic!("missing worksheet node"),
        };

        assert_eq!(worksheet.columns.len(), 2);
        assert_eq!(worksheet.merged_cells.len(), 1);
        assert_eq!(worksheet.cells.len(), 6);

        let cell_a1 = store
            .get(worksheet.cells[0])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell a1");
        assert!(matches!(cell_a1.value, CellValue::String(ref v) if v == "Hello"));

        let cell_b1 = store
            .get(worksheet.cells[1])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell b1");
        assert!(matches!(cell_b1.value, CellValue::Number(v) if (v - 42.0).abs() < f64::EPSILON));

        let cell_c1 = store
            .get(worksheet.cells[2])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell c1");
        assert!(matches!(cell_c1.value, CellValue::Boolean(true)));

        let cell_d1 = store
            .get(worksheet.cells[3])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell d1");
        assert!(matches!(cell_d1.value, CellValue::Error(CellError::Ref)));

        let cell_e1 = store
            .get(worksheet.cells[4])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell e1");
        assert!(cell_e1.formula.is_some());

        let cell_f1 = store
            .get(worksheet.cells[5])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell f1");
        assert!(matches!(cell_f1.value, CellValue::InlineString(ref v) if v == "Inline"));
    }

    #[test]
    fn test_parse_worksheet_properties() {
        let xml = r#"
        <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <dimension ref="A1:C5"/>
          <sheetPr>
            <tabColor rgb="FF00FF"/>
          </sheetPr>
          <pageMargins left="0.5" right="0.6" top="0.7" bottom="0.8" header="0.3" footer="0.4"/>
          <sheetData/>
        </worksheet>
        "#;

        let mut parser = XlsxParser::new();
        let sheet = SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };
        let mut zip = build_empty_zip();
        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                xml,
                &sheet,
                "xl/worksheets/sheet1.xml",
                &Relationships::default(),
                SheetKind::Worksheet,
            )
            .expect("worksheet");
        let store = parser.into_store();

        let worksheet = match store.get(ws_id) {
            Some(IRNode::Worksheet(ws)) => ws,
            _ => panic!("expected worksheet"),
        };
        assert_eq!(worksheet.dimension.as_deref(), Some("A1:C5"));
        assert_eq!(worksheet.tab_color.as_deref(), Some("rgb:FF00FF"));
        let margins = worksheet.page_margins.as_ref().expect("margins");
        assert_eq!(margins.left, Some(0.5));
        assert_eq!(margins.right, Some(0.6));
        assert_eq!(margins.top, Some(0.7));
        assert_eq!(margins.bottom, Some(0.8));
        assert_eq!(margins.header, Some(0.3));
        assert_eq!(margins.footer, Some(0.4));
    }

    #[test]
    fn test_parse_worksheet_drawing_pic_and_chart() {
        let sheet_xml = r#"
        <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetData/>
        </worksheet>
        "#;

        let drawing_xml = r#"
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                 xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <xdr:twoCellAnchor>
            <xdr:pic>
              <xdr:nvPicPr>
                <xdr:cNvPr id="1" name="Picture 1" descr="Alt text"/>
              </xdr:nvPicPr>
              <xdr:blipFill>
                <a:blip r:embed="rIdImg"/>
              </xdr:blipFill>
            </xdr:pic>
          </xdr:twoCellAnchor>
          <xdr:graphicFrame>
            <xdr:nvGraphicFramePr>
              <xdr:cNvPr id="2" name="Chart 1"/>
            </xdr:nvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rIdChart"/>
              </a:graphicData>
            </a:graphic>
          </xdr:graphicFrame>
        </xdr:wsDr>
        "#;

        let chart_xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:barChart>
              <c:ser><c:tx><c:v>2019</c:v></c:tx></c:ser>
              <c:ser><c:tx><c:v>2020</c:v></c:tx></c:ser>
            </c:barChart>
          </c:chart>
        </c:chartSpace>
        "#;

        let drawing_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image1.png"/>
          <Relationship Id="rIdChart"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="../charts/chart1.xml"/>
        </Relationships>
        "#;

        let sheet_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdDraw"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
            Target="../drawings/drawing1.xml"/>
        </Relationships>
        "#;

        let mut zip = build_zip_with_entries(vec![
            ("xl/drawings/drawing1.xml", drawing_xml),
            ("xl/drawings/_rels/drawing1.xml.rels", drawing_rels),
            ("xl/charts/chart1.xml", chart_xml),
        ]);

        let mut parser = XlsxParser::new();
        let sheet = SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };
        let rels = Relationships::parse(sheet_rels).expect("sheet rels");

        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                sheet_xml,
                &sheet,
                "xl/worksheets/sheet1.xml",
                &rels,
                SheetKind::Worksheet,
            )
            .expect("parse worksheet");
        let store = parser.into_store();
        let ws = match store.get(ws_id) {
            Some(IRNode::Worksheet(w)) => w,
            _ => panic!("missing worksheet"),
        };
        assert_eq!(ws.drawings.len(), 1);
        let drawing = match store.get(ws.drawings[0]) {
            Some(IRNode::WorksheetDrawing(d)) => d,
            _ => panic!("missing drawing"),
        };
        assert_eq!(drawing.shapes.len(), 2);
    }

    #[test]
    fn test_parse_chartsheet_chart() {
        let chartsheet_xml = r#"
        <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <chart r:id="rIdChart"/>
        </chartsheet>
        "#;
        let chart_xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:lineChart/>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdChart"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="../charts/chart1.xml"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels");
        let mut zip = build_zip_with_entries(vec![("xl/charts/chart1.xml", chart_xml)]);
        let mut parser = XlsxParser::new();
        let sheet = SheetInfo {
            name: "Chart1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };
        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                chartsheet_xml,
                &sheet,
                "xl/chartsheets/sheet1.xml",
                &rels,
                SheetKind::ChartSheet,
            )
            .expect("chartsheet");
        let store = parser.into_store();
        let ws = match store.get(ws_id) {
            Some(IRNode::Worksheet(w)) => w,
            _ => panic!("missing worksheet"),
        };
        assert_eq!(ws.drawings.len(), 1);
        let drawing = match store.get(ws.drawings[0]) {
            Some(IRNode::WorksheetDrawing(d)) => d,
            _ => panic!("missing drawing"),
        };
        assert_eq!(drawing.shapes.len(), 1);
        let shape = match store.get(drawing.shapes[0]) {
            Some(IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.shape_type, ShapeType::Chart);
        assert_eq!(shape.media_target.as_deref(), Some("xl/charts/chart1.xml"));
    }

    #[test]
    fn test_parse_dialogsheet_kind() {
        let dialog_xml = r#"
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1" t="str"><v>Hello</v></c>
            </row>
          </sheetData>
        </worksheet>
        "#;
        let mut parser = XlsxParser::new();
        let mut zip = build_empty_zip();
        let sheet = SheetInfo {
            name: "Dialog1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };
        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                dialog_xml,
                &sheet,
                "xl/dialogsheets/sheet1.xml",
                &Relationships::default(),
                SheetKind::DialogSheet,
            )
            .expect("dialogsheet");
        let store = parser.into_store();
        let ws = match store.get(ws_id) {
            Some(IRNode::Worksheet(w)) => w,
            _ => panic!("missing worksheet"),
        };
        assert_eq!(ws.kind, SheetKind::DialogSheet);
        assert_eq!(ws.cells.len(), 1);
    }

    #[test]
    fn test_parse_styles_minimal() {
        let xml = r#"
        <styleSheet>
          <numFmts count="1">
            <numFmt numFmtId="164" formatCode="0.00"/>
          </numFmts>
          <fonts count="1">
            <font>
              <sz val="11"/>
              <name val="Calibri"/>
              <b/>
              <color rgb="FF0000"/>
            </font>
          </fonts>
          <fills count="1">
            <fill>
              <patternFill patternType="solid">
                <fgColor rgb="FFFF00"/>
              </patternFill>
            </fill>
          </fills>
          <borders count="1">
            <border>
              <left style="thin"><color rgb="FF00FF"/></left>
            </border>
          </borders>
          <cellXfs count="1">
            <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"
                applyNumberFormat="1" applyAlignment="1" applyProtection="1" quotePrefix="1">
              <alignment horizontal="center" wrapText="1" indent="2" textRotation="45"
                         shrinkToFit="1" readingOrder="1"/>
              <protection locked="1" hidden="0"/>
            </xf>
          </cellXfs>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>
          </cellStyleXfs>
          <dxfs count="1">
            <dxf>
              <numFmt numFmtId="200" formatCode="0.00"/>
              <font><b/><color rgb="FF0000"/></font>
              <fill><patternFill patternType="solid"><fgColor rgb="00FF00"/></patternFill></fill>
            </dxf>
          </dxfs>
          <tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">
            <tableStyle name="TableStyleMedium2" pivot="0" table="1"/>
          </tableStyles>
        </styleSheet>
        "#;

        let styles = parse_styles(xml, "xl/styles.xml").expect("styles");
        assert_eq!(styles.number_formats.len(), 1);
        assert_eq!(styles.fonts.len(), 1);
        assert_eq!(styles.fills.len(), 1);
        assert_eq!(styles.borders.len(), 1);
        assert_eq!(styles.cell_xfs.len(), 1);
        assert_eq!(styles.cell_style_xfs.len(), 1);
        assert_eq!(styles.dxfs.len(), 1);
        assert!(styles.table_styles.is_some());
        assert_eq!(styles.fonts[0].name.as_deref(), Some("Calibri"));
        assert!(styles.fonts[0].bold);
        assert_eq!(styles.cell_xfs[0].apply_number_format, true);
        assert_eq!(styles.dxfs[0].num_fmt.as_ref().map(|n| n.id), Some(200));
        assert_eq!(
            styles
                .table_styles
                .as_ref()
                .unwrap()
                .default_table_style
                .as_deref(),
            Some("TableStyleMedium2")
        );
        assert_eq!(styles.table_styles.as_ref().unwrap().styles.len(), 1);
        assert_eq!(
            styles.table_styles.as_ref().unwrap().styles[0].name,
            "TableStyleMedium2"
        );
        assert_eq!(
            styles.cell_xfs[0]
                .alignment
                .as_ref()
                .and_then(|a| a.horizontal.as_deref()),
            Some("center")
        );
        assert_eq!(
            styles.cell_xfs[0].alignment.as_ref().and_then(|a| a.indent),
            Some(2)
        );
        assert_eq!(
            styles.cell_xfs[0]
                .alignment
                .as_ref()
                .and_then(|a| a.text_rotation),
            Some(45)
        );
        assert!(styles.cell_xfs[0]
            .alignment
            .as_ref()
            .map(|a| a.shrink_to_fit)
            .unwrap_or(false));
        assert_eq!(
            styles.cell_xfs[0]
                .protection
                .as_ref()
                .and_then(|p| p.locked),
            Some(true)
        );
    }

    #[test]
    fn test_parse_calc_chain() {
        let xml = r#"
        <calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <c r="A1" i="0" l="1" s="1"/>
          <c r="B2" i="2"/>
        </calcChain>
        "#;
        let chain = parse_calc_chain(xml, "xl/calcChain.xml").expect("calc chain");
        assert_eq!(chain.entries.len(), 2);
        assert_eq!(chain.entries[0].cell_ref, "A1");
        assert_eq!(chain.entries[0].level, Some(1));
        assert_eq!(chain.entries[0].new_value, Some(true));
        assert_eq!(chain.entries[1].cell_ref, "B2");
        assert_eq!(chain.entries[1].index, Some(2));
    }

    #[test]
    fn test_parse_sheet_comments() {
        let xml = r#"
        <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <authors>
            <author>Alice</author>
            <author>Bob</author>
          </authors>
          <commentList>
            <comment ref="A1" authorId="0">
              <text><r><t>Hello</t></r></text>
            </comment>
            <comment ref="B2" authorId="1">
              <text><t>World</t></text>
            </comment>
          </commentList>
        </comments>
        "#;
        let comments =
            parse_sheet_comments(xml, "xl/comments1.xml", Some("Sheet1")).expect("comments");
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[0].cell_ref, "A1");
        assert_eq!(comments[0].author.as_deref(), Some("Alice"));
        assert_eq!(comments[0].text, "Hello");
        assert_eq!(comments[1].cell_ref, "B2");
        assert_eq!(comments[1].author.as_deref(), Some("Bob"));
        assert_eq!(comments[1].text, "World");
    }

    #[test]
    fn test_parse_conditional_and_validation() {
        let xml = r#"
        <worksheet>
          <sheetData/>
          <conditionalFormatting sqref="A1:A10">
            <cfRule type="expression" priority="1">
              <formula>SUM(A1)&gt;10</formula>
            </cfRule>
          </conditionalFormatting>
          <dataValidations count="1">
            <dataValidation type="list" allowBlank="1" sqref="B1">
              <formula1>"Yes,No"</formula1>
            </dataValidation>
          </dataValidations>
        </worksheet>
        "#;

        let mut parser = XlsxParser::new();
        let sheet = SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };
        let mut zip = build_empty_zip();
        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                xml,
                &sheet,
                "xl/worksheets/sheet1.xml",
                &Relationships::default(),
                SheetKind::Worksheet,
            )
            .expect("worksheet");
        let store = parser.into_store();

        let worksheet = match store.get(ws_id) {
            Some(IRNode::Worksheet(ws)) => ws,
            _ => panic!("expected worksheet"),
        };
        assert_eq!(worksheet.conditional_formats.len(), 1);
        assert_eq!(worksheet.data_validations.len(), 1);
    }

    #[test]
    fn test_parse_table_and_pivot_definition() {
        let table_xml = r#"
        <table name="Table1" displayName="Table1" ref="A1:B3" headerRowCount="1">
          <tableColumns count="2">
            <tableColumn id="1" name="Col1"/>
            <tableColumn id="2" name="Col2" totalsRowFunction="sum"/>
          </tableColumns>
        </table>
        "#;
        let table = parse_table_definition(table_xml, "xl/tables/table1.xml").expect("table");
        assert_eq!(table.columns.len(), 2);
        assert_eq!(table.ref_range.as_deref(), Some("A1:B3"));

        let pivot_xml = r#"
        <pivotTableDefinition name="Pivot1" cacheId="3">
          <location ref="D1:F10"/>
        </pivotTableDefinition>
        "#;
        let pivot = parse_pivot_table_definition(pivot_xml, "xl/pivotTables/pivotTable1.xml")
            .expect("pivot");
        assert_eq!(pivot.cache_id, Some(3));
        assert_eq!(pivot.ref_range.as_deref(), Some("D1:F10"));
    }

    #[test]
    fn test_parse_pivot_cache_records() {
        let records_xml = r#"
        <pivotCacheRecords count="2">
          <r><n v="1"/><s v="0"/><b v="1"/></r>
          <r><n v="2"/><s v="1"/><b v="0"/></r>
        </pivotCacheRecords>
        "#;

        let records =
            parse_pivot_cache_records(records_xml, "xl/pivotCache/pivotCacheRecords1.xml")
                .expect("records");
        assert_eq!(records.record_count, Some(2));
        assert_eq!(records.field_count, Some(3));
    }

    #[test]
    fn test_parse_xlm_macro_sheet() {
        let workbook_xml = r#"
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Macro1" sheetId="1" r:id="rId1"/>
          </sheets>
        </workbook>
        "#;

        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/macrosheet"
            Target="macrosheets/sheet1.xml"/>
        </Relationships>
        "#;

        let sheet_xml = r#"
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1"><f>EXEC(\"calc\")</f></c>
            </row>
          </sheetData>
        </worksheet>
        "#;

        let mut zip = build_zip_with_entries(vec![("xl/macrosheets/sheet1.xml", sheet_xml)]);
        let rels = Relationships::parse(rels_xml).expect("rels");

        let mut parser = XlsxParser::new();
        let root = parser
            .parse_workbook(&mut zip, workbook_xml, &rels, "xl/workbook.xml")
            .expect("workbook");
        let store = parser.into_store();
        let doc = match store.get(root) {
            Some(IRNode::Document(d)) => d,
            _ => panic!("missing document"),
        };
        assert_eq!(doc.security.xlm_macros.len(), 1);
        assert!(doc.security.xlm_macros[0].macro_cells.len() >= 1);
    }

    #[test]
    fn test_parse_xlm_auto_open_defined_name() {
        let workbook_xml = r#"
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <definedNames>
            <definedName name="_xlnm.Auto_Open">Macro1!$A$1</definedName>
          </definedNames>
          <sheets>
            <sheet name="Macro1" sheetId="1" r:id="rId1"/>
          </sheets>
        </workbook>
        "#;

        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/macrosheet"
            Target="macrosheets/sheet1.xml"/>
        </Relationships>
        "#;

        let sheet_xml = r#"
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1"><f>RUN(\"TEST\")</f></c>
            </row>
          </sheetData>
        </worksheet>
        "#;

        let mut zip = build_zip_with_entries(vec![("xl/macrosheets/sheet1.xml", sheet_xml)]);
        let rels = Relationships::parse(rels_xml).expect("rels");

        let mut parser = XlsxParser::new();
        let root = parser
            .parse_workbook(&mut zip, workbook_xml, &rels, "xl/workbook.xml")
            .expect("workbook");
        let store = parser.into_store();
        let doc = match store.get(root) {
            Some(IRNode::Document(d)) => d,
            _ => panic!("missing document"),
        };
        assert_eq!(doc.security.xlm_macros.len(), 1);
        assert!(doc.security.xlm_macros[0].has_auto_open);
    }

    #[test]
    fn test_parse_chart_xml() {
        let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart>
            <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:lineChart>
              <c:ser><c:tx><c:v>Q1</c:v></c:tx></c:ser>
            </c:lineChart>
          </c:chart>
        </c:chartSpace>
        "#;
        let mut parser = XlsxParser::new();
        let id = parser
            .parse_chart(xml, "xl/charts/chart2.xml")
            .expect("chart");
        let store = parser.into_store();
        let chart = match store.get(id) {
            Some(IRNode::ChartData(c)) => c,
            _ => panic!("missing chart"),
        };
        assert!(chart
            .chart_type
            .as_deref()
            .unwrap_or("")
            .contains("lineChart"));
        assert_eq!(chart.title.as_deref(), Some("Revenue"));
        assert_eq!(chart.series.len(), 1);
        assert_eq!(chart.series_data.len(), 1);
    }

    #[test]
    fn test_parse_connections_xml_targets() {
        let xml = r#"
        <connections>
          <connection id="1" name="Conn1" type="1">
            <webPr url="https://example.com/data"/>
          </connection>
          <connection id="2" name="Conn2" type="2">
            <dbPr connection="DatabaseName" command="SELECT * FROM foo" commandType="2"/>
          </connection>
        </connections>
        "#;
        let part = parse_connections_part(xml, "xl/connections.xml").expect("connections");
        assert_eq!(part.entries.len(), 2);
        assert_eq!(part.entries[0].connection_id, Some(1));
        assert_eq!(
            part.entries[0].url.as_deref(),
            Some("https://example.com/data")
        );
        assert_eq!(part.entries[1].connection_id, Some(2));
        assert_eq!(part.entries[1].connection.as_deref(), Some("DatabaseName"));
        assert_eq!(
            part.entries[1].command.as_deref(),
            Some("SELECT * FROM foo")
        );
        assert_eq!(part.entries[1].command_type, Some(2));
        let targets = connection_targets(&part);
        assert!(targets.contains(&"https://example.com/data".to_string()));
        assert!(targets.contains(&"DatabaseName".to_string()));
    }

    #[test]
    fn test_parse_sheet_metadata() {
        let xml = r#"
        <metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <metadataTypes count="1">
            <metadataType name="XLDynamicArray" minSupportedVersion="120000" copy="1" update="0"/>
          </metadataTypes>
          <cellMetadata count="2"/>
          <valueMetadata count="3"/>
        </metadata>
        "#;
        let meta = parse_sheet_metadata(xml, "xl/metadata.xml").expect("metadata");
        assert_eq!(meta.metadata_types.len(), 1);
        assert_eq!(
            meta.metadata_types[0].name.as_deref(),
            Some("XLDynamicArray")
        );
        assert_eq!(meta.cell_metadata_count, Some(2));
        assert_eq!(meta.value_metadata_count, Some(3));
    }

    #[test]
    fn test_parse_people_part() {
        let xml = r#"
        <ppl:people xmlns:ppl="http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes">
          <ppl:person ppl:id="p1" ppl:userId="user1" ppl:displayName="Alice" ppl:initials="A"/>
          <ppl:person ppl:id="p2" ppl:userId="user2" ppl:displayName="Bob"/>
        </ppl:people>
        "#;
        let mut parser = XlsxParser::new();
        let mut zip = build_zip_with_entries(vec![("xl/persons/person.xml", xml)]);
        let workbook_xml =
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
        let doc_id = parser
            .parse_workbook(
                &mut zip,
                workbook_xml,
                &Relationships::default(),
                "xl/workbook.xml",
            )
            .expect("workbook");
        let store = parser.into_store();

        let doc = match store.get(doc_id) {
            Some(IRNode::Document(d)) => d,
            _ => panic!("missing document"),
        };
        assert_eq!(doc.shared_parts.len(), 1);

        let people = match store.get(doc.shared_parts[0]) {
            Some(IRNode::PeoplePart(p)) => p,
            _ => panic!("missing people part"),
        };
        assert_eq!(people.people.len(), 2);
        assert_eq!(people.people[0].display_name.as_deref(), Some("Alice"));
        assert_eq!(people.people[1].display_name.as_deref(), Some("Bob"));
    }

    #[test]
    fn test_parse_external_link_part() {
        let xml = r#"
        <externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <externalBook r:id="rId1">
            <sheetNames>
              <sheetName val="SheetA"/>
              <sheetName val="SheetB"/>
            </sheetNames>
          </externalBook>
        </externalLink>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
            Target="file:///C:/data.xlsx"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels");
        let part = parse_external_link_part(xml, "xl/externalLinks/externalLink1.xml", Some(&rels))
            .expect("external link");
        assert_eq!(part.target.as_deref(), Some("file:///C:/data.xlsx"));
        assert_eq!(part.sheets.len(), 2);
        assert_eq!(part.sheets[0].name.as_deref(), Some("SheetA"));
    }

    #[test]
    fn test_parse_slicer_part() {
        let xml = r#"
        <slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                name="Slicer1" caption="Region" cache="1" />
        "#;
        let slicer = parse_slicer_part(xml, "xl/slicers/slicer1.xml").expect("slicer");
        assert_eq!(slicer.name.as_deref(), Some("Slicer1"));
        assert_eq!(slicer.caption.as_deref(), Some("Region"));
        assert_eq!(slicer.cache_id.as_deref(), Some("1"));
    }

    #[test]
    fn test_parse_timeline_part() {
        let xml = r#"
        <timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                  name="Timeline1" cache="2" />
        "#;
        let timeline = parse_timeline_part(xml, "xl/timelines/timeline1.xml").expect("timeline");
        assert_eq!(timeline.name.as_deref(), Some("Timeline1"));
        assert_eq!(timeline.cache_id.as_deref(), Some("2"));
    }

    #[test]
    fn test_parse_query_table_part() {
        let xml = r#"
        <queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    name="Query1" connectionId="7">
          <dbPr command="SELECT * FROM tbl"/>
          <webPr url="https://example.com/data"/>
        </queryTable>
        "#;
        let query = parse_query_table_part(xml, "xl/queryTables/queryTable1.xml").expect("query");
        assert_eq!(query.name.as_deref(), Some("Query1"));
        assert_eq!(query.connection_id.as_deref(), Some("7"));
        assert_eq!(query.command.as_deref(), Some("SELECT * FROM tbl"));
        assert_eq!(query.url.as_deref(), Some("https://example.com/data"));
    }

    fn build_empty_zip() -> crate::zip_handler::SecureZipReader<std::io::Cursor<Vec<u8>>> {
        build_zip_with_entries(Vec::new())
    }

    fn build_zip_with_entries(
        entries: Vec<(&str, &str)>,
    ) -> crate::zip_handler::SecureZipReader<std::io::Cursor<Vec<u8>>> {
        let mut data = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(&mut data));
            let options = zip::write::FileOptions::<()>::default();
            for (path, contents) in entries {
                writer.start_file(path, options).expect("start file");
                use std::io::Write;
                writer.write_all(contents.as_bytes()).expect("write file");
            }
            writer.finish().expect("finish zip");
        }
        crate::zip_handler::SecureZipReader::new(std::io::Cursor::new(data), Default::default())
            .expect("zip")
    }
}
