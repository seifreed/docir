use self::worksheet_parse_events::{
    handle_worksheet_common_tag, parse_conditional_formatting_empty, parse_data_validations,
};
use crate::ooxml::relationships::Relationships;
use crate::ooxml::xlsx::workbook::SheetInfo;
use crate::ooxml::xlsx::{
    parse_conditional_formatting, Cell, ColumnDefinition, ConditionalFormat, IRNode,
    MergedCellRange, ParseError, SheetKind, Worksheet, XlsxParser,
};
use crate::xml_utils::{
    is_end_event_local, local_name, reader_from_str, scan_xml_events_until_end, XmlScanControl,
};
use crate::zip_handler::PackageReader;
use docir_core::ir::DataValidation;
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

#[path = "worksheet_parse_events.rs"]
mod worksheet_parse_events;
#[path = "worksheet_parse_external.rs"]
mod worksheet_parse_external;

pub(crate) struct WorksheetParseAccum {
    columns: HashMap<u32, ColumnDefinition>,
    merged_cells: Vec<MergedCellRange>,
    cells: Vec<NodeId>,
    conditional_formats: Vec<NodeId>,
    data_validations: Vec<NodeId>,
}

impl WorksheetParseAccum {
    pub(crate) fn new() -> Self {
        Self {
            columns: HashMap::new(),
            merged_cells: Vec::new(),
            cells: Vec::new(),
            conditional_formats: Vec::new(),
            data_validations: Vec::new(),
        }
    }

    #[cfg(test)]
    pub(crate) fn is_cells_empty(&self) -> bool {
        self.cells.is_empty()
    }
}

impl XlsxParser {
    pub(crate) fn parse_worksheet(
        &mut self,
        zip: &mut impl PackageReader,
        xml: &str,
        sheet: &SheetInfo,
        sheet_path: &str,
        relationships: &Relationships,
        kind: SheetKind,
    ) -> Result<NodeId, ParseError> {
        let mut worksheet = Worksheet::new(sheet.name.clone(), sheet.sheet_id);
        worksheet.state = sheet.state;
        worksheet.relationship_id = Some(sheet.rel_id.clone());
        worksheet.kind = kind;
        worksheet.span = Some(SourceSpan::new(sheet_path));

        self.current_sheet_kind = Some(kind);
        self.current_sheet_name = Some(sheet.name.clone());
        self.current_sheet_state = Some(sheet.state);
        self.current_xlm_index = None;

        if kind == SheetKind::MacroSheet {
            self.begin_macro_sheet(sheet);
        }

        let mut accum = self.parse_worksheet_xml(xml, sheet_path, relationships, &mut worksheet)?;

        // Deterministic ordering
        let mut columns_sorted: Vec<ColumnDefinition> = accum.columns.into_values().collect();
        columns_sorted.sort_by_key(|c| c.index);
        accum
            .merged_cells
            .sort_by_key(|r| (r.start_row, r.start_col, r.end_row, r.end_col));

        worksheet.columns = columns_sorted;
        worksheet.merged_cells = accum.merged_cells;
        worksheet.cells = accum.cells;
        worksheet.conditional_formats = accum.conditional_formats;
        worksheet.data_validations = accum.data_validations;

        worksheet.drawings = self.load_worksheet_drawings(zip, sheet_path, relationships)?;
        worksheet.tables = self.load_worksheet_tables(zip, sheet_path, relationships)?;
        worksheet.pivot_tables = self.load_worksheet_pivots(zip, sheet_path, relationships)?;
        worksheet.comments =
            self.load_worksheet_comments(zip, sheet_path, relationships, &sheet.name)?;

        if kind == SheetKind::ChartSheet {
            if let Some(drawing) = self.parse_chartsheet(zip, xml, sheet_path, relationships)? {
                worksheet.drawings.push(drawing);
            }
        }

        self.current_sheet_kind = None;
        self.current_sheet_name = None;
        self.current_sheet_state = None;
        self.current_xlm_index = None;

        let ws_id = worksheet.id;
        self.store.insert(IRNode::Worksheet(worksheet));
        Ok(ws_id)
    }

    fn parse_worksheet_xml(
        &mut self,
        xml: &str,
        sheet_path: &str,
        relationships: &Relationships,
        worksheet: &mut Worksheet,
    ) -> Result<WorksheetParseAccum, ParseError> {
        let mut accum = WorksheetParseAccum::new();
        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();
        scan_xml_events_until_end(
            &mut reader,
            &mut buf,
            sheet_path,
            |event| is_end_event_local(event, b"worksheet"),
            |reader, event| {
                match event {
                    Event::Start(start) => {
                        self.matches_worksheet_start_event(
                            reader,
                            sheet_path,
                            relationships,
                            worksheet,
                            &mut accum,
                            start,
                        )?;
                    }
                    Event::Empty(start) => {
                        self.matches_worksheet_empty_event(
                            sheet_path,
                            relationships,
                            worksheet,
                            &mut accum,
                            start,
                        )?;
                    }
                    _ => {}
                }
                Ok(XmlScanControl::Continue)
            },
        )?;

        Ok(accum)
    }

    pub(crate) fn matches_worksheet_start_event(
        &mut self,
        reader: &mut quick_xml::Reader<&[u8]>,
        sheet_path: &str,
        relationships: &Relationships,
        worksheet: &mut Worksheet,
        accum: &mut WorksheetParseAccum,
        start: &BytesStart<'_>,
    ) -> Result<bool, ParseError> {
        if handle_worksheet_common_tag(start, sheet_path, relationships, worksheet, accum, self) {
            return Ok(true);
        }

        match local_name(start.name().as_ref()) {
            b"c" => {
                let cell = self.parse_cell(reader, start, sheet_path)?;
                self.insert_cell(cell, accum);
                return Ok(true);
            }
            b"conditionalFormatting" => {
                let fmt = parse_conditional_formatting(reader, start, sheet_path)?;
                self.insert_conditional_format(fmt, accum);
                return Ok(true);
            }
            b"dataValidations" => {
                let vals = parse_data_validations(reader, sheet_path)?;
                for val in vals {
                    self.insert_data_validation(val, accum);
                }
                return Ok(true);
            }
            _ => {}
        }

        Ok(false)
    }

    pub(crate) fn matches_worksheet_empty_event(
        &mut self,
        sheet_path: &str,
        relationships: &Relationships,
        worksheet: &mut Worksheet,
        accum: &mut WorksheetParseAccum,
        empty: &BytesStart<'_>,
    ) -> Result<(), ParseError> {
        if handle_worksheet_common_tag(empty, sheet_path, relationships, worksheet, accum, self) {
            return Ok(());
        }

        match local_name(empty.name().as_ref()) {
            b"c" => {
                let cell = self.parse_empty_cell(empty, sheet_path)?;
                self.insert_cell(cell, accum);
            }
            b"conditionalFormatting" => {
                let fmt = parse_conditional_formatting_empty(empty, sheet_path);
                self.insert_conditional_format(fmt, accum);
            }
            b"dataValidations" => {}
            _ => {}
        }

        Ok(())
    }

    fn insert_cell(&mut self, cell: Cell, accum: &mut WorksheetParseAccum) {
        let cell_id = cell.id;
        self.store.insert(IRNode::Cell(cell));
        accum.cells.push(cell_id);
    }

    fn insert_conditional_format(
        &mut self,
        fmt: ConditionalFormat,
        accum: &mut WorksheetParseAccum,
    ) {
        let id = fmt.id;
        self.store.insert(IRNode::ConditionalFormat(fmt));
        accum.conditional_formats.push(id);
    }

    fn insert_data_validation(
        &mut self,
        validation: DataValidation,
        accum: &mut WorksheetParseAccum,
    ) {
        let id = validation.id;
        self.store.insert(IRNode::DataValidation(validation));
        accum.data_validations.push(id);
    }
}
