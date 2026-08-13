use super::{
    Cell, CellValue, ContentTypes, Cursor, DocumentFormat, IRNode, OoxmlParser, ParseError,
    ParseMetrics, ParsedDocument, SecurityInfo, SheetKind, SheetState, SourceSpan, Worksheet,
    XlsxParser, column_to_letter, map_calamine_error,
};
use crate::ooxml::part_utils::read_xml_part_and_rels;
use crate::parse_utils::init_store_and_document;
use crate::zip_handler::PackageReader;

impl OoxmlParser {
    /// Parse an XLSX document.
    pub(super) fn parse_xlsx(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        content_types: &ContentTypes,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<ParsedDocument, ParseError> {
        let (workbook_xml, workbook_rels) = read_xml_part_and_rels(zip, main_part_path)?;

        let mut parser = XlsxParser::new();
        let root_id = parser.parse_workbook(zip, &workbook_xml, &workbook_rels, main_part_path)?;
        let mut store = parser.into_store();

        self.finalize_ooxml_document(zip, content_types, &mut store, root_id, metrics)?;

        Ok(self.build_parsed_document(root_id, DocumentFormat::Spreadsheet, store))
    }

    /// Parse an XLSB document using calamine for binary sheets.
    pub(super) fn parse_xlsb(
        &self,
        zip: &mut impl PackageReader,
        data: &[u8],
        content_types: &ContentTypes,
        metrics: &mut Option<ParseMetrics>,
    ) -> Result<ParsedDocument, ParseError> {
        use calamine::{Data, Reader, Xlsb};

        let mut workbook = Xlsb::new(Cursor::new(data))
            .map_err(|e| ParseError::InvalidFormat(format!("XLSB parse error: {}", e)))?;

        let (mut store, mut document) = init_store_and_document(DocumentFormat::Spreadsheet);
        document.span = Some(SourceSpan::new("xl/workbook.bin"));

        for (sheet_index, name) in (1_u32..).zip(workbook.sheet_names().to_vec()) {
            let range = workbook
                .worksheet_range(&name)
                .map_err(|err| xlsb_sheet_error(&name, err))?;
            let mut worksheet = Worksheet::new(name.clone(), sheet_index);
            worksheet.kind = SheetKind::Worksheet;
            worksheet.state = SheetState::Visible;
            worksheet.span = Some(SourceSpan::new("xl/workbook.bin"));

            let (start_row, start_col) = range.start().unwrap_or((0, 0));
            let mut cell_ids = Vec::new();
            for (row, col, value) in range.used_cells() {
                let (reference, abs_row, abs_col) = xlsb_cell((start_row, start_col), (row, col))?;
                let mut cell = Cell::new(reference, abs_col, abs_row);
                cell.value = match value {
                    Data::Empty => CellValue::Empty,
                    Data::String(s) => CellValue::String(s.to_string()),
                    Data::Float(f) => CellValue::Number(*f),
                    Data::Int(i) => CellValue::Number(*i as f64),
                    Data::Bool(b) => CellValue::Boolean(*b),
                    Data::DateTime(dt) => CellValue::DateTime(dt.as_f64()),
                    Data::DateTimeIso(s) => CellValue::String(s.to_string()),
                    Data::DurationIso(s) => CellValue::String(s.to_string()),
                    Data::Error(e) => CellValue::Error(map_calamine_error(e.clone())),
                };
                cell.span = Some(SourceSpan::new("xl/workbook.bin"));
                let cell_id = cell.id;
                store.insert(IRNode::Cell(cell));
                cell_ids.push(cell_id);
            }

            worksheet.cells = cell_ids;
            let sheet_id = worksheet.id;
            store.insert(IRNode::Worksheet(worksheet));
            document.content.push(sheet_id);
        }

        document.security = SecurityInfo::default();
        let doc_id = document.id;
        store.insert(IRNode::Document(document));

        self.finalize_ooxml_document(zip, content_types, &mut store, doc_id, metrics)?;

        Ok(self.build_parsed_document(doc_id, DocumentFormat::Spreadsheet, store))
    }
}

fn xlsb_cell(start: (u32, u32), offset: (usize, usize)) -> Result<(String, u32, u32), ParseError> {
    let row_offset = u32::try_from(offset.0).map_err(|_| {
        ParseError::InvalidStructure("XLSB row offset does not fit in u32".to_string())
    })?;
    let col_offset = u32::try_from(offset.1).map_err(|_| {
        ParseError::InvalidStructure("XLSB column offset does not fit in u32".to_string())
    })?;
    let row = start
        .0
        .checked_add(row_offset)
        .ok_or_else(|| ParseError::InvalidStructure("XLSB row coordinate overflow".to_string()))?;
    let col = start.1.checked_add(col_offset).ok_or_else(|| {
        ParseError::InvalidStructure("XLSB column coordinate overflow".to_string())
    })?;
    let row_number = row
        .checked_add(1)
        .ok_or_else(|| ParseError::InvalidStructure("XLSB row number overflow".to_string()))?;
    Ok((format!("{}{}", column_to_letter(col), row_number), row, col))
}

fn xlsb_sheet_error(name: &str, err: calamine::XlsbError) -> ParseError {
    ParseError::InvalidFormat(format!("XLSB sheet '{name}' parse error: {err}"))
}

#[cfg(test)]
mod tests {
    use super::{xlsb_cell, xlsb_sheet_error};
    use crate::ParseError;

    #[test]
    fn xlsb_sheet_error_reports_sheet_name() {
        let err = xlsb_sheet_error("Sheet1", calamine::XlsbError::FileNotFound("sheet".into()));

        match err {
            ParseError::InvalidFormat(message) => {
                assert!(message.contains("Sheet1"));
                assert!(message.contains("XLSB sheet"));
            }
            other => panic!("expected invalid XLSB sheet error, got {other:?}"),
        }
    }

    #[test]
    fn xlsb_cell_rejects_coordinate_overflow() {
        assert!(xlsb_cell((u32::MAX, 0), (0, 0)).is_err());
        assert!(xlsb_cell((0, u32::MAX), (0, 1)).is_err());
        assert_eq!(
            xlsb_cell((4, 1), (2, 3)).expect("cell"),
            ("E7".to_string(), 6, 4)
        );
    }
}
