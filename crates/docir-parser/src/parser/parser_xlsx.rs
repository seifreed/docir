use super::*;
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

        Ok(ParsedDocument {
            root_id,
            format: DocumentFormat::Spreadsheet,
            store,
            metrics: None,
        })
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

        let mut sheet_index: u32 = 1;
        for name in workbook.sheet_names().to_vec() {
            let range = match workbook.worksheet_range(&name) {
                Ok(r) => r,
                Err(_) => {
                    sheet_index += 1;
                    continue;
                }
            };
            let mut worksheet = Worksheet::new(name.clone(), sheet_index);
            worksheet.kind = SheetKind::Worksheet;
            worksheet.state = SheetState::Visible;
            worksheet.span = Some(SourceSpan::new("xl/workbook.bin"));

            let (start_row, start_col) = range.start().unwrap_or((0, 0));
            let mut cell_ids = Vec::new();
            for (row, col, value) in range.used_cells() {
                let abs_row = start_row + row as u32;
                let abs_col = start_col + col as u32;
                let reference = format!("{}{}", column_to_letter(abs_col), abs_row + 1);
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
            sheet_index += 1;
        }

        document.security = SecurityInfo::default();
        let doc_id = document.id;
        store.insert(IRNode::Document(document));

        self.finalize_ooxml_document(zip, content_types, &mut store, doc_id, metrics)?;

        Ok(ParsedDocument {
            root_id: doc_id,
            format: DocumentFormat::Spreadsheet,
            store,
            metrics: None,
        })
    }
}
