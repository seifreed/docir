use super::{
    attr_value, scan_xml_events_until_end, Event, IRNode, NodeId, OdfLimitCounter, OdfReader,
    ParseError, Table, TableCell, TableCellProperties, TableRow, XmlScanControl, ODF_CONTENT_XML,
};
use crate::odf::paragraph::parse_paragraph;
use docir_core::visitor::IrStore;
use quick_xml::events::BytesStart;

pub(super) fn parse_table(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut buf = Vec::new();
    let mut table = Table::new();
    let mut current_row: Option<TableRow> = None;

    scan_xml_events_until_end(
        reader,
        &mut buf,
        ODF_CONTENT_XML,
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"table:table"),
        |reader, event| {
            match event {
                Event::Start(e) => match e.name().as_ref() {
                    b"table:table-row" => {
                        current_row = Some(TableRow::new());
                    }
                    b"table:table-cell" => {
                        let cell_id = parse_table_cell(reader, e, store, limits)?;
                        if let Some(row) = current_row.as_mut() {
                            row.cells.push(cell_id);
                        }
                    }
                    _ => {}
                },
                Event::Empty(e) => match e.name().as_ref() {
                    b"table:table-row" => {
                        let row = TableRow::new();
                        let row_id = row.id;
                        store.insert(IRNode::TableRow(row));
                        table.rows.push(row_id);
                    }
                    b"table:table-cell" => {
                        let mut cell = TableCell::new();
                        if let Some(span) = attr_value(e, b"table:number-columns-spanned")
                            .and_then(|v| v.parse::<u32>().ok())
                        {
                            cell.properties = TableCellProperties {
                                grid_span: Some(span),
                                ..TableCellProperties::default()
                            };
                        }
                        let cell_id = cell.id;
                        store.insert(IRNode::TableCell(cell));
                        if let Some(row) = current_row.as_mut() {
                            row.cells.push(cell_id);
                        }
                    }
                    _ => {}
                },
                Event::End(e) if e.name().as_ref() == b"table:table-row" => {
                    if let Some(row) = current_row.take() {
                        let row_id = row.id;
                        store.insert(IRNode::TableRow(row));
                        table.rows.push(row_id);
                    }
                }
                _ => {}
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    let table_id = table.id;
    store.insert(IRNode::Table(table));
    Ok(table_id)
}

pub(super) fn parse_table_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut cell = TableCell::new();
    if let Some(span) =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok())
    {
        cell.properties = TableCellProperties {
            grid_span: Some(span),
            ..TableCellProperties::default()
        };
    }

    let mut buf = Vec::new();
    scan_xml_events_until_end(
        reader,
        &mut buf,
        ODF_CONTENT_XML,
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"table:table-cell"),
        |reader, event| {
            if let Event::Start(e) = event {
                if e.name().as_ref() == b"text:p" {
                    let paragraph_id = parse_paragraph(
                        reader,
                        e.name().as_ref(),
                        None,
                        None,
                        store,
                        &mut Vec::new(),
                        limits,
                    )?;
                    cell.content.push(paragraph_id);
                }
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    let cell_id = cell.id;
    store.insert(IRNode::TableCell(cell));
    Ok(cell_id)
}
