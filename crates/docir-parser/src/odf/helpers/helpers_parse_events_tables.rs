use super::{
    scan_xml_events_until_end, Event, IRNode, NodeId, OdfLimitCounter, OdfReader, ParseError,
    Table, TableCell, TableCellProperties, TableRow, XmlScanControl, ODF_CONTENT_XML,
};
use crate::odf::paragraph::parse_paragraph;
use crate::xml_utils::{attr_value_by_suffix, is_end_event_local, local_name};
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
        |event| is_end_event_local(event, b"table"),
        |reader, event| {
            match event {
                Event::Start(e) => match local_name(e.name().as_ref()) {
                    b"table-row" => {
                        current_row = Some(TableRow::new());
                    }
                    b"table-cell" => {
                        let cell_id = parse_table_cell(reader, e, store, limits)?;
                        if let Some(row) = current_row.as_mut() {
                            row.cells.push(cell_id);
                        }
                    }
                    _ => {}
                },
                Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"table-row" => {
                        let row = TableRow::new();
                        let row_id = row.id;
                        store.insert(IRNode::TableRow(row));
                        table.rows.push(row_id);
                    }
                    b"table-cell" => {
                        let mut cell = TableCell::new();
                        if let Some(span) = attr_value_by_suffix(e, &[b":number-columns-spanned"])
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
                Event::End(e) if local_name(e.name().as_ref()) == b"table-row" => {
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
    if let Some(span) = attr_value_by_suffix(start, &[b":number-columns-spanned"])
        .and_then(|v| v.parse::<u32>().ok())
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
        |event| is_end_event_local(event, b"table-cell"),
        |reader, event| {
            if let Event::Start(e) = event {
                if local_name(e.name().as_ref()) == b"p" {
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
