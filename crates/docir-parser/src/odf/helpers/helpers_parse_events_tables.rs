use super::{
    Event, IRNode, NodeId, ODF_CONTENT_XML, OdfLimitCounter, OdfReader, ParseError, Table,
    TableCell, TableCellProperties, TableRow, XmlScanControl, scan_xml_events_until_end,
};
use crate::odf::paragraph::parse_paragraph;
use crate::odf::text::build_paragraph;
use crate::xml_utils::{is_end_event_local, local_name, try_attr_value_by_suffix, xml_error};
use docir_core::visitor::IrStore;
use quick_xml::events::BytesStart;

const MAX_ODF_TABLE_DEPTH: usize = 100;

pub(super) fn parse_empty_table(store: &mut IrStore) -> NodeId {
    let table = Table::new();
    let table_id = table.id;
    store.insert(IRNode::Table(table));
    table_id
}

pub(super) fn parse_table(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    parse_table_at_depth(reader, store, limits, 0)
}

fn parse_table_at_depth(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    depth: usize,
) -> Result<NodeId, ParseError> {
    if depth > MAX_ODF_TABLE_DEPTH {
        return Err(ParseError::ResourceLimit(format!(
            "ODF table nesting depth exceeds maximum ({MAX_ODF_TABLE_DEPTH})"
        )));
    }
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
                        let cell_id = parse_table_cell(reader, e, store, limits, depth)?;
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
                        if let Some(span) = try_attr_value_by_suffix(
                            e,
                            &[b":number-columns-spanned"],
                            ODF_CONTENT_XML,
                        )?
                        .map(|v| parse_u32_attr(&v))
                        .transpose()?
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
    depth: usize,
) -> Result<NodeId, ParseError> {
    let mut cell = TableCell::new();
    if let Some(span) =
        try_attr_value_by_suffix(start, &[b":number-columns-spanned"], ODF_CONTENT_XML)?
            .map(|v| parse_u32_attr(&v))
            .transpose()?
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
            if let Event::Start(e) = event
                && local_name(e.name().as_ref()) == b"p"
            {
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
            } else if let Event::Start(e) = event
                && local_name(e.name().as_ref()) == b"table"
            {
                let next_depth = depth.checked_add(1).ok_or_else(|| {
                    ParseError::InvalidStructure("ODF table nesting depth overflow".to_string())
                })?;
                let table_id = parse_table_at_depth(reader, store, limits, next_depth)?;
                cell.content.push(table_id);
            } else if let Event::Empty(e) = event
                && local_name(e.name().as_ref()) == b"table"
            {
                cell.content.push(parse_empty_table(store));
            } else if let Event::Empty(e) = event
                && local_name(e.name().as_ref()) == b"p"
            {
                limits.bump_paragraphs(1)?;
                cell.content.push(build_paragraph(store, "", None, None));
            }
            Ok(XmlScanControl::Continue)
        },
    )?;

    let cell_id = cell.id;
    store.insert(IRNode::TableCell(cell));
    Ok(cell_id)
}

fn parse_u32_attr(value: &str) -> Result<u32, ParseError> {
    let parsed = value
        .parse()
        .map_err(|err| xml_error(ODF_CONTENT_XML, err))?;
    if parsed == 0 {
        return Err(xml_error(
            ODF_CONTENT_XML,
            "ODF cell span attributes must be positive",
        ));
    }
    Ok(parsed)
}

#[cfg(test)]
mod tests {
    use super::{MAX_ODF_TABLE_DEPTH, parse_table};
    use crate::odf::limits::OdfLimits;
    use crate::parser::ParserConfig;
    use docir_core::visitor::IrStore;
    use quick_xml::Reader;
    use quick_xml::events::Event;
    use std::io::Cursor;

    #[test]
    fn parse_table_rejects_excessive_nesting() {
        let mut xml = String::new();
        for _ in 0..=MAX_ODF_TABLE_DEPTH + 1 {
            xml.push_str("<table:table><table:table-row><table:table-cell>");
        }
        xml.push_str("<text:p>value</text:p>");
        for _ in 0..=MAX_ODF_TABLE_DEPTH + 1 {
            xml.push_str("</table:table-cell></table:table-row></table:table>");
        }

        let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
        let mut buf = Vec::new();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(_)) => {}
            other => panic!("expected table start, got {other:?}"),
        }
        let mut store = IrStore::new();
        let limits = OdfLimits::new(&ParserConfig::default(), false);

        let error = parse_table(&mut reader, &mut store, &limits)
            .expect_err("excessively nested tables must be rejected");
        assert!(error.to_string().contains("table nesting depth"));
    }
}
