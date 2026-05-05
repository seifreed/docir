use super::super::{parse_border, parse_paragraph_simple, span_from_reader, DocxParser};
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{attr_value, local_name, xml_error};
use docir_core::ir::{
    CellMargins, CellVerticalAlignment, MergeType, RowHeight, RowHeightRule, Table, TableAlignment,
    TableBorders, TableCell, TableCellProperties, TableRow, TableWidth, TableWidthType,
};
use docir_core::types::NodeId;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

pub(crate) fn parse_table(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<NodeId, ParseError> {
    let mut table = Table::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"tblPr" => {
                    parse_table_properties(reader, &mut table.properties)?;
                }
                b"tblGrid" => {
                    table.grid = parse_table_grid(reader)?;
                }
                b"tr" => {
                    let row_id = parse_table_row(parser, reader, rels)?;
                    table.rows.push(row_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"tbl" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    table.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = table.id;
    parser.store.insert(docir_core::ir::IRNode::Table(table));
    Ok(id)
}

pub(crate) fn parse_table_row(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<NodeId, ParseError> {
    let mut row = TableRow::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"trPr" => {
                    parse_table_row_properties(reader, &mut row.properties)?;
                }
                b"tc" => {
                    let cell_id = parse_table_cell(parser, reader, rels)?;
                    row.cells.push(cell_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"tr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    row.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = row.id;
    parser.store.insert(docir_core::ir::IRNode::TableRow(row));
    Ok(id)
}

pub(crate) fn parse_table_cell(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<NodeId, ParseError> {
    let mut cell = TableCell::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"tcPr" => {
                    parse_table_cell_properties(reader, &mut cell.properties)?;
                }
                b"p" => {
                    let para_id = parse_paragraph_simple(parser, reader, rels)?;
                    cell.content.push(para_id);
                }
                b"tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    cell.content.push(table_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"tc" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    cell.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = cell.id;
    parser.store.insert(docir_core::ir::IRNode::TableCell(cell));
    Ok(id)
}

pub(crate) fn parse_table_cell_properties(
    reader: &mut Reader<&[u8]>,
    props: &mut TableCellProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                apply_table_cell_property_event(reader, &e, props)?
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"tcPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn apply_table_cell_property_event(
    reader: &mut Reader<&[u8]>,
    event: &BytesStart<'_>,
    props: &mut TableCellProperties,
) -> Result<(), ParseError> {
    match local_name(event.name().as_ref()) {
        b"tcW" => {
            if let Some(val) = attr_value(event, b"w:w").and_then(|v| v.parse().ok()) {
                let width_type = match attr_value(event, b"w:type").as_deref() {
                    Some("dxa") => TableWidthType::Dxa,
                    Some("pct") => TableWidthType::Pct,
                    Some("auto") => TableWidthType::Auto,
                    _ => TableWidthType::Nil,
                };
                props.width = Some(TableWidth {
                    value: val,
                    width_type,
                });
            }
        }
        b"gridSpan" => {
            if let Some(val) = attr_value(event, b"w:val").and_then(|v| v.parse().ok()) {
                props.grid_span = Some(val);
            }
        }
        b"vMerge" => {
            let merge = match attr_value(event, b"w:val").as_deref() {
                Some("restart") => MergeType::Restart,
                Some("continue") => MergeType::Continue,
                _ => MergeType::Continue,
            };
            props.vertical_merge = Some(merge);
        }
        b"vAlign" => {
            if let Some(val) = attr_value(event, b"w:val") {
                props.vertical_align = match val.as_str() {
                    "center" => Some(CellVerticalAlignment::Center),
                    "bottom" => Some(CellVerticalAlignment::Bottom),
                    _ => Some(CellVerticalAlignment::Top),
                };
            }
        }
        b"tcBorders" => {
            if let Some(borders) = parse_table_borders(reader, b"tcBorders")? {
                props.borders = Some(borders);
            }
        }
        b"shd" => {
            if let Some(fill) = attr_value(event, b"w:fill") {
                props.shading = Some(fill);
            }
        }
        _ => {}
    }
    Ok(())
}

pub(crate) fn parse_table_properties(
    reader: &mut Reader<&[u8]>,
    props: &mut docir_core::ir::TableProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                apply_table_property_attrs(&e, props);
                match local_name(e.name().as_ref()) {
                    b"tblBorders" => {
                        if let Some(borders) = parse_table_borders(reader, b"tblBorders")? {
                            props.borders = Some(borders);
                        }
                    }
                    b"tblCellMar" => {
                        if let Some(margins) = parse_cell_margins(reader)? {
                            props.cell_margins = Some(margins);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                apply_table_property_attrs(&e, props);
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"tblPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn apply_table_property_attrs(event: &BytesStart<'_>, props: &mut docir_core::ir::TableProperties) {
    match local_name(event.name().as_ref()) {
        b"tblW" => {
            if let Some(val) = attr_value(event, b"w:w").and_then(|v| v.parse().ok()) {
                let width_type = match attr_value(event, b"w:type").as_deref() {
                    Some("dxa") => TableWidthType::Dxa,
                    Some("pct") => TableWidthType::Pct,
                    Some("auto") => TableWidthType::Auto,
                    _ => TableWidthType::Nil,
                };
                props.width = Some(TableWidth {
                    value: val,
                    width_type,
                });
            }
        }
        b"jc" => {
            if let Some(val) = attr_value(event, b"w:val") {
                props.alignment = match val.as_str() {
                    "center" => Some(TableAlignment::Center),
                    "right" => Some(TableAlignment::Right),
                    _ => Some(TableAlignment::Left),
                };
            }
        }
        b"tblStyle" => {
            if let Some(val) = attr_value(event, b"w:val") {
                props.style_id = Some(val);
            }
        }
        _ => {}
    }
}

pub(crate) fn parse_table_grid(
    reader: &mut Reader<&[u8]>,
) -> Result<Vec<docir_core::ir::GridColumn>, ParseError> {
    let mut buf = Vec::new();
    let mut grid = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == b"gridCol" {
                    if let Some(val) = attr_value(&e, b"w:w").and_then(|v| v.parse().ok()) {
                        grid.push(docir_core::ir::GridColumn { width: val });
                    }
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"tblGrid" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(grid)
}

pub(crate) fn parse_table_row_properties(
    reader: &mut Reader<&[u8]>,
    props: &mut docir_core::ir::TableRowProperties,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"trHeight" => {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        let rule = match attr_value(&e, b"w:hRule").as_deref() {
                            Some("exact") => RowHeightRule::Exact,
                            Some("atLeast") => RowHeightRule::AtLeast,
                            _ => RowHeightRule::Auto,
                        };
                        props.height = Some(RowHeight { value: val, rule });
                    }
                }
                b"tblHeader" => {
                    let is_header = !matches!(
                        attr_value(&e, b"w:val").as_deref(),
                        Some("0") | Some("false")
                    );
                    props.is_header = Some(is_header);
                }
                b"cantSplit" => {
                    let cant_split = !matches!(
                        attr_value(&e, b"w:val").as_deref(),
                        Some("0") | Some("false")
                    );
                    props.cant_split = Some(cant_split);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"trPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_table_borders(
    reader: &mut Reader<&[u8]>,
    end_tag: &[u8],
) -> Result<Option<TableBorders>, ParseError> {
    let mut buf = Vec::new();
    let mut borders = TableBorders::default();
    let mut has_any = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let border = parse_border(&e);
                if border.is_none() {
                    continue;
                }
                match local_name(e.name().as_ref()) {
                    b"top" => {
                        borders.top = border;
                        has_any = true;
                    }
                    b"bottom" => {
                        borders.bottom = border;
                        has_any = true;
                    }
                    b"left" => {
                        borders.left = border;
                        has_any = true;
                    }
                    b"right" => {
                        borders.right = border;
                        has_any = true;
                    }
                    b"insideH" => {
                        borders.inside_h = border;
                        has_any = true;
                    }
                    b"insideV" => {
                        borders.inside_v = border;
                        has_any = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == local_name(end_tag) {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    if has_any {
        Ok(Some(borders))
    } else {
        Ok(None)
    }
}

fn parse_cell_margins(reader: &mut Reader<&[u8]>) -> Result<Option<CellMargins>, ParseError> {
    let mut buf = Vec::new();
    let mut margins = CellMargins::default();
    let mut has_any = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let val = attr_value(&e, b"w:w").and_then(|v| v.parse().ok());
                match local_name(e.name().as_ref()) {
                    b"top" => {
                        margins.top = val;
                        has_any = true;
                    }
                    b"bottom" => {
                        margins.bottom = val;
                        has_any = true;
                    }
                    b"left" => {
                        margins.left = val;
                        has_any = true;
                    }
                    b"right" => {
                        margins.right = val;
                        has_any = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"tblCellMar" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/document.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }
    if has_any {
        Ok(Some(margins))
    } else {
        Ok(None)
    }
}
