use super::worksheet_parse::WorksheetParseAccum;
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::ooxml::xlsx::{Worksheet, XlsxParser};
use crate::xml_utils::reader_from_str;
use quick_xml::events::Event;

#[test]
fn matches_worksheet_start_event_returns_false_for_unknown_tag() -> Result<(), ParseError> {
    let mut parser = XlsxParser::new();
    let relationships = Relationships::default();
    let mut worksheet = Worksheet::new("sheet-1", 1);
    let mut accum = WorksheetParseAccum::new();

    let mut reader = reader_from_str("<foo></foo>");
    let mut buf = Vec::new();
    let start = match reader.read_event_into(&mut buf)? {
        Event::Start(start) => start,
        _ => panic!("expected start"),
    };

    assert!(!parser.matches_worksheet_start_event(
        &mut reader,
        "xl/worksheets/sheet1.xml",
        &relationships,
        &mut worksheet,
        &mut accum,
        &start
    )?);
    assert!(worksheet.dimension.is_none());
    assert!(accum.is_cells_empty());
    Ok(())
}

#[test]
fn matches_worksheet_start_event_reads_sheet_dimension() -> Result<(), ParseError> {
    let mut parser = XlsxParser::new();
    let relationships = Relationships::default();
    let mut worksheet = Worksheet::new("sheet-1", 1);
    let mut accum = WorksheetParseAccum::new();

    let mut reader = reader_from_str("<dimension ref=\"A1:B2\"></dimension>");
    let mut buf = Vec::new();
    let start = match reader.read_event_into(&mut buf)? {
        Event::Start(start) => start,
        _ => panic!("expected start"),
    };

    assert!(parser.matches_worksheet_start_event(
        &mut reader,
        "xl/worksheets/sheet1.xml",
        &relationships,
        &mut worksheet,
        &mut accum,
        &start
    )?);
    assert_eq!(worksheet.dimension, Some("A1:B2".to_string()));
    Ok(())
}

#[test]
fn matches_worksheet_empty_event_reads_sheet_dimension() -> Result<(), ParseError> {
    let mut parser = XlsxParser::new();
    let relationships = Relationships::default();
    let mut worksheet = Worksheet::new("sheet-1", 1);
    let mut accum = WorksheetParseAccum::new();

    let mut reader = reader_from_str(r#"<dimension ref="C3:D4"/>"#);
    let mut buf = Vec::new();
    let empty = match reader.read_event_into(&mut buf)? {
        Event::Empty(empty) => empty,
        _ => panic!("expected empty"),
    };

    parser.matches_worksheet_empty_event(
        "xl/worksheets/sheet1.xml",
        &relationships,
        &mut worksheet,
        &mut accum,
        &empty,
    )?;

    assert_eq!(worksheet.dimension, Some("C3:D4".to_string()));
    assert!(accum.is_cells_empty());
    Ok(())
}
