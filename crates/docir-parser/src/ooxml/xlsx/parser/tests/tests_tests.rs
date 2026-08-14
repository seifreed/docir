use crate::error::ParseError;
use crate::ooxml::xlsx::{
    parse_column, parse_conditional_formatting, parse_merge_cell, parse_threaded_comments,
};
use quick_xml::Reader;
use quick_xml::events::Event;
#[test]
fn test_parse_column_and_merge_helpers() {
    let mut columns = std::collections::HashMap::new();

    let mut reader =
        Reader::from_str(r#"<col min="2" max="4" width="12.5" hidden="1" customWidth="1"/>"#);
    let mut buf = Vec::new();
    let col = match reader.read_event_into(&mut buf).expect("col") {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    parse_column(&col, &mut columns, "xl/worksheets/sheet1.xml").expect("column");
    assert_eq!(columns.len(), 3);
    assert_eq!(columns.get(&1).and_then(|c| c.width), Some(12.5));
    assert!(columns.get(&3).map(|c| c.hidden).unwrap_or(false));
    assert!(columns.get(&3).map(|c| c.custom_width).unwrap_or(false));

    let mut ignored_reader = Reader::from_str(r#"<col max="3"/>"#);
    buf.clear();
    let ignored = match ignored_reader
        .read_event_into(&mut buf)
        .expect("ignored col")
    {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    assert!(parse_column(&ignored, &mut columns, "xl/worksheets/sheet1.xml").is_err());
    assert_eq!(columns.len(), 3, "incomplete columns are ignored");

    let mut merge_reader = Reader::from_str(r#"<mergeCell ref="A1:C3"/>"#);
    buf.clear();
    let merge = match merge_reader.read_event_into(&mut buf).expect("merge") {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let range = parse_merge_cell(&merge, "xl/worksheets/sheet1.xml")
        .expect("merge attrs")
        .expect("valid merge");
    assert_eq!((range.start_col, range.start_row), (0, 0));
    assert_eq!((range.end_col, range.end_row), (2, 2));

    let mut single_merge_reader = Reader::from_str(r#"<mergeCell ref="D10"/>"#);
    buf.clear();
    let single = match single_merge_reader
        .read_event_into(&mut buf)
        .expect("single merge")
    {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let single_range = parse_merge_cell(&single, "xl/worksheets/sheet1.xml")
        .expect("single merge attrs")
        .expect("single ref merge");
    assert_eq!((single_range.start_col, single_range.end_col), (3, 3));
    assert_eq!((single_range.start_row, single_range.end_row), (9, 9));

    let mut bad_merge_reader = Reader::from_str(r#"<mergeCell ref="not-a-cell"/>"#);
    buf.clear();
    let bad = match bad_merge_reader
        .read_event_into(&mut buf)
        .expect("bad merge")
    {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    assert!(parse_merge_cell(&bad, "xl/worksheets/sheet1.xml").is_err());

    let mut missing_reader = Reader::from_str(r#"<mergeCell/>"#);
    buf.clear();
    let missing = match missing_reader
        .read_event_into(&mut buf)
        .expect("missing merge ref")
    {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    assert!(parse_merge_cell(&missing, "xl/worksheets/sheet1.xml").is_err());

    let mut out_of_range_reader = Reader::from_str(r#"<mergeCell ref="XFE1:XFE2"/>"#);
    buf.clear();
    let out_of_range = match out_of_range_reader
        .read_event_into(&mut buf)
        .expect("out-of-range merge")
    {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let err = parse_merge_cell(&out_of_range, "xl/worksheets/sheet1.xml")
        .expect_err("merge outside worksheet limits must fail");
    assert!(matches!(err, ParseError::Xml { .. }));
}

#[test]
fn parse_column_and_merge_helpers_report_malformed_attributes() {
    let mut buf = Vec::new();
    let mut columns = std::collections::HashMap::new();

    let mut column_reader = Reader::from_str(r#"<col min="2" min="3" max="4"/>"#);
    let col = match column_reader.read_event_into(&mut buf).expect("column") {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let err = parse_column(&col, &mut columns, "xl/worksheets/sheet1.xml")
        .expect_err("duplicate column attrs must fail");
    assert!(matches!(
        err,
        ParseError::Xml {
            ref file,
            ..
        } if file == "xl/worksheets/sheet1.xml"
    ));

    for xml in [
        r#"<col min="bad" max="4"/>"#,
        r#"<col min="0" max="4"/>"#,
        r#"<col min="2" max="bad"/>"#,
        r#"<col min="2" max="0"/>"#,
        r#"<col min="4" max="2"/>"#,
        r#"<col min="1" max="4294967295"/>"#,
        r#"<col min="2" max="4" width="bad"/>"#,
        r#"<col min="2" max="4" width="NaN"/>"#,
    ] {
        let mut reader = Reader::from_str(xml);
        buf.clear();
        let col = match reader.read_event_into(&mut buf).expect("column") {
            Event::Empty(e) => e.into_owned(),
            other => panic!("unexpected event: {other:?}"),
        };
        let err = parse_column(&col, &mut columns, "xl/worksheets/sheet1.xml")
            .expect_err("malformed column numeric attrs must fail");
        assert!(matches!(
            err,
            ParseError::Xml {
                ref file,
                ..
            } if file == "xl/worksheets/sheet1.xml"
        ));
    }

    let mut merge_reader = Reader::from_str(r#"<mergeCell ref="A1:B2" ref="C1:D2"/>"#);
    buf.clear();
    let merge = match merge_reader.read_event_into(&mut buf).expect("merge") {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let err = parse_merge_cell(&merge, "xl/worksheets/sheet1.xml")
        .expect_err("duplicate merge attrs must fail");
    assert!(matches!(
        err,
        ParseError::Xml {
            ref file,
            ..
        } if file == "xl/worksheets/sheet1.xml"
    ));

    let xml = r#"<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1">"#;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    buf.clear();
    let start = match reader.read_event_into(&mut buf).expect("conditional start") {
        Event::Start(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    assert!(matches!(
        parse_conditional_formatting(&mut reader, &start, "xl/worksheets/sheet1.xml"),
        Err(ParseError::Xml { .. })
    ));
}

#[test]
fn test_parse_conditional_formatting_and_threaded_comments() {
    let xml = r#"
        <conditionalFormatting sqref="A1 B2:C3">
          <cfRule type="cellIs" priority="9" operator="between">
            <formula>1</formula>
            <formula>10</formula>
          </cfRule>
        </conditionalFormatting>
    "#;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let start = match reader.read_event_into(&mut buf).expect("conditional start") {
        Event::Start(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let conditional = parse_conditional_formatting(&mut reader, &start, "xl/worksheets/sheet1.xml")
        .expect("conditional");
    assert_eq!(conditional.ranges, vec!["A1", "B2:C3"]);
    assert_eq!(conditional.rules.len(), 1);
    assert_eq!(conditional.rules[0].rule_type, "cellIs");
    assert_eq!(conditional.rules[0].priority, Some(9));
    assert_eq!(conditional.rules[0].operator.as_deref(), Some("between"));
    assert_eq!(conditional.rules[0].formulae, vec!["1", "10"]);

    let threaded_xml = r#"
        <ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
          <threadedComment ref="C5" personId="{AA}">
            <text>Threaded note</text>
          </threadedComment>
        </ThreadedComments>
    "#;
    let threaded = parse_threaded_comments(
        threaded_xml,
        "xl/threadedComments/threadedComment1.xml",
        None,
    )
    .expect("threaded comments");
    assert_eq!(threaded.len(), 1);
    assert_eq!(threaded[0].cell_ref, "C5");
    assert_eq!(threaded[0].text, "Threaded note");
}

#[test]
fn parse_conditional_formatting_preserves_empty_rule() {
    let xml = r#"<conditionalFormatting sqref="A1">
        <cfRule type="aboveAverage" priority="3"/>
    </conditionalFormatting>"#;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let start = match reader.read_event_into(&mut buf).expect("conditional start") {
        Event::Start(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };

    let conditional = parse_conditional_formatting(&mut reader, &start, "xl/worksheets/sheet1.xml")
        .expect("empty conditional rule should parse");
    assert_eq!(conditional.rules.len(), 1);
    assert_eq!(conditional.rules[0].rule_type, "aboveAverage");
    assert_eq!(conditional.rules[0].priority, Some(3));
}

#[test]
fn parse_conditional_formatting_reports_malformed_attributes() {
    let xml = r#"
        <conditionalFormatting sqref="A1" sqref="B2">
          <cfRule type="cellIs" priority="1">
            <formula>1</formula>
          </cfRule>
        </conditionalFormatting>
    "#;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let start = match reader.read_event_into(&mut buf).expect("conditional start") {
        Event::Start(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let err = parse_conditional_formatting(&mut reader, &start, "xl/worksheets/sheet1.xml")
        .expect_err("duplicate conditional attrs must fail");
    assert!(matches!(
        err,
        ParseError::Xml {
            ref file,
            ..
        } if file == "xl/worksheets/sheet1.xml"
    ));

    let xml = r#"
        <conditionalFormatting sqref="A1">
          <cfRule type="cellIs" priority="bad">
            <formula>1</formula>
          </cfRule>
        </conditionalFormatting>
    "#;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    buf.clear();
    let start = match reader.read_event_into(&mut buf).expect("conditional start") {
        Event::Start(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let err = parse_conditional_formatting(&mut reader, &start, "xl/worksheets/sheet1.xml")
        .expect_err("malformed conditional priority must fail");
    assert!(matches!(
        err,
        ParseError::Xml {
            ref file,
            ..
        } if file == "xl/worksheets/sheet1.xml"
    ));
}

pub(crate) fn build_empty_zip() -> crate::zip_handler::SecureZipReader<std::io::Cursor<Vec<u8>>> {
    build_zip_with_entries(Vec::new())
}

pub(crate) fn build_zip_with_entries(
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
