use super::*;
use crate::ooxml::docx::document::table::parse_table;
#[test]
fn test_parse_table_returns_xml_error_for_malformed_input() {
    let xml = r#"
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tr>
            <w:tc>
              <w:p><w:r><w:t>Broken
            </w:tc>
          </w:tr>
        </w:tbl>
        "#;
    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut err = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:tbl" => {
                err = parse_table(&mut parser, &mut reader, &rels).err();
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error before table parse: {}", e),
            _ => {}
        }
        buf.clear();
    }

    let err = err.expect("expected parse_table to fail");
    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "word/document.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_table_applies_auto_width_and_row_defaults() {
    let xml = r#"
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tblPr>
            <w:tblW w:w="0" w:type="auto"/>
            <w:jc w:val="left"/>
            <w:tblBorders>
              <w:left w:val="single" w:sz="4" w:color="111111"/>
              <w:right w:val="single" w:sz="4" w:color="222222"/>
              <w:bottom w:val="single" w:sz="4" w:color="333333"/>
            </w:tblBorders>
          </w:tblPr>
          <w:tr>
            <w:trPr>
              <w:trHeight w:val="280" w:hRule="customRule"/>
              <w:tblHeader w:val="false"/>
              <w:cantSplit/>
            </w:trPr>
            <w:tc>
              <w:p><w:r><w:t>Cell</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
        "#;

    let (parser, table_id) = parse_single_table(xml);
    let store = parser.into_store();
    let table = match store.get(table_id) {
        Some(docir_core::ir::IRNode::Table(t)) => t,
        _ => panic!("missing table"),
    };

    assert!(matches!(
        table.properties.width.as_ref().map(|w| w.width_type),
        Some(docir_core::ir::TableWidthType::Auto)
    ));
    assert!(matches!(
        table.properties.alignment,
        Some(docir_core::ir::TableAlignment::Left)
    ));
    assert!(table
        .properties
        .borders
        .as_ref()
        .and_then(|b| b.left.as_ref())
        .is_some());
    assert!(table
        .properties
        .borders
        .as_ref()
        .and_then(|b| b.right.as_ref())
        .is_some());
    assert!(table
        .properties
        .borders
        .as_ref()
        .and_then(|b| b.bottom.as_ref())
        .is_some());

    let row = match store.get(table.rows[0]) {
        Some(docir_core::ir::IRNode::TableRow(r)) => r,
        _ => panic!("missing row"),
    };
    assert!(matches!(
        row.properties.height.as_ref().map(|h| h.rule),
        Some(docir_core::ir::RowHeightRule::Auto)
    ));
    assert_eq!(row.properties.is_header, Some(false));
    assert_eq!(row.properties.cant_split, Some(true));
}
#[test]
fn test_parse_table_cell_applies_width_auto_continue_merge_and_top_align() {
    let xml = r#"
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tr>
            <w:tc>
              <w:tcPr>
                <w:tcW w:w="1800" w:type="auto"/>
                <w:vMerge w:val="continue"/>
                <w:vAlign w:val="unexpected"/>
                <w:tcBorders>
                  <w:top w:val="single" w:sz="4" w:color="AA0000"/>
                  <w:right w:val="single" w:sz="4" w:color="00AA00"/>
                </w:tcBorders>
              </w:tcPr>
              <w:p><w:r><w:t>X</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
        "#;

    let (parser, table_id) = parse_single_table(xml);
    let store = parser.into_store();
    let table = match store.get(table_id) {
        Some(docir_core::ir::IRNode::Table(t)) => t,
        _ => panic!("missing table"),
    };
    let row = match store.get(table.rows[0]) {
        Some(docir_core::ir::IRNode::TableRow(r)) => r,
        _ => panic!("missing row"),
    };
    let cell = match store.get(row.cells[0]) {
        Some(docir_core::ir::IRNode::TableCell(c)) => c,
        _ => panic!("missing cell"),
    };

    assert!(matches!(
        cell.properties.width.as_ref().map(|w| w.width_type),
        Some(docir_core::ir::TableWidthType::Auto)
    ));
    assert!(matches!(
        cell.properties.vertical_merge,
        Some(docir_core::ir::MergeType::Continue)
    ));
    assert!(matches!(
        cell.properties.vertical_align,
        Some(docir_core::ir::CellVerticalAlignment::Top)
    ));
    assert!(cell
        .properties
        .borders
        .as_ref()
        .and_then(|b| b.top.as_ref())
        .is_some());
    assert!(cell
        .properties
        .borders
        .as_ref()
        .and_then(|b| b.right.as_ref())
        .is_some());
}
#[test]
fn test_parse_table_properties_ignore_unknown_border_children_and_keep_margin_shell() {
    let xml = r#"
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tblPr>
            <w:tblW w:w="4200" w:type="dxa"/>
            <w:tblBorders>
              <w:between w:val="single" w:sz="4" w:color="AA00AA"/>
            </w:tblBorders>
            <w:tblCellMar>
              <w:top w:w="bad"/>
            </w:tblCellMar>
          </w:tblPr>
          <w:tr>
            <w:tc><w:p><w:r><w:t>M</w:t></w:r></w:p></w:tc>
          </w:tr>
        </w:tbl>
        "#;

    let (parser, table_id) = parse_single_table(xml);
    let store = parser.into_store();
    let table = match store.get(table_id) {
        Some(docir_core::ir::IRNode::Table(t)) => t,
        _ => panic!("missing table"),
    };

    assert_eq!(table.properties.width.as_ref().map(|w| w.value), Some(4200));
    assert!(
        table.properties.borders.is_none(),
        "unknown tblBorders children should not produce recognized border sides"
    );

    let margins = table
        .properties
        .cell_margins
        .as_ref()
        .expect("tblCellMar should be retained when recognized side tags exist");
    assert_eq!(margins.top, None, "invalid numeric margin is ignored");
}
#[test]
fn test_parse_table_cell_ignores_invalid_numeric_attrs_and_keeps_nil_border() {
    let xml = r#"
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tr>
            <w:tc>
              <w:tcPr>
                <w:tcW w:w="oops" w:type="dxa"/>
                <w:gridSpan w:val="oops"/>
                <w:vMerge w:val="bogus"/>
                <w:tcBorders>
                  <w:left w:val="nil" w:color="auto"/>
                  <w:between w:val="single" w:sz="6" w:color="AA00AA"/>
                </w:tcBorders>
              </w:tcPr>
              <w:p><w:r><w:t>Cell</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
        "#;

    let (parser, table_id) = parse_single_table(xml);
    let store = parser.into_store();
    let table = match store.get(table_id) {
        Some(docir_core::ir::IRNode::Table(t)) => t,
        _ => panic!("missing table"),
    };
    let row = match store.get(table.rows[0]) {
        Some(docir_core::ir::IRNode::TableRow(r)) => r,
        _ => panic!("missing row"),
    };
    let cell = match store.get(row.cells[0]) {
        Some(docir_core::ir::IRNode::TableCell(c)) => c,
        _ => panic!("missing cell"),
    };

    assert!(cell.properties.width.is_none());
    assert_eq!(cell.properties.grid_span, None);
    assert!(matches!(
        cell.properties.vertical_merge,
        Some(docir_core::ir::MergeType::Continue)
    ));

    let borders = cell
        .properties
        .borders
        .as_ref()
        .expect("expected cell borders");
    let left = borders.left.as_ref().expect("left border should be parsed");
    assert!(matches!(left.style, docir_core::ir::BorderStyle::None));
    assert_eq!(left.color, None, "auto color should normalize to None");
    assert!(borders.inside_h.is_none());
    assert!(borders.inside_v.is_none());
}
