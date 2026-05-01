use crate::error::ParseError;
use crate::ooxml::docx::document::table::table_parse::{
    parse_table, parse_table_cell_properties, parse_table_grid, parse_table_properties,
    parse_table_row, parse_table_row_properties,
};
use crate::ooxml::docx::DocxParser;
use crate::ooxml::relationships::Relationships;

mod tests {
    use super::*;
    use docir_core::ir::{
        CellVerticalAlignment, MergeType, RowHeightRule, TableAlignment, TableCellProperties,
        TableProperties, TableRowProperties, TableWidthType,
    };
    use quick_xml::events::Event;
    use quick_xml::Reader;

    fn reader_from(xml: &str) -> Reader<&[u8]> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        reader
    }

    fn seek_start(reader: &mut Reader<&[u8]>, tag: &[u8]) {
        let mut buf = Vec::new();
        loop {
            let event = reader
                .read_event_into(&mut buf)
                .expect("unexpected xml read error");
            match event {
                Event::Start(e) if e.name().as_ref() == tag => break,
                Event::Eof => panic!("start tag not found"),
                _ => {}
            }
            buf.clear();
        }
    }

    #[test]
    fn parse_table_properties_reads_width_alignment_style_borders_and_margins() {
        let xml = r#"
            <w:tblPr>
              <w:tblW w:w="7200" w:type="dxa"/>
              <w:jc w:val="center"/>
              <w:tblStyle w:val="TableGrid"/>
              <w:tblBorders>
                <w:top w:val="single" w:sz="8" w:color="000000"/>
                <w:bottom w:val="single" w:sz="8" w:color="000000"/>
              </w:tblBorders>
              <w:tblCellMar>
                <w:left w:w="120"/>
                <w:right w:w="120"/>
              </w:tblCellMar>
            </w:tblPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableProperties::default();
        seek_start(&mut reader, b"w:tblPr");

        parse_table_properties(&mut reader, &mut props).expect("table properties should parse");
        let width = props.width.expect("width should be parsed");
        assert_eq!(width.value, 7200);
        assert_eq!(width.width_type, TableWidthType::Dxa);
        assert_eq!(props.alignment, Some(TableAlignment::Center));
        assert_eq!(props.style_id.as_deref(), Some("TableGrid"));
        assert!(props.borders.is_some());
        let margins = props.cell_margins.expect("cell margins should be parsed");
        assert_eq!(margins.left, Some(120));
        assert_eq!(margins.right, Some(120));
    }

    #[test]
    fn parse_table_grid_reads_grid_columns_and_ignores_invalid_widths() {
        let xml = r#"
            <w:tblGrid>
              <w:gridCol w:w="1200"/>
              <w:gridCol w:w="bad"/>
              <w:gridCol w:w="2400"/>
            </w:tblGrid>
        "#;
        let mut reader = reader_from(xml);
        seek_start(&mut reader, b"w:tblGrid");

        let grid = parse_table_grid(&mut reader).expect("grid should parse");
        assert_eq!(grid.len(), 2);
        assert_eq!(grid[0].width, 1200);
        assert_eq!(grid[1].width, 2400);
    }

    #[test]
    fn parse_table_row_properties_reads_height_header_and_split_flags() {
        let xml = r#"
            <w:trPr>
              <w:trHeight w:val="300" w:hRule="exact"/>
              <w:tblHeader w:val="1"/>
              <w:cantSplit w:val="false"/>
            </w:trPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableRowProperties::default();
        seek_start(&mut reader, b"w:trPr");

        parse_table_row_properties(&mut reader, &mut props).expect("row properties should parse");
        let height = props.height.expect("row height should parse");
        assert_eq!(height.value, 300);
        assert_eq!(height.rule, RowHeightRule::Exact);
        assert_eq!(props.is_header, Some(true));
        assert_eq!(props.cant_split, Some(false));
    }

    #[test]
    fn parse_table_cell_properties_reads_width_span_merge_alignment_shading_and_borders() {
        let xml = r#"
            <w:tcPr>
              <w:tcW w:w="1500" w:type="pct"/>
              <w:gridSpan w:val="2"/>
              <w:vMerge w:val="restart"/>
              <w:vAlign w:val="bottom"/>
              <w:shd w:fill="FFFF00"/>
              <w:tcBorders>
                <w:left w:val="single" w:sz="8" w:color="FF0000"/>
              </w:tcBorders>
            </w:tcPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableCellProperties::default();
        seek_start(&mut reader, b"w:tcPr");

        parse_table_cell_properties(&mut reader, &mut props).expect("cell properties should parse");
        let width = props.width.expect("cell width should parse");
        assert_eq!(width.value, 1500);
        assert_eq!(width.width_type, TableWidthType::Pct);
        assert_eq!(props.grid_span, Some(2));
        assert_eq!(props.vertical_merge, Some(MergeType::Restart));
        assert_eq!(props.vertical_align, Some(CellVerticalAlignment::Bottom));
        assert_eq!(props.shading.as_deref(), Some("FFFF00"));
        assert!(props.borders.is_some());
    }

    #[test]
    fn parse_table_properties_uses_nil_width_and_left_alignment_fallback() {
        let xml = r#"
            <w:tblPr>
              <w:tblW w:w="1234" w:type="mystery"/>
              <w:jc w:val="unknown"/>
            </w:tblPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableProperties::default();
        seek_start(&mut reader, b"w:tblPr");

        parse_table_properties(&mut reader, &mut props).expect("table properties parse");
        let width = props.width.expect("width");
        assert_eq!(width.value, 1234);
        assert_eq!(width.width_type, TableWidthType::Nil);
        assert_eq!(props.alignment, Some(TableAlignment::Left));
    }

    #[test]
    fn parse_table_row_properties_handles_atleast_and_default_booleans() {
        let xml = r#"
            <w:trPr>
              <w:trHeight w:val="240" w:hRule="atLeast"/>
              <w:tblHeader/>
              <w:cantSplit w:val="0"/>
            </w:trPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableRowProperties::default();
        seek_start(&mut reader, b"w:trPr");

        parse_table_row_properties(&mut reader, &mut props).expect("row props parse");
        let height = props.height.expect("row height");
        assert_eq!(height.value, 240);
        assert_eq!(height.rule, RowHeightRule::AtLeast);
        assert_eq!(props.is_header, Some(true));
        assert_eq!(props.cant_split, Some(false));
    }

    #[test]
    fn parse_table_cell_properties_defaults_merge_and_alignment_and_ignores_invalid_width() {
        let xml = r#"
            <w:tcPr>
              <w:tcW w:w="bad" w:type="dxa"/>
              <w:vMerge/>
              <w:vAlign w:val="unknown"/>
              <w:tcBorders>
                <w:insideH w:val="single" w:sz="4" w:color="00AA00"/>
                <w:insideV w:val="single" w:sz="6" w:color="AA0000"/>
              </w:tcBorders>
            </w:tcPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableCellProperties::default();
        seek_start(&mut reader, b"w:tcPr");

        parse_table_cell_properties(&mut reader, &mut props).expect("cell props parse");
        assert!(props.width.is_none(), "invalid width should be ignored");
        assert_eq!(props.vertical_merge, Some(MergeType::Continue));
        assert_eq!(props.vertical_align, Some(CellVerticalAlignment::Top));
        let borders = props.borders.expect("borders");
        assert!(borders.inside_h.is_some());
        assert!(borders.inside_v.is_some());
    }

    #[test]
    fn parse_table_properties_ignores_unknown_border_and_margin_entries() {
        let xml = r#"
            <w:tblPr>
              <w:tblBorders>
                <w:between w:val="single" w:sz="4" w:color="AA00AA"/>
              </w:tblBorders>
              <w:tblCellMar>
                <w:start w:w="120"/>
              </w:tblCellMar>
            </w:tblPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableProperties::default();
        seek_start(&mut reader, b"w:tblPr");

        parse_table_properties(&mut reader, &mut props).expect("table properties parse");
        assert!(
            props.borders.is_none(),
            "unknown border sides should be ignored"
        );
        assert!(
            props.cell_margins.is_none(),
            "unknown margin sides should not create margin object"
        );
    }

    #[test]
    fn parse_table_cell_properties_parses_nil_border_and_auto_color_fallback() {
        let xml = r#"
            <w:tcPr>
              <w:tcBorders>
                <w:left w:val="nil" w:sz="8" w:color="auto"/>
              </w:tcBorders>
            </w:tcPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableCellProperties::default();
        seek_start(&mut reader, b"w:tcPr");

        parse_table_cell_properties(&mut reader, &mut props).expect("cell props parse");
        let borders = props.borders.expect("cell borders");
        let left = borders.left.expect("left border");
        assert!(matches!(left.style, docir_core::ir::BorderStyle::None));
        assert_eq!(
            left.color, None,
            "auto border color should normalize to None"
        );
        assert_eq!(left.width, Some(8));
    }

    #[test]
    fn parse_table_row_properties_defaults_height_rule_to_auto_when_missing() {
        let xml = r#"
            <w:trPr>
              <w:trHeight w:val="180"/>
            </w:trPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableRowProperties::default();
        seek_start(&mut reader, b"w:trPr");

        parse_table_row_properties(&mut reader, &mut props).expect("row props parse");
        let height = props.height.expect("row height");
        assert_eq!(height.value, 180);
        assert_eq!(height.rule, RowHeightRule::Auto);
    }

    #[test]
    fn parse_table_builds_rows_cells_and_nested_table_content() {
        let xml = r#"
            <w:tbl>
              <w:tblGrid>
                <w:gridCol w:w="1200"/>
              </w:tblGrid>
              <w:tr>
                <w:tc>
                  <w:p><w:r><w:t>outer</w:t></w:r></w:p>
                  <w:tbl>
                    <w:tr>
                      <w:tc>
                        <w:p><w:r><w:t>inner</w:t></w:r></w:p>
                      </w:tc>
                    </w:tr>
                  </w:tbl>
                </w:tc>
              </w:tr>
            </w:tbl>
        "#;
        let mut reader = reader_from(xml);
        let rels = Relationships::default();
        let mut parser = DocxParser::new();
        seek_start(&mut reader, b"w:tbl");
        let table_id = parse_table(&mut parser, &mut reader, &rels).expect("parse table");
        let store = parser.into_store();
        let table = match store.get(table_id) {
            Some(docir_core::ir::IRNode::Table(table)) => table,
            _ => panic!("expected table"),
        };
        assert_eq!(table.grid.len(), 1);
        assert_eq!(table.rows.len(), 1);
        let row = match store.get(table.rows[0]) {
            Some(docir_core::ir::IRNode::TableRow(row)) => row,
            _ => panic!("expected row"),
        };
        assert_eq!(row.cells.len(), 1);
        let cell = match store.get(row.cells[0]) {
            Some(docir_core::ir::IRNode::TableCell(cell)) => cell,
            _ => panic!("expected table cell"),
        };
        assert_eq!(
            cell.content.len(),
            2,
            "cell should include paragraph and nested table"
        );
        assert!(matches!(
            store.get(cell.content[1]),
            Some(docir_core::ir::IRNode::Table(_))
        ));
    }

    #[test]
    fn parse_table_ignores_unknown_child_elements() {
        let xml = r#"
            <w:tbl>
              <w:unknown/>
              <w:tblGrid>
                <w:gridCol w:w="1200"/>
              </w:tblGrid>
              <w:tr>
                <w:tc>
                  <w:p><w:r><w:t>value</w:t></w:r></w:p>
                </w:tc>
              </w:tr>
            </w:tbl>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = reader_from(xml);

        seek_start(&mut reader, b"w:tbl");
        let table_id = parse_table(&mut parser, &mut reader, &rels).expect("parse table");
        let store = parser.into_store();
        let table = match store.get(table_id) {
            Some(docir_core::ir::IRNode::Table(table)) => table,
            _ => panic!("expected table"),
        };
        assert_eq!(table.rows.len(), 1);
    }

    #[test]
    fn parse_table_row_ignores_unknown_row_children() {
        let xml = r#"
            <w:tr>
              <w:unknown/>
            </w:tr>
        "#;
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = reader_from(xml);
        let row = parse_table_row(&mut parser, &mut reader, &rels).expect("parse row");
        let store = parser.into_store();
        let row = match store.get(row) {
            Some(docir_core::ir::IRNode::TableRow(row)) => row,
            _ => panic!("expected table row"),
        };
        assert_eq!(row.cells.len(), 0);
    }

    #[test]
    fn parse_table_cell_properties_ignores_unknown_properties() {
        let xml = r#"
            <w:tcPr>
              <w:unknown>
                <w:placeholder/>
              </w:unknown>
            </w:tcPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableCellProperties::default();
        seek_start(&mut reader, b"w:tcPr");
        parse_table_cell_properties(&mut reader, &mut props).expect("table cell props");
        assert!(props.width.is_none());
        assert!(props.vertical_merge.is_none());
        assert!(props.borders.is_none());
        assert!(props.shading.is_none());
    }

    #[test]
    fn parse_table_row_properties_ignores_unknown_items() {
        let xml = r#"
            <w:trPr>
              <w:unknown/>
            </w:trPr>
        "#;
        let mut reader = reader_from(xml);
        let mut props = TableRowProperties::default();
        seek_start(&mut reader, b"w:trPr");
        parse_table_row_properties(&mut reader, &mut props).expect("table row props");
        assert!(props.height.is_none());
        assert_eq!(props.is_header, None);
        assert_eq!(props.cant_split, None);
    }

    #[test]
    fn parse_table_returns_xml_error_for_malformed_table() {
        let xml = r#"
            <w:tbl>
              <w:tr><w:tc><w:p><w:r><w:t>bad</w:r></w:p></w:tc></w:tr>
            </w:tbl>
        "#;
        let mut reader = reader_from(xml);
        let rels = Relationships::default();
        let mut parser = DocxParser::new();
        seek_start(&mut reader, b"w:tbl");
        let err =
            parse_table(&mut parser, &mut reader, &rels).expect_err("malformed table should error");
        match err {
            ParseError::Xml { file, .. } => {
                assert_eq!(file, "word/document.xml");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
