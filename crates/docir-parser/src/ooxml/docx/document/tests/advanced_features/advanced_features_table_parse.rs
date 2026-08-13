use super::parse_single_table;

#[test]
fn test_parse_table_grid_and_properties() {
    let xml = r#"
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tblPr>
            <w:tblW w:w="5000" w:type="dxa"/>
            <w:jc w:val="center"/>
            <w:tblStyle w:val="TableStyle1"/>
            <w:tblBorders>
              <w:top w:val="single" w:sz="8" w:color="FF0000"/>
            </w:tblBorders>
            <w:tblCellMar>
              <w:top w:w="100"/>
              <w:left w:w="120"/>
            </w:tblCellMar>
          </w:tblPr>
          <w:tblGrid>
            <w:gridCol w:w="2400"/>
            <w:gridCol w:w="2600"/>
          </w:tblGrid>
          <w:tr>
            <w:trPr>
              <w:trHeight w:val="300" w:hRule="exact"/>
              <w:tblHeader/>
              <w:cantSplit w:val="1"/>
            </w:trPr>
            <w:tc>
              <w:tcPr>
                <w:shd w:fill="FFFF00"/>
              </w:tcPr>
              <w:p><w:r><w:t>A</w:t></w:r></w:p>
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
    assert_eq!(table.grid.len(), 2);
    assert_eq!(table.grid[0].width, 2400);
    assert_eq!(table.grid[1].width, 2600);
    let props = &table.properties;
    assert_eq!(props.width.as_ref().map(|w| w.value), Some(5000));
    assert!(matches!(
        props.alignment,
        Some(docir_core::ir::TableAlignment::Center)
    ));
    assert_eq!(props.style_id.as_deref(), Some("TableStyle1"));
    assert_eq!(props.cell_margins.as_ref().and_then(|m| m.top), Some(100));
    assert_eq!(props.cell_margins.as_ref().and_then(|m| m.left), Some(120));
    assert!(
        props
            .borders
            .as_ref()
            .and_then(|b| b.top.as_ref())
            .is_some()
    );

    let row = match store.get(table.rows[0]) {
        Some(docir_core::ir::IRNode::TableRow(r)) => r,
        _ => panic!("missing row"),
    };
    assert_eq!(row.properties.height.as_ref().map(|h| h.value), Some(300));
    assert!(matches!(
        row.properties.height.as_ref().map(|h| h.rule),
        Some(docir_core::ir::RowHeightRule::Exact)
    ));
    assert_eq!(row.properties.is_header, Some(true));
    assert_eq!(row.properties.cant_split, Some(true));

    let cell = match store.get(row.cells[0]) {
        Some(docir_core::ir::IRNode::TableCell(c)) => c,
        _ => panic!("missing cell"),
    };
    assert_eq!(cell.properties.shading.as_deref(), Some("FFFF00"));
}

#[test]
fn test_parse_table_cell_and_row_property_variants() {
    let xml = r#"
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tblPr>
            <w:tblW w:w="7500" w:type="pct"/>
            <w:jc w:val="right"/>
            <w:tblBorders>
              <w:insideH w:val="single" w:sz="4" w:color="00AA00"/>
              <w:insideV w:val="single" w:sz="6" w:color="AA0000"/>
            </w:tblBorders>
            <w:tblCellMar>
              <w:top w:w="20"/>
              <w:bottom w:w="30"/>
              <w:left w:w="40"/>
              <w:right w:w="50"/>
            </w:tblCellMar>
          </w:tblPr>
          <w:tblGrid>
            <w:gridCol w:w="2400"/>
            <w:gridCol w:w="2600"/>
          </w:tblGrid>
          <w:tr>
            <w:trPr>
              <w:trHeight w:val="360" w:hRule="atLeast"/>
              <w:tblHeader w:val="0"/>
              <w:cantSplit w:val="false"/>
            </w:trPr>
            <w:tc>
              <w:tcPr>
                <w:tcW w:w="2400" w:type="pct"/>
                <w:gridSpan w:val="2"/>
                <w:vMerge w:val="restart"/>
                <w:vAlign w:val="center"/>
                <w:tcBorders>
                  <w:insideH w:val="single" w:sz="2" w:color="00FF00"/>
                  <w:insideV w:val="single" w:sz="2" w:color="0000FF"/>
                </w:tcBorders>
                <w:shd w:fill="CCCCCC"/>
              </w:tcPr>
              <w:p/>
              <w:tbl>
                <w:tblPr>
                  <w:tblW w:w="1200" w:type="dxa"/>
                </w:tblPr>
                <w:tr>
                  <w:tc>
                    <w:p><w:r><w:t>Nested</w:t></w:r></w:p>
                  </w:tc>
                </w:tr>
              </w:tbl>
            </w:tc>
            <w:tc>
              <w:tcPr>
                <w:tcW w:w="100" w:type="bogus"/>
                <w:vMerge/>
                <w:vAlign w:val="bottom"/>
              </w:tcPr>
              <w:p><w:r><w:t>B</w:t></w:r></w:p>
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

    assert_eq!(table.properties.width.as_ref().map(|w| w.value), Some(7500));
    assert!(matches!(
        table.properties.width.as_ref().map(|w| w.width_type),
        Some(docir_core::ir::TableWidthType::Pct)
    ));
    assert!(matches!(
        table.properties.alignment,
        Some(docir_core::ir::TableAlignment::Right)
    ));
    assert!(
        table
            .properties
            .borders
            .as_ref()
            .and_then(|b| b.inside_h.as_ref())
            .is_some()
    );
    assert!(
        table
            .properties
            .borders
            .as_ref()
            .and_then(|b| b.inside_v.as_ref())
            .is_some()
    );
    assert_eq!(
        table.properties.cell_margins.as_ref().and_then(|m| m.top),
        Some(20)
    );
    assert_eq!(
        table
            .properties
            .cell_margins
            .as_ref()
            .and_then(|m| m.bottom),
        Some(30)
    );
    assert_eq!(
        table.properties.cell_margins.as_ref().and_then(|m| m.left),
        Some(40)
    );
    assert_eq!(
        table.properties.cell_margins.as_ref().and_then(|m| m.right),
        Some(50)
    );

    let row = match store.get(table.rows[0]) {
        Some(docir_core::ir::IRNode::TableRow(r)) => r,
        _ => panic!("missing row"),
    };
    assert!(matches!(
        row.properties.height.as_ref().map(|h| h.rule),
        Some(docir_core::ir::RowHeightRule::AtLeast)
    ));
    assert_eq!(row.properties.is_header, Some(false));
    assert_eq!(row.properties.cant_split, Some(false));

    let cell_a = match store.get(row.cells[0]) {
        Some(docir_core::ir::IRNode::TableCell(c)) => c,
        _ => panic!("missing first cell"),
    };
    assert_eq!(cell_a.properties.grid_span, Some(2));
    assert!(matches!(
        cell_a.properties.vertical_merge,
        Some(docir_core::ir::MergeType::Restart)
    ));
    assert!(matches!(
        cell_a.properties.vertical_align,
        Some(docir_core::ir::CellVerticalAlignment::Center)
    ));
    assert!(
        cell_a
            .properties
            .borders
            .as_ref()
            .and_then(|b| b.inside_h.as_ref())
            .is_some()
    );
    assert!(
        cell_a
            .properties
            .borders
            .as_ref()
            .and_then(|b| b.inside_v.as_ref())
            .is_some()
    );
    assert_eq!(cell_a.properties.shading.as_deref(), Some("CCCCCC"));
    assert!(
        cell_a
            .content
            .iter()
            .any(|id| matches!(store.get(*id), Some(docir_core::ir::IRNode::Table(_))))
    );
    assert!(
        cell_a
            .content
            .iter()
            .any(|id| matches!(store.get(*id), Some(docir_core::ir::IRNode::Paragraph(_))))
    );

    let cell_b = match store.get(row.cells[1]) {
        Some(docir_core::ir::IRNode::TableCell(c)) => c,
        _ => panic!("missing second cell"),
    };
    assert!(matches!(
        cell_b.properties.width.as_ref().map(|w| w.width_type),
        Some(docir_core::ir::TableWidthType::Nil)
    ));
    assert!(matches!(
        cell_b.properties.vertical_merge,
        Some(docir_core::ir::MergeType::Continue)
    ));
    assert!(matches!(
        cell_b.properties.vertical_align,
        Some(docir_core::ir::CellVerticalAlignment::Bottom)
    ));
}
