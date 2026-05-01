use super::{build_empty_zip, Relationships, SheetKind, SheetState};
use crate::ooxml::xlsx::workbook::SheetInfo;
use crate::ooxml::xlsx::{
    parse_calc_chain, parse_pivot_cache_records, parse_pivot_table_definition,
    parse_sheet_comments, parse_styles, parse_table_definition, XlsxParser,
};
use docir_core::ir::{Cell, IRNode};
use docir_core::types::NodeId;
use docir_core::visitor::IrStore;
use docir_core::SpreadsheetStyles;
#[test]
fn test_parse_dialogsheet_kind() {
    let dialog_xml = r#"
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1" t="str"><v>Hello</v></c>
            </row>
          </sheetData>
        </worksheet>
        "#;
    let mut parser = XlsxParser::new();
    let mut zip = build_empty_zip();
    let sheet = SheetInfo {
        name: "Dialog1".to_string(),
        sheet_id: 1,
        rel_id: "rId1".to_string(),
        state: SheetState::Visible,
    };
    let ws_id = parser
        .parse_worksheet(
            &mut zip,
            dialog_xml,
            &sheet,
            "xl/dialogsheets/sheet1.xml",
            &Relationships::default(),
            SheetKind::DialogSheet,
        )
        .expect("dialogsheet");
    let store = parser.into_store();
    let ws = match store.get(ws_id) {
        Some(IRNode::Worksheet(w)) => w,
        _ => panic!("missing worksheet"),
    };
    assert_eq!(ws.kind, SheetKind::DialogSheet);
    assert_eq!(ws.cells.len(), 1);
}

#[test]
fn test_parse_styles_minimal() {
    let xml = r#"
        <styleSheet>
          <numFmts count="1">
            <numFmt numFmtId="164" formatCode="0.00"/>
          </numFmts>
          <fonts count="1">
            <font>
              <sz val="11"/>
              <name val="Calibri"/>
              <b/>
              <color rgb="FF0000"/>
            </font>
          </fonts>
          <fills count="1">
            <fill>
              <patternFill patternType="solid">
                <fgColor rgb="FFFF00"/>
              </patternFill>
            </fill>
          </fills>
          <borders count="1">
            <border>
              <left style="thin"><color rgb="FF00FF"/></left>
            </border>
          </borders>
          <cellXfs count="1">
            <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"
                applyNumberFormat="1" applyAlignment="1" applyProtection="1" quotePrefix="1">
              <alignment horizontal="center" wrapText="1" indent="2" textRotation="45"
                         shrinkToFit="1" readingOrder="1"/>
              <protection locked="1" hidden="0"/>
            </xf>
          </cellXfs>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>
          </cellStyleXfs>
          <dxfs count="1">
            <dxf>
              <numFmt numFmtId="200" formatCode="0.00"/>
              <font><b/><color rgb="FF0000"/></font>
              <fill><patternFill patternType="solid"><fgColor rgb="00FF00"/></patternFill></fill>
            </dxf>
          </dxfs>
          <tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">
            <tableStyle name="TableStyleMedium2" pivot="0" table="1"/>
          </tableStyles>
        </styleSheet>
        "#;

    let styles = parse_styles(xml, "xl/styles.xml").expect("styles");
    assert_style_counts(&styles);
    assert_style_content(&styles);
}

fn assert_style_counts(styles: &SpreadsheetStyles) {
    assert_eq!(styles.number_formats.len(), 1);
    assert_eq!(styles.fonts.len(), 1);
    assert_eq!(styles.fills.len(), 1);
    assert_eq!(styles.borders.len(), 1);
    assert_eq!(styles.cell_xfs.len(), 1);
    assert_eq!(styles.cell_style_xfs.len(), 1);
    assert_eq!(styles.dxfs.len(), 1);
    assert!(styles.table_styles.is_some());
}

fn assert_style_content(styles: &SpreadsheetStyles) {
    assert_eq!(styles.fonts[0].name.as_deref(), Some("Calibri"));
    assert!(styles.fonts[0].bold);
    assert!(styles.cell_xfs[0].apply_number_format);
    assert_eq!(styles.dxfs[0].num_fmt.as_ref().map(|n| n.id), Some(200));
    assert_eq!(
        styles
            .table_styles
            .as_ref()
            .unwrap()
            .default_table_style
            .as_deref(),
        Some("TableStyleMedium2")
    );
    assert_eq!(styles.table_styles.as_ref().unwrap().styles.len(), 1);
    assert_eq!(
        styles.table_styles.as_ref().unwrap().styles[0].name,
        "TableStyleMedium2"
    );
    assert_eq!(
        styles.cell_xfs[0]
            .alignment
            .as_ref()
            .and_then(|a| a.horizontal.as_deref()),
        Some("center")
    );
    assert_eq!(
        styles.cell_xfs[0].alignment.as_ref().and_then(|a| a.indent),
        Some(2)
    );
    assert_eq!(
        styles.cell_xfs[0]
            .alignment
            .as_ref()
            .and_then(|a| a.text_rotation),
        Some(45)
    );
    assert!(styles.cell_xfs[0]
        .alignment
        .as_ref()
        .map(|a| a.shrink_to_fit)
        .unwrap_or(false));
    assert_eq!(
        styles.cell_xfs[0]
            .protection
            .as_ref()
            .and_then(|p| p.locked),
        Some(true)
    );
}

pub(crate) fn get_cell<'a>(store: &'a IrStore, id: NodeId, label: &str) -> &'a Cell {
    store
        .get(id)
        .and_then(|n| match n {
            IRNode::Cell(c) => Some(c),
            _ => None,
        })
        .unwrap_or_else(|| panic!("{label}"))
}

#[test]
fn test_parse_calc_chain() {
    let xml = r#"
        <calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <c r="A1" i="0" l="1" s="1"/>
          <c r="B2" i="2"/>
        </calcChain>
        "#;
    let chain = parse_calc_chain(xml, "xl/calcChain.xml").expect("calc chain");
    assert_eq!(chain.entries.len(), 2);
    assert_eq!(chain.entries[0].cell_ref, "A1");
    assert_eq!(chain.entries[0].level, Some(1));
    assert_eq!(chain.entries[0].new_value, Some(true));
    assert_eq!(chain.entries[1].cell_ref, "B2");
    assert_eq!(chain.entries[1].index, Some(2));
}

#[test]
fn test_parse_sheet_comments() {
    let xml = r#"
        <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <authors>
            <author>Alice</author>
            <author>Bob</author>
          </authors>
          <commentList>
            <comment ref="A1" authorId="0">
              <text><r><t>Hello</t></r></text>
            </comment>
            <comment ref="B2" authorId="1">
              <text><t>World</t></text>
            </comment>
          </commentList>
        </comments>
        "#;
    let comments = parse_sheet_comments(xml, "xl/comments1.xml", Some("Sheet1")).expect("comments");
    assert_eq!(comments.len(), 2);
    assert_eq!(comments[0].cell_ref, "A1");
    assert_eq!(comments[0].author.as_deref(), Some("Alice"));
    assert_eq!(comments[0].text, "Hello");
    assert_eq!(comments[1].cell_ref, "B2");
    assert_eq!(comments[1].author.as_deref(), Some("Bob"));
    assert_eq!(comments[1].text, "World");
}

#[test]
fn test_parse_conditional_and_validation() {
    let xml = r#"
        <worksheet>
          <sheetData/>
          <conditionalFormatting sqref="A1:A10">
            <cfRule type="expression" priority="1">
              <formula>SUM(A1)&gt;10</formula>
            </cfRule>
          </conditionalFormatting>
          <dataValidations count="1">
            <dataValidation type="list" allowBlank="1" sqref="B1">
              <formula1>"Yes,No"</formula1>
            </dataValidation>
          </dataValidations>
        </worksheet>
        "#;

    let mut parser = XlsxParser::new();
    let sheet = SheetInfo {
        name: "Sheet1".to_string(),
        sheet_id: 1,
        rel_id: "rId1".to_string(),
        state: SheetState::Visible,
    };
    let mut zip = build_empty_zip();
    let ws_id = parser
        .parse_worksheet(
            &mut zip,
            xml,
            &sheet,
            "xl/worksheets/sheet1.xml",
            &Relationships::default(),
            SheetKind::Worksheet,
        )
        .expect("worksheet");
    let store = parser.into_store();

    let worksheet = match store.get(ws_id) {
        Some(IRNode::Worksheet(ws)) => ws,
        _ => panic!("expected worksheet"),
    };
    assert_eq!(worksheet.conditional_formats.len(), 1);
    assert_eq!(worksheet.data_validations.len(), 1);
}

#[test]
fn test_parse_table_and_pivot_definition() {
    let table_xml = r#"
        <table name="Table1" displayName="Table1" ref="A1:B3" headerRowCount="1">
          <tableColumns count="2">
            <tableColumn id="1" name="Col1"/>
            <tableColumn id="2" name="Col2" totalsRowFunction="sum"/>
          </tableColumns>
        </table>
        "#;
    let table = parse_table_definition(table_xml, "xl/tables/table1.xml").expect("table");
    assert_eq!(table.columns.len(), 2);
    assert_eq!(table.ref_range.as_deref(), Some("A1:B3"));

    let pivot_xml = r#"
        <pivotTableDefinition name="Pivot1" cacheId="3">
          <location ref="D1:F10"/>
        </pivotTableDefinition>
        "#;
    let pivot =
        parse_pivot_table_definition(pivot_xml, "xl/pivotTables/pivotTable1.xml").expect("pivot");
    assert_eq!(pivot.cache_id, Some(3));
    assert_eq!(pivot.ref_range.as_deref(), Some("D1:F10"));
}

#[test]
fn test_parse_pivot_cache_records() {
    let records_xml = r#"
        <pivotCacheRecords count="2">
          <r><n v="1"/><s v="0"/><b v="1"/></r>
          <r><n v="2"/><s v="1"/><b v="0"/></r>
        </pivotCacheRecords>
        "#;

    let records = parse_pivot_cache_records(records_xml, "xl/pivotCache/pivotCacheRecords1.xml")
        .expect("records");
    assert_eq!(records.record_count, Some(2));
    assert_eq!(records.field_count, Some(3));
}
