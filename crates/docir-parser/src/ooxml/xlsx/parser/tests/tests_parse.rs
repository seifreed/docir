use super::get_cell;
use super::{
    build_empty_zip, build_zip_with_entries, ParseError, Relationships, SheetKind, SheetState,
};
use crate::ooxml::xlsx::relationships::classify_relationship;
use crate::ooxml::xlsx::workbook::{
    auto_open_target_from_defined_name, parse_workbook_info, SheetInfo,
};
use crate::ooxml::xlsx::{ShapeType, XlsxParser};
use docir_core::ir::IRNode;
use docir_core::CellError;
use docir_core::CellValue;
struct WorksheetDrawingFixture {
    sheet_xml: &'static str,
    drawing_xml: &'static str,
    chart_xml: &'static str,
    drawing_rels: &'static str,
    sheet_rels: &'static str,
}

fn worksheet_drawing_fixture() -> WorksheetDrawingFixture {
    WorksheetDrawingFixture {
        sheet_xml: r#"
        <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetData/>
        </worksheet>
        "#,
        drawing_xml: r#"
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                 xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <xdr:twoCellAnchor>
            <xdr:pic>
              <xdr:nvPicPr>
                <xdr:cNvPr id="1" name="Picture 1" descr="Alt text"/>
              </xdr:nvPicPr>
              <xdr:blipFill>
                <a:blip r:embed="rIdImg"/>
              </xdr:blipFill>
            </xdr:pic>
          </xdr:twoCellAnchor>
          <xdr:graphicFrame>
            <xdr:nvGraphicFramePr>
              <xdr:cNvPr id="2" name="Chart 1"/>
            </xdr:nvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rIdChart"/>
              </a:graphicData>
            </a:graphic>
          </xdr:graphicFrame>
        </xdr:wsDr>
        "#,
        chart_xml: r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:barChart>
              <c:ser><c:tx><c:v>2019</c:v></c:tx></c:ser>
              <c:ser><c:tx><c:v>2020</c:v></c:tx></c:ser>
            </c:barChart>
          </c:chart>
        </c:chartSpace>
        "#,
        drawing_rels: r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image1.png"/>
          <Relationship Id="rIdChart"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="../charts/chart1.xml"/>
        </Relationships>
        "#,
        sheet_rels: r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdDraw"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
            Target="../drawings/drawing1.xml"/>
        </Relationships>
        "#,
    }
}

#[test]
fn test_parse_workbook_info_sheets() {
    let xml = r#"
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <bookViews>
            <workbookView activeTab="1" firstSheet="0" showHorizontalScroll="1"
                          showVerticalScroll="0" showSheetTabs="1" tabRatio="400"
                          windowWidth="12000" windowHeight="8000" xWindow="120" yWindow="240"/>
          </bookViews>
          <sheets>
            <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
            <sheet name="Hidden" sheetId="2" r:id="rId2" state="hidden"/>
            <sheet name="VeryHidden" sheetId="3" r:id="rId3" state="veryHidden"/>
          </sheets>
        </workbook>
        "#;

    let info = parse_workbook_info(xml).expect("parse workbook info");
    assert_eq!(info.sheets.len(), 3);
    assert_eq!(info.sheets[0].name, "Sheet1");
    assert_eq!(info.sheets[1].state, SheetState::Hidden);
    assert_eq!(info.sheets[2].state, SheetState::VeryHidden);
    let props = info.workbook_properties.expect("workbook props");
    assert_eq!(props.active_tab, Some(1));
    assert_eq!(props.first_sheet, Some(0));
    assert_eq!(props.show_horizontal_scroll, Some(true));
    assert_eq!(props.show_vertical_scroll, Some(false));
    assert_eq!(props.show_sheet_tabs, Some(true));
    assert_eq!(props.tab_ratio, Some(400));
    assert_eq!(props.window_width, Some(12000));
    assert_eq!(props.window_height, Some(8000));
    assert_eq!(props.x_window, Some(120));
    assert_eq!(props.y_window, Some(240));
}

#[test]
fn test_parse_workbook_info_defined_names_and_props() {
    let xml = r#"
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <workbookPr date1904="true"/>
          <bookViews>
            <workbookView activeTab="2" firstSheet="1"/>
          </bookViews>
          <calcPr calcMode="manual" fullCalcOnLoad="1" calcOnSave="0"/>
          <workbookProtection lockStructure="1"/>
          <definedNames>
            <definedName name="AUTO_OPEN" localSheetId="0" hidden="1" comment="entry">'Macro Sheet'!$A$1</definedName>
            <definedName name="EmptyName" localSheetId="1" hidden="true" comment="empty"/>
          </definedNames>
          <pivotCaches>
            <pivotCache cacheId="4" r:id="rIdPivot"/>
            <pivotCache cacheId="bad" r:id="rIdIgnored"/>
          </pivotCaches>
          <sheets>
            <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
          </sheets>
        </workbook>
        "#;

    let info = parse_workbook_info(xml).expect("parse workbook info");
    assert_eq!(info.sheets.len(), 1);
    assert_eq!(info.defined_names.len(), 2);
    assert_eq!(info.pivot_cache_refs.len(), 1);
    assert_eq!(info.pivot_cache_refs[0].cache_id, 4);
    assert_eq!(info.pivot_cache_refs[0].rel_id, "rIdPivot");

    let auto_open = &info.defined_names[0];
    assert_eq!(auto_open.name, "AUTO_OPEN");
    assert_eq!(auto_open.local_sheet_id, Some(0));
    assert!(auto_open.hidden);
    assert_eq!(auto_open.comment.as_deref(), Some("entry"));
    assert_eq!(
        auto_open_target_from_defined_name(auto_open),
        Some(Some("Macro Sheet".to_string()))
    );

    let empty_name = &info.defined_names[1];
    assert_eq!(empty_name.value, "");
    assert_eq!(empty_name.local_sheet_id, Some(1));
    assert!(empty_name.hidden);
    assert_eq!(
        auto_open_target_from_defined_name(empty_name),
        None,
        "non AUTO_OPEN names must not be treated as macro triggers"
    );

    let props = info.workbook_properties.expect("workbook props");
    assert_eq!(props.date1904, Some(true));
    assert_eq!(props.active_tab, Some(2));
    assert_eq!(props.first_sheet, Some(1));
    assert_eq!(props.calc_mode.as_deref(), Some("manual"));
    assert_eq!(props.calc_full, Some(true));
    assert_eq!(props.calc_on_save, Some(false));
    assert!(props.workbook_protected);
}

#[test]
fn test_process_external_relationships_classification_and_dedup() {
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdHyper"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            Target="https://example.com"
            TargetMode="External"/>
          <Relationship Id="rIdImage"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="https://example.com/image.png"
            TargetMode="External"/>
          <Relationship Id="rIdOle"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
            Target="https://example.com/object.bin"
            TargetMode="External"/>
          <Relationship Id="rIdExternalData"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
            Target="https://example.com/data"
            TargetMode="External"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");

    let mut parser = XlsxParser::new();
    parser.process_external_relationships(&rels, "xl/workbook.xml");
    parser.process_external_relationships(&rels, "xl/workbook.xml");

    assert_eq!(
        parser.security_info.external_refs.len(),
        4,
        "same file/id pair should be deduplicated"
    );
    assert_eq!(
        classify_relationship(".../hyperlink"),
        docir_core::security::ExternalRefType::Hyperlink
    );
    assert_eq!(
        classify_relationship(".../image"),
        docir_core::security::ExternalRefType::Image
    );
    assert_eq!(
        classify_relationship(".../slideMaster"),
        docir_core::security::ExternalRefType::SlideMaster
    );
    assert_eq!(
        classify_relationship(".../oleObject"),
        docir_core::security::ExternalRefType::OleLink
    );
    assert_eq!(
        classify_relationship(".../externalLink"),
        docir_core::security::ExternalRefType::DataConnection
    );
    assert_eq!(
        classify_relationship(".../other"),
        docir_core::security::ExternalRefType::Other
    );
}

#[test]
fn test_parse_worksheet_cells_additional_value_branches() {
    let xml = r#"
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1" t="str"><v>Plain text</v></c>
              <c r="B1" t="d"><v>45123.5</v></c>
              <c r="C1" t="d"><v>not-a-number</v></c>
              <c r="D1" t="s"><v>99</v></c>
              <c r="E1"/>
            </row>
          </sheetData>
        </worksheet>
        "#;

    let mut parser = XlsxParser::new();
    parser.shared_strings = vec!["only-one".to_string()];

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
        .expect("parse worksheet");
    let store = parser.into_store();
    let worksheet = match store.get(ws_id) {
        Some(IRNode::Worksheet(ws)) => ws,
        _ => panic!("missing worksheet node"),
    };

    let c_a1 = get_cell(&store, worksheet.cells[0], "cell a1");
    assert!(matches!(c_a1.value, CellValue::String(ref v) if v == "Plain text"));

    let c_b1 = get_cell(&store, worksheet.cells[1], "cell b1");
    assert!(matches!(c_b1.value, CellValue::DateTime(v) if (v - 45123.5).abs() < f64::EPSILON));

    let c_c1 = get_cell(&store, worksheet.cells[2], "cell c1");
    assert!(matches!(c_c1.value, CellValue::String(ref v) if v == "not-a-number"));

    let c_d1 = get_cell(&store, worksheet.cells[3], "cell d1");
    assert!(matches!(c_d1.value, CellValue::SharedString(99)));

    let c_e1 = get_cell(&store, worksheet.cells[4], "cell e1");
    assert!(matches!(c_e1.value, CellValue::Empty));
}

#[test]
fn test_parse_worksheet_cells() {
    let xml = r#"
        <worksheet>
          <cols>
            <col min="1" max="2" width="10" customWidth="1"/>
          </cols>
          <sheetData>
            <row r="1">
              <c r="A1" t="s"><v>0</v></c>
              <c r="B1"><v>42</v></c>
              <c r="C1" t="b"><v>1</v></c>
              <c r="D1" t="e"><v>#REF!</v></c>
              <c r="E1"><f>SUM(A1:B1)</f><v>3</v></c>
              <c r="F1"><is><t>Inline</t></is></c>
            </row>
          </sheetData>
          <mergeCells>
            <mergeCell ref="A1:B1"/>
          </mergeCells>
        </worksheet>
        "#;

    let mut parser = XlsxParser::new();
    parser.shared_strings = vec!["Hello".to_string()];

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
        .expect("parse worksheet");
    let store = parser.into_store();

    let worksheet = match store.get(ws_id) {
        Some(IRNode::Worksheet(ws)) => ws,
        _ => panic!("missing worksheet node"),
    };

    assert_eq!(worksheet.columns.len(), 2);
    assert_eq!(worksheet.merged_cells.len(), 1);
    assert_eq!(worksheet.cells.len(), 6);

    let cell_a1 = get_cell(&store, worksheet.cells[0], "cell a1");
    assert!(matches!(cell_a1.value, CellValue::String(ref v) if v == "Hello"));

    let cell_b1 = get_cell(&store, worksheet.cells[1], "cell b1");
    assert!(matches!(cell_b1.value, CellValue::Number(v) if (v - 42.0).abs() < f64::EPSILON));

    let cell_c1 = get_cell(&store, worksheet.cells[2], "cell c1");
    assert!(matches!(cell_c1.value, CellValue::Boolean(true)));

    let cell_d1 = get_cell(&store, worksheet.cells[3], "cell d1");
    assert!(matches!(cell_d1.value, CellValue::Error(CellError::Ref)));

    let cell_e1 = get_cell(&store, worksheet.cells[4], "cell e1");
    assert!(cell_e1.formula.is_some());

    let cell_f1 = get_cell(&store, worksheet.cells[5], "cell f1");
    assert!(matches!(cell_f1.value, CellValue::InlineString(ref v) if v == "Inline"));
}

#[test]
fn test_parse_worksheet_properties() {
    let xml = r#"
        <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <dimension ref="A1:C5"/>
          <sheetPr>
            <tabColor rgb="FF00FF"/>
          </sheetPr>
          <pageMargins left="0.5" right="0.6" top="0.7" bottom="0.8" header="0.3" footer="0.4"/>
          <sheetData/>
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
    assert_eq!(worksheet.dimension.as_deref(), Some("A1:C5"));
    assert_eq!(worksheet.tab_color.as_deref(), Some("rgb:FF00FF"));
    let margins = worksheet.page_margins.as_ref().expect("margins");
    assert_eq!(margins.left, Some(0.5));
    assert_eq!(margins.right, Some(0.6));
    assert_eq!(margins.top, Some(0.7));
    assert_eq!(margins.bottom, Some(0.8));
    assert_eq!(margins.header, Some(0.3));
    assert_eq!(margins.footer, Some(0.4));
}

#[test]
fn test_parse_worksheet_drawing_pic_and_chart() {
    let fixture = worksheet_drawing_fixture();

    let mut zip = build_zip_with_entries(vec![
        ("xl/drawings/drawing1.xml", fixture.drawing_xml),
        ("xl/drawings/_rels/drawing1.xml.rels", fixture.drawing_rels),
        ("xl/charts/chart1.xml", fixture.chart_xml),
    ]);

    let mut parser = XlsxParser::new();
    let sheet = SheetInfo {
        name: "Sheet1".to_string(),
        sheet_id: 1,
        rel_id: "rId1".to_string(),
        state: SheetState::Visible,
    };
    let rels = Relationships::parse(fixture.sheet_rels).expect("sheet rels");

    let ws_id = parser
        .parse_worksheet(
            &mut zip,
            fixture.sheet_xml,
            &sheet,
            "xl/worksheets/sheet1.xml",
            &rels,
            SheetKind::Worksheet,
        )
        .expect("parse worksheet");
    let store = parser.into_store();
    let ws = match store.get(ws_id) {
        Some(IRNode::Worksheet(w)) => w,
        _ => panic!("missing worksheet"),
    };
    assert_eq!(ws.drawings.len(), 1);
    let drawing = match store.get(ws.drawings[0]) {
        Some(IRNode::WorksheetDrawing(d)) => d,
        _ => panic!("missing drawing"),
    };
    assert_eq!(drawing.shapes.len(), 2);
}

#[test]
fn test_parse_drawing_tracks_external_image_reference() {
    let drawing_xml = r#"
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <xdr:twoCellAnchor>
            <xdr:pic>
              <xdr:nvPicPr>
                <xdr:cNvPr id="1" name="External Picture" descr="External image"></xdr:cNvPr>
              </xdr:nvPicPr>
              <xdr:blipFill>
                <a:blip r:embed="rIdExtImg"></a:blip>
              </xdr:blipFill>
            </xdr:pic>
          </xdr:twoCellAnchor>
        </xdr:wsDr>
        "#;
    let drawing_rels = Relationships::parse(
        r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdExtImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="https://example.com/image.png"
            TargetMode="External"/>
        </Relationships>
        "#,
    )
    .expect("drawing rels");

    let mut parser = XlsxParser::new();
    let mut zip = build_empty_zip();
    let drawing_id = parser
        .parse_drawing(
            drawing_xml,
            "xl/drawings/drawing1.xml",
            &drawing_rels,
            &mut zip,
        )
        .expect("parse drawing");

    assert_eq!(parser.security_info.external_refs.len(), 1);

    let store = parser.into_store();
    let drawing = match store.get(drawing_id) {
        Some(IRNode::WorksheetDrawing(d)) => d,
        _ => panic!("missing drawing"),
    };
    assert_eq!(drawing.shapes.len(), 1);
    let shape = match store.get(drawing.shapes[0]) {
        Some(IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(
        shape.media_target.as_deref(),
        Some("xl/drawings/https:/example.com/image.png")
    );
    assert_eq!(shape.alt_text.as_deref(), Some("External image"));
}

#[test]
fn test_parse_drawing_returns_xml_error_on_malformed_xml() {
    let drawing_xml = r#"
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <xdr:pic>
        </xdr:wsDr>
        "#;
    let mut parser = XlsxParser::new();
    let mut zip = build_empty_zip();

    let err = parser
        .parse_drawing(
            drawing_xml,
            "xl/drawings/drawing1.xml",
            &Relationships::default(),
            &mut zip,
        )
        .expect_err("expected malformed drawing xml to fail");
    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "xl/drawings/drawing1.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_chartsheet_chart() {
    let chartsheet_xml = r#"
        <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <chart r:id="rIdChart"/>
        </chartsheet>
        "#;
    let chart_xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:lineChart/>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdChart"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="../charts/chart1.xml"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");
    let mut zip = build_zip_with_entries(vec![("xl/charts/chart1.xml", chart_xml)]);
    let mut parser = XlsxParser::new();
    let sheet = SheetInfo {
        name: "Chart1".to_string(),
        sheet_id: 1,
        rel_id: "rId1".to_string(),
        state: SheetState::Visible,
    };
    let ws_id = parser
        .parse_worksheet(
            &mut zip,
            chartsheet_xml,
            &sheet,
            "xl/chartsheets/sheet1.xml",
            &rels,
            SheetKind::ChartSheet,
        )
        .expect("chartsheet");
    let store = parser.into_store();
    let ws = match store.get(ws_id) {
        Some(IRNode::Worksheet(w)) => w,
        _ => panic!("missing worksheet"),
    };
    assert_eq!(ws.drawings.len(), 1);
    let drawing = match store.get(ws.drawings[0]) {
        Some(IRNode::WorksheetDrawing(d)) => d,
        _ => panic!("missing drawing"),
    };
    assert_eq!(drawing.shapes.len(), 1);
    let shape = match store.get(drawing.shapes[0]) {
        Some(IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(shape.shape_type, ShapeType::Chart);
    assert_eq!(shape.media_target.as_deref(), Some("xl/charts/chart1.xml"));
}
