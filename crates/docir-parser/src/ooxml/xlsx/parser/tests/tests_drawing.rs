use super::{
    ParseError, Relationships, SheetKind, SheetState, build_empty_zip, build_zip_with_entries,
};
use crate::ooxml::xlsx::workbook::SheetInfo;
use crate::ooxml::xlsx::{ShapeType, XlsxParser};
use docir_core::ir::IRNode;

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
        Some("https://example.com/image.png")
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
