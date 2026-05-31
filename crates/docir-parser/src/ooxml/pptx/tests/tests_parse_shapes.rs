use super::*;
use docir_core::ir::IRNode;

#[test]
fn test_parse_slide_shapes() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               show="0">
          <p:cSld name="Title Slide">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="1" name="Title"/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="100" y="200"/>
                    <a:ext cx="300" cy="400"/>
                  </a:xfrm>
                </p:spPr>
                <p:txBody>
                  <a:p>
                    <a:r>
                      <a:rPr b="1" sz="2400"/>
                      <a:t>Hello</a:t>
                    </a:r>
                  </a:p>
                </p:txBody>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let slide_id = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide1.xml",
            &Relationships::default(),
            (None, None),
        )
        .expect("parse slide");
    let store = parser.into_store();

    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };

    assert_eq!(slide.number, 1);
    assert!(slide.hidden);
    assert_eq!(slide.name.as_deref(), Some("Title Slide"));
    assert_eq!(slide.shapes.len(), 1);

    let shape = match store.get(slide.shapes[0]) {
        Some(IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };

    assert_eq!(shape.name.as_deref(), Some("Title"));
    assert_eq!(shape.transform.x, 100);
    assert_eq!(shape.transform.y, 200);
    assert_eq!(shape.transform.width, 300);
    assert_eq!(shape.transform.height, 400);
    assert!(shape.text.is_some());
}

#[test]
fn test_parse_shape_reports_truncated_xml() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="1" name="Title"/>
                </p:nvSpPr>
    "#;

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let err = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/broken-shape.xml",
            &Relationships::default(),
            (None, None),
        )
        .expect_err("truncated shape XML must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/slides/broken-shape.xml"),
        other => panic!("expected XML error, got {other:?}"),
    }
}

#[test]
fn test_parse_group_shape_reports_truncated_xml() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:grpSp>
                <p:nvGrpSpPr>
                  <p:cNvPr id="2" name="Group 1"/>
                </p:nvGrpSpPr>
    "#;

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let err = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/broken-group-shape.xml",
            &Relationships::default(),
            (None, None),
        )
        .expect_err("truncated group shape XML must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/slides/broken-group-shape.xml"),
        other => panic!("expected XML error, got {other:?}"),
    }
}

#[test]
fn test_parse_group_properties_reports_truncated_xml() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:grpSp>
                <p:grpSpPr>
    "#;

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let err = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/broken-group-props.xml",
            &Relationships::default(),
            (None, None),
        )
        .expect_err("truncated group properties XML must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/slides/broken-group-props.xml"),
        other => panic!("expected XML error, got {other:?}"),
    }
}

#[test]
fn test_parse_slide_accepts_alternate_namespace_prefixes() {
    let slide_xml = r#"
        <deck:sld xmlns:deck="http://schemas.openxmlformats.org/presentationml/2006/main"
                  xmlns:draw="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  show="0">
          <deck:cSld name="Prefixed Slide">
            <deck:spTree>
              <deck:sp>
                <deck:nvSpPr>
                  <deck:cNvPr id="1" name="Title" descr="Title alt"/>
                </deck:nvSpPr>
                <deck:spPr>
                  <draw:xfrm>
                    <draw:off x="100" y="200"/>
                    <draw:ext cx="300" cy="400"/>
                  </draw:xfrm>
                  <draw:prstGeom prst="ellipse"/>
                </deck:spPr>
                <deck:txBody>
                  <draw:p>
                    <draw:pPr algn="ctr"/>
                    <draw:r>
                      <draw:rPr b="1" sz="2400"/>
                      <draw:t>Hello</draw:t>
                      <draw:latin typeface="Aptos"/>
                    </draw:r>
                  </draw:p>
                </deck:txBody>
              </deck:sp>
              <deck:pic>
                <deck:nvPicPr>
                  <deck:cNvPr id="2" name="Picture 1" descr="Alt text"/>
                </deck:nvPicPr>
                <deck:blipFill>
                  <draw:blip rel:embed="rIdImage"/>
                </deck:blipFill>
                <deck:spPr>
                  <draw:xfrm>
                    <draw:off x="10" y="20"/>
                    <draw:ext cx="30" cy="40"/>
                  </draw:xfrm>
                </deck:spPr>
              </deck:pic>
              <deck:graphicFrame>
                <deck:nvGraphicFramePr>
                  <deck:cNvPr id="3" name="Table 1"/>
                </deck:nvGraphicFramePr>
                <draw:graphic>
                  <draw:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                    <draw:tbl>
                      <draw:tblGrid><draw:gridCol w="1000"/></draw:tblGrid>
                      <draw:tr>
                        <draw:tc>
                          <draw:txBody>
                            <draw:p><draw:r><draw:t>Cell</draw:t></draw:r></draw:p>
                          </draw:txBody>
                        </draw:tc>
                      </draw:tr>
                    </draw:tbl>
                  </draw:graphicData>
                </draw:graphic>
              </deck:graphicFrame>
            </deck:spTree>
          </deck:cSld>
          <deck:transition spd="fast" advClick="1" advTm="500">
            <deck:fade/>
          </deck:transition>
          <deck:timing>
            <deck:tnLst>
              <deck:par>
                <deck:anim dur="300" presetID="1" presetClass="entr">
                  <deck:tgtEl><deck:spTgt spid="1"/></deck:tgtEl>
                </deck:anim>
              </deck:par>
            </deck:tnLst>
          </deck:timing>
        </deck:sld>
        "#;
    let rels = Relationships::parse(
        r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImage"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image1.png"/>
        </Relationships>
        "#,
    )
    .expect("rels");

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let slide_id = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide1.xml",
            &rels,
            (None, None),
        )
        .expect("parse prefixed slide");
    let store = parser.into_store();

    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(slide)) => slide,
        _ => panic!("missing slide"),
    };
    assert!(slide.hidden);
    assert_eq!(slide.name.as_deref(), Some("Prefixed Slide"));
    assert_eq!(slide.shapes.len(), 3);
    assert!(slide.transition.is_some());
    assert_eq!(slide.animations.len(), 1);
    assert_eq!(slide.animations[0].target.as_deref(), Some("1"));

    let text_shape = match store.get(slide.shapes[0]) {
        Some(IRNode::Shape(shape)) => shape,
        _ => panic!("missing text shape"),
    };
    assert_eq!(text_shape.name.as_deref(), Some("Title"));
    assert_eq!(text_shape.alt_text.as_deref(), Some("Title alt"));
    assert_eq!(text_shape.shape_type, ShapeType::Ellipse);
    assert_eq!(text_shape.transform.x, 100);
    assert_eq!(text_shape.transform.width, 300);
    assert_eq!(
        shape_text_to_plain(text_shape.text.as_ref().expect("shape text")).as_str(),
        "Hello"
    );

    let picture = match store.get(slide.shapes[1]) {
        Some(IRNode::Shape(shape)) => shape,
        _ => panic!("missing picture"),
    };
    assert_eq!(picture.shape_type, ShapeType::Picture);
    assert_eq!(
        picture.media_target.as_deref(),
        Some("ppt/media/image1.png")
    );
    assert_eq!(picture.alt_text.as_deref(), Some("Alt text"));

    let table_shape = match store.get(slide.shapes[2]) {
        Some(IRNode::Shape(shape)) => shape,
        _ => panic!("missing table shape"),
    };
    assert_eq!(table_shape.shape_type, ShapeType::Table);
    let table_id = table_shape.table.expect("table id");
    let table = match store.get(table_id) {
        Some(IRNode::Table(table)) => table,
        _ => panic!("missing table"),
    };
    assert_eq!(table.grid.len(), 1);
    assert_eq!(table.rows.len(), 1);
}

#[test]
fn test_parse_pic_with_media_target() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <p:cSld>
            <p:spTree>
              <p:pic>
                <p:nvPicPr>
                  <p:cNvPr id="2" name="Picture 1" descr="Alt text"/>
                </p:nvPicPr>
                <p:blipFill>
                  <a:blip r:embed="rId2"/>
                </p:blipFill>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="10" y="20"/>
                    <a:ext cx="300" cy="400"/>
                  </a:xfrm>
                </p:spPr>
              </p:pic>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;

    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId2"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image2.png"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels parse");

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let slide_id = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide1.xml",
            &rels,
            (None, None),
        )
        .expect("parse slide");
    let store = parser.into_store();

    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    let shape = match store.get(slide.shapes[0]) {
        Some(IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(shape.shape_type, ShapeType::Picture);
    assert_eq!(shape.media_target.as_deref(), Some("ppt/media/image2.png"));
    assert_eq!(shape.alt_text.as_deref(), Some("Alt text"));
}

#[test]
fn test_parse_graphic_frame_chart() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <p:cSld>
            <p:spTree>
              <p:graphicFrame>
                <p:nvGraphicFramePr>
                  <p:cNvPr id="3" name="Chart 1"/>
                </p:nvGraphicFramePr>
                <p:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="1000" cy="800"/>
                </p:xfrm>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                    <c:chart r:id="rId3"/>
                  </a:graphicData>
                </a:graphic>
              </p:graphicFrame>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;

    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId3"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="../charts/chart1.xml"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels parse");

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let slide_id = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide1.xml",
            &rels,
            (None, None),
        )
        .expect("parse slide");
    let store = parser.into_store();

    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    let shape = match store.get(slide.shapes[0]) {
        Some(IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(shape.shape_type, ShapeType::Chart);
    assert_eq!(shape.media_target.as_deref(), Some("ppt/charts/chart1.xml"));
}
