use super::*;
use docir_core::ir::IRNode;

#[test]
fn test_parse_notes_slide_text() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld>
            <p:spTree/>
          </p:cSld>
        </p:sld>
        "#;
    let notes_xml = r#"
        <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:sp>
                <p:txBody>
                  <a:p>
                    <a:r><a:t>First note</a:t></a:r>
                  </a:p>
                  <a:p>
                    <a:r><a:t>Second note</a:t></a:r>
                  </a:p>
                </p:txBody>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:notes>
        "#;

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let notes_text = parse_notes_slide(
        notes_xml,
        "ppt/notesSlides/notesSlide1.xml",
        &Relationships::default(),
        &mut parser,
        &mut zip,
    )
    .unwrap()
    .1;
    let slide_id = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide1.xml",
            &Relationships::default(),
            (Some(&notes_text), None),
        )
        .expect("parse slide");
    let store = parser.into_store();

    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert_eq!(slide.notes.as_deref(), Some("First note Second note"));
}

#[test]
fn test_parse_shapes_from_xml_covers_sp_group_and_table_paths() {
    let xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="1" name="Title" descr="alt text"/>
                  <p:cNvSpPr/>
                  <p:nvPr/>
                </p:nvSpPr>
                <p:spPr/>
                <p:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p><a:r><a:t>Hello</a:t></a:r></a:p>
                </p:txBody>
                <a:hlinkClick rel:id="rIdExt"/>
              </p:sp>
              <p:grpSp>
                <p:nvGrpSpPr><p:cNvPr id="2" name="Group 1"/></p:nvGrpSpPr>
                <p:grpSpPr>
                  <a:xfrm>
                    <a:off x="10" y="20"/>
                    <a:ext cx="30" cy="40"/>
                  </a:xfrm>
                </p:grpSpPr>
              </p:grpSp>
              <p:graphicFrame>
                <p:nvGraphicFramePr><p:cNvPr id="3" name="Table 1"/></p:nvGraphicFramePr>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                    <a:tbl>
                      <a:tblGrid><a:gridCol w="1000"/></a:tblGrid>
                      <a:tr>
                        <a:tc>
                          <a:txBody>
                            <a:bodyPr/><a:lstStyle/>
                            <a:p><a:r><a:t>Cell</a:t></a:r></a:p>
                          </a:txBody>
                        </a:tc>
                      </a:tr>
                    </a:tbl>
                  </a:graphicData>
                </a:graphic>
              </p:graphicFrame>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;

    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdExt"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            Target="https://example.test"
            TargetMode="External"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let shape_ids = parser
        .parse_shapes_from_xml(xml, "ppt/slides/slideX.xml", &rels, &mut zip)
        .expect("parse shapes");
    assert_eq!(shape_ids.len(), 3);

    let store = parser.into_store();
    let first = match store.get(shape_ids[0]) {
        Some(IRNode::Shape(shape)) => shape,
        _ => panic!("first shape missing"),
    };
    assert_eq!(first.name.as_deref(), Some("Title"));
    assert_eq!(first.alt_text.as_deref(), Some("alt text"));
    assert_eq!(first.hyperlink.as_deref(), Some("https://example.test"));

    let mut saw_group = false;
    let mut saw_table_shape = false;
    for id in shape_ids {
        let shape = match store.get(id) {
            Some(IRNode::Shape(shape)) => shape,
            _ => continue,
        };
        if matches!(shape.shape_type, ShapeType::Group) {
            saw_group = true;
        }
        if matches!(shape.shape_type, ShapeType::Table) {
            saw_table_shape = true;
        }
    }
    assert!(saw_group);
    assert!(saw_table_shape);
}

#[test]
fn test_parse_shapes_from_xml_reports_xml_error_for_malformed_input() {
    let malformed_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree><p:sp></p:spTree></p:cSld>
        </p:sld>
        "#;
    let mut parser = PptxParser::new();
    let rels = Relationships::default();
    let mut zip = build_empty_zip();

    let err = parser
        .parse_shapes_from_xml(
            malformed_xml,
            "ppt/slides/slide-bad-shapes.xml",
            &rels,
            &mut zip,
        )
        .expect_err("malformed shapes xml should fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/slides/slide-bad-shapes.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_notes_slide_reports_xml_error_for_malformed_input() {
    let malformed = r#"
        <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:sp><p:txBody><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>broken</a:t></a:r></a:p>
            </p:spTree>
          </p:cSld>
        </p:notes>
    "#;
    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let err = parse_notes_slide(
        malformed,
        "ppt/notesSlides/notesSlide-bad.xml",
        &Relationships::default(),
        &mut parser,
        &mut zip,
    )
    .expect_err("malformed notes should fail");
    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/notesSlides/notesSlide-bad.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_slide_transition_and_animation() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:transition spd="fast" advClick="1" advTm="500">
            <p:fade/>
          </p:transition>
          <p:timing>
            <p:tnLst>
              <p:par>
                <p:anim dur="300" presetID="1" presetClass="entr">
                  <p:tgtEl><p:spTgt spid="4"/></p:tgtEl>
                </p:anim>
              </p:par>
            </p:tnLst>
          </p:timing>
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
        .expect("slide");
    let store = parser.into_store();
    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert!(slide.transition.is_some());
    assert_eq!(slide.animations.len(), 1);
    assert_eq!(slide.animations[0].target.as_deref(), Some("4"));
}
