use super::*;
use docir_core::ir::{IRNode, ShapeType, TextAlignment};

#[test]
fn test_parse_smartart_part_counts() {
    let xml = r#"
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <dgm:ptLst>
            <dgm:pt/>
            <dgm:pt/>
          </dgm:ptLst>
          <dgm:cxnLst>
            <dgm:cxn/>
          </dgm:cxnLst>
          <dgm:relIds r:dm="rId1" r:lo="rId2"/>
        </dgm:dataModel>
        "#;
    let part = parse_smartart_part(xml, "ppt/diagrams/data1.xml").expect("smartart");
    assert_eq!(part.point_count, Some(2));
    assert_eq!(part.connection_count, Some(1));
    assert_eq!(part.rel_ids.len(), 2);
    assert!(part.rel_ids.contains(&"rId1".to_string()));
    assert!(part.rel_ids.contains(&"rId2".to_string()));
}

#[test]
fn test_parse_presentation_builder_end_to_end() {
    let presentation_xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        firstSlideNum="3">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rIdSlide1"/>
            <p:sldId id="257" r:id="rIdMissing"/>
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
          <p:notesSz cx="6858000" cy="9144000"/>
          <p:showPr showType="speaker" loop="1"/>
        </p:presentation>
        "#;

    let presentation_rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdSlide1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
            Target="slides/slide1.xml"/>
          <Relationship Id="rIdMaster"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
            Target="slideMasters/slideMaster1.xml"/>
          <Relationship Id="rIdNotesMaster"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
            Target="notesMasters/notesMaster1.xml"/>
          <Relationship Id="rIdHandout"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"
            Target="handoutMasters/handoutMaster1.xml"/>
          <Relationship Id="rIdExternal"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            Target="https://example.com/deck"
            TargetMode="External"/>
        </Relationships>
        "#;

    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld name="Main Slide">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="1" name="Body"/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="100" y="200"/>
                    <a:ext cx="300" cy="400"/>
                  </a:xfrm>
                </p:spPr>
                <p:txBody>
                  <a:p>
                    <a:pPr algn="ctr"/>
                    <a:r>
                      <a:rPr b="1" i="1" sz="1800"/>
                      <a:latin typeface="Calibri"/>
                      <a:t>Hello</a:t>
                    </a:r>
                    <a:br/>
                    <a:r>
                      <a:t>World</a:t>
                    </a:r>
                  </a:p>
                </p:txBody>
                <a:hlinkClick r:id="rIdHyper"/>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;

    let slide_rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdLayout"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
            Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdNotes"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
            Target="../notesSlides/notesSlide1.xml"/>
          <Relationship Id="rIdHyper"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            Target="https://example.com/slide"
            TargetMode="External"/>
        </Relationships>
        "#;

    let notes_slide_xml = r#"
        <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld name="Notes 1">
            <p:spTree>
              <p:sp>
                <p:txBody>
                  <a:p><a:r><a:t>Speaker note</a:t></a:r></a:p>
                </p:txBody>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:notes>
        "#;

    let slide_master_xml = r#"
        <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld name="Master A">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="10" name="MasterShape"/>
                </p:nvSpPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sldMaster>
        "#;

    let slide_master_rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdLayout1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
            Target="../slideLayouts/slideLayout1.xml"/>
        </Relationships>
        "#;

    let slide_layout_xml = r#"
        <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                     type="title" matchingName="Title Layout" preserve="1">
          <p:cSld name="Layout A">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="20" name="LayoutShape"/>
                </p:nvSpPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sldLayout>
        "#;

    let notes_master_xml = r#"
        <p:notesMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld name="NotesMaster A"><p:spTree/></p:cSld>
        </p:notesMaster>
        "#;
    let handout_master_xml = r#"
        <p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld name="HandoutMaster A"><p:spTree/></p:cSld>
        </p:handoutMaster>
        "#;

    let mut zip = build_zip_with_entries(vec![
        ("ppt/slides/slide1.xml", slide_xml),
        ("ppt/slides/_rels/slide1.xml.rels", slide_rels_xml),
        ("ppt/notesSlides/notesSlide1.xml", notes_slide_xml),
        ("ppt/slideMasters/slideMaster1.xml", slide_master_xml),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            slide_master_rels_xml,
        ),
        ("ppt/slideLayouts/slideLayout1.xml", slide_layout_xml),
        ("ppt/notesMasters/notesMaster1.xml", notes_master_xml),
        ("ppt/handoutMasters/handoutMaster1.xml", handout_master_xml),
        (
            "ppt/presProps.xml",
            r#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" autoCompressPictures="1"/>"#,
        ),
    ]);
    let presentation_rels = Relationships::parse(presentation_rels_xml).expect("rels");

    let mut parser = PptxParser::new();
    let root = parser
        .parse_presentation(
            &mut zip,
            presentation_xml,
            &presentation_rels,
            "ppt/presentation.xml",
        )
        .expect("parse presentation");
    let store = parser.into_store();

    let document = match store.get(root) {
        Some(IRNode::Document(d)) => d,
        _ => panic!("missing document"),
    };
    assert_eq!(
        document.content.len(),
        1,
        "missing slide rel must be skipped"
    );
    assert!(
        !document.shared_parts.is_empty(),
        "masters/layouts or optional parts should be loaded"
    );
    assert!(
        !document.security.external_refs.is_empty(),
        "external relationships should be tracked as security refs"
    );
    assert!(
        !document.diagnostics.is_empty(),
        "missing relationships should generate diagnostics"
    );
}

#[test]
fn test_parse_slide_text_alignment_and_runs() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:sp>
                <p:nvSpPr><p:cNvPr id="1" name="Text"/></p:nvSpPr>
                <p:txBody>
                  <a:p>
                    <a:pPr algn="l"></a:pPr>
                    <a:r>
                      <a:rPr b="1" i="1" sz="2200"></a:rPr>
                      <a:latin typeface="Consolas"></a:latin>
                      <a:t>Left</a:t>
                    </a:r>
                    <a:br></a:br>
                    <a:r><a:t>Line</a:t></a:r>
                  </a:p>
                  <a:p><a:pPr algn="r"></a:pPr><a:r><a:t>Right</a:t></a:r></a:p>
                  <a:p><a:pPr algn="just"></a:pPr><a:r><a:t>Justify</a:t></a:r></a:p>
                  <a:p><a:pPr algn="dist"></a:pPr><a:r><a:t>Distribute</a:t></a:r></a:p>
                  <a:p><a:pPr algn="unknown"></a:pPr><a:r><a:t>Unknown</a:t></a:r></a:p>
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
    let shape = match store.get(slide.shapes[0]) {
        Some(IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    let text = shape.text.as_ref().expect("text");
    assert_eq!(text.paragraphs.len(), 5);
    assert_eq!(text.paragraphs[0].alignment, Some(TextAlignment::Left));
    assert_eq!(text.paragraphs[1].alignment, Some(TextAlignment::Right));
    assert_eq!(text.paragraphs[2].alignment, Some(TextAlignment::Justify));
    assert_eq!(
        text.paragraphs[3].alignment,
        Some(TextAlignment::Distribute)
    );
    assert_eq!(text.paragraphs[4].alignment, None);
    assert_eq!(
        text.paragraphs[0].runs[0].font_family.as_deref(),
        Some("Consolas")
    );
    assert_eq!(text.paragraphs[0].runs[0].font_size, Some(2200));
    assert_eq!(text.paragraphs[0].runs[0].bold, Some(true));
    assert_eq!(text.paragraphs[0].runs[0].italic, Some(true));
    assert_eq!(
        shape_text_to_plain(text),
        "Left\nLine\nRight\nJustify\nDistribute\nUnknown"
    );
}

#[test]
fn test_parse_graphic_frame_transform_invalid_numbers_fallback_to_zero() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:graphicFrame>
                <p:nvGraphicFramePr>
                  <p:cNvPr id="3" name="Frame"/>
                </p:nvGraphicFramePr>
                <p:xfrm>
                  <a:off x="invalid" y="-9"/>
                  <a:ext cx="invalid" cy="400"/>
                </p:xfrm>
              </p:graphicFrame>
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
    let shape = match store.get(slide.shapes[0]) {
        Some(IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(shape.transform.x, 0);
    assert_eq!(shape.transform.y, -9);
    assert_eq!(shape.transform.width, 0);
    assert_eq!(shape.transform.height, 400);
}

#[test]
fn test_parse_metadata_table_styles_tags_and_shape_mapping() {
    let table_styles_xml = r#"
        <a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{default}">
          <a:tblStyle styleId="{default}" name="Medium Style 2"/>
          <a:tblStyle name="MissingIdShouldBeIgnored"/>
        </a:tblStyleLst>
    "#;
    let styles =
        parse_table_styles(table_styles_xml, "ppt/tableStyles.xml").expect("table styles parse");
    assert_eq!(styles.default_style_id.as_deref(), Some("{default}"));
    assert_eq!(styles.styles.len(), 1);
    assert_eq!(styles.styles[0].style_id, "{default}");
    assert_eq!(styles.styles[0].name.as_deref(), Some("Medium Style 2"));

    let tags_xml = r#"
        <p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:tag name="Department" val="Finance"/>
          <p:tag val="missing name"/>
        </p:tagLst>
    "#;
    let tags = parse_presentation_tags(tags_xml, "ppt/tags/tag1.xml").expect("tags parse");
    assert_eq!(tags.len(), 1);
    assert_eq!(tags[0].name, "Department");
    assert_eq!(tags[0].value.as_deref(), Some("Finance"));

    assert!(matches!(map_shape_type("rect"), ShapeType::Rectangle));
    assert!(matches!(map_shape_type("roundRect"), ShapeType::RoundRect));
    assert!(matches!(map_shape_type("ellipse"), ShapeType::Ellipse));
    assert!(matches!(map_shape_type("triangle"), ShapeType::Triangle));
    assert!(matches!(map_shape_type("line"), ShapeType::Line));
    assert!(matches!(map_shape_type("arrow"), ShapeType::Arrow));
    assert!(matches!(map_shape_type("freeform"), ShapeType::Custom));
}
