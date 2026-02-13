
use super::*;
use crate::zip_handler::SecureZipReader;
use docir_core::ir::IRNode;
use std::io::Cursor;

#[test]
fn test_parse_slide_list() {
    let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId r:id="rId1"/>
            <p:sldId r:id="rId2"/>
          </p:sldIdLst>
        </p:presentation>
        "#;

    let slides = parse_slide_list(xml).expect("parse slide list");
    assert_eq!(slides, vec!["rId1", "rId2"]);
}

#[test]
fn test_parse_presentation_info() {
    let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        firstSlideNum="5">
          <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
          <p:notesSz cx="6858000" cy="9144000"/>
          <p:showPr showType="kiosk" loop="1" showNarration="0" showAnimation="1" useTimings="1"/>
        </p:presentation>
        "#;
    let info = parse_presentation_info(xml, "ppt/presentation.xml")
        .expect("info")
        .expect("info present");
    assert_eq!(info.first_slide_num, Some(5));
    assert_eq!(info.slide_size.as_ref().unwrap().cx, 9144000);
    assert_eq!(info.notes_size.as_ref().unwrap().cy, 9144000);
    assert_eq!(info.show_type.as_deref(), Some("kiosk"));
    assert_eq!(info.show_loop, Some(true));
    assert_eq!(info.show_narration, Some(false));
    assert_eq!(info.show_animation, Some(true));
    assert_eq!(info.use_timings, Some(true));
}

#[test]
fn test_parse_presentation_and_view_properties_extended() {
    let pres_xml = r#"
        <p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                          removePersonalInfoOnSave="1"
                          showInkAnnotation="0"/>
        "#;
    let view_xml = r#"
        <p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  lastView="slideSorterView"
                  showHiddenSlides="1"
                  showGuides="0"
                  showGrid="1"
                  showOutlineIcons="1">
          <p:zoom percent="85"/>
        </p:viewPr>
        "#;
    let props = parse_presentation_properties(pres_xml, "ppt/presProps.xml").expect("pres props");
    assert_eq!(props.remove_personal_info_on_save, Some(true));
    assert_eq!(props.show_ink_annotation, Some(false));

    let view = parse_view_properties(view_xml, "ppt/viewProps.xml").expect("view props");
    assert_eq!(view.last_view.as_deref(), Some("slideSorterView"));
    assert_eq!(view.show_hidden_slides, Some(true));
    assert_eq!(view.show_guides, Some(false));
    assert_eq!(view.show_grid, Some(true));
    assert_eq!(view.show_outline_icons, Some(true));
    assert_eq!(view.zoom, Some(85));
}

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
            None,
            None,
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
            None,
            None,
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
            None,
            None,
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
            Some(&notes_text),
            None,
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
fn test_parse_master_and_layout_shapes() {
    let master_xml = r#"
        <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                     preserve="1" showMasterSp="1" showMasterPhAnim="0">
          <p:cSld name="Master 1">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="1" name="MasterShape"/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="100" cy="100"/>
                  </a:xfrm>
                </p:spPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sldMaster>
        "#;
    let layout_xml = r#"
        <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                     type="title" matchingName="Title" preserve="1" showMasterSp="1" showMasterPhAnim="0">
          <p:cSld name="Layout 1">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="2" name="LayoutShape"/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="200" cy="200"/>
                  </a:xfrm>
                </p:spPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sldLayout>
        "#;

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let master_shapes = parser
        .parse_shapes_from_xml(
            master_xml,
            "ppt/slideMasters/slideMaster1.xml",
            &Relationships::default(),
            &mut zip,
        )
        .expect("parse master shapes");
    let layout_id = parser
        .parse_slide_layout(
            layout_xml,
            "ppt/slideLayouts/slideLayout1.xml",
            &Relationships::default(),
            &mut zip,
        )
        .expect("parse layout");
    let mut master = docir_core::ir::SlideMaster::new();
    master.name = extract_c_sld_name(master_xml);
    let meta = parse_slide_master_meta(master_xml, "ppt/slideMasters/slideMaster1.xml")
        .expect("master meta");
    master.preserve = meta.preserve;
    master.show_master_sp = meta.show_master_sp;
    master.show_master_ph_anim = meta.show_master_ph_anim;
    master.shapes = master_shapes;
    master.layouts = vec![layout_id];
    let master_id = master.id;
    parser.store.insert(IRNode::SlideMaster(master));

    let store = parser.into_store();
    let master_node = match store.get(master_id) {
        Some(IRNode::SlideMaster(m)) => m,
        _ => panic!("missing master"),
    };
    assert_eq!(master_node.name.as_deref(), Some("Master 1"));
    assert_eq!(master_node.preserve, Some(true));
    assert_eq!(master_node.show_master_sp, Some(true));
    assert_eq!(master_node.show_master_ph_anim, Some(false));
    assert_eq!(master_node.shapes.len(), 1);
    assert_eq!(master_node.layouts.len(), 1);

    let layout_node = match store.get(layout_id) {
        Some(IRNode::SlideLayout(l)) => l,
        _ => panic!("missing layout"),
    };
    assert_eq!(layout_node.layout_type.as_deref(), Some("title"));
    assert_eq!(layout_node.matching_name.as_deref(), Some("Title"));
    assert_eq!(layout_node.preserve, Some(true));
    assert_eq!(layout_node.show_master_sp, Some(true));
    assert_eq!(layout_node.show_master_ph_anim, Some(false));
}

#[test]
fn test_parse_notes_and_handout_master_shapes() {
    let notes_master_xml = r#"
        <p:notesMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld name="NotesMaster 1">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="10" name="NotesShape"/>
                </p:nvSpPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:notesMaster>
        "#;
    let handout_master_xml = r#"
        <p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld name="HandoutMaster 1">
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="11" name="HandoutShape"/>
                </p:nvSpPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:handoutMaster>
        "#;

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let notes_shapes = parser
        .parse_shapes_from_xml(
            notes_master_xml,
            "ppt/notesMasters/notesMaster1.xml",
            &Relationships::default(),
            &mut zip,
        )
        .expect("parse notes master shapes");
    let handout_shapes = parser
        .parse_shapes_from_xml(
            handout_master_xml,
            "ppt/handoutMasters/handoutMaster1.xml",
            &Relationships::default(),
            &mut zip,
        )
        .expect("parse handout master shapes");

    let mut notes_master = docir_core::ir::NotesMaster::new();
    notes_master.name = extract_c_sld_name(notes_master_xml);
    notes_master.shapes = notes_shapes;
    let notes_id = notes_master.id;
    parser.store.insert(IRNode::NotesMaster(notes_master));

    let mut handout_master = docir_core::ir::HandoutMaster::new();
    handout_master.name = extract_c_sld_name(handout_master_xml);
    handout_master.shapes = handout_shapes;
    let handout_id = handout_master.id;
    parser.store.insert(IRNode::HandoutMaster(handout_master));

    let store = parser.into_store();
    let notes = match store.get(notes_id) {
        Some(IRNode::NotesMaster(m)) => m,
        _ => panic!("missing notes master"),
    };
    let handout = match store.get(handout_id) {
        Some(IRNode::HandoutMaster(m)) => m,
        _ => panic!("missing handout master"),
    };
    assert_eq!(notes.name.as_deref(), Some("NotesMaster 1"));
    assert_eq!(notes.shapes.len(), 1);
    assert_eq!(handout.name.as_deref(), Some("HandoutMaster 1"));
    assert_eq!(handout.shapes.len(), 1);
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
            None,
            None,
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

#[test]
fn test_parse_slide_timing_media() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld><p:spTree/></p:cSld>
          <p:timing>
            <p:audio r:link="rIdAudio" dur="5000"/>
            <p:video r:embed="rIdVideo" dur="12000"/>
          </p:timing>
        </p:sld>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdAudio"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
            Target="../media/audio1.wav"/>
          <Relationship Id="rIdVideo"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
            Target="../media/video1.mp4"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");
    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let slide_id = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide1.xml",
            &rels,
            None,
            None,
        )
        .expect("slide");
    let store = parser.into_store();
    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert_eq!(slide.animations.len(), 2);
    assert_eq!(
        slide.animations[0].target.as_deref(),
        Some("ppt/media/audio1.wav")
    );
    assert_eq!(
        slide.animations[1].target.as_deref(),
        Some("ppt/media/video1.mp4")
    );
}

#[test]
fn test_parse_slide_comments() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
        </p:sld>
        "#;
    let comments_xml = r#"
        <p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cm authorId="1" dt="2024-01-01T00:00:00Z">
            <p:txBody><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Note</a:t></a:r></a:p></p:txBody>
          </p:cm>
        </p:cmLst>
        "#;
    let authors_xml = r#"
        <p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cmAuthor id="1" name="Alice" initials="AL"/>
        </p:cmAuthorLst>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdC"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            Target="../comments/comment1.xml"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");
    let mut zip = build_zip_with_entries(vec![("ppt/comments/comment1.xml", comments_xml)]);
    let mut parser = PptxParser::new();
    let authors = parse_comment_authors(authors_xml, "ppt/commentAuthors.xml").expect("authors");
    parser.set_comment_authors(&authors);
    let slide_id = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide1.xml",
            &rels,
            None,
            None,
        )
        .expect("slide");
    let store = parser.into_store();
    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert_eq!(slide.comments.len(), 1);
    let comment = match store.get(slide.comments[0]) {
        Some(IRNode::PptxComment(c)) => c,
        _ => panic!("missing comment"),
    };
    assert_eq!(comment.text, "Note");
    assert_eq!(comment.author_name.as_deref(), Some("Alice"));
    assert_eq!(comment.author_initials.as_deref(), Some("AL"));
}

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
fn test_parse_graphic_frame_ole_object() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:graphicFrame>
                <p:nvGraphicFramePr>
                  <p:cNvPr id="5" name="OLE 1"/>
                </p:nvGraphicFramePr>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
                    <p:oleObj r:id="rId5"/>
                  </a:graphicData>
                </a:graphic>
              </p:graphicFrame>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId5"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
            Target="../embeddings/oleObject1.bin"/>
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
            None,
            None,
        )
        .expect("parse slide");
    let store = parser.into_store();

    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert_eq!(slide.shapes.len(), 1);
    let shape = match store.get(slide.shapes[0]) {
        Some(IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(shape.shape_type, ShapeType::OleObject);
    assert_eq!(
        shape.media_target.as_deref(),
        Some("ppt/embeddings/oleObject1.bin")
    );
}

#[test]
fn test_parse_pic_external_media_reference() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:pic>
                <p:blipFill>
                  <a:blip r:embed="rIdAudio"/>
                </p:blipFill>
              </p:pic>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdAudio"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
            Target="https://example.com/audio.wav"
            TargetMode="External"/>
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
            None,
            None,
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
    assert_eq!(shape.shape_type, ShapeType::Audio);
    assert_eq!(
        shape.media_target.as_deref(),
        Some("https://example.com/audio.wav")
    );
}

#[test]
fn test_parse_pic_linked_media_reference() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:pic>
                <p:blipFill>
                  <a:blip r:link="rIdVideo"/>
                </p:blipFill>
              </p:pic>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdVideo"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
            Target="https://example.com/video.mp4"
            TargetMode="External"/>
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
            None,
            None,
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
    assert_eq!(shape.shape_type, ShapeType::Video);
    assert_eq!(
        shape.media_target.as_deref(),
        Some("https://example.com/video.mp4")
    );
}

#[test]
fn test_parse_pic_embed_and_link_external() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:pic>
                <p:blipFill>
                  <a:blip r:embed="rIdImg" r:link="rIdExt"/>
                </p:blipFill>
              </p:pic>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image2.png"/>
          <Relationship Id="rIdExt"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
            Target="https://example.com/video.mp4"
            TargetMode="External"/>
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
            None,
            None,
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

    let mut ext_targets = Vec::new();
    for node in store.values() {
        if let IRNode::ExternalReference(ext) = node {
            ext_targets.push(ext.target.clone());
        }
    }
    assert!(ext_targets
        .iter()
        .any(|t| t == "https://example.com/video.mp4"));
}

#[test]
fn test_parse_table_in_graphic_frame() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:graphicFrame>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                    <a:tbl>
                      <a:tblGrid>
                        <a:gridCol w="3000"/>
                        <a:gridCol w="3000"/>
                      </a:tblGrid>
                      <a:tr>
                        <a:tc><a:txBody><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody></a:tc>
                        <a:tc><a:txBody><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody></a:tc>
                      </a:tr>
                    </a:tbl>
                  </a:graphicData>
                </a:graphic>
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
            None,
            None,
        )
        .expect("slide");
    let store = parser.into_store();
    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    let shape = match store.get(slide.shapes[0]) {
        Some(IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(shape.shape_type, ShapeType::Table);
    let table_id = shape.table.expect("table id");
    let table = match store.get(table_id) {
        Some(IRNode::Table(t)) => t,
        _ => panic!("missing table"),
    };
    assert_eq!(table.rows.len(), 1);
    assert_eq!(table.grid.len(), 2);
}

fn build_empty_zip() -> SecureZipReader<Cursor<Vec<u8>>> {
    build_zip_with_entries(Vec::new())
}

fn build_zip_with_entries(entries: Vec<(&str, &str)>) -> SecureZipReader<Cursor<Vec<u8>>> {
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
    SecureZipReader::new(std::io::Cursor::new(data), Default::default()).expect("zip")
}
