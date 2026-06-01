use super::*;
use docir_core::ir::IRNode;

fn build_master_layout_fixture() -> (&'static str, &'static str) {
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
    (master_xml, layout_xml)
}

#[test]
fn test_parse_master_and_layout_shapes() {
    let (master_xml, layout_xml) = build_master_layout_fixture();

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
    master.name = extract_c_sld_name(master_xml, "ppt/slideMasters/slideMaster1.xml")
        .expect("extract master name");
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
fn test_parse_slide_layout_reports_malformed_csld_name_attributes() {
    let layout_xml = r#"
        <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld name="Layout 1" name="Layout 2">
            <p:spTree/>
          </p:cSld>
        </p:sldLayout>
        "#;

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    match parser
        .parse_slide_layout(
            layout_xml,
            "ppt/slideLayouts/slideLayout1.xml",
            &Relationships::default(),
            &mut zip,
        )
        .expect_err("malformed cSld attributes must fail")
    {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/slideLayouts/slideLayout1.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
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
    notes_master.name = extract_c_sld_name(notes_master_xml, "ppt/notesMasters/notesMaster1.xml")
        .expect("extract notes master name");
    notes_master.shapes = notes_shapes;
    let notes_id = notes_master.id;
    parser.store.insert(IRNode::NotesMaster(notes_master));

    let mut handout_master = docir_core::ir::HandoutMaster::new();
    handout_master.name =
        extract_c_sld_name(handout_master_xml, "ppt/handoutMasters/handoutMaster1.xml")
            .expect("extract handout master name");
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
