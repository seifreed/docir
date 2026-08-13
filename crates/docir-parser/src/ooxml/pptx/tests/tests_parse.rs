use super::*;

#[test]
fn test_parse_slide_list() {
    let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId rel:id="rId1"/>
            <p:sldId rel:id="rId2"/>
          </p:sldIdLst>
        </p:presentation>
        "#;

    let slides = parse_slide_list(xml).expect("parse slide list");
    assert_eq!(slides, vec!["rId1", "rId2"]);
}

#[test]
fn test_parse_slide_list_accepts_alternate_namespace_prefixes() {
    let xml = r#"
        <pres:presentation xmlns:pres="http://schemas.openxmlformats.org/presentationml/2006/main"
                           xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <pres:sldIdLst>
            <pres:sldId rel:id="rId1"/>
            <pres:sldId rel:id="rId2"/>
          </pres:sldIdLst>
        </pres:presentation>
        "#;

    let slides = parse_slide_list(xml).expect("parse prefixed slide list");
    assert_eq!(slides, vec!["rId1", "rId2"]);
}

#[test]
fn test_parse_slide_list_returns_xml_error_for_malformed_input() {
    let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId r:id="rId1">
          </p:sldIdLst>
        </p:presentation>
        "#;
    let err = parse_slide_list(xml).expect_err("expected malformed slide list error");
    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/presentation.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_slide_list_reports_malformed_relationship_attributes() {
    let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId r:id="rId1" r:id="rId2"/>
          </p:sldIdLst>
        </p:presentation>
        "#;

    let err = parse_slide_list(xml).expect_err("malformed slide relationship attrs must fail");
    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/presentation.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_slide_list_rejects_truncated_root() {
    let xml = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId r:id="rId1"/>"#;
    assert!(parse_slide_list(xml).is_err());
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
fn test_parse_presentation_info_returns_xml_error_for_malformed_show_properties() {
    let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:showPr showType="speaker">
        </p:presentation>
        "#;
    let err =
        parse_presentation_info(xml, "ppt/presentation.xml").expect_err("expected malformed xml");
    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/presentation.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_presentation_info_reports_malformed_show_attributes() {
    let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:showPr showType="speaker" showType="kiosk"/>
        </p:presentation>
        "#;

    let err = parse_presentation_info(xml, "ppt/presentation.xml")
        .expect_err("malformed show attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/presentation.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_presentation_info_reports_malformed_size_attributes() {
    let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:sldSz cx="9144000" cx="duplicate" cy="6858000"/>
        </p:presentation>
        "#;

    let err = parse_presentation_info(xml, "ppt/presentation.xml")
        .expect_err("malformed size attributes must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/presentation.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_presentation_info_reports_malformed_first_slide_number() {
    let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        firstSlideNum="bad">
          <p:sldSz cx="9144000" cy="6858000"/>
        </p:presentation>
        "#;

    let err = parse_presentation_info(xml, "ppt/presentation.xml")
        .expect_err("malformed first slide number must fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/presentation.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_presentation_info_rejects_truncated_root() {
    let xml = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="9144000" cy="6858000">"#;
    assert!(parse_presentation_info(xml, "ppt/presentation.xml").is_err());
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

fn assert_pptx_metadata_xml_error<T>(result: Result<T, ParseError>, expected_file: &str) {
    match result {
        Err(ParseError::Xml { file, .. }) => assert_eq!(file, expected_file),
        Err(other) => panic!("unexpected error: {other:?}"),
        Ok(_) => panic!("expected pptx metadata xml error"),
    }
}

#[test]
fn test_parse_pptx_metadata_reports_malformed_attributes() {
    assert_pptx_metadata_xml_error(
        parse_presentation_properties(
            r#"<p:presentationPr xmlns:p="x" rtl="1" rtl="0"/>"#,
            "ppt/presProps.xml",
        ),
        "ppt/presProps.xml",
    );

    assert_pptx_metadata_xml_error(
        parse_view_properties(
            r#"<p:viewPr xmlns:p="x"><p:zoom percent="85" percent="90"/></p:viewPr>"#,
            "ppt/viewProps.xml",
        ),
        "ppt/viewProps.xml",
    );

    assert_pptx_metadata_xml_error(
        parse_table_styles(
            r#"<a:tblStyleLst xmlns:a="x"><a:tblStyle styleId="{a}" styleId="{b}"/></a:tblStyleLst>"#,
            "ppt/tableStyles.xml",
        ),
        "ppt/tableStyles.xml",
    );

    assert_pptx_metadata_xml_error(
        parse_presentation_tags(
            r#"<p:tagLst xmlns:p="x"><p:tag name="Department" name="Finance"/></p:tagLst>"#,
            "ppt/tags/tag1.xml",
        ),
        "ppt/tags/tag1.xml",
    );

    assert_pptx_metadata_xml_error(
        parse_smartart_part(
            r#"<dgm:dataModel xmlns:dgm="x" xmlns:r="r"><dgm:relIds r:dm="a" r:dm="b"/></dgm:dataModel>"#,
            "ppt/diagrams/data1.xml",
        ),
        "ppt/diagrams/data1.xml",
    );

    assert_pptx_metadata_xml_error(
        parse_slide_layout_meta(
            r#"<p:sldLayout xmlns:p="x" preserve="1" preserve="0"/>"#,
            "ppt/slideLayouts/slideLayout1.xml",
        ),
        "ppt/slideLayouts/slideLayout1.xml",
    );

    assert_pptx_metadata_xml_error(
        parse_slide_master_meta(
            r#"<p:sldMaster xmlns:p="x" preserve="1" preserve="0"/>"#,
            "ppt/slideMasters/slideMaster1.xml",
        ),
        "ppt/slideMasters/slideMaster1.xml",
    );
}

#[test]
fn test_parse_view_properties_reports_malformed_zoom_value() {
    assert_pptx_metadata_xml_error(
        parse_view_properties(
            r#"<p:viewPr xmlns:p="x"><p:zoom percent="bad"/></p:viewPr>"#,
            "ppt/viewProps.xml",
        ),
        "ppt/viewProps.xml",
    );
}

#[test]
fn test_parse_pptx_metadata_accepts_alternate_namespace_prefixes() {
    let presentation_xml = r#"
        <pres:presentation xmlns:pres="http://schemas.openxmlformats.org/presentationml/2006/main"
                           firstSlideNum="7">
          <pres:sldSz cx="9144000" cy="6858000" type="screen16x9"/>
          <pres:notesSz cx="6858000" cy="9144000"/>
          <pres:showPr showType="kiosk" loop="1" showNarration="0" showAnimation="1" useTimings="1"/>
        </pres:presentation>
        "#;
    let info = parse_presentation_info(presentation_xml, "ppt/presentation.xml")
        .expect("presentation info")
        .expect("info present");
    assert_eq!(info.first_slide_num, Some(7));
    assert_eq!(info.slide_size.as_ref().unwrap().cx, 9144000);
    assert_eq!(
        info.slide_size.as_ref().unwrap().size_type.as_deref(),
        Some("screen16x9")
    );
    assert_eq!(info.notes_size.as_ref().unwrap().cy, 9144000);
    assert_eq!(info.show_type.as_deref(), Some("kiosk"));
    assert_eq!(info.show_loop, Some(true));

    let pres_props_xml = r#"
        <props:presentationPr xmlns:props="http://schemas.openxmlformats.org/presentationml/2006/main"
                              removePersonalInfoOnSave="1"
                              showInkAnnotation="0"/>
        "#;
    let props =
        parse_presentation_properties(pres_props_xml, "ppt/presProps.xml").expect("pres props");
    assert_eq!(props.remove_personal_info_on_save, Some(true));
    assert_eq!(props.show_ink_annotation, Some(false));

    let view_xml = r#"
        <view:viewPr xmlns:view="http://schemas.openxmlformats.org/presentationml/2006/main"
                     lastView="slideSorterView"
                     showHiddenSlides="1">
          <view:zoom percent="85"/>
        </view:viewPr>
        "#;
    let view = parse_view_properties(view_xml, "ppt/viewProps.xml").expect("view props");
    assert_eq!(view.last_view.as_deref(), Some("slideSorterView"));
    assert_eq!(view.show_hidden_slides, Some(true));
    assert_eq!(view.zoom, Some(85));

    let table_styles_xml = r#"
        <draw:tblStyleLst xmlns:draw="http://schemas.openxmlformats.org/drawingml/2006/main"
                          def="{default-style}">
          <draw:tblStyle styleId="{style-1}" name="Table Style"/>
        </draw:tblStyleLst>
        "#;
    let styles = parse_table_styles(table_styles_xml, "ppt/tableStyles.xml").expect("table styles");
    assert_eq!(styles.default_style_id.as_deref(), Some("{default-style}"));
    assert_eq!(styles.styles.len(), 1);
    assert_eq!(styles.styles[0].name.as_deref(), Some("Table Style"));

    let layout_xml = r#"
        <layout:sldLayout xmlns:layout="http://schemas.openxmlformats.org/presentationml/2006/main"
                          type="titleOnly"
                          matchingName="Title Only"
                          preserve="0"
                          showMasterSp="true"
                          showMasterPhAnim="0"/>
        "#;
    let layout = parse_slide_layout_meta(layout_xml, "ppt/slideLayouts/slideLayout1.xml")
        .expect("layout meta");
    assert_eq!(layout.layout_type.as_deref(), Some("titleOnly"));
    assert_eq!(layout.matching_name.as_deref(), Some("Title Only"));

    let master_xml = r#"
        <master:sldMaster xmlns:master="http://schemas.openxmlformats.org/presentationml/2006/main"
                          preserve="1"
                          showMasterSp="1"
                          showMasterPhAnim="0"/>
        "#;
    let master = parse_slide_master_meta(master_xml, "ppt/slideMasters/slideMaster1.xml")
        .expect("master meta");
    assert_eq!(master.preserve, Some(true));
    assert_eq!(master.show_master_sp, Some(true));
    assert_eq!(master.show_master_ph_anim, Some(false));
}
