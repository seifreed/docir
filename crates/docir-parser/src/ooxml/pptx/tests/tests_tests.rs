use super::*;
use crate::error::ParseError;
use crate::zip_handler::SecureZipReader;
use docir_core::IRNode;
use std::io::Cursor;

#[test]
fn test_parse_metadata_meta_and_smartart_variants() {
    let pres_xml = r#"
        <p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                          autoCompressPictures="0"
                          compatMode="compat15"
                          rtl="true"
                          showSpecialPlsOnTitleSld="1"/>
    "#;
    let props = parse_presentation_properties(pres_xml, "ppt/presProps.xml").expect("pres props");
    assert_eq!(props.auto_compress_pictures, Some(false));
    assert_eq!(props.compat_mode.as_deref(), Some("compat15"));
    assert_eq!(props.rtl, Some(true));
    assert_eq!(props.show_special_placeholders, Some(true));

    let view_xml = r#"
        <p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  showComments="true"
                  showHiddenSlides="0">
          <p:zoom percent="invalid"/>
        </p:viewPr>
    "#;
    let view = parse_view_properties(view_xml, "ppt/viewProps.xml").expect("view props");
    assert_eq!(view.show_comments, Some(true));
    assert_eq!(view.show_hidden_slides, Some(false));
    assert_eq!(view.zoom, None, "invalid zoom value should be ignored");

    let layout_meta_xml = r#"
        <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     type="titleOnly"
                     matchingName="Title Only"
                     preserve="0"
                     showMasterSp="true"
                     showMasterPhAnim="0"/>
    "#;
    let layout_meta = parse_slide_layout_meta(layout_meta_xml, "ppt/slideLayouts/slideLayout1.xml")
        .expect("layout meta");
    assert_eq!(layout_meta.layout_type.as_deref(), Some("titleOnly"));
    assert_eq!(layout_meta.matching_name.as_deref(), Some("Title Only"));
    assert_eq!(layout_meta.preserve, Some(false));
    assert_eq!(layout_meta.show_master_sp, Some(true));
    assert_eq!(layout_meta.show_master_ph_anim, Some(false));

    let smartart_xml = r#"
        <dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <dgm:pt/>
          <dgm:cxn/>
          <dgm:relIds r:dm="rIdDm" r:lo="" r:qs="rIdQs" r:cs="rIdCs"/>
        </dgm:layoutDef>
    "#;
    let layout_part =
        parse_smartart_part(smartart_xml, "ppt/diagrams/layout1.xml").expect("smartart layout");
    assert_eq!(layout_part.kind, "layout");
    assert_eq!(layout_part.point_count, Some(1));
    assert_eq!(layout_part.connection_count, Some(1));
    assert_eq!(layout_part.rel_ids, vec!["rIdDm", "rIdQs", "rIdCs"]);

    let style_part =
        parse_smartart_part("<dgm:styleDef xmlns:dgm='x'/>", "ppt/diagrams/style1.xml")
            .expect("smartart style");
    assert_eq!(style_part.kind, "style");

    let colors_part =
        parse_smartart_part("<dgm:colorsDef xmlns:dgm='x'/>", "ppt/diagrams/colors1.xml")
            .expect("smartart colors");
    assert_eq!(colors_part.kind, "colors");

    let data_part = parse_smartart_part("<dgm:dataModel xmlns:dgm='x'/>", "ppt/diagrams/data1.xml")
        .expect("smartart data");
    assert_eq!(data_part.kind, "data");
}

#[test]
fn test_parse_metadata_tolerates_truncated_parts() {
    let bad_pres = r#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:bad>"#;
    let props = parse_presentation_properties(bad_pres, "ppt/presProps.xml")
        .expect("presentation props parser should be tolerant here");
    assert!(props.auto_compress_pictures.is_none());
    assert!(props.compat_mode.is_none());

    let bad_view = r#"<p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:zoom>"#;
    let view =
        parse_view_properties(bad_view, "ppt/viewProps.xml").expect("view parser tolerates eof");
    assert!(view.last_view.is_none());
    assert!(view.zoom.is_none());

    let bad_styles = r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblStyle>"#;
    let styles = parse_table_styles(bad_styles, "ppt/tableStyles.xml")
        .expect("table styles parser tolerates eof");
    assert!(styles.default_style_id.is_none());
    assert!(styles.styles.is_empty());

    let bad_tags =
        r#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag"#;
    let err =
        parse_presentation_tags(bad_tags, "ppt/tags/tag1.xml").expect_err("malformed tags fail");
    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/tags/tag1.xml"),
        other => panic!("unexpected error: {other:?}"),
    }

    let bad_smartart = r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:pt>"#;
    let smartart = parse_smartart_part(bad_smartart, "ppt/diagrams/data1.xml")
        .expect("smartart parser tolerates eof");
    assert_eq!(smartart.kind, "data");
}

#[test]
fn test_parse_shapes_from_xml_handles_empty_shape_nodes_and_unknown_children() {
    let xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:sp>
                <p:nvSpPr><p:cNvPr id="1" name="Empty Shape"/></p:nvSpPr>
                <p:spPr/>
                <p:txBody>
                  <a:bodyPr/>
                  <a:p><a:r><a:t></a:t></a:r></a:p>
                </p:txBody>
              </p:sp>
              <p:unknown/>
            </p:spTree>
          </p:cSld>
        </p:sld>
        "#;

    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let slide_id = parser
        .parse_slide(
            &mut zip,
            xml,
            1,
            "ppt/slides/slide-empty-shape.xml",
            &Relationships::default(),
            (None, None),
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
    assert_eq!(shape.name.as_deref(), Some("Empty Shape"));
    assert_eq!(shape.shape_type, ShapeType::TextBox);
}

#[test]
fn test_parse_pptx_table_handles_empty_columns_empty_cells_and_unknown_children() {
    let tbl_xml = r#"
        <a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:tblGrid>
            <a:gridCol/>
            <a:gridCol w="2000"/>
            <a:gridCol w="bad"/>
          </a:tblGrid>
          <a:tr>
            <a:tc>
              <a:txBody><a:p><a:r><a:t>One</a:t></a:r></a:p></a:txBody>
            </a:tc>
            <a:tc/>
            <a:badChild/>
          </a:tr>
        </a:tbl>
        "#;

    let mut reader = quick_xml::Reader::from_str(tbl_xml);
    reader.config_mut().trim_text(true);
    let mut parser = PptxParser::new();
    let table = parser
        .parse_pptx_table(&mut reader, "ppt/tables/table1.xml")
        .expect("parse table");
    assert_eq!(table.grid.len(), 1);
    assert_eq!(table.rows.len(), 1);

    let table_store = parser.into_store();
    let row = match table_store.get(table.rows[0]) {
        Some(IRNode::TableRow(row)) => row,
        _ => panic!("missing row"),
    };
    assert_eq!(row.cells.len(), 1);
}

#[test]
fn test_parse_presentation_info_ignores_invalid_size_values() {
    let xml = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:sldSz cx="bad" cy="123"/>
        </p:presentation>
        "#;

    let info = parse_presentation_info(xml, "ppt/presentation.xml").expect("presentation info");
    assert!(
        info.is_none(),
        "invalid dimensions should keep presentation info absent"
    );
}

#[test]
fn test_parse_slide_master_meta_tolerates_missing_root_and_parses_flags() {
    let master_xml = r#"
        <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     preserve="0"
                     showMasterSp="1"
                     showMasterPhAnim="false"/>
    "#;
    let meta = parse_slide_master_meta(master_xml, "ppt/slideMasters/slideMaster1.xml")
        .expect("master meta");
    assert_eq!(meta.preserve, Some(false));
    assert_eq!(meta.show_master_sp, Some(true));
    assert_eq!(meta.show_master_ph_anim, Some(false));

    let missing = parse_slide_master_meta("<p:other/>", "ppt/slideMasters/missing.xml")
        .expect("missing master meta");
    assert!(missing.preserve.is_none());
    assert!(missing.show_master_sp.is_none());
    assert!(missing.show_master_ph_anim.is_none());
}

pub(crate) fn build_empty_zip() -> SecureZipReader<Cursor<Vec<u8>>> {
    build_zip_with_entries(Vec::new())
}

pub(crate) fn build_zip_with_entries(
    entries: Vec<(&str, &str)>,
) -> SecureZipReader<Cursor<Vec<u8>>> {
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
