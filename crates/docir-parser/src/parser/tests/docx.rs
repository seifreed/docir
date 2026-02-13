use super::super::*;
use super::helpers::{create_docx_with_body, create_docx_with_relationships, create_minimal_docx};
use docir_core::ir::{DocumentMetadata, PropertyValue};

fn docx_sections_fixture() -> (
    &'static str,
    &'static str,
    &'static str,
    &'static str,
    &'static str,
) {
    let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p>
              <w:pPr>
                <w:sectPr>
                  <w:headerReference w:type="default" r:id="rIdHeader1"/>
                  <w:footerReference w:type="default" r:id="rIdFooter1"/>
                  <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
                </w:sectPr>
              </w:pPr>
              <w:r><w:t>Section1</w:t></w:r>
            </w:p>
            <w:p><w:r><w:t>Section2</w:t></w:r></w:p>
            <w:sectPr>
              <w:headerReference w:type="default" r:id="rIdHeader2"/>
            </w:sectPr>
          </w:body>
        </w:document>"#;

    let header1 = r#"
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p><w:r><w:t>Header1</w:t></w:r></w:p>
        </w:hdr>"#;
    let footer1 = r#"
        <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p><w:r><w:t>Footer1</w:t></w:r></w:p>
        </w:ftr>"#;
    let header2 = r#"
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p><w:r><w:t>Header2</w:t></w:r></w:p>
        </w:hdr>"#;

    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdHeader1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
            Target="header1.xml"/>
          <Relationship Id="rIdFooter1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
            Target="footer1.xml"/>
          <Relationship Id="rIdHeader2"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
            Target="header2.xml"/>
        </Relationships>"#;

    (body, header1, footer1, header2, rels)
}

#[test]
fn test_parse_custom_properties() {
    let xml = r#"
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
          <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="AuthorId">
            <vt:i4>123</vt:i4>
          </property>
          <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Flag">
            <vt:bool>true</vt:bool>
          </property>
          <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Title">
            <vt:lpwstr>Doc</vt:lpwstr>
          </property>
        </Properties>
        "#;
    let mut metadata = DocumentMetadata::new();
    let parser = OoxmlParser::new();
    parser.parse_custom_properties(xml, &mut metadata);
    assert_eq!(metadata.custom_properties.len(), 3);
    assert!(matches!(
        metadata.custom_properties[0].value,
        PropertyValue::Integer(123)
    ));
    assert!(matches!(
        metadata.custom_properties[1].value,
        PropertyValue::Boolean(true)
    ));
    assert!(matches!(
        metadata.custom_properties[2].value,
        PropertyValue::String(_)
    ));
}

#[test]
fn test_parse_minimal_docx() {
    let path = create_minimal_docx(true);
    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse minimal docx");
    assert!(parsed.document().is_some());
    std::fs::remove_file(path).ok();
}

#[test]
fn test_missing_main_part_fails_validation() {
    let path = create_minimal_docx(false);
    let parser = OoxmlParser::new();
    let err = parser
        .parse_file(&path)
        .expect_err("missing part should fail");
    match err {
        ParseError::MissingPart(_) => {}
        other => panic!("unexpected error: {other:?}"),
    }
    std::fs::remove_file(path).ok();
}

#[test]
fn test_docx_paragraph_and_run_properties() {
    let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr>
                <w:jc w:val="center"/>
                <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>
                <w:ind w:left="720" w:firstLine="360"/>
                <w:outlineLvl w:val="2"/>
              </w:pPr>
              <w:r>
                <w:rPr>
                  <w:b/>
                  <w:i/>
                  <w:color w:val="FF0000"/>
                  <w:highlight w:val="yellow"/>
                  <w:sz w:val="24"/>
                </w:rPr>
                <w:t>Hello</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>"#;
    let path = create_docx_with_body(body);
    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse docx");

    let mut para_props = None;
    let mut run_props = None;
    for node in parsed.store.values() {
        if let IRNode::Paragraph(p) = node {
            para_props = Some(p.properties.clone());
        } else if let IRNode::Run(r) = node {
            run_props = Some(r.properties.clone());
        }
    }

    let para_props = para_props.expect("paragraph props");
    assert!(matches!(
        para_props.alignment,
        Some(docir_core::ir::TextAlignment::Center)
    ));
    assert_eq!(
        para_props.spacing.as_ref().and_then(|s| s.before),
        Some(120)
    );
    assert_eq!(para_props.spacing.as_ref().and_then(|s| s.after), Some(240));
    assert_eq!(para_props.spacing.as_ref().and_then(|s| s.line), Some(360));
    assert_eq!(
        para_props.indentation.as_ref().and_then(|i| i.left),
        Some(720)
    );
    assert_eq!(
        para_props.indentation.as_ref().and_then(|i| i.first_line),
        Some(360)
    );
    assert_eq!(para_props.outline_level, Some(2));

    let run_props = run_props.expect("run props");
    assert_eq!(run_props.bold, Some(true));
    assert_eq!(run_props.italic, Some(true));
    assert_eq!(run_props.color.as_deref(), Some("FF0000"));
    assert_eq!(run_props.highlight.as_deref(), Some("yellow"));
    assert_eq!(run_props.font_size, Some(24));

    std::fs::remove_file(path).ok();
}

#[test]
fn test_docx_sections_with_headers_and_footers() {
    let (body, header1, footer1, header2, rels) = docx_sections_fixture();

    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/header1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
          <Override PartName="/word/footer1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
          <Override PartName="/word/header2.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
        </Types>"#;

    let path = create_docx_with_relationships(
        body,
        rels,
        content_types,
        &[
            ("word/header1.xml", header1),
            ("word/footer1.xml", footer1),
            ("word/header2.xml", header2),
        ],
    );

    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse docx");
    let doc = parsed.document().expect("doc");
    assert_eq!(doc.content.len(), 2);

    let mut section_headers = Vec::new();
    for section_id in &doc.content {
        if let Some(IRNode::Section(section)) = parsed.store.get(*section_id) {
            section_headers.push(section.headers.len());
        }
    }
    assert_eq!(section_headers, vec![1, 1]);

    std::fs::remove_file(path).ok();
}

#[test]
fn test_docx_content_controls_inline_and_block() {
    let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:sdt>
              <w:sdtPr>
                <w:alias w:val="BlockControl"/>
                <w:tag w:val="block-tag"/>
              </w:sdtPr>
              <w:sdtContent>
                <w:p><w:r><w:t>Block content</w:t></w:r></w:p>
              </w:sdtContent>
            </w:sdt>
            <w:p>
              <w:r><w:t>Before</w:t></w:r>
              <w:sdt>
                <w:sdtPr>
                  <w:alias w:val="InlineControl"/>
                  <w:tag w:val="inline-tag"/>
                </w:sdtPr>
                <w:sdtContent>
                  <w:r><w:t>Inline content</w:t></w:r>
                </w:sdtContent>
              </w:sdt>
              <w:r><w:t>After</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>"#;

    let path = create_docx_with_body(body);
    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse docx");

    let mut controls = Vec::new();
    for node in parsed.store.values() {
        if let IRNode::ContentControl(ctrl) = node {
            controls.push((ctrl.alias.clone(), ctrl.tag.clone(), ctrl.content.len()));
        }
    }
    controls.sort_by(|a, b| a.0.cmp(&b.0));

    assert_eq!(controls.len(), 2);
    assert_eq!(controls[0].0.as_deref(), Some("BlockControl"));
    assert_eq!(controls[0].1.as_deref(), Some("block-tag"));
    assert!(controls[0].2 >= 1);
    assert_eq!(controls[1].0.as_deref(), Some("InlineControl"));
    assert_eq!(controls[1].1.as_deref(), Some("inline-tag"));
    assert!(controls[1].2 >= 1);

    std::fs::remove_file(path).ok();
}

#[test]
fn test_docx_chart_shape_linkage() {
    let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p>
              <w:r>
                <w:drawing>
                  <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <c:chart r:id="rIdChart1"/>
                    </a:graphicData>
                  </a:graphic>
                </w:drawing>
              </w:r>
            </w:p>
          </w:body>
        </w:document>
        "#;

    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdChart1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="charts/chart1.xml"/>
        </Relationships>
        "#;

    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/charts/chart1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
        </Types>"#;

    let chart_xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart>
            <c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:barChart/>
          </c:chart>
        </c:chartSpace>
        "#;

    let path = create_docx_with_relationships(
        body,
        rels_xml,
        content_types,
        &[("word/charts/chart1.xml", chart_xml)],
    );

    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse docx");
    let mut shape_chart_ids = Vec::new();
    for node in parsed.store.values() {
        if let IRNode::Shape(shape) = node {
            if let Some(chart_id) = shape.chart_id {
                shape_chart_ids.push(chart_id);
            }
        }
    }
    assert_eq!(shape_chart_ids.len(), 1);
    let chart = match parsed.store.get(shape_chart_ids[0]) {
        Some(IRNode::ChartData(c)) => c,
        _ => panic!("missing chart node"),
    };
    assert_eq!(chart.title.as_deref(), Some("Sales"));

    std::fs::remove_file(path).ok();
}
