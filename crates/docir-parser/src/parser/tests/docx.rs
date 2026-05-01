use super::super::*;
use super::helpers::{create_docx_with_body, create_docx_with_relationships, create_minimal_docx};
use docir_core::ir::{DocumentMetadata, PropertyValue};
use std::fs::File;
use std::io::Write;
use zip::write::FileOptions;

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

#[test]
fn test_docx_shared_parts_collects_multiple_part_kinds() {
    let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Shared parts</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;

    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>"#;

    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="png" ContentType="image/png"/>
          <Default Extension="jpeg" ContentType="image/jpeg"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/theme/theme1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
          <Override PartName="/word/charts/chart1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
          <Override PartName="/word/diagrams/data1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>
          <Override PartName="/word/people.xml"
            ContentType="application/vnd.openxmlformats-officedocument.people+xml"/>
          <Override PartName="/word/webExtensions/webextension1.xml"
            ContentType="application/vnd.ms-office.webextension+xml"/>
          <Override PartName="/word/webExtensions/taskpanes.xml"
            ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>
          <Override PartName="/word/embeddings/ext1.xml"
            ContentType="application/vnd.ms-office.extension+xml"/>
          <Override PartName="/_xmlsignatures/sig1.xml"
            ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>
        </Types>"#;

    let theme_xml = r#"
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
          <a:themeElements>
            <a:clrScheme name="Office">
              <a:dk1><a:srgbClr val="000000"/></a:dk1>
            </a:clrScheme>
            <a:fontScheme name="Office Font">
              <a:majorFont><a:latin typeface="Calibri"/></a:majorFont>
              <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
            </a:fontScheme>
          </a:themeElements>
        </a:theme>"#;

    let chart_xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:barChart/>
          </c:chart>
        </c:chartSpace>"#;

    let smartart_xml = r#"
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <dgm:pt/>
          <dgm:cxn/>
          <dgm:relIds r:dm="rIdDm" r:lo="rIdLo"/>
        </dgm:dataModel>"#;

    let people_xml = r#"
        <w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
          <w15:person w15:id="0" w15:userId="alice@example.com" w15:displayName="Alice" w15:initials="AL"/>
        </w15:people>"#;

    let web_extension_xml = r#"
        <we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="ext-1">
          <we:storeReference store="OMEX" storeType="OMEX" id="store-1" version="1.0"/>
          <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
        </we:webextension>"#;

    let taskpanes_xml = r#"
        <wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <wetp:taskpane dockState="right" visibility="true" width="320" row="0">
            <wetp:webextensionref r:id="rIdWebExt1"/>
          </wetp:taskpane>
        </wetp:taskpanes>"#;

    let custom_xml = r#"<root xmlns="urn:test"><child/></root>"#;

    let signature_xml = r#"
        <ds:Signature xmlns:ds="http://www.w3.org/2000/09/xmldsig#" Id="sig-1">
          <ds:SignedInfo>
            <ds:SignatureMethod Algorithm="rsa-sha256"></ds:SignatureMethod>
            <ds:DigestMethod Algorithm="sha256"></ds:DigestMethod>
          </ds:SignedInfo>
          <ds:KeyInfo><ds:X509Data><ds:X509SubjectName>CN=Signer</ds:X509SubjectName></ds:X509Data></ds:KeyInfo>
        </ds:Signature>"#;

    let path = create_docx_with_relationships(
        body,
        rels_xml,
        content_types,
        &[
            ("word/theme/theme1.xml", theme_xml),
            ("word/charts/chart1.xml", chart_xml),
            ("word/diagrams/data1.xml", smartart_xml),
            ("word/people.xml", people_xml),
            ("word/webExtensions/webextension1.xml", web_extension_xml),
            ("word/webExtensions/taskpanes.xml", taskpanes_xml),
            ("customXml/item1.xml", custom_xml),
            ("_xmlsignatures/sig1.xml", signature_xml),
            ("word/embeddings/ext1.xml", "<ext/>"),
            ("word/media/image1.png", "PNG"),
            ("docProps/thumbnail.jpeg", "JPEG"),
        ],
    );

    let parser = OoxmlParser::new();
    let parsed = parser
        .parse_file(&path)
        .expect("parse docx with shared parts");
    let doc = parsed.document().expect("document");
    assert!(
        !doc.shared_parts.is_empty(),
        "shared parts should have been collected"
    );

    let mut has_relationship_graph = false;
    let mut has_theme = false;
    let mut has_media = false;
    let mut has_chart = false;
    let mut has_smartart = false;
    let mut has_people = false;
    let mut has_web_extension = false;
    let mut has_taskpane = false;
    let mut has_custom_xml = false;
    let mut has_signature = false;
    let mut has_extension_part = false;

    for id in &doc.shared_parts {
        match parsed.store.get(*id) {
            Some(IRNode::RelationshipGraph(_)) => has_relationship_graph = true,
            Some(IRNode::Theme(_)) => has_theme = true,
            Some(IRNode::MediaAsset(_)) => has_media = true,
            Some(IRNode::ChartData(_)) => has_chart = true,
            Some(IRNode::SmartArtPart(_)) => has_smartart = true,
            Some(IRNode::PeoplePart(_)) => has_people = true,
            Some(IRNode::WebExtension(_)) => has_web_extension = true,
            Some(IRNode::WebExtensionTaskpane(_)) => has_taskpane = true,
            Some(IRNode::CustomXmlPart(_)) => has_custom_xml = true,
            Some(IRNode::DigitalSignature(_)) => has_signature = true,
            Some(IRNode::ExtensionPart(_)) => has_extension_part = true,
            _ => {}
        }
    }

    assert!(has_relationship_graph);
    assert!(has_theme);
    assert!(has_media);
    assert!(has_chart);
    assert!(has_smartart);
    assert!(has_people);
    assert!(has_web_extension);
    assert!(has_taskpane);
    assert!(has_custom_xml);
    assert!(has_signature);
    assert!(has_extension_part);

    std::fs::remove_file(path).ok();
}

#[test]
fn test_invalid_word_main_part_location_fails_validation() {
    let mut path = std::env::temp_dir();
    path.push("docir_word_invalid_main_part.docx");

    let file = std::fs::File::create(&path).expect("create temp docx");
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::<()>::default();

    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>"#;
    let package_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="xl/workbook.xml"/>
        </Relationships>"#;
    let workbook_xml =
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.trim().as_bytes()).unwrap();
    zip.add_directory("_rels/", options).unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(package_rels.trim().as_bytes()).unwrap();
    zip.add_directory("xl/", options).unwrap();
    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.trim().as_bytes()).unwrap();
    zip.finish().unwrap();

    let parser = OoxmlParser::new();
    let err = parser
        .parse_file(&path)
        .expect_err("word main part under xl/ must fail");
    match err {
        ParseError::InvalidFormat(msg) => {
            assert!(msg.contains("Expected Word document under word/"));
        }
        other => panic!("unexpected error: {other:?}"),
    }

    std::fs::remove_file(path).ok();
}

#[test]
fn test_format_parser_parse_reader_and_metrics_enabled() {
    let path = create_minimal_docx(true);
    let cfg = ParserConfig {
        enable_metrics: true,
        ..ParserConfig::default()
    };
    let parser = OoxmlParser::with_config(cfg);
    let reader = File::open(&path).expect("open temp docx");
    let parsed = <OoxmlParser as crate::format::FormatParser>::parse_reader(&parser, reader)
        .expect("trait parse_reader should succeed");

    let metrics = parsed.metrics.as_ref().expect("metrics expected");
    assert!(metrics.main_parse_ms > 0 || parsed.document().is_some());
    assert!(parsed.document().is_some());

    std::fs::remove_file(path).ok();
}

#[test]
fn test_docx_optional_parts_and_fallback_paths_attach_spans() {
    let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p>
              <w:pPr>
                <w:sectPr>
                  <w:headerReference w:type="default" r:id="rIdHeader1"/>
                  <w:footerReference w:type="default" r:id="rIdFooter1"/>
                </w:sectPr>
              </w:pPr>
              <w:r><w:t>Hello</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>"#;

    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdHeader1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
            Target="header1.xml"/>
          <Relationship Id="rIdFooter1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
            Target="footer1.xml"/>
          <Relationship Id="rIdEndnotes"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
            Target="endnotes.xml"/>
          <Relationship Id="rIdSettings"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
            Target="settings.xml"/>
          <Relationship Id="rIdWebSettings"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
            Target="webSettings.xml"/>
        </Relationships>"#;

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
          <Override PartName="/word/endnotes.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
          <Override PartName="/word/settings.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
          <Override PartName="/word/webSettings.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
          <Override PartName="/word/fontTable.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
          <Override PartName="/word/comments.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
          <Override PartName="/word/stylesWithEffects.xml"
            ContentType="application/vnd.ms-word.stylesWithEffects+xml"/>
          <Override PartName="/word/commentsExtended.xml"
            ContentType="application/vnd.ms-word.commentsExtended+xml"/>
          <Override PartName="/word/commentsIds.xml"
            ContentType="application/vnd.ms-word.commentsIds+xml"/>
          <Override PartName="/word/glossary/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>
        </Types>"#;

    let path = create_docx_with_relationships(
        body,
        rels_xml,
        content_types,
        &[
            (
                "word/header1.xml",
                r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>"#,
            ),
            (
                "word/footer1.xml",
                r#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>F</w:t></w:r></w:p></w:ftr>"#,
            ),
            (
                "word/endnotes.xml",
                r#"<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="1"><w:p><w:r><w:t>E</w:t></w:r></w:p></w:endnote></w:endnotes>"#,
            ),
            (
                "word/settings.xml",
                r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/></w:settings>"#,
            ),
            (
                "word/webSettings.xml",
                r#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:optimizeForBrowser/></w:webSettings>"#,
            ),
            (
                "word/fontTable.xml",
                r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="Calibri"/></w:fonts>"#,
            ),
            (
                "word/comments.xml",
                r#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="a"><w:p><w:r><w:t>C</w:t></w:r></w:p></w:comment></w:comments>"#,
            ),
            (
                "word/stylesWithEffects.xml",
                r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#,
            ),
            (
                "word/commentsExtended.xml",
                r#"<w16cid:commentsEx xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w16cid:commentExt w:id="0" w16cid:paraId="00112233" w:done="1"/></w16cid:commentsEx>"#,
            ),
            (
                "word/commentsIds.xml",
                r#"<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w16cid:commentId w:id="0" w16cid:paraId="00112233"/></w16cid:commentsIds>"#,
            ),
            (
                "word/glossary/document.xml",
                r#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docPart><w:docPartPr><w:name w:val="g1"/><w:gallery w:val="autoText"/></w:docPartPr><w:docPartBody><w:p><w:r><w:t>Glossary</w:t></w:r></w:p></w:docPartBody></w:docPart></w:glossaryDocument>"#,
            ),
        ],
    );

    let parser = OoxmlParser::new();
    let parsed = parser
        .parse_file(&path)
        .expect("parse enriched docx should succeed");
    let doc = parsed.document().expect("document");

    assert!(
        !doc.comments.is_empty(),
        "comments fallback path should parse word/comments.xml"
    );
    assert!(
        doc.font_table.is_some(),
        "font table fallback should resolve"
    );
    assert!(
        doc.styles_with_effects.is_some(),
        "stylesWithEffects should parse"
    );
    assert!(
        doc.comments_extended.is_some(),
        "commentsExtended should parse when part exists"
    );
    assert!(
        doc.comment_id_map.is_some(),
        "commentsIds should parse when part exists"
    );
    assert!(
        !doc.endnotes.is_empty(),
        "endnotes relationship should parse"
    );

    let mut has_glossary_part = false;
    for shared_id in &doc.shared_parts {
        if let Some(IRNode::GlossaryDocument(_)) = parsed.store.get(*shared_id) {
            has_glossary_part = true;
            break;
        }
    }
    assert!(has_glossary_part, "glossary document should be linked");

    for comment_id in &doc.comments {
        let comment = match parsed.store.get(*comment_id) {
            Some(IRNode::Comment(c)) => c,
            other => panic!("expected comment node, got {other:?}"),
        };
        assert_eq!(
            comment.span.as_ref().map(|s| s.file_path.as_str()),
            Some("word/comments.xml")
        );
    }

    if let Some(font_id) = doc.font_table {
        let table = match parsed.store.get(font_id) {
            Some(IRNode::FontTable(t)) => t,
            other => panic!("expected font table node, got {other:?}"),
        };
        assert_eq!(
            table.span.as_ref().map(|s| s.file_path.as_str()),
            Some("word/fontTable.xml")
        );
    }

    for endnote_id in &doc.endnotes {
        let note = match parsed.store.get(*endnote_id) {
            Some(IRNode::Endnote(n)) => n,
            other => panic!("expected endnote node, got {other:?}"),
        };
        assert_eq!(
            note.span.as_ref().map(|s| s.file_path.as_str()),
            Some("word/endnotes.xml")
        );
    }

    std::fs::remove_file(path).ok();
}
