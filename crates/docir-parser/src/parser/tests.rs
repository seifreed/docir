use super::*;
use docir_core::types::NodeType;
use docir_core::IrSummary;
use std::fs::File;
use std::io::Write;
use std::time::{SystemTime, UNIX_EPOCH};
use zip::write::FileOptions;

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
fn test_pptx_animation_media_asset_link() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld><p:spTree/></p:cSld>
          <p:timing>
            <p:audio r:link="rIdAudio" dur="5000"/>
          </p:timing>
        </p:sld>
        "#;
    let slide_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdAudio"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
            Target="../media/audio1.wav"/>
        </Relationships>
        "#;

    let path = create_pptx_with_media(slide_xml, slide_rels);
    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse pptx");
    let doc = parsed.document().expect("doc");
    let slide_id = doc.content[0];
    let slide = match parsed.store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert_eq!(slide.animations.len(), 1);
    let anim = &slide.animations[0];
    let media_id = anim.media_asset.expect("media asset link");
    let media = match parsed.store.get(media_id) {
        Some(IRNode::MediaAsset(m)) => m,
        _ => panic!("missing media asset"),
    };
    assert_eq!(media.path, "ppt/media/audio1.wav");
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

fn create_minimal_docx(include_document: bool) -> std::path::PathBuf {
    let mut path = std::env::temp_dir();
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    path.push(format!("docir_minimal_{nanos}.docx"));

    let file = File::create(&path).expect("create temp docx");
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::<()>::default();

    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>"#;

    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;

    let document = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.trim().as_bytes()).unwrap();
    zip.add_directory("_rels/", options).unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(rels.trim().as_bytes()).unwrap();
    zip.add_directory("word/", options).unwrap();
    if include_document {
        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(document.trim().as_bytes()).unwrap();
    }
    zip.finish().unwrap();
    path
}

fn create_docx_with_body(body_xml: &str) -> std::path::PathBuf {
    let mut path = std::env::temp_dir();
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    path.push(format!("docir_body_{nanos}.docx"));

    let file = File::create(&path).expect("create temp docx");
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::<()>::default();

    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>"#;

    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.trim().as_bytes()).unwrap();
    zip.add_directory("_rels/", options).unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(rels.trim().as_bytes()).unwrap();
    zip.add_directory("word/", options).unwrap();
    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(body_xml.trim().as_bytes()).unwrap();
    zip.finish().unwrap();
    path
}

fn create_docx_with_relationships(
    body_xml: &str,
    rels_xml: &str,
    content_types: &str,
    extra_files: &[(&str, &str)],
) -> std::path::PathBuf {
    let mut path = std::env::temp_dir();
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    path.push(format!("docir_docx_{nanos}.docx"));

    let file = File::create(&path).expect("create temp docx");
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::<()>::default();

    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.trim().as_bytes()).unwrap();
    zip.add_directory("_rels/", options).unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(rels.trim().as_bytes()).unwrap();
    zip.add_directory("word/", options).unwrap();
    zip.add_directory("word/_rels/", options).unwrap();
    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    zip.write_all(rels_xml.trim().as_bytes()).unwrap();
    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(body_xml.trim().as_bytes()).unwrap();

    for (path, xml) in extra_files {
        zip.start_file(*path, options).unwrap();
        zip.write_all(xml.trim().as_bytes()).unwrap();
    }

    zip.finish().unwrap();
    path
}

fn build_odf_zip_custom(
    mimetype: &str,
    content_xml: &str,
    manifest_xml: &str,
    extra_files: &[(&str, &[u8])],
) -> Vec<u8> {
    let mut buffer = Vec::new();
    let cursor = std::io::Cursor::new(&mut buffer);
    let mut zip = zip::ZipWriter::new(cursor);
    let stored = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();

    zip.start_file("META-INF/manifest.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(manifest_xml.as_bytes()).unwrap();

    zip.start_file("content.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(content_xml.as_bytes()).unwrap();

    zip.start_file("meta.xml", FileOptions::<()>::default())
        .unwrap();
    zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
  <office:meta>
    <dc:title>Parity</dc:title>
  </office:meta>
</office:document-meta>
"#,
        )
        .unwrap();

    for (path, bytes) in extra_files {
        zip.start_file(*path, FileOptions::<()>::default()).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap();
    buffer
}

#[test]
fn test_parity_docx_odt_core_nodes() {
    let docx_body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
            <w:tbl>
              <w:tr>
                <w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
            <w:p>
              <w:hyperlink r:id="rId3"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
            </w:p>
            <w:p>
              <w:commentRangeStart w:id="0"/>
              <w:r><w:t>Commented</w:t></w:r>
              <w:commentRangeEnd w:id="0"/>
            </w:p>
            <w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>
            <w:p>
              <w:ins w:author="A" w:date="2024-01-01T00:00:00Z">
                <w:r><w:t>Inserted</w:t></w:r>
              </w:ins>
            </w:p>
          </w:body>
        </w:document>
        "#;

    let docx_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId3"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            Target="https://example.com" TargetMode="External"/>
          <Relationship Id="rId4"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            Target="comments.xml"/>
          <Relationship Id="rId5"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
            Target="footnotes.xml"/>
        </Relationships>
        "#;

    let docx_content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/comments.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
          <Override PartName="/word/footnotes.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
        </Types>
        "#;

    let comments_xml = r#"
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:comment w:id="0" w:author="Bob" w:date="2024-01-01T00:00:00Z">
            <w:p><w:r><w:t>Comment text</w:t></w:r></w:p>
          </w:comment>
        </w:comments>
        "#;

    let footnotes_xml = r#"
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:footnote w:id="1">
            <w:p><w:r><w:t>Footnote</w:t></w:r></w:p>
          </w:footnote>
        </w:footnotes>
        "#;

    let docx_path = create_docx_with_relationships(
        docx_body,
        docx_rels,
        docx_content_types,
        &[
            ("word/comments.xml", comments_xml),
            ("word/footnotes.xml", footnotes_xml),
        ],
    );

    let odt_content = r#"
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p>Hello</text:p>
      <table:table>
        <table:table-row>
          <table:table-cell><text:p>Cell</text:p></table:table-cell>
        </table:table-row>
      </table:table>
      <text:p><text:a xlink:href="https://example.com">Link</text:a></text:p>
      <office:annotation>
        <text:p>Comment text</text:p>
      </office:annotation>
      <text:note text:note-class="footnote">
        <text:note-body><text:p>Footnote</text:p></text:note-body>
      </text:note>
      <text:tracked-changes>
        <text:changed-region>
          <text:insertion>
            <text:p>Inserted</text:p>
          </text:insertion>
        </text:changed-region>
      </text:tracked-changes>
    </office:text>
  </office:body>
</office:document-content>
        "#;

    let odt_manifest = r#"
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
        "#;

    let odt_zip = build_odf_zip_custom(
        "application/vnd.oasis.opendocument.text",
        odt_content,
        odt_manifest,
        &[],
    );

    let parser = DocumentParser::new();
    let docx = parser.parse_file(&docx_path).expect("docx parse");
    let odt = parser
        .parse_reader(std::io::Cursor::new(odt_zip))
        .expect("odt parse");

    let docx_summary = IrSummary::from_store(&docx.store);
    let odt_summary = IrSummary::from_store(&odt.store);

    assert!(docx_summary.run_texts.contains("Hello"));
    assert!(odt_summary.run_texts.contains("Hello"));
    assert!(docx_summary.run_texts.contains("Cell"));
    assert!(odt_summary.run_texts.contains("Cell"));

    assert_eq!(
        docx_summary.count(NodeType::Table),
        odt_summary.count(NodeType::Table)
    );
    assert_eq!(
        docx_summary.count(NodeType::Comment),
        odt_summary.count(NodeType::Comment)
    );
    assert_eq!(
        docx_summary.count(NodeType::Footnote),
        odt_summary.count(NodeType::Footnote)
    );
    assert_eq!(
        docx_summary.count(NodeType::Revision),
        odt_summary.count(NodeType::Revision)
    );
    assert_eq!(
        docx_summary.count(NodeType::ExternalReference),
        odt_summary.count(NodeType::ExternalReference)
    );

    std::fs::remove_file(docx_path).ok();
}

fn create_pptx_with_media(slide_xml: &str, slide_rels: &str) -> std::path::PathBuf {
    let mut path = std::env::temp_dir();
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    path.push(format!("docir_pptx_{nanos}.pptx"));

    let file = File::create(&path).expect("create temp pptx");
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::<()>::default();

    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="wav" ContentType="audio/wav"/>
          <Override PartName="/ppt/presentation.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
          <Override PartName="/ppt/slides/slide1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
        </Types>"#;

    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="ppt/presentation.xml"/>
        </Relationships>"#;

    let presentation = r#"
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId r:id="rId1"/>
          </p:sldIdLst>
        </p:presentation>"#;

    let pres_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
            Target="slides/slide1.xml"/>
        </Relationships>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.trim().as_bytes()).unwrap();
    zip.add_directory("_rels/", options).unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(rels.trim().as_bytes()).unwrap();
    zip.add_directory("ppt/", options).unwrap();
    zip.add_directory("ppt/_rels/", options).unwrap();
    zip.start_file("ppt/presentation.xml", options).unwrap();
    zip.write_all(presentation.trim().as_bytes()).unwrap();
    zip.start_file("ppt/_rels/presentation.xml.rels", options)
        .unwrap();
    zip.write_all(pres_rels.trim().as_bytes()).unwrap();
    zip.add_directory("ppt/slides/", options).unwrap();
    zip.add_directory("ppt/slides/_rels/", options).unwrap();
    zip.start_file("ppt/slides/slide1.xml", options).unwrap();
    zip.write_all(slide_xml.trim().as_bytes()).unwrap();
    zip.start_file("ppt/slides/_rels/slide1.xml.rels", options)
        .unwrap();
    zip.write_all(slide_rels.trim().as_bytes()).unwrap();
    zip.add_directory("ppt/media/", options).unwrap();
    zip.start_file("ppt/media/audio1.wav", options).unwrap();
    zip.write_all(b"RIFFDATA").unwrap();
    zip.finish().unwrap();

    path
}

#[test]
fn test_content_type_inventory_diagnostics() {
    let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;
    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/extra.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        </Types>"#;
    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;
    let path =
        create_docx_with_relationships(body, rels, content_types, &[("word/extra.xml", "<root/>")]);

    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse docx");
    let doc = parsed.document().expect("doc");
    assert!(!doc.diagnostics.is_empty());
    let mut found = false;
    for diag_id in &doc.diagnostics {
        if let Some(IRNode::Diagnostics(diag)) = parsed.store.get(*diag_id) {
            if diag
                .entries
                .iter()
                .any(|e| e.code == "CONTENT_TYPE_INVENTORY" && e.message.contains("word/extra.xml"))
            {
                found = true;
            }
        }
    }
    assert!(found);
    std::fs::remove_file(path).ok();
}

#[test]
fn test_content_type_unknown_diagnostics() {
    let body = r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;
    let content_types = r#"
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>"#;
    let rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="word/document.xml"/>
        </Relationships>"#;
    let path =
        create_docx_with_relationships(body, rels, content_types, &[("word/unknown.bin", "BIN")]);

    let parser = OoxmlParser::new();
    let parsed = parser.parse_file(&path).expect("parse docx");
    let mut found = false;
    for node in parsed.store.values() {
        if let IRNode::Diagnostics(diag) = node {
            if diag
                .entries
                .iter()
                .any(|e| e.code == "CONTENT_TYPE_UNKNOWN" && e.message.contains("word/unknown.bin"))
            {
                found = true;
            }
        }
    }
    assert!(found);
    std::fs::remove_file(path).ok();
}

#[test]
fn test_document_parser_enforces_max_input_size() {
    let mut config = ParserConfig::default();
    config.max_input_size = 32;
    let parser = DocumentParser::with_config(config);
    let data = vec![b'A'; 128];
    let err = parser
        .parse_reader(std::io::Cursor::new(data))
        .expect_err("expected size limit error");
    assert!(matches!(err, ParseError::ResourceLimit(_)));
}
