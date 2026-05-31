use super::super::*;
use super::helpers::{create_docx_with_relationships, create_minimal_docx};
use std::fs::File;
use std::io::Write;
use zip::write::FileOptions;

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
