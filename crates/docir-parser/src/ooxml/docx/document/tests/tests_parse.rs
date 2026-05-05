use super::*;
use docir_core::ir::{
    LineNumberRestart, PageOrientation, RevisionType, SectionType, TextAlignment, UnderlineStyle,
};

#[test]
fn test_parse_glossary_document() {
    let xml = r#"
        <w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:docParts>
            <w:docPart>
              <w:docPartPr>
                <w:name w:val="BlockOne"/>
                <w:gallery w:val="QuickParts"/>
              </w:docPartPr>
              <w:docPartBody>
                <w:p>
                  <w:r><w:t>Hello</w:t></w:r>
                </w:p>
              </w:docPartBody>
            </w:docPart>
          </w:docParts>
        </w:glossaryDocument>
        "#;

    let mut parser = DocxParser::new();
    let glossary_id = parser
        .parse_glossary_document(xml, &Relationships::default())
        .expect("parse glossary");
    let store = parser.into_store();
    let glossary = match store.get(glossary_id) {
        Some(docir_core::ir::IRNode::GlossaryDocument(doc)) => doc,
        _ => panic!("missing glossary"),
    };
    assert_eq!(glossary.entries.len(), 1);
    let entry_id = glossary.entries[0];
    let entry = match store.get(entry_id) {
        Some(docir_core::ir::IRNode::GlossaryEntry(e)) => e,
        _ => panic!("missing glossary entry"),
    };
    assert_eq!(entry.name.as_deref(), Some("BlockOne"));
    assert_eq!(entry.gallery.as_deref(), Some("QuickParts"));
    assert_eq!(entry.content.len(), 1);
}
#[test]
fn test_parse_section_properties_extended() {
    let xml = r#"
        <w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pgSz w:w="12240" w:h="15840" w:orient="landscape"/>
          <w:pgMar w:top="720" w:bottom="720" w:left="720" w:right="720" w:gutter="180"/>
          <w:cols w:num="2" w:space="720" w:sep="1"/>
          <w:type w:val="continuous"/>
          <w:titlePg/>
          <w:pgNumType w:start="3" w:fmt="upperRoman"/>
          <w:lnNumType w:start="1" w:countBy="2" w:distance="240" w:restart="newPage"/>
          <w:pgBorders>
            <w:top w:val="single" w:sz="8" w:color="FF0000"/>
          </w:pgBorders>
          <w:textDirection w:val="tbRl"/>
        </w:sectPr>
        "#;

    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:sectPr" => break,
            Ok(Event::Eof) => panic!("no sectPr"),
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }

    let section = apply_section_refs(&mut reader, None).expect("section");
    let props = section.properties;
    assert_eq!(props.page_width, Some(12240));
    assert_eq!(props.page_height, Some(15840));
    assert_eq!(props.orientation, Some(PageOrientation::Landscape));
    assert_eq!(props.columns, Some(2));
    assert_eq!(props.column_spacing, Some(720));
    assert_eq!(props.column_separator, Some(true));
    assert_eq!(props.section_type, Some(SectionType::Continuous));
    assert_eq!(props.title_page, Some(true));
    assert_eq!(props.text_direction.as_deref(), Some("tbRl"));
    assert_eq!(props.margins.as_ref().and_then(|m| m.gutter), Some(180));
    assert_eq!(props.page_numbering.as_ref().and_then(|n| n.start), Some(3));
    assert_eq!(
        props
            .page_numbering
            .as_ref()
            .and_then(|n| n.format.as_deref()),
        Some("upperRoman")
    );
    assert_eq!(
        props.line_numbering.as_ref().and_then(|n| n.count_by),
        Some(2)
    );
    assert_eq!(
        props.line_numbering.as_ref().and_then(|n| n.restart),
        Some(LineNumberRestart::NewPage)
    );
    assert!(props
        .page_borders
        .as_ref()
        .and_then(|b| b.top.as_ref())
        .is_some());
}
#[test]
fn test_parse_styles_with_table_props() {
    let xml = r#"
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:styleId="MyStyle">
            <w:name w:val="My Style"/>
            <w:rPr>
              <w:b/>
              <w:u w:val="single"/>
              <w:color w:val="FF0000"/>
            </w:rPr>
            <w:pPr>
              <w:jc w:val="center"/>
              <w:spacing w:before="120"/>
            </w:pPr>
            <w:tblPr>
              <w:tblW w:w="5000" w:type="dxa"/>
            </w:tblPr>
          </w:style>
        </w:styles>
        "#;

    let mut parser = DocxParser::new();
    let styles_id = parser.parse_styles(xml).expect("styles");
    let store = parser.into_store();
    let styles = match store.get(styles_id) {
        Some(docir_core::ir::IRNode::StyleSet(s)) => s,
        _ => panic!("missing styles"),
    };
    let style = &styles.styles[0];
    assert_eq!(style.name.as_deref(), Some("My Style"));
    let run_props = style.run_props.as_ref().expect("run props");
    assert_eq!(run_props.bold, Some(true));
    assert_eq!(run_props.underline, Some(UnderlineStyle::Single));
    assert_eq!(run_props.color.as_deref(), Some("FF0000"));
    let para_props = style.paragraph_props.as_ref().expect("para props");
    assert_eq!(para_props.alignment, Some(TextAlignment::Center));
    assert_eq!(
        para_props.spacing.as_ref().and_then(|s| s.before),
        Some(120)
    );
    assert!(style
        .table_props
        .as_ref()
        .and_then(|t| t.width.as_ref())
        .is_some());
}
#[test]
fn test_parse_font_table_collects_font_entries() {
    let xml = r#"
        <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:font w:name="Calibri">
            <w:altName w:val="Body"/>
            <w:charset w:val="204"/>
            <w:family w:val="swiss"/>
            <w:panose1 w:val="020F0502020204030204"/>
          </w:font>
          <w:font w:name="NoDetails"></w:font>
        </w:fonts>
        "#;

    let mut parser = DocxParser::new();
    let font_table_id = parser.parse_font_table(xml).expect("font table");
    let store = parser.into_store();
    let table = match store.get(font_table_id) {
        Some(docir_core::ir::IRNode::FontTable(t)) => t,
        _ => panic!("missing font table"),
    };

    assert_eq!(table.fonts.len(), 2);
    assert_eq!(table.fonts[0].name, "Calibri");
    assert_eq!(table.fonts[0].alt_name.as_deref(), Some("Body"));
    assert_eq!(table.fonts[0].charset, Some(204));
    assert_eq!(table.fonts[0].family.as_deref(), Some("swiss"));
    assert_eq!(
        table.fonts[0].panose.as_deref(),
        Some("020F0502020204030204")
    );

    assert_eq!(table.fonts[1].name, "NoDetails");
    assert_eq!(table.fonts[1].alt_name, None);
    assert_eq!(table.fonts[1].charset, None);
    assert_eq!(table.fonts[1].family, None);
    assert_eq!(table.fonts[1].panose, None);
}
#[test]
fn test_parse_font_table_returns_xml_error_on_malformed_input() {
    let xml = r#"
        <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:font w:name="Broken">
        </w:fonts>
        "#;
    let mut parser = DocxParser::new();
    let err = parser
        .parse_font_table(xml)
        .expect_err("expected malformed fontTable to fail");
    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "word/fontTable.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_numbering_level_props() {
    let xml = r#"
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="1">
            <w:lvl w:ilvl="0">
              <w:start w:val="1"/>
              <w:numFmt w:val="decimal"/>
              <w:lvlText w:val="%1."/>
              <w:lvlJc w:val="right"/>
              <w:suff w:val="space"/>
              <w:pPr>
                <w:spacing w:after="200"/>
              </w:pPr>
              <w:rPr>
                <w:i/>
              </w:rPr>
            </w:lvl>
          </w:abstractNum>
        </w:numbering>
        "#;

    let mut parser = DocxParser::new();
    let numbering_id = parser.parse_numbering(xml).expect("numbering");
    let store = parser.into_store();
    let numbering = match store.get(numbering_id) {
        Some(docir_core::ir::IRNode::NumberingSet(n)) => n,
        _ => panic!("missing numbering"),
    };
    let level = &numbering.abstract_nums[0].levels[0];
    assert_eq!(level.alignment, Some(TextAlignment::Right));
    assert_eq!(level.suffix.as_deref(), Some("space"));
    assert_eq!(
        level
            .paragraph_props
            .as_ref()
            .and_then(|p| p.spacing.as_ref())
            .and_then(|s| s.after),
        Some(200)
    );
    assert_eq!(level.run_props.as_ref().and_then(|r| r.italic), Some(true));
}

#[test]
fn test_parse_drawing_smartart_targets() {
    let xml = r#"
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:drawing>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:relIds r:dm="rId1" r:lo="rId2" r:cs="rId3"/>
              </a:graphicData>
            </a:graphic>
          </w:drawing>
        </w:r>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
            Target="diagrams/data1.xml"/>
          <Relationship Id="rId2"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
            Target="diagrams/layout1.xml"/>
          <Relationship Id="rId3"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"
            Target="diagrams/colors1.xml"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");

    let mut reader = reader_from_str(xml);
    let mut parser = DocxParser::new();

    let mut buf = Vec::new();
    let mut run = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                run = Some(parse_run(&mut parser, &mut reader, &rels).expect("run"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }

    let run = run.expect("run parsed");
    assert_eq!(run.embedded.len(), 1);
    let store = parser.into_store();
    let shape = match store.get(run.embedded[0]) {
        Some(docir_core::ir::IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(shape.related_targets.len(), 3);
    assert!(shape
        .related_targets
        .contains(&"word/diagrams/data1.xml".to_string()));
    assert!(shape
        .related_targets
        .contains(&"word/diagrams/layout1.xml".to_string()));
    assert!(shape
        .related_targets
        .contains(&"word/diagrams/colors1.xml".to_string()));
}

#[test]
fn test_parse_drawing_normalizes_targets() {
    let xml = r#"
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:drawing>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <a:blip r:embed="rIdImg"/>
                <dgm:relIds r:dm="rId1" r:lo="rId2"/>
              </a:graphicData>
            </a:graphic>
          </w:drawing>
        </w:r>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image1.png"/>
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
            Target="../diagrams/data1.xml"/>
          <Relationship Id="rId2"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
            Target="./diagrams/layout1.xml"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");

    let mut reader = reader_from_str(xml);
    let mut parser = DocxParser::new();

    let mut buf = Vec::new();
    let mut run = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                run = Some(parse_run(&mut parser, &mut reader, &rels).expect("run"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }
    let run = run.expect("run parse");
    let store = parser.into_store();
    let shape = match store.get(run.embedded[0]) {
        Some(docir_core::ir::IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(
        shape.media_target.as_deref(),
        Some("word/diagrams/data1.xml")
    );
    assert!(shape
        .related_targets
        .contains(&"word/diagrams/data1.xml".to_string()));
    assert!(shape
        .related_targets
        .contains(&"word/diagrams/layout1.xml".to_string()));
}

#[test]
fn test_parse_drawing_text_and_hyperlink() {
    let xml = r#"
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:drawing>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <a:prstGeom prst="rect"/>
                <a:blip rel:embed="rIdImg"/>
                <a:txBody>
                  <a:p>
                    <a:r>
                      <a:rPr b="1"/>
                      <a:t>Hello</a:t>
                    </a:r>
                  </a:p>
                </a:txBody>
                <a:hlinkClick rel:id="rIdLink"/>
              </a:graphicData>
            </a:graphic>
          </w:drawing>
        </w:r>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="media/image1.png"/>
          <Relationship Id="rIdLink"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            Target="https://example.com"
            TargetMode="External"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");

    let mut reader = reader_from_str(xml);
    let mut parser = DocxParser::new();

    let mut buf = Vec::new();
    let mut run = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                run = Some(parse_run(&mut parser, &mut reader, &rels).expect("run"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }
    let run = run.expect("run parse");
    let store = parser.into_store();
    let shape = match store.get(run.embedded[0]) {
        Some(docir_core::ir::IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(shape.shape_type, docir_core::ir::ShapeType::Rectangle);
    assert_eq!(shape.hyperlink.as_deref(), Some("https://example.com"));
    let text = shape.text.as_ref().expect("shape text");
    assert_eq!(text.paragraphs.len(), 1);
    assert_eq!(text.paragraphs[0].runs[0].text, "Hello");
}

#[test]
fn test_parse_drawing_applies_position_offsets_and_text_style_details() {
    let xml = r#"
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:drawing>
            <wp:extent cx="1234" cy="5678"/>
            <wp:posOffset>100</wp:posOffset>
            <wp:posOffset>200</wp:posOffset>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rIdChart"></c:chart>
                <a:txBody>
                  <a:p>
                    <a:pPr algn="dist"></a:pPr>
                    <a:r>
                      <a:rPr b="1" i="1" sz="1800"></a:rPr>
                      <a:latin typeface="Calibri"></a:latin>
                      <a:t>One</a:t>
                    </a:r>
                    <a:br></a:br>
                    <a:r><a:t>Two</a:t></a:r>
                  </a:p>
                </a:txBody>
              </a:graphicData>
            </a:graphic>
          </w:drawing>
        </w:r>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdChart"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="charts/chart1.xml"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");

    let mut reader = reader_from_str(xml);
    let mut parser = DocxParser::new();
    let mut run = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                run = Some(parse_run(&mut parser, &mut reader, &rels).expect("run"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }

    let run = run.expect("run parse");
    assert_eq!(run.embedded.len(), 1);
    let store = parser.into_store();
    let shape = match store.get(run.embedded[0]) {
        Some(docir_core::ir::IRNode::Shape(s)) => s,
        _ => panic!("missing shape"),
    };
    assert_eq!(shape.shape_type, docir_core::ir::ShapeType::Chart);
    assert_eq!(shape.transform.width, 1234);
    assert_eq!(shape.transform.height, 5678);
    assert_eq!(shape.transform.x, 100);
    assert_eq!(shape.transform.y, 200);
    let text = shape.text.as_ref().expect("shape text");
    assert_eq!(text.paragraphs.len(), 1);
    assert_eq!(
        text.paragraphs[0].alignment,
        Some(TextAlignment::Distribute)
    );
    assert_eq!(
        text.paragraphs[0].runs[0].font_family.as_deref(),
        Some("Calibri")
    );
    assert_eq!(text.paragraphs[0].runs[0].font_size, Some(1800));
    assert_eq!(text.paragraphs[0].runs[0].bold, Some(true));
    assert_eq!(text.paragraphs[0].runs[0].italic, Some(true));
    assert_eq!(text.paragraphs[0].runs[1].text, "\n");
}

#[test]
fn test_parse_drawing_with_missing_relationship_is_ignored() {
    let xml = r#"
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:drawing>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <a:blip r:embed="rIdMissing"></a:blip>
              </a:graphicData>
            </a:graphic>
          </w:drawing>
        </w:r>
        "#;

    let mut reader = reader_from_str(xml);
    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut run = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:r" => {
                run = Some(parse_run(&mut parser, &mut reader, &rels).expect("run"));
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }
    let run = run.expect("run parse");
    assert!(
        run.embedded.is_empty(),
        "missing rel should skip drawing shape"
    );
}

#[test]
fn test_parse_revisions_move_and_format() {
    let xml = r#"
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:moveFrom w:author="Alice">
            <w:r><w:t>Old</w:t></w:r>
          </w:moveFrom>
          <w:moveTo w:author="Bob">
            <w:r><w:t>New</w:t></w:r>
          </w:moveTo>
          <w:rPrChange w:author="Carol">
            <w:rPr><w:b/></w:rPr>
          </w:rPrChange>
        </w:p>
        "#;

    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);

    let mut para = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:p" => {
                let parsed = parse_paragraph(&mut parser, &mut reader, &rels, None).expect("para");
                para = Some(parsed);
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml error: {}", e),
            _ => {}
        }
        buf.clear();
    }

    let para = para.expect("para parsed");
    let store = parser.into_store();
    let mut types = Vec::new();
    let para_node = match store.get(para.id) {
        Some(docir_core::ir::IRNode::Paragraph(p)) => p,
        _ => panic!("missing paragraph"),
    };
    for id in &para_node.runs {
        if let Some(docir_core::ir::IRNode::Revision(rev)) = store.get(*id) {
            types.push(rev.change_type);
        }
    }
    types.sort_by_key(|t| format!("{t:?}"));
    assert!(types.contains(&RevisionType::MoveFrom));
    assert!(types.contains(&RevisionType::MoveTo));
    assert!(types.contains(&RevisionType::FormatChange));
}

#[test]
fn test_parse_comments_and_notes_metadata() {
    let comments_xml = r#"
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:comment w:id="1" w:author="Alice" w:date="2020-01-01T00:00:00Z"
                     w:initials="AL" w:parentId="0" w:paraId="ABC" w:done="1">
            <w:p><w:r><w:t>Note</w:t></w:r></w:p>
          </w:comment>
        </w:comments>
        "#;
    let notes_xml = r#"
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:footnote w:id="2" w:type="separator">
            <w:p><w:r><w:t>---</w:t></w:r></w:p>
          </w:footnote>
        </w:footnotes>
        "#;
    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let comment_ids = parser
        .parse_comments(comments_xml, &rels)
        .expect("comments");
    let note_ids = parser
        .parse_notes(notes_xml, NoteKind::Footnote, &rels)
        .expect("notes");
    let store = parser.into_store();

    let comment = match store.get(comment_ids[0]) {
        Some(docir_core::ir::IRNode::Comment(c)) => c,
        _ => panic!("missing comment"),
    };
    assert_eq!(comment.author.as_deref(), Some("Alice"));
    assert_eq!(comment.initials.as_deref(), Some("AL"));
    assert_eq!(comment.parent_id.as_deref(), Some("0"));
    assert_eq!(comment.para_id.as_deref(), Some("ABC"));
    assert_eq!(comment.done, Some(true));

    let footnote = match store.get(note_ids[0]) {
        Some(docir_core::ir::IRNode::Footnote(n)) => n,
        _ => panic!("missing footnote"),
    };
    assert_eq!(footnote.note_type.as_deref(), Some("separator"));
}
