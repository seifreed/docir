use super::*;
use docir_core::ir::TextAlignment;

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
    assert!(
        shape
            .related_targets
            .contains(&"word/diagrams/data1.xml".to_string())
    );
    assert!(
        shape
            .related_targets
            .contains(&"word/diagrams/layout1.xml".to_string())
    );
    assert!(
        shape
            .related_targets
            .contains(&"word/diagrams/colors1.xml".to_string())
    );
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
    assert!(
        shape
            .related_targets
            .contains(&"word/diagrams/data1.xml".to_string())
    );
    assert!(
        shape
            .related_targets
            .contains(&"word/diagrams/layout1.xml".to_string())
    );
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
