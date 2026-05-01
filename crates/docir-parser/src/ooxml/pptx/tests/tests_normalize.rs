use super::*;
use crate::error::ParseError;
use docir_core::ir::{IRNode, ShapeType};

#[test]
fn test_parse_slide_returns_xml_error_for_malformed_slide_xml() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree><p:sp></p:spTree></p:cSld>
        </p:sld>
        "#;
    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();

    let err = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide-malformed.xml",
            &Relationships::default(),
            (None, None),
        )
        .expect_err("malformed slide xml should fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/slides/slide-malformed.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn test_parse_slide_transition_with_invalid_numeric_attrs_keeps_defaults() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:transition spd="slow" advClick="false" advTm="bad" dur="oops"></p:transition>
        </p:sld>
        "#;
    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();

    let slide_id = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide-transition.xml",
            &Relationships::default(),
            (None, None),
        )
        .expect("slide parse should succeed");

    let store = parser.into_store();
    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    let transition = slide.transition.as_ref().expect("transition");
    assert_eq!(transition.speed.as_deref(), Some("slow"));
    assert_eq!(transition.advance_on_click, Some(false));
    assert_eq!(transition.advance_after_ms, None);
    assert_eq!(transition.duration_ms, None);
    assert_eq!(transition.transition_type, None);
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
            (None, None),
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
            (None, None),
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
fn test_parse_pic_with_missing_relationship_keeps_picture_without_target() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree>
              <p:pic>
                <p:nvPicPr><p:cNvPr id="2" name="Pic"/></p:nvPicPr>
                <p:blipFill><a:blip r:embed="rIdMissing"/></p:blipFill>
              </p:pic>
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
            "ppt/slides/slide-missing-pic-rel.xml",
            &Relationships::default(),
            (None, None),
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
    assert_eq!(shape.shape_type, ShapeType::Picture);
    assert!(shape.relationship_id.is_none());
    assert!(shape.media_target.is_none());
}

#[test]
fn test_parse_slide_ignores_whitespace_notes_but_keeps_notes_slide_id() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
        </p:sld>
        "#;
    let mut parser = PptxParser::new();
    let mut zip = build_empty_zip();
    let notes_slide_id = docir_core::types::NodeId::new();

    let slide_id = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide1.xml",
            &Relationships::default(),
            (Some(" \n\t "), Some(notes_slide_id)),
        )
        .expect("slide");

    let store = parser.into_store();
    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert!(slide.notes.is_none(), "blank notes should be ignored");
    assert_eq!(slide.notes_slide, Some(notes_slide_id));
}

#[test]
fn test_parse_slide_timing_media_start_event_and_orphan_shape_target() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld><p:spTree/></p:cSld>
          <p:timing>
            <p:tgtEl><p:spTgt spid="999"/></p:tgtEl>
            <p:audio r:link="rIdAudio" dur="900"></p:audio>
          </p:timing>
        </p:sld>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdAudio"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
            Target="../media/audio-start.wav"/>
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
            (None, None),
        )
        .expect("slide");

    let store = parser.into_store();
    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert_eq!(slide.animations.len(), 1);
    assert_eq!(slide.animations[0].duration_ms, Some(900));
    assert_eq!(
        slide.animations[0].target.as_deref(),
        Some("ppt/media/audio-start.wav")
    );
}

#[test]
fn test_parse_slide_timing_media_missing_relation_uses_rel_id_fallback() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld><p:spTree/></p:cSld>
          <p:timing>
            <p:audio r:link="rIdMissing" dur="2500"/>
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
            (None, None),
        )
        .expect("slide");

    let store = parser.into_store();
    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert_eq!(slide.animations.len(), 1);
    assert_eq!(slide.animations[0].duration_ms, Some(2500));
    assert_eq!(slide.animations[0].target.as_deref(), Some("rIdMissing"));
}

#[test]
fn test_parse_slide_comments_with_missing_malformed_target_is_ignored() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
        </p:sld>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdC"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            Target="../../../../comments/comment1.xml"/>
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
            (None, None),
        )
        .expect("slide");

    let store = parser.into_store();
    let slide = match store.get(slide_id) {
        Some(IRNode::Slide(s)) => s,
        _ => panic!("missing slide"),
    };
    assert!(
        slide.comments.is_empty(),
        "missing resolved comments part should be ignored"
    );
}

#[test]
fn test_parse_slide_comments_malformed_xml_reports_comments_file() {
    let slide_xml = r#"
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
        </p:sld>
        "#;
    let malformed_comments_xml = r#"
        <p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cm authorId="1"><p:txBody>
        </p:cmLst>
        "#;
    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdC"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            Target="../comments/comment1.xml"/>
        </Relationships>
        "#;
    let rels = Relationships::parse(rels_xml).expect("rels");
    let mut zip =
        build_zip_with_entries(vec![("ppt/comments/comment1.xml", malformed_comments_xml)]);
    let mut parser = PptxParser::new();

    let err = parser
        .parse_slide(
            &mut zip,
            slide_xml,
            1,
            "ppt/slides/slide1.xml",
            &rels,
            (None, None),
        )
        .expect_err("malformed comments should fail");

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "ppt/comments/comment1.xml"),
        other => panic!("unexpected error: {other:?}"),
    }
}
