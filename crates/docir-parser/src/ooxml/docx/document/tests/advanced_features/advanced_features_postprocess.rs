use super::*;
use crate::ooxml::relationships::{rel_type, Relationship, TargetMode};
use docir_core::ir::RevisionType;
#[test]
fn test_parse_sdt_block_collects_paragraph_table_and_nested_sdt() {
    let xml = r#"
        <w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:sdtPr>
            <w:dropDownList/>
          </w:sdtPr>
          <w:sdtContent>
            <w:p><w:r><w:t>Top</w:t></w:r></w:p>
            <w:tbl>
              <w:tr>
                <w:tc><w:p><w:r><w:t>T1</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
            <w:sdt>
              <w:sdtPr><w:date/></w:sdtPr>
              <w:sdtContent>
                <w:p><w:r><w:t>Nested</w:t></w:r></w:p>
              </w:sdtContent>
            </w:sdt>
          </w:sdtContent>
        </w:sdt>
    "#;

    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let sdt_id = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:sdt" => {
                break parse_sdt(&mut parser, &mut reader, &rels, SdtMode::Block)
                    .expect("parse block sdt");
            }
            Ok(Event::Eof) => panic!("missing sdt"),
            Err(e) => panic!("xml error: {e}"),
            _ => {}
        }
        buf.clear();
    };

    let store = parser.into_store();
    let control = match store.get(sdt_id) {
        Some(docir_core::ir::IRNode::ContentControl(c)) => c,
        _ => panic!("expected top-level content control"),
    };

    assert_eq!(control.control_type.as_deref(), Some("dropDownList"));
    assert_eq!(control.content.len(), 3);
    assert!(matches!(
        store.get(control.content[0]),
        Some(docir_core::ir::IRNode::Paragraph(_))
    ));
    assert!(matches!(
        store.get(control.content[1]),
        Some(docir_core::ir::IRNode::Table(_))
    ));

    let nested = match store.get(control.content[2]) {
        Some(docir_core::ir::IRNode::ContentControl(c)) => c,
        _ => panic!("expected nested content control"),
    };
    assert_eq!(nested.control_type.as_deref(), Some("date"));
    assert_eq!(nested.content.len(), 1);
    assert!(matches!(
        store.get(nested.content[0]),
        Some(docir_core::ir::IRNode::Paragraph(_))
    ));
}

#[test]
fn test_parse_sdt_inline_control_type_variants() {
    let cases = [
        ("comboBox", "comboBox"),
        ("text", "text"),
        ("picture", "picture"),
    ];

    for (marker, expected_type) in cases {
        let xml = format!(
            r#"
            <w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:sdtPr><w:{marker}/></w:sdtPr>
              <w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent>
            </w:sdt>
            "#
        );

        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let mut reader = reader_from_str(&xml);
        let mut buf = Vec::new();
        let sdt_id = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.name().as_ref() == b"w:sdt" => {
                    break parse_sdt(&mut parser, &mut reader, &rels, SdtMode::Inline)
                        .expect("parse inline sdt");
                }
                Ok(Event::Eof) => panic!("missing sdt"),
                Err(e) => panic!("xml error: {e}"),
                _ => {}
            }
            buf.clear();
        };

        let store = parser.into_store();
        let control = match store.get(sdt_id) {
            Some(docir_core::ir::IRNode::ContentControl(c)) => c,
            _ => panic!("expected content control"),
        };
        assert_eq!(control.control_type.as_deref(), Some(expected_type));
    }
}
#[test]
fn test_parse_hyperlink_preserves_existing_fragment_and_internal_mode() {
    let mut rels = Relationships::default();
    let rel = Relationship {
        id: "rIdFrag".to_string(),
        rel_type: rel_type::HYPERLINK.to_string(),
        target: "https://example.test/path#existing".to_string(),
        target_mode: TargetMode::Internal,
    };
    rels.by_type
        .entry(rel.rel_type.clone())
        .or_default()
        .push(rel.id.clone());
    rels.by_id.insert(rel.id.clone(), rel);

    let xml = r#"
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          r:id="rIdFrag" w:anchor="ignored">
          <w:r><w:t>link</w:t></w:r>
        </w:hyperlink>
    "#;

    let mut parser = DocxParser::new();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let link_id = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:hyperlink" => {
                break parse_hyperlink(&mut parser, &mut reader, &rels, &e).expect("parse link");
            }
            Ok(Event::Eof) => panic!("missing hyperlink"),
            Err(e) => panic!("xml error: {e}"),
            _ => {}
        }
        buf.clear();
    };

    let store = parser.into_store();
    let link = match store.get(link_id) {
        Some(docir_core::ir::IRNode::Hyperlink(h)) => h,
        _ => panic!("expected hyperlink"),
    };
    assert_eq!(link.target, "https://example.test/path#existing");
    assert!(!link.is_external);
}
#[test]
fn test_parse_sdt_inline_fldsimple_without_instruction_keeps_field_unparsed() {
    let xml = r#"
        <w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:sdtContent>
            <w:fldSimple>
              <w:r><w:t>value</w:t></w:r>
            </w:fldSimple>
          </w:sdtContent>
        </w:sdt>
    "#;

    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let sdt_id = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:sdt" => {
                break parse_sdt(&mut parser, &mut reader, &rels, SdtMode::Inline)
                    .expect("parse inline sdt");
            }
            Ok(Event::Eof) => panic!("missing sdt"),
            Err(e) => panic!("xml error: {e}"),
            _ => {}
        }
        buf.clear();
    };

    let store = parser.into_store();
    let control = match store.get(sdt_id) {
        Some(docir_core::ir::IRNode::ContentControl(c)) => c,
        _ => panic!("expected content control"),
    };
    assert_eq!(control.content.len(), 1);
    let field = match store.get(control.content[0]) {
        Some(docir_core::ir::IRNode::Field(f)) => f,
        _ => panic!("expected field"),
    };
    assert!(field.instruction.is_none());
    assert!(field.instruction_parsed.is_none());
    assert_eq!(field.runs.len(), 1);
}
#[test]
fn test_parse_hyperlink_appends_anchor_for_external_relationship_target() {
    let mut rels = Relationships::default();
    let rel = Relationship {
        id: "rIdExt".to_string(),
        rel_type: rel_type::HYPERLINK.to_string(),
        target: "https://example.test/base".to_string(),
        target_mode: TargetMode::External,
    };
    rels.by_type
        .entry(rel.rel_type.clone())
        .or_default()
        .push(rel.id.clone());
    rels.by_id.insert(rel.id.clone(), rel);

    let xml = r#"
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          r:id="rIdExt" w:anchor="sec1" w:tooltip="tip">
          <w:r><w:t>link</w:t></w:r>
        </w:hyperlink>
    "#;

    let mut parser = DocxParser::new();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let link_id = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:hyperlink" => {
                break parse_hyperlink(&mut parser, &mut reader, &rels, &e).expect("parse link");
            }
            Ok(Event::Eof) => panic!("missing hyperlink"),
            Err(e) => panic!("xml error: {e}"),
            _ => {}
        }
        buf.clear();
    };

    let store = parser.into_store();
    let link = match store.get(link_id) {
        Some(docir_core::ir::IRNode::Hyperlink(h)) => h,
        _ => panic!("expected hyperlink"),
    };
    assert_eq!(link.target, "https://example.test/base#sec1");
    assert!(link.is_external);
    assert_eq!(link.relationship_id.as_deref(), Some("rIdExt"));
    assert_eq!(link.tooltip.as_deref(), Some("tip"));
    assert_eq!(link.runs.len(), 1);
}

#[test]
fn test_parse_sdt_inline_collects_end_markers_delete_revision_and_skips_missing_ids() {
    let xml = r#"
        <w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:sdtContent>
            <w:commentRangeEnd></w:commentRangeEnd>
            <w:commentRangeEnd w:id="9"></w:commentRangeEnd>
            <w:commentReference w:id="9"></w:commentReference>
            <w:bookmarkEnd></w:bookmarkEnd>
            <w:bookmarkEnd w:id="8"></w:bookmarkEnd>
            <w:del w:id="12" w:author="Bot">
              <w:r><w:t>removed</w:t></w:r>
            </w:del>
          </w:sdtContent>
        </w:sdt>
    "#;

    let mut parser = DocxParser::new();
    let rels = Relationships::default();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let sdt_id = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:sdt" => {
                break parse_sdt(&mut parser, &mut reader, &rels, SdtMode::Inline)
                    .expect("parse inline sdt");
            }
            Ok(Event::Eof) => panic!("missing sdt"),
            Err(e) => panic!("xml error: {e}"),
            _ => {}
        }
        buf.clear();
    };

    let store = parser.into_store();
    let control = match store.get(sdt_id) {
        Some(docir_core::ir::IRNode::ContentControl(c)) => c,
        _ => panic!("expected content control"),
    };
    // Missing-id entries are skipped, id-bearing entries + revision are retained.
    assert_eq!(control.content.len(), 4);

    assert!(matches!(
        store.get(control.content[0]),
        Some(docir_core::ir::IRNode::CommentRangeEnd(_))
    ));
    assert!(matches!(
        store.get(control.content[1]),
        Some(docir_core::ir::IRNode::CommentReference(_))
    ));
    assert!(matches!(
        store.get(control.content[2]),
        Some(docir_core::ir::IRNode::BookmarkEnd(_))
    ));
    let revision = match store.get(control.content[3]) {
        Some(docir_core::ir::IRNode::Revision(r)) => r,
        _ => panic!("expected revision"),
    };
    assert!(matches!(revision.change_type, RevisionType::Delete));
    assert_eq!(revision.author.as_deref(), Some("Bot"));
    assert_eq!(revision.content.len(), 1);
}

#[test]
fn test_parse_hyperlink_with_unresolved_rel_uses_anchor_target_only() {
    let rels = Relationships::default();
    let xml = r#"
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          r:id="rIdMissing" w:anchor="LocalRef" w:tooltip="fallback">
          <w:r><w:t>link</w:t></w:r>
        </w:hyperlink>
    "#;

    let mut parser = DocxParser::new();
    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let link_id = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"w:hyperlink" => {
                break parse_hyperlink(&mut parser, &mut reader, &rels, &e).expect("parse link");
            }
            Ok(Event::Eof) => panic!("missing hyperlink"),
            Err(e) => panic!("xml error: {e}"),
            _ => {}
        }
        buf.clear();
    };

    let store = parser.into_store();
    let link = match store.get(link_id) {
        Some(docir_core::ir::IRNode::Hyperlink(h)) => h,
        _ => panic!("expected hyperlink"),
    };
    assert_eq!(link.target, "#LocalRef");
    assert!(!link.is_external);
    assert!(link.relationship_id.is_none());
    assert_eq!(link.tooltip.as_deref(), Some("fallback"));
    assert_eq!(link.runs.len(), 1);
}
