mod tests {
    use crate::error::ParseError;
    use crate::hwp::section::{
        note_kind_from_local, parse_hwpx_section, revision_type_from_local, HwpxNoteKind,
    };
    use docir_core::ir::{IRNode, RevisionType};
    use docir_core::visitor::IrStore;
    use std::collections::HashMap;

    #[test]
    fn note_and_revision_kind_mappings_cover_known_and_unknown_tags() {
        assert_eq!(
            note_kind_from_local(b"comment"),
            Some(HwpxNoteKind::Comment)
        );
        assert_eq!(
            note_kind_from_local(b"annotation"),
            Some(HwpxNoteKind::Comment)
        );
        assert_eq!(
            note_kind_from_local(b"footnote"),
            Some(HwpxNoteKind::Footnote)
        );
        assert_eq!(
            note_kind_from_local(b"endnote"),
            Some(HwpxNoteKind::Endnote)
        );
        assert_eq!(note_kind_from_local(b"other"), None);

        assert_eq!(revision_type_from_local(b"ins"), Some(RevisionType::Insert));
        assert_eq!(
            revision_type_from_local(b"move-from"),
            Some(RevisionType::MoveFrom)
        );
        assert_eq!(
            revision_type_from_local(b"format-change"),
            Some(RevisionType::FormatChange)
        );
        assert_eq!(revision_type_from_local(b"none"), None);
    }

    #[test]
    fn parse_hwpx_section_collects_comments_notes_revisions_and_shapes() {
        let xml = r#"
            <hp:section xmlns:hp="http://www.hancom.co.kr/hwpml" xmlns:xlink="http://www.w3.org/1999/xlink">
              <hp:p styleId="Body">
                <hp:t>Hello</hp:t>
                <hp:comment id="c1" author="Ana" date="2025-01-01">
                  <hp:p><hp:t>comment text</hp:t></hp:p>
                </hp:comment>
                <hp:commentRef ref="c1" />
                <hp:footnote id="f1"><hp:p><hp:t>foot</hp:t></hp:p></hp:footnote>
                <hp:endnote id="e1"><hp:p><hp:t>end</hp:t></hp:p></hp:endnote>
                <hp:ins author="Bob"><hp:t>inserted</hp:t></hp:ins>
                <hp:pic xlink:href="BinData/image1.png" />
              </hp:p>
              <hp:list><hp:li><hp:t>item</hp:t></hp:li></hp:list>
              <hp:table>
                <hp:row>
                  <hp:cell><hp:p><hp:t>A1</hp:t></hp:p></hp:cell>
                </hp:row>
              </hp:table>
            </hp:section>
        "#;

        let mut store = IrStore::new();
        let mut comments = Vec::new();
        let mut footnotes = Vec::new();
        let mut endnotes = Vec::new();
        let mut media_lookup = HashMap::new();
        let media = docir_core::ir::MediaAsset::new(
            "BinData/image1.png".to_string(),
            docir_core::ir::MediaType::Image,
            3,
        );
        let media_id = media.id;
        store.insert(IRNode::MediaAsset(media));
        media_lookup.insert("BinData/image1.png".to_string(), media_id);

        let content = parse_hwpx_section(
            xml,
            "Contents/section0.xml",
            &mut store,
            &mut comments,
            &mut footnotes,
            &mut endnotes,
            &media_lookup,
        )
        .expect("section parse");

        assert!(!content.is_empty());
        assert_eq!(comments.len(), 1);
        assert_eq!(footnotes.len(), 1);
        assert_eq!(endnotes.len(), 1);
        assert!(store
            .values()
            .any(|n| matches!(n, IRNode::Revision(rev) if rev.revision_id.is_none())));
        assert!(store
            .values()
            .any(|n| matches!(n, IRNode::CommentReference(_))));
        assert!(store.values().any(|n| matches!(n, IRNode::Shape(_))));
        assert!(store.values().any(|n| matches!(n, IRNode::Table(_))));
    }

    #[test]
    fn parse_hwpx_section_returns_xml_error_on_malformed_input() {
        let xml = "<hp:section><hp:p><hp:t>broken</hp:p></hp:section>";
        let mut store = IrStore::new();
        let mut comments = Vec::new();
        let mut footnotes = Vec::new();
        let mut endnotes = Vec::new();
        let media_lookup = HashMap::new();

        let err = parse_hwpx_section(
            xml,
            "Contents/section0.xml",
            &mut store,
            &mut comments,
            &mut footnotes,
            &mut endnotes,
            &media_lookup,
        )
        .expect_err("malformed xml must fail");

        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "Contents/section0.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_hwpx_section_generates_note_ids_and_nested_list_numbering() {
        let xml = r#"
            <hp:section xmlns:hp="http://www.hancom.co.kr/hwpml">
              <hp:note writer="Writer A" created="2026-02-01">
                <hp:t>auto comment</hp:t>
              </hp:note>
              <hp:footnote>
                <hp:t>auto footnote</hp:t>
              </hp:footnote>
              <hp:endnote>
                <hp:t>auto endnote</hp:t>
              </hp:endnote>
              <hp:commentRef />
              <hp:list>
                <hp:li>
                  <hp:p><hp:t>outer</hp:t></hp:p>
                  <hp:list>
                    <hp:li><hp:p><hp:t>inner</hp:t></hp:p></hp:li>
                  </hp:list>
                </hp:li>
              </hp:list>
            </hp:section>
        "#;

        let mut store = IrStore::new();
        let mut comments = Vec::new();
        let mut footnotes = Vec::new();
        let mut endnotes = Vec::new();
        let media_lookup = HashMap::new();

        let content = parse_hwpx_section(
            xml,
            "Contents/section1.xml",
            &mut store,
            &mut comments,
            &mut footnotes,
            &mut endnotes,
            &media_lookup,
        )
        .expect("section parse");

        assert!(!content.is_empty());
        assert_eq!(comments.len(), 1);
        assert_eq!(footnotes.len(), 1);
        assert_eq!(endnotes.len(), 1);
        assert!(!store
            .values()
            .any(|n| matches!(n, IRNode::CommentReference(_))));

        let Some(IRNode::Comment(comment)) = store.get(comments[0]) else {
            panic!("expected comment");
        };
        assert_eq!(comment.comment_id, "hwpx-comment-1");
        assert_eq!(comment.author.as_deref(), Some("Writer A"));
        assert_eq!(comment.date.as_deref(), Some("2026-02-01"));

        let Some(IRNode::Footnote(footnote)) = store.get(footnotes[0]) else {
            panic!("expected footnote");
        };
        assert_eq!(footnote.footnote_id, "hwpx-footnote-1");

        let Some(IRNode::Endnote(endnote)) = store.get(endnotes[0]) else {
            panic!("expected endnote");
        };
        assert_eq!(endnote.endnote_id, "hwpx-endnote-1");

        let mut seen_level0 = false;
        let mut seen_level1 = false;
        for node in store.values() {
            if let IRNode::Paragraph(paragraph) = node {
                if let Some(numbering) = &paragraph.properties.numbering {
                    if numbering.level == 0 {
                        seen_level0 = true;
                    }
                    if numbering.level == 1 {
                        seen_level1 = true;
                    }
                }
            }
        }
        assert!(seen_level0, "expected top-level list numbering");
        assert!(seen_level1, "expected nested list numbering");
    }
}
