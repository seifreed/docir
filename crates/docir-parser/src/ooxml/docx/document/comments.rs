use super::{DocxParser, NoteKind, parse_block_until};
use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{local_name, try_attr_value, xml_error};
use docir_core::ir::{
    Comment, CommentExtension, CommentExtensionSet, CommentIdMap, CommentIdMapEntry, Endnote,
    Footnote, IRNode,
};
use docir_core::types::NodeId;
use quick_xml::Reader;
use quick_xml::events::Event;

impl DocxParser {
    /// Public API entrypoint: parse_comments.
    pub fn parse_comments(
        &mut self,
        xml: &str,
        rels: &Relationships,
    ) -> Result<Vec<NodeId>, ParseError> {
        parse_comments_like(self, xml, rels, None)
    }

    /// Public API entrypoint: parse_notes.
    pub fn parse_notes(
        &mut self,
        xml: &str,
        kind: NoteKind,
        rels: &Relationships,
    ) -> Result<Vec<NodeId>, ParseError> {
        parse_comments_like(self, xml, rels, Some(kind))
    }

    /// Public API entrypoint: parse_comments_extended.
    pub fn parse_comments_extended(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut set = CommentExtensionSet::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e))
                    if local_name(e.name().as_ref()) == b"commentExt" =>
                {
                    let comment_id = try_attr_value(&e, b"w:id", "word/commentsExtended.xml")?
                        .unwrap_or_default();
                    let entry = CommentExtension {
                        comment_id,
                        para_id: try_attr_value(&e, b"w16cid:paraId", "word/commentsExtended.xml")?,
                        parent_para_id: try_attr_value(
                            &e,
                            b"w16cid:parentParaId",
                            "word/commentsExtended.xml",
                        )?,
                        done: try_attr_value(&e, b"w:done", "word/commentsExtended.xml")?
                            .map(|v| v == "1" || v.eq_ignore_ascii_case("true")),
                    };
                    set.entries.push(entry);
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(xml_error("word/commentsExtended.xml", e));
                }
                _ => {}
            }
            buf.clear();
        }

        let id = set.id;
        self.store.insert(IRNode::CommentExtensionSet(set));
        Ok(id)
    }

    /// Public API entrypoint: parse_comments_ids.
    pub fn parse_comments_ids(&mut self, xml: &str) -> Result<NodeId, ParseError> {
        let mut map = CommentIdMap::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e))
                    if local_name(e.name().as_ref()) == b"commentId" =>
                {
                    let entry = CommentIdMapEntry {
                        comment_id: try_attr_value(&e, b"w:id", "word/commentsIds.xml")?
                            .unwrap_or_default(),
                        para_id: try_attr_value(&e, b"w16cid:paraId", "word/commentsIds.xml")?,
                        parent_para_id: try_attr_value(
                            &e,
                            b"w16cid:parentParaId",
                            "word/commentsIds.xml",
                        )?,
                    };
                    map.mappings.push(entry);
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(xml_error("word/commentsIds.xml", e));
                }
                _ => {}
            }
            buf.clear();
        }

        let id = map.id;
        self.store.insert(IRNode::CommentIdMap(map));
        Ok(id)
    }
}

fn parse_comments_like(
    parser: &mut DocxParser,
    xml: &str,
    rels: &Relationships,
    kind: Option<NoteKind>,
) -> Result<Vec<NodeId>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    reader.config_mut().check_end_names = false;
    let mut buf = Vec::new();
    let mut nodes = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"comment" => {
                    let comment_id =
                        try_attr_value(&e, b"w:id", "word/comments.xml")?.unwrap_or_default();
                    let mut comment = Comment::new(comment_id);
                    comment.author = try_attr_value(&e, b"w:author", "word/comments.xml")?;
                    comment.initials = try_attr_value(&e, b"w:initials", "word/comments.xml")?;
                    comment.parent_id = try_attr_value(&e, b"w:parentId", "word/comments.xml")?;
                    comment.para_id = try_attr_value(&e, b"w:paraId", "word/comments.xml")?;
                    if let Some(val) = try_attr_value(&e, b"w:done", "word/comments.xml")? {
                        let v = val.as_str();
                        comment.done = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                    }
                    comment.date = try_attr_value(&e, b"w:date", "word/comments.xml")?;
                    comment.content = parse_block_until(parser, &mut reader, rels, b"comment")?;
                    let id = comment.id;
                    parser.store.insert(IRNode::Comment(comment));
                    nodes.push(id);
                }
                b"footnote" => {
                    if matches!(kind, Some(NoteKind::Footnote)) {
                        let note_id =
                            try_attr_value(&e, b"w:id", "word/comments.xml")?.unwrap_or_default();
                        let mut note = Footnote::new(note_id);
                        note.note_type = try_attr_value(&e, b"w:type", "word/comments.xml")?;
                        note.content = parse_block_until(parser, &mut reader, rels, b"footnote")?;
                        let id = note.id;
                        parser.store.insert(IRNode::Footnote(note));
                        nodes.push(id);
                    }
                }
                b"endnote" => {
                    if matches!(kind, Some(NoteKind::Endnote)) {
                        let note_id =
                            try_attr_value(&e, b"w:id", "word/comments.xml")?.unwrap_or_default();
                        let mut note = Endnote::new(note_id);
                        note.note_type = try_attr_value(&e, b"w:type", "word/comments.xml")?;
                        note.content = parse_block_until(parser, &mut reader, rels, b"endnote")?;
                        let id = note.id;
                        parser.store.insert(IRNode::Endnote(note));
                        nodes.push(id);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error("word/comments.xml", e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(nodes)
}
