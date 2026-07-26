use super::HeaderFooterSpec;
use super::{
    DocxAnnotationParts, DocxWordParts, IRNode, NodeId, OoxmlParser, PackageReader, ParseError,
    Relationships, rel_type,
};
use crate::ooxml::docx::DocxParser;
use crate::ooxml::part_utils::read_relationships_optional;
use crate::ooxml::part_utils::{read_xml_part, read_xml_part_by_rel};
use docir_core::types::SourceSpan;
use docir_core::visitor::IrStore;
use std::collections::HashMap;

type DocxStylePartIds = (Option<NodeId>, Option<NodeId>, Option<NodeId>);
type DocxSettingsPartIds = (Option<NodeId>, Option<NodeId>, Option<NodeId>);

impl OoxmlParser {
    pub(crate) fn parse_docx_word_parts(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Result<DocxWordParts, ParseError> {
        let (styles_id, styles_with_effects_id, numbering_id) =
            self.parse_docx_style_parts(zip, main_part_path, doc_rels, parser)?;
        let (comments, footnotes, endnotes, comments_ext_id, comments_id_map_id) =
            self.parse_docx_annotation_parts(zip, main_part_path, doc_rels, parser)?;
        let (settings_id, web_settings_id, font_table_id) =
            self.parse_docx_settings_parts(zip, main_part_path, doc_rels, parser)?;

        let glossary_id = match read_xml_part(zip, "word/glossary/document.xml")? {
            Some(xml) => Some(parser.parse_glossary_document(&xml, doc_rels)?),
            None => None,
        };

        Ok(DocxWordParts {
            styles_id,
            styles_with_effects_id,
            numbering_id,
            comments,
            footnotes,
            endnotes,
            settings_id,
            web_settings_id,
            font_table_id,
            comments_ext_id,
            comments_id_map_id,
            glossary_id,
        })
    }

    fn parse_docx_style_parts(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Result<DocxStylePartIds, ParseError> {
        let styles_id = self.parse_docx_part_by_rel_with_span_result(
            zip,
            main_part_path,
            doc_rels,
            rel_type::STYLES,
            parser,
            |parser, part_path, xml| {
                let id = parser.parse_styles(xml)?;
                if let Some(IRNode::StyleSet(set)) = parser.store_mut().get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
                Ok(id)
            },
        )?;

        let styles_with_effects_id = self.parse_docx_part_by_path_with_span_result(
            zip,
            "word/stylesWithEffects.xml",
            parser,
            |parser, _part_path, xml| parser.parse_styles_with_effects(xml),
            |store, id, part_path| {
                if let Some(IRNode::StyleSet(set)) = store.get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
            },
        )?;

        let numbering_id = self.parse_docx_part_by_rel_with_span_result(
            zip,
            main_part_path,
            doc_rels,
            rel_type::NUMBERING,
            parser,
            |parser, part_path, xml| {
                let id = parser.parse_numbering(xml)?;
                if let Some(IRNode::NumberingSet(set)) = parser.store_mut().get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
                Ok(id)
            },
        )?;

        Ok((styles_id, styles_with_effects_id, numbering_id))
    }

    fn parse_docx_annotation_parts(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Result<DocxAnnotationParts, ParseError> {
        let comments = self.parse_docx_comments(zip, main_part_path, doc_rels, parser)?;
        let footnotes = self.parse_docx_notes(
            zip,
            main_part_path,
            doc_rels,
            parser,
            rel_type::FOOTNOTES,
            crate::ooxml::docx::document::NoteKind::Footnote,
        )?;
        let endnotes = self.parse_docx_notes(
            zip,
            main_part_path,
            doc_rels,
            parser,
            rel_type::ENDNOTES,
            crate::ooxml::docx::document::NoteKind::Endnote,
        )?;

        let comments_ext_id = self.parse_docx_part_by_path_with_span_result(
            zip,
            "word/commentsExtended.xml",
            parser,
            |parser, _part_path, xml| parser.parse_comments_extended(xml),
            |store, id, part_path| {
                if let Some(IRNode::CommentExtensionSet(set)) = store.get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
            },
        )?;

        let comments_id_map_id = self.parse_docx_part_by_path_with_span_result(
            zip,
            "word/commentsIds.xml",
            parser,
            |parser, _part_path, xml| parser.parse_comments_ids(xml),
            |store, id, part_path| {
                if let Some(IRNode::CommentIdMap(map)) = store.get_mut(id) {
                    map.span = Some(SourceSpan::new(part_path));
                }
            },
        )?;

        Ok((
            comments,
            footnotes,
            endnotes,
            comments_ext_id,
            comments_id_map_id,
        ))
    }

    fn parse_docx_settings_parts(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Result<DocxSettingsPartIds, ParseError> {
        let settings_id = self.parse_docx_part_by_rel_with_span_result(
            zip,
            main_part_path,
            doc_rels,
            rel_type::SETTINGS,
            parser,
            |parser, part_path, xml| {
                let id = parser.parse_settings(xml)?;
                if let Some(IRNode::WordSettings(settings)) = parser.store_mut().get_mut(id) {
                    settings.span = Some(SourceSpan::new(part_path));
                }
                Ok(id)
            },
        )?;

        let web_settings_id = self.parse_docx_part_by_rel_with_span_result(
            zip,
            main_part_path,
            doc_rels,
            rel_type::WEB_SETTINGS,
            parser,
            |parser, part_path, xml| {
                let id = parser.parse_web_settings(xml)?;
                if let Some(IRNode::WebSettings(settings)) = parser.store_mut().get_mut(id) {
                    settings.span = Some(SourceSpan::new(part_path));
                }
                Ok(id)
            },
        )?;

        let font_table_id = self.parse_docx_font_table(zip, main_part_path, doc_rels, parser)?;
        Ok((settings_id, web_settings_id, font_table_id))
    }

    pub(crate) fn parse_docx_headers_footers(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Result<HashMap<String, NodeId>, ParseError> {
        let mut map = HashMap::new();

        self.parse_docx_header_footer_kind(
            zip,
            main_part_path,
            doc_rels,
            parser,
            HeaderFooterSpec {
                rel_type: rel_type::HEADER,
                kind: crate::ooxml::docx::document::HeaderFooterKind::Header,
            },
            &mut map,
        )?;
        self.parse_docx_header_footer_kind(
            zip,
            main_part_path,
            doc_rels,
            parser,
            HeaderFooterSpec {
                rel_type: rel_type::FOOTER,
                kind: crate::ooxml::docx::document::HeaderFooterKind::Footer,
            },
            &mut map,
        )?;

        Ok(map)
    }

    fn parse_docx_header_footer_kind(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
        spec: HeaderFooterSpec,
        map: &mut HashMap<String, NodeId>,
    ) -> Result<(), ParseError> {
        for rel in doc_rels.get_by_type(spec.rel_type) {
            let part_path = Relationships::resolve_target(main_part_path, &rel.target);
            let rels = read_relationships_optional(zip, &part_path)?;
            let xml = zip.read_file_string(&part_path)?;
            let node_id = parser.parse_header_footer(&xml, &part_path, spec.kind, &rels)?;
            map.insert(rel.id.clone(), node_id);
        }

        Ok(())
    }

    fn parse_docx_comments(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Result<Vec<NodeId>, ParseError> {
        if let Some(rel) = doc_rels.get_first_by_type(rel_type::COMMENTS) {
            let part_path = Relationships::resolve_target(main_part_path, &rel.target);
            let rels = read_relationships_optional(zip, &part_path)?;
            let xml = zip.read_file_string(&part_path)?;
            let ids = parser.parse_comments(&xml, &rels)?;
            set_comment_spans(parser, &ids, &part_path);
            return Ok(ids);
        }

        if !zip.contains("word/comments.xml") {
            return Ok(Vec::new());
        }
        let rels = read_relationships_optional(zip, "word/comments.xml")?;
        let xml = zip.read_file_string("word/comments.xml")?;
        let ids = parser.parse_comments(&xml, &rels)?;
        set_comment_spans(parser, &ids, "word/comments.xml");
        Ok(ids)
    }

    fn parse_docx_notes(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
        rel_type: &str,
        kind: crate::ooxml::docx::document::NoteKind,
    ) -> Result<Vec<NodeId>, ParseError> {
        let Some(rel) = doc_rels.get_first_by_type(rel_type) else {
            return Ok(Vec::new());
        };
        let part_path = Relationships::resolve_target(main_part_path, &rel.target);
        let rels = read_relationships_optional(zip, &part_path)?;
        let xml = zip.read_file_string(&part_path)?;
        let ids = parser.parse_notes(&xml, kind, &rels)?;
        set_note_spans(parser, &ids, kind, &part_path);
        Ok(ids)
    }

    fn parse_docx_font_table(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Result<Option<NodeId>, ParseError> {
        let mut font_table_id = self.parse_docx_part_by_rel_with_span_result(
            zip,
            main_part_path,
            doc_rels,
            rel_type::FONT_TABLE,
            parser,
            |parser, part_path, xml| {
                let id = parser.parse_font_table(xml)?;
                if let Some(IRNode::FontTable(table)) = parser.store_mut().get_mut(id) {
                    table.span = Some(SourceSpan::new(part_path));
                }
                Ok(id)
            },
        )?;

        if font_table_id.is_none() && zip.contains("word/fontTable.xml") {
            let xml = zip.read_file_string("word/fontTable.xml")?;
            let id = parser.parse_font_table(&xml)?;
            if let Some(IRNode::FontTable(table)) = parser.store_mut().get_mut(id) {
                table.span = Some(SourceSpan::new("word/fontTable.xml"));
            }
            font_table_id = Some(id);
        }

        Ok(font_table_id)
    }

    fn parse_docx_part_by_rel_with_span_result<F>(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        rel_type: &str,
        parser: &mut DocxParser,
        parse: F,
    ) -> Result<Option<NodeId>, ParseError>
    where
        F: FnOnce(&mut DocxParser, &str, &str) -> Result<NodeId, ParseError>,
    {
        let Some((part_path, xml)) = read_xml_part_by_rel(zip, main_part_path, doc_rels, rel_type)?
        else {
            return Ok(None);
        };
        parse(parser, &part_path, &xml).map(Some)
    }

    fn parse_docx_part_by_path_with_span_result<F, S>(
        &self,
        zip: &mut impl PackageReader,
        part_path: &str,
        parser: &mut DocxParser,
        parse: F,
        set_span: S,
    ) -> Result<Option<NodeId>, ParseError>
    where
        F: FnOnce(&mut DocxParser, &str, &str) -> Result<NodeId, ParseError>,
        S: FnOnce(&mut IrStore, NodeId, &str),
    {
        let Some(xml) = read_xml_part(zip, part_path)? else {
            return Ok(None);
        };
        let id = parse(parser, part_path, &xml)?;
        set_span(parser.store_mut(), id, part_path);
        Ok(Some(id))
    }
}

fn set_comment_spans(parser: &mut DocxParser, ids: &[NodeId], part_path: &str) {
    for id in ids {
        if let Some(IRNode::Comment(comment)) = parser.store_mut().get_mut(*id) {
            comment.span = Some(SourceSpan::new(part_path));
        }
    }
}

fn set_note_spans(
    parser: &mut DocxParser,
    ids: &[NodeId],
    kind: crate::ooxml::docx::document::NoteKind,
    part_path: &str,
) {
    for id in ids {
        match kind {
            crate::ooxml::docx::document::NoteKind::Footnote => {
                if let Some(IRNode::Footnote(note)) = parser.store_mut().get_mut(*id) {
                    note.span = Some(SourceSpan::new(part_path));
                }
            }
            crate::ooxml::docx::document::NoteKind::Endnote => {
                if let Some(IRNode::Endnote(note)) = parser.store_mut().get_mut(*id) {
                    note.span = Some(SourceSpan::new(part_path));
                }
            }
        }
    }
}
