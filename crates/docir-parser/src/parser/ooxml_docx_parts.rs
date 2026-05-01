use super::HeaderFooterSpec;
use super::{
    rel_type, DocxAnnotationParts, DocxWordParts, IRNode, NodeId, OoxmlParser, PackageReader,
    ParseError, Relationships,
};
use crate::ooxml::docx::DocxParser;
use crate::ooxml::part_utils::read_relationships_optional;
use crate::ooxml::part_utils::{read_xml_part, read_xml_part_by_rel};
use docir_core::types::SourceSpan;
use docir_core::visitor::IrStore;
use std::collections::HashMap;

impl OoxmlParser {
    pub(crate) fn parse_docx_word_parts(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> DocxWordParts {
        let (styles_id, styles_with_effects_id, numbering_id) =
            self.parse_docx_style_parts(zip, main_part_path, doc_rels, parser);
        let (comments, footnotes, endnotes, comments_ext_id, comments_id_map_id) =
            self.parse_docx_annotation_parts(zip, main_part_path, doc_rels, parser);
        let (settings_id, web_settings_id, font_table_id) =
            self.parse_docx_settings_parts(zip, main_part_path, doc_rels, parser);

        let glossary_id =
            self.parse_docx_part_by_path(zip, "word/glossary/document.xml", |_, xml| {
                parser.parse_glossary_document(xml, doc_rels).ok()
            });

        DocxWordParts {
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
        }
    }

    fn parse_docx_style_parts(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> (Option<NodeId>, Option<NodeId>, Option<NodeId>) {
        let styles_id = self.parse_docx_part_by_rel_with_span(
            zip,
            main_part_path,
            doc_rels,
            rel_type::STYLES,
            parser,
            |parser, part_path, xml| {
                let id = parser.parse_styles(xml).ok()?;
                if let Some(IRNode::StyleSet(set)) = parser.store_mut().get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            },
        );

        let styles_with_effects_id = self.parse_docx_part_by_path_with_span(
            zip,
            "word/stylesWithEffects.xml",
            parser,
            |parser, _part_path, xml| parser.parse_styles_with_effects(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::StyleSet(set)) = store.get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        let numbering_id = self.parse_docx_part_by_rel_with_span(
            zip,
            main_part_path,
            doc_rels,
            rel_type::NUMBERING,
            parser,
            |parser, part_path, xml| {
                let id = parser.parse_numbering(xml).ok()?;
                if let Some(IRNode::NumberingSet(set)) = parser.store_mut().get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            },
        );

        (styles_id, styles_with_effects_id, numbering_id)
    }

    fn parse_docx_annotation_parts(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> DocxAnnotationParts {
        let comments = self.parse_docx_comments(zip, main_part_path, doc_rels, parser);
        let footnotes = self.parse_docx_notes(
            zip,
            main_part_path,
            doc_rels,
            parser,
            rel_type::FOOTNOTES,
            crate::ooxml::docx::document::NoteKind::Footnote,
        );
        let endnotes = self.parse_docx_notes(
            zip,
            main_part_path,
            doc_rels,
            parser,
            rel_type::ENDNOTES,
            crate::ooxml::docx::document::NoteKind::Endnote,
        );

        let comments_ext_id = self.parse_docx_part_by_path_with_span(
            zip,
            "word/commentsExtended.xml",
            parser,
            |parser, _part_path, xml| parser.parse_comments_extended(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::CommentExtensionSet(set)) = store.get_mut(id) {
                    set.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        let comments_id_map_id = self.parse_docx_part_by_path_with_span(
            zip,
            "word/commentsIds.xml",
            parser,
            |parser, _part_path, xml| parser.parse_comments_ids(xml).ok(),
            |store, id, part_path| {
                if let Some(IRNode::CommentIdMap(map)) = store.get_mut(id) {
                    map.span = Some(SourceSpan::new(part_path));
                }
            },
        );

        (
            comments,
            footnotes,
            endnotes,
            comments_ext_id,
            comments_id_map_id,
        )
    }

    fn parse_docx_settings_parts(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> (Option<NodeId>, Option<NodeId>, Option<NodeId>) {
        let settings_id = self.parse_docx_part_by_rel_with_span(
            zip,
            main_part_path,
            doc_rels,
            rel_type::SETTINGS,
            parser,
            |parser, part_path, xml| {
                let id = parser.parse_settings(xml).ok()?;
                if let Some(IRNode::WordSettings(settings)) = parser.store_mut().get_mut(id) {
                    settings.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            },
        );

        let web_settings_id = self.parse_docx_part_by_rel_with_span(
            zip,
            main_part_path,
            doc_rels,
            rel_type::WEB_SETTINGS,
            parser,
            |parser, part_path, xml| {
                let id = parser.parse_web_settings(xml).ok()?;
                if let Some(IRNode::WebSettings(settings)) = parser.store_mut().get_mut(id) {
                    settings.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            },
        );

        let font_table_id = self.parse_docx_font_table(zip, main_part_path, doc_rels, parser);
        (settings_id, web_settings_id, font_table_id)
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
            let rels = read_relationships_optional(zip, &part_path);
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
    ) -> Vec<NodeId> {
        let comments = doc_rels
            .get_first_by_type(rel_type::COMMENTS)
            .and_then(|rel| {
                let part_path = Relationships::resolve_target(main_part_path, &rel.target);
                let rels = read_relationships_optional(zip, &part_path);
                zip.read_file_string(&part_path).ok().and_then(|xml| {
                    let ids = parser.parse_comments(&xml, &rels).ok()?;
                    for id in &ids {
                        if let Some(IRNode::Comment(comment)) = parser.store_mut().get_mut(*id) {
                            comment.span = Some(SourceSpan::new(&part_path));
                        }
                    }
                    Some(ids)
                })
            })
            .unwrap_or_default();

        if !comments.is_empty() || !zip.contains("word/comments.xml") {
            return comments;
        }

        let rels = read_relationships_optional(zip, "word/comments.xml");
        if let Ok(xml) = zip.read_file_string("word/comments.xml") {
            if let Ok(ids) = parser.parse_comments(&xml, &rels) {
                for id in &ids {
                    if let Some(IRNode::Comment(comment)) = parser.store_mut().get_mut(*id) {
                        comment.span = Some(SourceSpan::new("word/comments.xml"));
                    }
                }
                return ids;
            }
        }

        comments
    }

    fn parse_docx_notes(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
        rel_type: &str,
        kind: crate::ooxml::docx::document::NoteKind,
    ) -> Vec<NodeId> {
        doc_rels
            .get_first_by_type(rel_type)
            .and_then(|rel| {
                let part_path = Relationships::resolve_target(main_part_path, &rel.target);
                let rels = read_relationships_optional(zip, &part_path);
                zip.read_file_string(&part_path).ok().and_then(|xml| {
                    let ids = parser.parse_notes(&xml, kind, &rels).ok()?;
                    for id in &ids {
                        match kind {
                            crate::ooxml::docx::document::NoteKind::Footnote => {
                                if let Some(IRNode::Footnote(note)) =
                                    parser.store_mut().get_mut(*id)
                                {
                                    note.span = Some(SourceSpan::new(&part_path));
                                }
                            }
                            crate::ooxml::docx::document::NoteKind::Endnote => {
                                if let Some(IRNode::Endnote(note)) = parser.store_mut().get_mut(*id)
                                {
                                    note.span = Some(SourceSpan::new(&part_path));
                                }
                            }
                        }
                    }
                    Some(ids)
                })
            })
            .unwrap_or_default()
    }

    fn parse_docx_font_table(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        parser: &mut DocxParser,
    ) -> Option<NodeId> {
        let mut font_table_id = self.parse_docx_part_by_rel_with_span(
            zip,
            main_part_path,
            doc_rels,
            rel_type::FONT_TABLE,
            parser,
            |parser, part_path, xml| {
                let id = parser.parse_font_table(xml).ok()?;
                if let Some(IRNode::FontTable(table)) = parser.store_mut().get_mut(id) {
                    table.span = Some(SourceSpan::new(part_path));
                }
                Some(id)
            },
        );

        if font_table_id.is_none() && zip.contains("word/fontTable.xml") {
            if let Ok(xml) = zip.read_file_string("word/fontTable.xml") {
                if let Ok(id) = parser.parse_font_table(&xml) {
                    if let Some(IRNode::FontTable(table)) = parser.store_mut().get_mut(id) {
                        table.span = Some(SourceSpan::new("word/fontTable.xml"));
                    }
                    font_table_id = Some(id);
                }
            }
        }

        font_table_id
    }

    fn parse_docx_part_by_rel_with_span<F>(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        rel_type: &str,
        parser: &mut DocxParser,
        parse: F,
    ) -> Option<NodeId>
    where
        F: FnOnce(&mut DocxParser, &str, &str) -> Option<NodeId>,
    {
        let (part_path, xml) =
            self.read_xml_part_by_rel_optional(zip, main_part_path, doc_rels, rel_type)?;
        parse(parser, &part_path, &xml)
    }

    fn parse_docx_part_by_path<F>(
        &self,
        zip: &mut impl PackageReader,
        part_path: &str,
        mut parse: F,
    ) -> Option<NodeId>
    where
        F: FnMut(&str, &str) -> Option<NodeId>,
    {
        let xml = self.read_xml_part_optional(zip, part_path)?;
        parse(part_path, &xml)
    }

    fn parse_docx_part_by_path_with_span<F, S>(
        &self,
        zip: &mut impl PackageReader,
        part_path: &str,
        parser: &mut DocxParser,
        parse: F,
        set_span: S,
    ) -> Option<NodeId>
    where
        F: FnOnce(&mut DocxParser, &str, &str) -> Option<NodeId>,
        S: FnOnce(&mut IrStore, NodeId, &str),
    {
        let xml = self.read_xml_part_optional(zip, part_path)?;
        let id = parse(parser, part_path, &xml)?;
        set_span(parser.store_mut(), id, part_path);
        Some(id)
    }

    fn read_xml_part_by_rel_optional(
        &self,
        zip: &mut impl PackageReader,
        main_part_path: &str,
        doc_rels: &Relationships,
        rel_type: &str,
    ) -> Option<(String, String)> {
        read_xml_part_by_rel(zip, main_part_path, doc_rels, rel_type)
            .ok()
            .flatten()
    }

    fn read_xml_part_optional(
        &self,
        zip: &mut impl PackageReader,
        part_path: &str,
    ) -> Option<String> {
        read_xml_part(zip, part_path).ok().flatten()
    }
}
