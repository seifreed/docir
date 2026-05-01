use super::{
    parse_comment_authors, parse_presentation_properties, parse_presentation_tags,
    parse_smartart_part, parse_table_styles, parse_view_properties, parse_xml_part_with_span,
    Document, IRNode, ParseError, PptxParser, SourceSpan,
};
use crate::ooxml::part_utils::insert_shared_part;
use crate::zip_handler::PackageReader;

impl PptxParser {
    pub(super) fn load_presentation_parts(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
    ) -> Result<(), ParseError> {
        // Presentation properties
        if let Some(props) = parse_xml_part_with_span(
            zip,
            "ppt/presProps.xml",
            parse_presentation_properties,
            |props, path| props.span = Some(SourceSpan::new(path)),
        )? {
            let id = props.id;
            insert_shared_part(
                &mut self.store,
                document,
                IRNode::PresentationProperties(props),
                id,
            );
        }

        // View properties
        if let Some(props) = parse_xml_part_with_span(
            zip,
            "ppt/viewProps.xml",
            parse_view_properties,
            |props, path| props.span = Some(SourceSpan::new(path)),
        )? {
            let id = props.id;
            insert_shared_part(&mut self.store, document, IRNode::ViewProperties(props), id);
        }

        // Table styles
        if let Some(styles) = parse_xml_part_with_span(
            zip,
            "ppt/tableStyles.xml",
            parse_table_styles,
            |styles, path| styles.span = Some(SourceSpan::new(path)),
        )? {
            let id = styles.id;
            insert_shared_part(&mut self.store, document, IRNode::TableStyleSet(styles), id);
        }

        // Comment authors
        if zip.contains("ppt/commentAuthors.xml") {
            let authors_xml = zip.read_file_string("ppt/commentAuthors.xml")?;
            let authors = parse_comment_authors(&authors_xml, "ppt/commentAuthors.xml")?;
            self.set_comment_authors(&authors);
            for author in authors {
                let mut author = author;
                author.span = Some(SourceSpan::new("ppt/commentAuthors.xml"));
                let id = author.id;
                insert_shared_part(
                    &mut self.store,
                    document,
                    IRNode::PptxCommentAuthor(author),
                    id,
                );
            }
        }

        // Tags
        let tag_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("ppt/tags/") && p.ends_with(".xml"))
            .collect();
        for tag_path in tag_paths {
            let tag_xml = zip.read_file_string(&tag_path)?;
            let tags = parse_presentation_tags(&tag_xml, &tag_path)?;
            for tag in tags {
                let id = tag.id;
                insert_shared_part(&mut self.store, document, IRNode::PresentationTag(tag), id);
            }
        }

        // People part (coauthoring)
        if let Some(people) = parse_xml_part_with_span(
            zip,
            "ppt/people.xml",
            crate::ooxml::shared::parse_people_part,
            |people, path| people.span = Some(SourceSpan::new(path)),
        )? {
            let id = people.id;
            insert_shared_part(&mut self.store, document, IRNode::PeoplePart(people), id);
        }

        // SmartArt parts
        let diagram_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("ppt/diagrams/") && p.ends_with(".xml"))
            .collect();
        for path in diagram_paths {
            let xml = zip.read_file_string(&path)?;
            let part = parse_smartart_part(&xml, &path)?;
            let id = part.id;
            insert_shared_part(&mut self.store, document, IRNode::SmartArtPart(part), id);
        }

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::zip_handler::SecureZipReader;
    use docir_core::types::DocumentFormat;
    use std::io::Cursor;

    fn build_zip_with_entries(entries: Vec<(&str, &str)>) -> SecureZipReader<Cursor<Vec<u8>>> {
        let mut data = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut data));
            let options = zip::write::FileOptions::<()>::default();
            for (path, contents) in entries {
                writer.start_file(path, options).expect("start file");
                use std::io::Write;
                writer.write_all(contents.as_bytes()).expect("write file");
            }
            writer.finish().expect("finish zip");
        }
        SecureZipReader::new(Cursor::new(data), Default::default()).expect("zip")
    }

    #[test]
    fn load_presentation_parts_collects_optional_parts() {
        let mut parser = PptxParser::new();
        let mut document = Document::new(DocumentFormat::Presentation);
        let mut zip = build_zip_with_entries(vec![
            (
                "ppt/presProps.xml",
                r#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" autoCompressPictures="1"/>"#,
            ),
            (
                "ppt/viewProps.xml",
                r#"<p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" showGrid="1"/>"#,
            ),
            (
                "ppt/tableStyles.xml",
                r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{DEF}"><a:tblStyle styleId="{DEF}" name="Default"/></a:tblStyleLst>"#,
            ),
            (
                "ppt/commentAuthors.xml",
                r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Alice" initials="A"/></p:cmAuthorLst>"#,
            ),
            (
                "ppt/tags/tag1.xml",
                r#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag name="env" val="prod"/></p:tagLst>"#,
            ),
            (
                "ppt/people.xml",
                r#"<ppl:people xmlns:ppl="http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes"><ppl:person ppl:id="p1" ppl:userId="u1" ppl:displayName="Alice"/></ppl:people>"#,
            ),
            (
                "ppt/diagrams/data1.xml",
                r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:pt/><dgm:cxn/></dgm:dataModel>"#,
            ),
        ]);

        parser
            .load_presentation_parts(&mut document, &mut zip)
            .expect("load parts");

        assert!(document.shared_parts.iter().any(|id| matches!(
            parser.store.get(*id),
            Some(IRNode::PresentationProperties(_))
        )));
        assert!(document
            .shared_parts
            .iter()
            .any(|id| matches!(parser.store.get(*id), Some(IRNode::ViewProperties(_)))));
        assert!(document
            .shared_parts
            .iter()
            .any(|id| matches!(parser.store.get(*id), Some(IRNode::TableStyleSet(_)))));
        assert!(document
            .shared_parts
            .iter()
            .any(|id| matches!(parser.store.get(*id), Some(IRNode::PptxCommentAuthor(_)))));
        assert!(document
            .shared_parts
            .iter()
            .any(|id| matches!(parser.store.get(*id), Some(IRNode::PresentationTag(_)))));
        assert!(document
            .shared_parts
            .iter()
            .any(|id| matches!(parser.store.get(*id), Some(IRNode::PeoplePart(_)))));
        assert!(document
            .shared_parts
            .iter()
            .any(|id| matches!(parser.store.get(*id), Some(IRNode::SmartArtPart(_)))));
    }
}
