use super::*;
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
