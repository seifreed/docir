use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::{local_name, read_event};
use docir_core::ir::Section;
use docir_core::types::NodeId;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

use super::inline::{SdtMode, parse_revision_block, parse_sdt};
use super::paragraph::{parse_empty_paragraph, parse_paragraph, parse_paragraph_simple};
use super::table::{parse_empty_table, parse_table};
use super::{DocxParser, apply_section_refs};
use docir_core::ir::RevisionType;

pub(crate) fn parse_body_sections(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<Vec<Section>, ParseError> {
    let mut sections: Vec<Section> = Vec::new();
    let mut current = Section::new();
    let mut buf = Vec::new();

    loop {
        match read_event(reader, &mut buf, "word/document.xml")? {
            Event::Start(e) => handle_body_start_event(
                parser,
                reader,
                rels,
                header_footer_map,
                &mut sections,
                &mut current,
                &e,
            )?,
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"p" => current.content.push(parse_empty_paragraph(parser, reader)),
                b"tbl" => current.content.push(parse_empty_table(parser)),
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"body" => {
                break;
            }
            Event::Eof => {
                return Err(crate::xml_utils::xml_error(
                    "word/document.xml",
                    "unexpected end of body",
                ));
            }
            _ => {}
        }
        buf.clear();
    }

    if section_has_content(&current) || sections.is_empty() {
        sections.push(current);
    }

    Ok(sections)
}

fn handle_body_start_event(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
    sections: &mut Vec<Section>,
    current: &mut Section,
    e: &BytesStart<'_>,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"p" => {
            let para = parse_paragraph(parser, reader, rels, header_footer_map)?;
            current.content.push(para.id);
            if let Some(section_ref) = para.section_ref {
                current.headers = section_ref.headers;
                current.footers = section_ref.footers;
                current.properties = section_ref.properties;
                flush_section(sections, current);
            }
        }
        b"tbl" => current.content.push(parse_table(parser, reader, rels)?),
        b"sdt" => current
            .content
            .push(parse_sdt(parser, reader, rels, SdtMode::Block)?),
        b"sectPr" => {
            let section_ref = apply_section_refs(reader, header_footer_map)?;
            current.headers = section_ref.headers;
            current.footers = section_ref.footers;
            current.properties = section_ref.properties;
            flush_section(sections, current);
        }
        b"ins" => push_revision(parser, reader, rels, current, e, RevisionType::Insert)?,
        b"del" => push_revision(parser, reader, rels, current, e, RevisionType::Delete)?,
        b"moveFrom" => push_revision(parser, reader, rels, current, e, RevisionType::MoveFrom)?,
        b"moveTo" => push_revision(parser, reader, rels, current, e, RevisionType::MoveTo)?,
        b"pPrChange" | b"rPrChange" => {
            push_revision(parser, reader, rels, current, e, RevisionType::FormatChange)?;
        }
        _ => {}
    }
    Ok(())
}

fn push_revision(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    current: &mut Section,
    e: &BytesStart<'_>,
    revision_type: RevisionType,
) -> Result<(), ParseError> {
    let rev_id = parse_revision_block(parser, reader, rels, e, revision_type)?;
    current.content.push(rev_id);
    Ok(())
}

fn flush_section(sections: &mut Vec<Section>, current: &mut Section) {
    let mut next = Section::new();
    std::mem::swap(current, &mut next);
    sections.push(next);
}

fn section_has_content(section: &Section) -> bool {
    !section.content.is_empty()
        || !section.headers.is_empty()
        || !section.footers.is_empty()
        || section.properties.page_width.is_some()
        || section.properties.page_height.is_some()
        || section.properties.orientation.is_some()
        || section.properties.margins.is_some()
        || section.properties.columns.is_some()
}

pub(crate) fn parse_block_until(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    end_tag: &[u8],
) -> Result<Vec<NodeId>, ParseError> {
    let mut content = Vec::new();
    let mut buf = Vec::new();

    loop {
        match read_event(reader, &mut buf, "word/document.xml")? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"p" => {
                    let para_id = parse_paragraph_simple(parser, reader, rels)?;
                    content.push(para_id);
                }
                b"tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    content.push(table_id);
                }
                b"sdt" => {
                    let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Block)?;
                    content.push(sdt_id);
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"p" => content.push(parse_empty_paragraph(parser, reader)),
                b"tbl" => content.push(parse_empty_table(parser)),
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == local_name(end_tag) => {
                break;
            }
            Event::Eof => {
                return Err(crate::xml_utils::xml_error(
                    "word/document.xml",
                    "unexpected end of block",
                ));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(content)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::relationships::Relationships;
    use crate::xml_utils::reader_from_str;
    use docir_core::types::NodeId;

    #[test]
    fn parse_body_sections_returns_default_section_for_empty_body() {
        let xml = r#"<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:body>"#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();

        let sections = parse_body_sections(&mut parser, &mut reader, &rels, None)
            .expect("empty body should parse");
        assert_eq!(sections.len(), 1);
        assert!(sections[0].content.is_empty());
    }

    #[test]
    fn parse_body_sections_preserves_empty_paragraph() {
        let xml = r#"<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:body>"#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let sections =
            parse_body_sections(&mut parser, &mut reader, &Relationships::default(), None)
                .expect("empty paragraph should parse");
        let store = parser.into_store();

        assert_eq!(sections.len(), 1);
        assert_eq!(sections[0].content.len(), 1);
        assert!(matches!(
            store.get(sections[0].content[0]),
            Some(docir_core::ir::IRNode::Paragraph(_))
        ));
    }

    #[test]
    fn parse_body_sections_collects_paragraph_content() {
        let xml = r#"
            <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:p><w:r><w:t>First</w:t></w:r></w:p>
              <w:p><w:r><w:t>Second</w:t></w:r></w:p>
            </w:body>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();

        let sections = parse_body_sections(&mut parser, &mut reader, &rels, None)
            .expect("body with paragraphs should parse");
        assert_eq!(sections.len(), 1);
        assert_eq!(sections[0].content.len(), 2);
    }

    #[test]
    fn parse_block_until_collects_nodes_and_stops_at_end_tag() {
        let xml = r#"
            <w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:p><w:r><w:t>One</w:t></w:r></w:p>
              <w:p><w:r><w:t>Two</w:t></w:r></w:p>
            </w:comment>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();

        let content = parse_block_until(&mut parser, &mut reader, &rels, b"w:comment")
            .expect("comment block should parse");
        assert_eq!(content.len(), 2);
    }

    #[test]
    fn parse_body_sections_rejects_missing_body_end() {
        let xml = r#"<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Broken</w:t></w:r></w:p>"#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let err = parse_body_sections(&mut parser, &mut reader, &Relationships::default(), None)
            .expect_err("truncated body must fail");

        assert!(matches!(err, ParseError::Xml { .. }));
    }

    #[test]
    fn parse_block_until_rejects_missing_end_tag() {
        let xml = r#"<w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Broken</w:t></w:r></w:p>"#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let err = parse_block_until(
            &mut parser,
            &mut reader,
            &Relationships::default(),
            b"w:comment",
        )
        .expect_err("truncated block must fail");

        assert!(matches!(err, ParseError::Xml { .. }));
    }

    #[test]
    fn parse_body_sections_flushes_section_on_sectpr() {
        let xml = r#"
            <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:p><w:r><w:t>Before break</w:t></w:r></w:p>
              <w:sectPr>
                <w:pgSz w:w="12240" w:h="15840"/>
              </w:sectPr>
              <w:p><w:r><w:t>After break</w:t></w:r></w:p>
            </w:body>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();

        let sections = parse_body_sections(&mut parser, &mut reader, &rels, None)
            .expect("body with section properties should parse");
        assert_eq!(sections.len(), 2);
        assert_eq!(sections[0].content.len(), 1);
        assert_eq!(sections[1].content.len(), 1);
    }

    #[test]
    fn parse_body_sections_collects_table_nodes() {
        let xml = r#"
            <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:tbl>
                <w:tr>
                  <w:tc>
                    <w:p><w:r><w:t>Cell</w:t></w:r></w:p>
                  </w:tc>
                </w:tr>
              </w:tbl>
              <w:p><w:r><w:t>Tail</w:t></w:r></w:p>
            </w:body>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();

        let sections = parse_body_sections(&mut parser, &mut reader, &rels, None)
            .expect("body with table should parse");
        assert_eq!(sections.len(), 1);
        assert_eq!(sections[0].content.len(), 2);
    }

    #[test]
    fn parse_body_sections_preserves_empty_table() {
        let xml = r#"
            <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:tbl/>
            </w:body>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();

        let sections =
            parse_body_sections(&mut parser, &mut reader, &Relationships::default(), None)
                .expect("empty table should parse");
        let store = parser.into_store();

        assert_eq!(sections[0].content.len(), 1);
        assert!(matches!(
            store.get(sections[0].content[0]),
            Some(docir_core::ir::IRNode::Table(table)) if table.rows.is_empty()
        ));
    }

    #[test]
    fn parse_body_sections_collects_sdt_and_revision_blocks() {
        let xml = r#"
            <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:sdt>
                <w:sdtContent>
                  <w:p><w:r><w:t>Inside SDT</w:t></w:r></w:p>
                </w:sdtContent>
              </w:sdt>
              <w:ins w:id="1" w:author="Alice">
                <w:r><w:t>Inserted</w:t></w:r>
              </w:ins>
              <w:del w:id="2" w:author="Bob">
                <w:r><w:delText>Deleted</w:delText></w:r>
              </w:del>
              <w:moveFrom w:id="3" w:author="Carol">
                <w:r><w:t>Moved from</w:t></w:r>
              </w:moveFrom>
              <w:moveTo w:id="4" w:author="Dave">
                <w:r><w:t>Moved to</w:t></w:r>
              </w:moveTo>
              <w:rPrChange w:id="5" w:author="Eve">
                <w:r><w:t>Format change</w:t></w:r>
              </w:rPrChange>
            </w:body>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();

        let sections =
            parse_body_sections(&mut parser, &mut reader, &rels, None).expect("body should parse");
        assert_eq!(sections.len(), 1);
        assert_eq!(
            sections[0].content.len(),
            6,
            "sdt + 5 revision blocks should be recorded"
        );
    }

    #[test]
    fn parse_body_sections_maps_header_and_footer_refs_from_sectpr() {
        let xml = r#"
            <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:sectPr>
                <w:headerReference w:type="default" r:id="rIdHeader"/>
                <w:footerReference w:type="default" r:id="rIdFooter"/>
              </w:sectPr>
            </w:body>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();

        let mut map = HashMap::new();
        let header_id = NodeId::new();
        let footer_id = NodeId::new();
        map.insert("rIdHeader".to_string(), header_id);
        map.insert("rIdFooter".to_string(), footer_id);

        let sections = parse_body_sections(&mut parser, &mut reader, &rels, Some(&map))
            .expect("body with sectPr refs should parse");
        assert_eq!(
            sections.len(),
            1,
            "sectPr-only body yields one populated section"
        );
        assert_eq!(sections[0].headers.len(), 1);
        assert_eq!(sections[0].footers.len(), 1);
        assert_eq!(sections[0].headers[0], header_id);
        assert_eq!(sections[0].footers[0], footer_id);
    }

    #[test]
    fn parse_block_until_collects_table_and_sdt_nodes() {
        let xml = r#"
            <w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:tbl>
                <w:tr>
                  <w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>
                </w:tr>
              </w:tbl>
              <w:sdt>
                <w:sdtContent>
                  <w:p><w:r><w:t>Inside SDT</w:t></w:r></w:p>
                </w:sdtContent>
              </w:sdt>
            </w:comment>
        "#;
        let mut reader = reader_from_str(xml);
        let mut parser = DocxParser::new();
        let rels = Relationships::default();
        let content = parse_block_until(&mut parser, &mut reader, &rels, b"w:comment")
            .expect("comment block should parse");
        assert_eq!(content.len(), 2);
    }
}
