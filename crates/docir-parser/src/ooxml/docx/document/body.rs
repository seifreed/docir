use crate::error::ParseError;
use crate::ooxml::relationships::Relationships;
use crate::xml_utils::read_event;
use docir_core::ir::Section;
use docir_core::types::NodeId;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;

use super::inline::{parse_revision_block, parse_sdt, SdtMode};
use super::paragraph::{parse_paragraph, parse_paragraph_simple};
use super::table::parse_table;
use super::{apply_section_refs, DocxParser};
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
            Event::Start(e) => match e.name().as_ref() {
                b"w:p" => {
                    let para = parse_paragraph(parser, reader, rels, header_footer_map)?;
                    current.content.push(para.id);
                    if let Some(section_ref) = para.section_ref {
                        current.headers = section_ref.headers;
                        current.footers = section_ref.footers;
                        current.properties = section_ref.properties;
                        sections.push(current);
                        current = Section::new();
                    }
                }
                b"w:tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    current.content.push(table_id);
                }
                b"w:sdt" => {
                    let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Block)?;
                    current.content.push(sdt_id);
                }
                b"w:sectPr" => {
                    let section_ref = apply_section_refs(reader, header_footer_map)?;
                    current.headers = section_ref.headers;
                    current.footers = section_ref.footers;
                    current.properties = section_ref.properties;
                    sections.push(current);
                    current = Section::new();
                }
                b"w:ins" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::Insert)?;
                    current.content.push(rev_id);
                }
                b"w:del" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::Delete)?;
                    current.content.push(rev_id);
                }
                b"w:moveFrom" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::MoveFrom)?;
                    current.content.push(rev_id);
                }
                b"w:moveTo" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::MoveTo)?;
                    current.content.push(rev_id);
                }
                b"w:pPrChange" | b"w:rPrChange" => {
                    let rev_id =
                        parse_revision_block(parser, reader, rels, &e, RevisionType::FormatChange)?;
                    current.content.push(rev_id);
                }
                _ => {}
            },
            Event::End(e) => {
                if e.name().as_ref() == b"w:body" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if !current.content.is_empty()
        || !current.headers.is_empty()
        || !current.footers.is_empty()
        || current.properties.page_width.is_some()
        || current.properties.page_height.is_some()
        || current.properties.orientation.is_some()
        || current.properties.margins.is_some()
        || current.properties.columns.is_some()
        || sections.is_empty()
    {
        sections.push(current);
    }

    Ok(sections)
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
            Event::Start(e) => match e.name().as_ref() {
                b"w:p" => {
                    let para_id = parse_paragraph_simple(parser, reader, rels)?;
                    content.push(para_id);
                }
                b"w:tbl" => {
                    let table_id = parse_table(parser, reader, rels)?;
                    content.push(table_id);
                }
                b"w:sdt" => {
                    let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Block)?;
                    content.push(sdt_id);
                }
                _ => {}
            },
            Event::End(e) => {
                if e.name().as_ref() == end_tag {
                    break;
                }
            }
            Event::Eof => break,
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
