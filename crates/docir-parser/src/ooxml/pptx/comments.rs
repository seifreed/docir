use crate::error::ParseError;
use crate::xml_utils::{local_name, lossy_attr_value, track_xml_root_event, xml_error};
use docir_core::ir::{PptxComment, PptxCommentAuthor};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

pub(crate) fn parse_comment_authors(
    xml: &str,
    path: &str,
) -> Result<Vec<PptxCommentAuthor>, ParseError> {
    let mut authors: Vec<PptxCommentAuthor> = Vec::new();
    let mut reader = Reader::from_str(xml);
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut root_name = None;
    let mut root_depth = 0;
    let mut root_closed = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        track_xml_root_event(
            &event,
            &mut root_name,
            &mut root_depth,
            &mut root_closed,
            path,
        )?;
        match event {
            Event::Start(e) if local_name(e.name().as_ref()) == b"cmAuthorLst" => {}
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"cmAuthor" => {
                let mut author_id = None;
                let mut name = None;
                let mut initials = None;
                for attr in e.attributes() {
                    let attr = attr.map_err(|err| xml_error(path, err))?;
                    match attr.key.as_ref() {
                        b"id" => {
                            author_id = Some(
                                lossy_attr_value(&attr)
                                    .parse::<u32>()
                                    .map_err(|err| xml_error(path, err))?,
                            );
                        }
                        b"name" => name = Some(lossy_attr_value(&attr).to_string()),
                        b"initials" => initials = Some(lossy_attr_value(&attr).to_string()),
                        _ => {}
                    }
                }
                if let Some(author_id) = author_id {
                    authors.push(PptxCommentAuthor {
                        id: NodeId::new(),
                        author_id,
                        name,
                        initials,
                        span: Some(SourceSpan::new(path)),
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(authors)
}

pub(crate) fn parse_comments(
    xml: &str,
    path: &str,
    authors: &HashMap<u32, (Option<String>, Option<String>)>,
) -> Result<Vec<PptxComment>, ParseError> {
    let mut comments: Vec<PptxComment> = Vec::new();
    let mut reader = Reader::from_str(xml);
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();

    let mut current: Option<PptxComment> = None;
    let mut in_text = false;
    let mut text_buf = String::new();
    let mut root_name = None;
    let mut root_depth = 0;
    let mut root_closed = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        track_xml_root_event(
            &event,
            &mut root_name,
            &mut root_depth,
            &mut root_closed,
            path,
        )?;
        match event {
            Event::Start(e) => {
                if local_name(e.name().as_ref()) == b"cmLst" {
                } else if local_name(e.name().as_ref()) == b"cm" {
                    current = Some(comment_from_start(&e, path)?);
                    text_buf.clear();
                } else if local_name(e.name().as_ref()) == b"t" {
                    in_text = true;
                }
            }
            Event::Text(e) if in_text => {
                text_buf.push_str(
                    &crate::xml_utils::decoded_text(&e).map_err(|err| xml_error(path, err))?,
                );
            }
            Event::GeneralRef(e) if in_text => {
                text_buf.push_str(
                    &crate::xml_utils::decoded_general_ref(&e)
                        .map_err(|err| xml_error(path, err))?,
                );
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == b"t" {
                    in_text = false;
                    flush_comment_text(&mut current, &mut text_buf);
                } else if local_name(e.name().as_ref()) == b"cm" {
                    finish_comment(&mut current, authors, &mut comments);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(comments)
}

fn comment_from_start(e: &BytesStart<'_>, path: &str) -> Result<PptxComment, ParseError> {
    let mut author_id = None;
    let mut dt = None;
    for attr in e.attributes() {
        let attr = attr.map_err(|err| xml_error(path, err))?;
        match attr.key.as_ref() {
            b"authorId" => {
                author_id = Some(
                    lossy_attr_value(&attr)
                        .parse::<u32>()
                        .map_err(|err| xml_error(path, err))?,
                );
            }
            b"dt" => dt = Some(lossy_attr_value(&attr).to_string()),
            _ => {}
        }
    }
    Ok(PptxComment {
        id: NodeId::new(),
        author_id,
        author_name: None,
        author_initials: None,
        datetime: dt,
        text: String::new(),
        span: Some(SourceSpan::new(path)),
    })
}

fn flush_comment_text(current: &mut Option<PptxComment>, text_buf: &mut String) {
    if text_buf.is_empty() {
        return;
    }
    if let Some(cur) = current.as_mut() {
        if !cur.text.is_empty() {
            cur.text.push(' ');
        }
        cur.text.push_str(text_buf);
    }
    text_buf.clear();
}

fn finish_comment(
    current: &mut Option<PptxComment>,
    authors: &HashMap<u32, (Option<String>, Option<String>)>,
    comments: &mut Vec<PptxComment>,
) {
    let Some(mut cur) = current.take() else {
        return;
    };
    if let Some(author_id) = cur.author_id
        && let Some((name, initials)) = authors.get(&author_id)
    {
        cur.author_name = name.clone();
        cur.author_initials = initials.clone();
    }
    comments.push(cur);
}

#[cfg(test)]
mod tests {
    use super::{parse_comment_authors, parse_comments};
    use crate::error::ParseError;
    use std::collections::HashMap;

    fn assert_xml_error(result: Result<impl std::fmt::Debug, ParseError>, path: &str) {
        match result.expect_err("malformed comments XML must fail") {
            ParseError::Xml { file, .. } => assert_eq!(file, path),
            other => panic!("expected XML error, got {other:?}"),
        }
    }

    #[test]
    fn parse_comment_authors_reports_truncated_xml() {
        let xml = r#"
            <p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cmAuthor id="1" name="Alice" initials="AL"/>
        "#;

        assert_xml_error(
            parse_comment_authors(xml, "ppt/commentAuthors.xml"),
            "ppt/commentAuthors.xml",
        );
    }

    #[test]
    fn parse_comments_reports_truncated_xml() {
        let xml = r#"
            <p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cm authorId="1">
                <p:text><p:t>Note</p:t></p:text>
              </p:cm>
        "#;
        let authors = HashMap::new();

        assert_xml_error(
            parse_comments(xml, "ppt/comments/comment1.xml", &authors),
            "ppt/comments/comment1.xml",
        );
    }

    #[test]
    fn parse_comments_reports_truncated_comment_body_xml() {
        let xml = r#"
            <p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cm authorId="1">
                <p:text><p:t>Note</p:t></p:text>
        "#;
        let authors = HashMap::new();

        assert_xml_error(
            parse_comments(xml, "ppt/comments/comment1.xml", &authors),
            "ppt/comments/comment1.xml",
        );
    }

    #[test]
    fn parse_comments_rejects_mismatched_nested_end_tag() {
        let xml = r#"<p:cmLst><p:cm></p:cmx></p:cmLst>"#;
        let authors = HashMap::new();

        assert_xml_error(
            parse_comments(xml, "ppt/comments/comment1.xml", &authors),
            "ppt/comments/comment1.xml",
        );
    }

    #[test]
    fn parse_pptx_comments_rejects_truncated_nested_root() {
        let authors_xml = r#"<p:cmAuthorLst><p:cmAuthorLst></p:cmAuthorLst>"#;
        assert_xml_error(
            parse_comment_authors(authors_xml, "ppt/commentAuthors.xml"),
            "ppt/commentAuthors.xml",
        );

        let comments_xml = r#"<p:cmLst><p:cmLst></p:cmLst>"#;
        assert_xml_error(
            parse_comments(comments_xml, "ppt/comments/comment1.xml", &HashMap::new()),
            "ppt/comments/comment1.xml",
        );
    }

    #[test]
    fn parse_comment_authors_reports_malformed_attributes() {
        let xml = r#"
            <p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cmAuthor id="1" name="Alice" name="Duplicate"/>
            </p:cmAuthorLst>
        "#;

        assert_xml_error(
            parse_comment_authors(xml, "ppt/commentAuthors.xml"),
            "ppt/commentAuthors.xml",
        );

        let xml = r#"
            <p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cmAuthor id="bad" name="Alice" initials="AL"/>
            </p:cmAuthorLst>
        "#;

        assert_xml_error(
            parse_comment_authors(xml, "ppt/commentAuthors.xml"),
            "ppt/commentAuthors.xml",
        );
    }

    #[test]
    fn parse_comments_reports_malformed_comment_attributes() {
        let xml = r#"
            <p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cm authorId="1" authorId="2">
                <p:text><p:t>Note</p:t></p:text>
              </p:cm>
            </p:cmLst>
        "#;
        let authors = HashMap::new();

        assert_xml_error(
            parse_comments(xml, "ppt/comments/comment1.xml", &authors),
            "ppt/comments/comment1.xml",
        );

        let xml = r#"
            <p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cm authorId="bad">
                <p:text><p:t>Note</p:t></p:text>
              </p:cm>
            </p:cmLst>
        "#;

        assert_xml_error(
            parse_comments(xml, "ppt/comments/comment1.xml", &authors),
            "ppt/comments/comment1.xml",
        );
    }
}
