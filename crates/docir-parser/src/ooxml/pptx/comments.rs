use crate::error::ParseError;
use crate::xml_utils::{local_name, lossy_attr_value, xml_error};
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
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut root_closed = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"cmAuthorLst" => {}
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"cmAuthorLst" => {
                root_closed = true;
            }
            Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"cmAuthorLst" => {
                root_closed = true;
            }
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if local_name(e.name().as_ref()) == b"cmAuthor" =>
            {
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
            Ok(Event::Eof) if root_closed => break,
            Ok(Event::Eof) => {
                return Err(xml_error(path, "unexpected EOF in comment authors XML"));
            }
            Err(e) => {
                return Err(xml_error(path, e));
            }
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
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut current: Option<PptxComment> = None;
    let mut in_text = false;
    let mut text_buf = String::new();
    let mut root_closed = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if local_name(e.name().as_ref()) == b"cmLst" {
                } else if local_name(e.name().as_ref()) == b"cm" {
                    current = Some(comment_from_start(&e, path)?);
                    text_buf.clear();
                } else if local_name(e.name().as_ref()) == b"t" {
                    in_text = true;
                }
            }
            Ok(Event::Text(e)) if in_text => {
                text_buf.push_str(
                    &crate::xml_utils::decoded_text(&e).map_err(|err| xml_error(path, err))?,
                );
            }
            Ok(Event::GeneralRef(e)) if in_text => {
                text_buf.push_str(
                    &crate::xml_utils::decoded_general_ref(&e)
                        .map_err(|err| xml_error(path, err))?,
                );
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"t" {
                    in_text = false;
                    flush_comment_text(&mut current, &mut text_buf);
                } else if local_name(e.name().as_ref()) == b"cm" {
                    finish_comment(&mut current, authors, &mut comments);
                } else if local_name(e.name().as_ref()) == b"cmLst" {
                    root_closed = true;
                }
            }
            Ok(Event::Eof) if root_closed && current.is_none() && !in_text => break,
            Ok(Event::Eof) => {
                return Err(xml_error(path, "unexpected EOF in comments XML"));
            }
            Err(e) => {
                return Err(xml_error(path, e));
            }
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
