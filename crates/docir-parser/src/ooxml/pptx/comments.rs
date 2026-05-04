use crate::error::ParseError;
use crate::xml_utils::lossy_attr_value;
use docir_core::ir::{PptxComment, PptxCommentAuthor};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;

pub(crate) fn parse_comment_authors(
    xml: &str,
    path: &str,
) -> Result<Vec<PptxCommentAuthor>, ParseError> {
    let mut authors: Vec<PptxCommentAuthor> = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref().ends_with(b"cmAuthor") {
                    let mut author_id = None;
                    let mut name = None;
                    let mut initials = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"id" => author_id = lossy_attr_value(&attr).parse::<u32>().ok(),
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
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
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

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref().ends_with(b"cm") {
                    let mut author_id = None;
                    let mut dt = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"authorId" => author_id = lossy_attr_value(&attr).parse::<u32>().ok(),
                            b"dt" => dt = Some(lossy_attr_value(&attr).to_string()),
                            _ => {}
                        }
                    }
                    current = Some(PptxComment {
                        id: NodeId::new(),
                        author_id,
                        author_name: None,
                        author_initials: None,
                        datetime: dt,
                        text: String::new(),
                        span: Some(SourceSpan::new(path)),
                    });
                    text_buf.clear();
                } else if e.name().as_ref().ends_with(b"t") {
                    in_text = true;
                }
            }
            Ok(Event::Text(e)) => {
                if in_text {
                    text_buf.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref().ends_with(b"t") {
                    in_text = false;
                    if !text_buf.is_empty() {
                        if let Some(cur) = current.as_mut() {
                            if !cur.text.is_empty() {
                                cur.text.push(' ');
                            }
                            cur.text.push_str(&text_buf);
                        }
                        text_buf.clear();
                    }
                } else if e.name().as_ref().ends_with(b"cm") {
                    if let Some(mut cur) = current.take() {
                        if let Some(author_id) = cur.author_id {
                            if let Some((name, initials)) = authors.get(&author_id) {
                                cur.author_name = name.clone();
                                cur.author_initials = initials.clone();
                            }
                        }
                        comments.push(cur);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(comments)
}
