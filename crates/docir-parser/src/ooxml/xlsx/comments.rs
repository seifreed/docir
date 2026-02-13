use crate::error::ParseError;
use crate::xml_utils::attr_value;
use docir_core::ir::SheetComment;
use quick_xml::events::Event;
use quick_xml::Reader;

pub(super) enum CommentFlavor {
    Legacy,
    Threaded,
}

pub(super) fn parse_sheet_comments_impl(
    xml: &str,
    path: &str,
    sheet_name: Option<&str>,
    flavor: CommentFlavor,
) -> Result<Vec<SheetComment>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut authors: Vec<String> = Vec::new();
    let mut in_author = false;
    let mut in_comment = false;
    let mut in_text = false;
    let mut current_ref: Option<String> = None;
    let mut current_author: Option<String> = None;
    let mut current_text = String::new();

    let mut out = Vec::new();
    loop {
        match crate::xml_utils::read_event(&mut reader, &mut buf, path)? {
            Event::Start(e) => match e.name().as_ref() {
                b"author" if matches!(flavor, CommentFlavor::Legacy) => in_author = true,
                b"comment" if matches!(flavor, CommentFlavor::Legacy) => {
                    in_comment = true;
                    current_ref = attr_value(&e, b"ref");
                    current_author = attr_value(&e, b"authorId");
                    current_text.clear();
                }
                b"threadedComment" if matches!(flavor, CommentFlavor::Threaded) => {
                    in_comment = true;
                    current_ref = attr_value(&e, b"ref");
                    current_author =
                        attr_value(&e, b"authorId").or_else(|| attr_value(&e, b"personId"));
                    current_text.clear();
                }
                b"text" | b"t" => {
                    if in_comment {
                        in_text = true;
                    }
                }
                _ => {}
            },
            Event::Text(e) => {
                let text = e.unescape().unwrap_or_default().to_string();
                if in_author {
                    authors.push(text);
                } else if in_text {
                    current_text.push_str(&text);
                }
            }
            Event::End(e) => match e.name().as_ref() {
                b"author" => in_author = false,
                b"text" | b"t" => in_text = false,
                b"comment" if matches!(flavor, CommentFlavor::Legacy) => {
                    if let Some(cell_ref) = current_ref.take() {
                        let mut comment =
                            SheetComment::new(cell_ref, current_text.trim().to_string());
                        comment.sheet_name = sheet_name.map(|s| s.to_string());
                        let author_id = current_author.take().and_then(|v| v.parse::<usize>().ok());
                        if let Some(id) = author_id {
                            comment.author = authors.get(id).cloned();
                        }
                        out.push(comment);
                    }
                    in_comment = false;
                }
                b"threadedComment" if matches!(flavor, CommentFlavor::Threaded) => {
                    if let Some(cell_ref) = current_ref.take() {
                        let mut comment =
                            SheetComment::new(cell_ref, current_text.trim().to_string());
                        comment.sheet_name = sheet_name.map(|s| s.to_string());
                        comment.author = current_author.take();
                        out.push(comment);
                    }
                    in_comment = false;
                }
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}
