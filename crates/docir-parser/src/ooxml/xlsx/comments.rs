use crate::error::ParseError;
use crate::xml_utils::{local_name, track_xml_document_event, try_attr_value, xml_error};
use docir_core::ir::SheetComment;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

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
    let mut state = CommentParseState::default();
    let mut depth = 0usize;
    let mut root_closed = false;
    loop {
        let event = crate::xml_utils::read_event(&mut reader, &mut buf, path)?;
        if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
            break;
        }
        match event {
            Event::Start(e) => handle_comment_start(&e, path, &flavor, &mut state)?,
            Event::Text(e) => {
                let text =
                    crate::xml_utils::decoded_text(&e).map_err(|err| xml_error(path, err))?;
                if state.in_author {
                    state.authors.push(text);
                } else if state.in_text {
                    state.current_text.push_str(&text);
                }
            }
            Event::GeneralRef(e) => {
                let text = crate::xml_utils::decoded_general_ref(&e)
                    .map_err(|err| xml_error(path, err))?;
                if state.in_author {
                    state.authors.push(text);
                } else if state.in_text {
                    state.current_text.push_str(&text);
                }
            }
            Event::End(e) => handle_comment_end(
                local_name(e.name().as_ref()),
                &flavor,
                path,
                sheet_name,
                &mut state,
            )?,
            _ => {}
        }
        buf.clear();
    }

    Ok(state.out)
}

#[derive(Default)]
struct CommentParseState {
    authors: Vec<String>,
    in_author: bool,
    in_comment: bool,
    in_text: bool,
    current_ref: Option<String>,
    current_author: Option<String>,
    current_text: String,
    out: Vec<SheetComment>,
}

fn handle_comment_start(
    e: &BytesStart<'_>,
    path: &str,
    flavor: &CommentFlavor,
    state: &mut CommentParseState,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"author" if matches!(flavor, CommentFlavor::Legacy) => state.in_author = true,
        b"comment" if matches!(flavor, CommentFlavor::Legacy) => {
            state.in_comment = true;
            state.current_ref = try_attr_value(e, b"ref", path)?;
            state.current_author = try_attr_value(e, b"authorId", path)?;
            state.current_text.clear();
        }
        b"threadedComment" if matches!(flavor, CommentFlavor::Threaded) => {
            state.in_comment = true;
            state.current_ref = try_attr_value(e, b"ref", path)?;
            state.current_author = match try_attr_value(e, b"authorId", path)? {
                Some(author_id) => Some(author_id),
                None => try_attr_value(e, b"personId", path)?,
            };
            state.current_text.clear();
        }
        b"text" | b"t" if state.in_comment => state.in_text = true,
        _ => {}
    }
    Ok(())
}

fn handle_comment_end(
    name: &[u8],
    flavor: &CommentFlavor,
    path: &str,
    sheet_name: Option<&str>,
    state: &mut CommentParseState,
) -> Result<(), ParseError> {
    match name {
        b"author" => state.in_author = false,
        b"text" | b"t" => state.in_text = false,
        b"comment" if matches!(flavor, CommentFlavor::Legacy) => {
            finish_legacy_comment(path, sheet_name, state)?;
            state.in_comment = false;
        }
        b"threadedComment" if matches!(flavor, CommentFlavor::Threaded) => {
            finish_threaded_comment(sheet_name, state);
            state.in_comment = false;
        }
        _ => {}
    }
    Ok(())
}

fn finish_legacy_comment(
    path: &str,
    sheet_name: Option<&str>,
    state: &mut CommentParseState,
) -> Result<(), ParseError> {
    if let Some(cell_ref) = state.current_ref.take() {
        let mut comment = SheetComment::new(cell_ref, state.current_text.trim().to_string());
        comment.sheet_name = sheet_name.map(|s| s.to_string());
        if let Some(raw_author_id) = state.current_author.take() {
            let id = raw_author_id
                .parse::<usize>()
                .map_err(|err| xml_error(path, err))?;
            comment.author = state.authors.get(id).cloned();
        }
        state.out.push(comment);
    }
    Ok(())
}

fn finish_threaded_comment(sheet_name: Option<&str>, state: &mut CommentParseState) {
    if let Some(cell_ref) = state.current_ref.take() {
        let mut comment = SheetComment::new(cell_ref, state.current_text.trim().to_string());
        comment.sheet_name = sheet_name.map(|s| s.to_string());
        comment.author = state.current_author.take();
        state.out.push(comment);
    }
}
