use super::{
    Endnote, Event, Footnote, IRNode, NodeId, ODF_CONTENT_XML, OdfLimitCounter, OdfReader,
    ParseError, Revision, RevisionType, scan_xml_events_until_end,
};
use crate::odf::paragraph::parse_paragraph;
use crate::odf::text::build_paragraph;
use crate::xml_utils::{local_name, xml_error};
use docir_core::ir::Comment;
use docir_core::visitor::IrStore;
use quick_xml::events::BytesStart;

pub(super) fn parse_annotation(
    reader: &mut OdfReader<'_>,
    comment_id: &str,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut comment = Comment::new(comment_id);
    let mut buf = Vec::new();
    let mut current = None;

    #[derive(Clone, Copy)]
    enum AnnotationField {
        Creator,
        Date,
    }

    scan_xml_events_until_end(
        reader,
        &mut buf,
        ODF_CONTENT_XML,
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"annotation"),
        |reader, event| {
            match event {
                Event::Start(e) => match local_name(e.name().as_ref()) {
                    b"creator" => current = Some(AnnotationField::Creator),
                    b"date" => current = Some(AnnotationField::Date),
                    b"p" => {
                        let paragraph_id = parse_paragraph(
                            reader,
                            e.name().as_ref(),
                            None,
                            None,
                            store,
                            &mut Vec::new(),
                            limits,
                        )?;
                        comment.content.push(paragraph_id);
                    }
                    _ => {}
                },
                Event::Empty(e) if local_name(e.name().as_ref()) == b"p" => {
                    limits.bump_paragraphs(1)?;
                    comment.content.push(build_paragraph(store, "", None, None));
                }
                Event::Text(e) => {
                    if let Some(field) = current {
                        let value = crate::xml_utils::decoded_text(e)
                            .map_err(|err| xml_error(ODF_CONTENT_XML, err))?;
                        match field {
                            AnnotationField::Creator => comment.author = Some(value),
                            AnnotationField::Date => comment.date = Some(value),
                        }
                    }
                }
                Event::GeneralRef(e) => {
                    if let Some(field) = current {
                        let value = crate::xml_utils::decoded_general_ref(e)
                            .map_err(|err| xml_error(ODF_CONTENT_XML, err))?;
                        match field {
                            AnnotationField::Creator => comment.author = Some(value),
                            AnnotationField::Date => comment.date = Some(value),
                        }
                    }
                }
                Event::End(e) => {
                    if matches!(local_name(e.name().as_ref()), b"creator" | b"date") {
                        current = None;
                    }
                }
                _ => {}
            }
            Ok(super::XmlScanControl::Continue)
        },
    )?;

    let comment_id = comment.id;
    store.insert(IRNode::Comment(comment));
    Ok(comment_id)
}

pub(super) fn parse_note(
    reader: &mut OdfReader<'_>,
    note_id: &str,
    note_class: &str,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut buf = Vec::new();
    let mut content: Vec<NodeId> = Vec::new();

    scan_xml_events_until_end(
        reader,
        &mut buf,
        ODF_CONTENT_XML,
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"note"),
        |reader, event| {
            if let Event::Start(e) = event
                && local_name(e.name().as_ref()) == b"p"
            {
                let paragraph_id = parse_paragraph(
                    reader,
                    e.name().as_ref(),
                    None,
                    None,
                    store,
                    &mut Vec::new(),
                    limits,
                )?;
                content.push(paragraph_id);
            } else if let Event::Empty(e) = event
                && local_name(e.name().as_ref()) == b"p"
            {
                limits.bump_paragraphs(1)?;
                content.push(build_paragraph(store, "", None, None));
            }
            Ok(super::XmlScanControl::Continue)
        },
    )?;

    if note_class == "endnote" {
        let mut endnote = Endnote::new(note_id);
        endnote.content = content;
        let id = endnote.id;
        store.insert(IRNode::Endnote(endnote));
        Ok(id)
    } else {
        let mut footnote = Footnote::new(note_id);
        footnote.content = content;
        let id = footnote.id;
        store.insert(IRNode::Footnote(footnote));
        Ok(id)
    }
}

pub(super) fn parse_tracked_changes(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<Vec<NodeId>, ParseError> {
    let mut buf = Vec::new();
    let mut revisions = Vec::new();
    let mut current_revision: Option<Revision> = None;
    let mut current_field: Option<ChangeInfoField> = None;

    scan_xml_events_until_end(
        reader,
        &mut buf,
        ODF_CONTENT_XML,
        |event| matches!(event, Event::End(e) if local_name(e.name().as_ref()) == b"tracked-changes"),
        |reader, event| {
            match event {
                Event::Start(e) => handle_tracked_change_start(
                    reader,
                    e,
                    store,
                    limits,
                    &mut current_revision,
                    &mut current_field,
                )?,
                Event::Text(e) => {
                    let value = crate::xml_utils::decoded_text(e)
                        .map_err(|err| xml_error(ODF_CONTENT_XML, err))?;
                    apply_tracked_change_field(&mut current_revision, current_field, value);
                }
                Event::GeneralRef(e) => {
                    let value = crate::xml_utils::decoded_general_ref(e)
                        .map_err(|err| xml_error(ODF_CONTENT_XML, err))?;
                    apply_tracked_change_field(&mut current_revision, current_field, value);
                }
                Event::Empty(e) if local_name(e.name().as_ref()) == b"p" => {
                    if let Some(rev) = current_revision.as_mut() {
                        limits.bump_paragraphs(1)?;
                        rev.content.push(build_paragraph(store, "", None, None));
                    }
                }
                Event::End(e) => match local_name(e.name().as_ref()) {
                    b"insertion" | b"deletion" => {
                        if let Some(rev) = current_revision.take() {
                            let id = rev.id;
                            store.insert(IRNode::Revision(rev));
                            revisions.push(id);
                        }
                    }
                    b"creator" | b"date" => current_field = None,
                    _ => {}
                },
                _ => {}
            }
            Ok(super::XmlScanControl::Continue)
        },
    )?;

    Ok(revisions)
}

#[derive(Clone, Copy)]
enum ChangeInfoField {
    Author,
    Date,
}

fn handle_tracked_change_start(
    reader: &mut OdfReader<'_>,
    e: &BytesStart<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
    current_revision: &mut Option<Revision>,
    current_field: &mut Option<ChangeInfoField>,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"changed-region" => *current_revision = None,
        b"insertion" => *current_revision = Some(Revision::new(RevisionType::Insert)),
        b"deletion" => *current_revision = Some(Revision::new(RevisionType::Delete)),
        b"creator" => *current_field = Some(ChangeInfoField::Author),
        b"date" => *current_field = Some(ChangeInfoField::Date),
        b"p" => {
            if let Some(rev) = current_revision.as_mut() {
                let paragraph_id = parse_paragraph(
                    reader,
                    e.name().as_ref(),
                    None,
                    None,
                    store,
                    &mut Vec::new(),
                    limits,
                )?;
                rev.content.push(paragraph_id);
            }
        }
        _ => {}
    }
    Ok(())
}

fn apply_tracked_change_field(
    current_revision: &mut Option<Revision>,
    current_field: Option<ChangeInfoField>,
    value: String,
) {
    if let Some(rev) = current_revision.as_mut()
        && let Some(field) = current_field
    {
        match field {
            ChangeInfoField::Author => rev.author = Some(value),
            ChangeInfoField::Date => rev.date = Some(value),
        }
    }
}
