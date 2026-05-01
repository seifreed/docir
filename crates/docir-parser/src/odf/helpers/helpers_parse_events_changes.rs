use super::{
    scan_xml_events_until_end, Endnote, Event, Footnote, IRNode, NodeId, OdfLimitCounter,
    OdfReader, ParseError, Revision, RevisionType, ODF_CONTENT_XML,
};
use crate::odf::paragraph::parse_paragraph;
use docir_core::ir::Comment;
use docir_core::visitor::IrStore;

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
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"office:annotation"),
        |reader, event| {
            match event {
                Event::Start(e) => match e.name().as_ref() {
                    b"dc:creator" => current = Some(AnnotationField::Creator),
                    b"dc:date" => current = Some(AnnotationField::Date),
                    b"text:p" => {
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
                Event::Text(e) => {
                    if let Some(field) = current {
                        let value = e.unescape().unwrap_or_default().to_string();
                        match field {
                            AnnotationField::Creator => comment.author = Some(value),
                            AnnotationField::Date => comment.date = Some(value),
                        }
                    }
                }
                Event::End(e) => {
                    if matches!(e.name().as_ref(), b"dc:creator" | b"dc:date") {
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
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"text:note"),
        |reader, event| {
            if let Event::Start(e) = event {
                if e.name().as_ref() == b"text:p" {
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
                }
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

    enum ChangeInfoField {
        Author,
        Date,
    }

    scan_xml_events_until_end(
        reader,
        &mut buf,
        ODF_CONTENT_XML,
        |event| matches!(event, Event::End(e) if e.name().as_ref() == b"text:tracked-changes"),
        |reader, event| {
            match event {
                Event::Start(e) => match e.name().as_ref() {
                    b"text:changed-region" => {
                        current_revision = None;
                    }
                    b"text:insertion" => {
                        current_revision = Some(Revision::new(RevisionType::Insert));
                    }
                    b"text:deletion" => {
                        current_revision = Some(Revision::new(RevisionType::Delete));
                    }
                    b"dc:creator" => current_field = Some(ChangeInfoField::Author),
                    b"dc:date" => current_field = Some(ChangeInfoField::Date),
                    b"text:p" => {
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
                },
                Event::Text(e) => {
                    if let Some(rev) = current_revision.as_mut() {
                        if let Some(field) = &current_field {
                            let value = e.unescape().unwrap_or_default().to_string();
                            match field {
                                ChangeInfoField::Author => rev.author = Some(value),
                                ChangeInfoField::Date => rev.date = Some(value),
                            }
                        }
                    }
                }
                Event::End(e) => match e.name().as_ref() {
                    b"text:insertion" | b"text:deletion" => {
                        if let Some(rev) = current_revision.take() {
                            let id = rev.id;
                            store.insert(IRNode::Revision(rev));
                            revisions.push(id);
                        }
                    }
                    b"dc:creator" | b"dc:date" => current_field = None,
                    _ => {}
                },
                _ => {}
            }
            Ok(super::XmlScanControl::Continue)
        },
    )?;

    Ok(revisions)
}
