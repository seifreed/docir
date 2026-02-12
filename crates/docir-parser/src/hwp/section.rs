use super::*;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum HwpxNoteKind {
    Comment,
    Footnote,
    Endnote,
}

struct HwpxNoteState {
    kind: HwpxNoteKind,
    id: String,
    author: Option<String>,
    date: Option<String>,
    parent: Option<String>,
    content: Vec<NodeId>,
    current_para: Option<Paragraph>,
}

fn note_kind_from_local(local: &[u8]) -> Option<HwpxNoteKind> {
    match local {
        b"comment" | b"annotation" | b"note" => Some(HwpxNoteKind::Comment),
        b"footnote" => Some(HwpxNoteKind::Footnote),
        b"endnote" => Some(HwpxNoteKind::Endnote),
        _ => None,
    }
}

fn revision_type_from_local(local: &[u8]) -> Option<RevisionType> {
    match local {
        b"ins" | b"insert" => Some(RevisionType::Insert),
        b"del" | b"delete" => Some(RevisionType::Delete),
        b"moveFrom" | b"move-from" => Some(RevisionType::MoveFrom),
        b"moveTo" | b"move-to" => Some(RevisionType::MoveTo),
        b"formatChange" | b"format-change" => Some(RevisionType::FormatChange),
        _ => None,
    }
}

pub(super) fn parse_hwpx_section(
    xml: &str,
    source: &str,
    store: &mut IrStore,
    comments: &mut Vec<NodeId>,
    footnotes: &mut Vec<NodeId>,
    endnotes: &mut Vec<NodeId>,
    media_lookup: &HashMap<String, NodeId>,
) -> Result<Vec<NodeId>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut content: Vec<NodeId> = Vec::new();
    let mut current_para: Option<Paragraph> = None;
    let mut current_table: Option<Table> = None;
    let mut current_row: Option<TableRow> = None;
    let mut current_cell: Option<TableCell> = None;
    let mut in_text = false;
    let mut run_props: Option<RunProperties> = None;
    let mut list_level: u32 = 0;
    let mut revision_stack: Vec<Revision> = Vec::new();
    let mut note_stack: Vec<HwpxNoteState> = Vec::new();
    let mut comment_counter: u32 = 0;
    let mut footnote_counter: u32 = 0;
    let mut endnote_counter: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if let Some(kind) = note_kind_from_local(local) {
                    let id = attr_any(
                        &e,
                        &[b"id", b"commentId", b"comment-id", b"refId", b"ref-id"],
                    )
                    .unwrap_or_else(|| match kind {
                        HwpxNoteKind::Comment => {
                            comment_counter = comment_counter.saturating_add(1);
                            format!("hwpx-comment-{}", comment_counter)
                        }
                        HwpxNoteKind::Footnote => {
                            footnote_counter = footnote_counter.saturating_add(1);
                            format!("hwpx-footnote-{}", footnote_counter)
                        }
                        HwpxNoteKind::Endnote => {
                            endnote_counter = endnote_counter.saturating_add(1);
                            format!("hwpx-endnote-{}", endnote_counter)
                        }
                    });
                    note_stack.push(HwpxNoteState {
                        kind,
                        id,
                        author: attr_any(&e, &[b"author", b"writer"]),
                        date: attr_any(&e, &[b"date", b"created", b"time"]),
                        parent: attr_any(&e, &[b"parent", b"parentId", b"parent-id"]),
                        content: Vec::new(),
                        current_para: None,
                    });
                    continue;
                }
                if matches!(
                    local,
                    b"commentRef" | b"comment-ref" | b"annotationRef" | b"noteRef" | b"note-ref"
                ) {
                    if let Some(comment_id) = attr_any(&e, &[b"id", b"ref", b"refId", b"ref-id"]) {
                        let mut node = CommentReference::new(comment_id);
                        node.span = Some(SourceSpan::new(source));
                        let node_id = node.id;
                        store.insert(IRNode::CommentReference(node));
                        push_node_to_hwpx_context(
                            node_id,
                            &mut current_para,
                            &mut note_stack,
                            source,
                        );
                    }
                    continue;
                }
                if let Some(change_type) = revision_type_from_local(local) {
                    let mut revision = Revision::new(change_type);
                    revision.revision_id =
                        attr_any(&e, &[b"id", b"revId", b"revisionId", b"revision-id"]);
                    revision.author = attr_any(&e, &[b"author", b"writer"]);
                    revision.date = attr_any(&e, &[b"date", b"created", b"time"]);
                    revision.span = Some(SourceSpan::new(source));
                    revision_stack.push(revision);
                    continue;
                }
                if let Some(shape_id) = parse_hwpx_shape(&e, local, source, media_lookup, store) {
                    push_node_to_hwpx_context(shape_id, &mut current_para, &mut note_stack, source);
                    continue;
                }
                match local {
                    b"p" => {
                        if let Some(note) = note_stack.last_mut() {
                            finalize_note_paragraph(note, store);
                            let mut para = Paragraph::new();
                            para.span = Some(SourceSpan::new(source));
                            note.current_para = Some(para);
                        } else {
                            finalize_paragraph_hwpx(
                                &mut current_para,
                                &mut current_cell,
                                &mut content,
                                store,
                            );
                            let mut para = Paragraph::new();
                            para.span = Some(SourceSpan::new(source));
                            if let Some(style_id) =
                                attr_any(&e, &[b"styleId", b"style-id", b"style"])
                            {
                                para.style_id = Some(style_id);
                            }
                            if list_level > 0 {
                                para.properties.numbering = Some(NumberingInfo {
                                    num_id: 1,
                                    level: list_level - 1,
                                    format: None,
                                });
                            }
                            current_para = Some(para);
                        }
                    }
                    b"t" => {
                        in_text = true;
                    }
                    b"r" | b"run" | b"span" => {
                        run_props = Some(run_properties_from_attrs(&e));
                    }
                    b"tbl" | b"table" => {
                        finalize_paragraph_hwpx(
                            &mut current_para,
                            &mut current_cell,
                            &mut content,
                            store,
                        );
                        if current_table.is_none() {
                            current_table = Some(Table::new());
                        }
                    }
                    b"tr" | b"row" => {
                        if current_table.is_some() {
                            current_row = Some(TableRow::new());
                        }
                    }
                    b"tc" | b"cell" => {
                        if current_row.is_some() {
                            current_cell = Some(TableCell::new());
                        }
                    }
                    b"list" | b"ul" | b"ol" => {
                        list_level = list_level.saturating_add(1);
                    }
                    b"li" | b"list-item" => {
                        if note_stack.is_empty() && current_para.is_none() {
                            let mut para = Paragraph::new();
                            para.span = Some(SourceSpan::new(source));
                            if list_level > 0 {
                                para.properties.numbering = Some(NumberingInfo {
                                    num_id: 1,
                                    level: list_level - 1,
                                    format: None,
                                });
                            }
                            current_para = Some(para);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if matches!(
                    local,
                    b"commentRef" | b"comment-ref" | b"annotationRef" | b"noteRef" | b"note-ref"
                ) {
                    if let Some(comment_id) = attr_any(&e, &[b"id", b"ref", b"refId", b"ref-id"]) {
                        let mut node = CommentReference::new(comment_id);
                        node.span = Some(SourceSpan::new(source));
                        let node_id = node.id;
                        store.insert(IRNode::CommentReference(node));
                        push_node_to_hwpx_context(
                            node_id,
                            &mut current_para,
                            &mut note_stack,
                            source,
                        );
                    }
                    continue;
                }
                if let Some(shape_id) = parse_hwpx_shape(&e, local, source, media_lookup, store) {
                    push_node_to_hwpx_context(shape_id, &mut current_para, &mut note_stack, source);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name().as_ref().to_vec();
                let local = local_name(&name);
                if let Some(kind) = note_kind_from_local(local) {
                    if let Some(mut note) = note_stack.pop() {
                        finalize_note_paragraph(&mut note, store);
                        match kind {
                            HwpxNoteKind::Comment => {
                                let mut comment = Comment::new(note.id);
                                comment.author = note.author;
                                comment.date = note.date;
                                comment.parent_id = note.parent;
                                comment.content = note.content;
                                comment.span = Some(SourceSpan::new(source));
                                let comment_id = comment.id;
                                store.insert(IRNode::Comment(comment));
                                comments.push(comment_id);
                            }
                            HwpxNoteKind::Footnote => {
                                let mut footnote = Footnote::new(note.id);
                                footnote.content = note.content;
                                footnote.span = Some(SourceSpan::new(source));
                                let footnote_id = footnote.id;
                                store.insert(IRNode::Footnote(footnote));
                                footnotes.push(footnote_id);
                            }
                            HwpxNoteKind::Endnote => {
                                let mut endnote = Endnote::new(note.id);
                                endnote.content = note.content;
                                endnote.span = Some(SourceSpan::new(source));
                                let endnote_id = endnote.id;
                                store.insert(IRNode::Endnote(endnote));
                                endnotes.push(endnote_id);
                            }
                        }
                    }
                    continue;
                }
                if let Some(_change_type) = revision_type_from_local(local) {
                    if let Some(revision) = revision_stack.pop() {
                        let revision_id = revision.id;
                        store.insert(IRNode::Revision(revision));
                        push_node_to_hwpx_context(
                            revision_id,
                            &mut current_para,
                            &mut note_stack,
                            source,
                        );
                    }
                    continue;
                }
                match local {
                    b"p" => {
                        if let Some(note) = note_stack.last_mut() {
                            finalize_note_paragraph(note, store);
                        } else {
                            finalize_paragraph_hwpx(
                                &mut current_para,
                                &mut current_cell,
                                &mut content,
                                store,
                            );
                        }
                    }
                    b"t" => {
                        in_text = false;
                    }
                    b"r" | b"run" | b"span" => {
                        run_props = None;
                    }
                    b"tc" | b"cell" => {
                        finalize_cell_hwpx(&mut current_cell, &mut current_row, store);
                    }
                    b"tr" | b"row" => {
                        finalize_row_hwpx(&mut current_row, &mut current_table, store);
                    }
                    b"tbl" | b"table" => {
                        finalize_table_hwpx(
                            &mut current_table,
                            &mut current_cell,
                            &mut content,
                            store,
                        );
                    }
                    b"list" | b"ul" | b"ol" => {
                        list_level = list_level.saturating_sub(1);
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                if in_text {
                    let text = e.unescape().unwrap_or_default().to_string();
                    if !text.is_empty() {
                        let props = run_props.clone().unwrap_or_default();
                        let mut run = Run::with_properties(text, props);
                        run.span = Some(SourceSpan::new(source));
                        let run_id = run.id;
                        store.insert(IRNode::Run(run));
                        if let Some(revision) = revision_stack.last_mut() {
                            revision.content.push(run_id);
                        } else {
                            push_node_to_hwpx_context(
                                run_id,
                                &mut current_para,
                                &mut note_stack,
                                source,
                            );
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: source.to_string(),
                    message: e.to_string(),
                })
            }
            _ => {}
        }
        buf.clear();
    }

    while let Some(mut note) = note_stack.pop() {
        finalize_note_paragraph(&mut note, store);
        match note.kind {
            HwpxNoteKind::Comment => {
                let mut comment = Comment::new(note.id);
                comment.author = note.author;
                comment.date = note.date;
                comment.parent_id = note.parent;
                comment.content = note.content;
                comment.span = Some(SourceSpan::new(source));
                let comment_id = comment.id;
                store.insert(IRNode::Comment(comment));
                comments.push(comment_id);
            }
            HwpxNoteKind::Footnote => {
                let mut footnote = Footnote::new(note.id);
                footnote.content = note.content;
                footnote.span = Some(SourceSpan::new(source));
                let footnote_id = footnote.id;
                store.insert(IRNode::Footnote(footnote));
                footnotes.push(footnote_id);
            }
            HwpxNoteKind::Endnote => {
                let mut endnote = Endnote::new(note.id);
                endnote.content = note.content;
                endnote.span = Some(SourceSpan::new(source));
                let endnote_id = endnote.id;
                store.insert(IRNode::Endnote(endnote));
                endnotes.push(endnote_id);
            }
        }
    }

    while let Some(revision) = revision_stack.pop() {
        let revision_id = revision.id;
        store.insert(IRNode::Revision(revision));
        push_node_to_hwpx_context(revision_id, &mut current_para, &mut note_stack, source);
    }

    finalize_paragraph_hwpx(&mut current_para, &mut current_cell, &mut content, store);
    finalize_table_hwpx(&mut current_table, &mut current_cell, &mut content, store);

    Ok(content)
}

fn finalize_paragraph_hwpx(
    current_para: &mut Option<Paragraph>,
    current_cell: &mut Option<TableCell>,
    content: &mut Vec<NodeId>,
    store: &mut IrStore,
) {
    if let Some(para) = current_para.take() {
        let para_id = para.id;
        store.insert(IRNode::Paragraph(para));
        if let Some(cell) = current_cell.as_mut() {
            cell.content.push(para_id);
        } else {
            content.push(para_id);
        }
    }
}

fn finalize_note_paragraph(note: &mut HwpxNoteState, store: &mut IrStore) {
    if let Some(para) = note.current_para.take() {
        let para_id = para.id;
        store.insert(IRNode::Paragraph(para));
        note.content.push(para_id);
    }
}

fn push_node_to_hwpx_context(
    node_id: NodeId,
    current_para: &mut Option<Paragraph>,
    note_stack: &mut Vec<HwpxNoteState>,
    source: &str,
) {
    if let Some(note) = note_stack.last_mut() {
        if note.current_para.is_none() {
            let mut para = Paragraph::new();
            para.span = Some(SourceSpan::new(source));
            note.current_para = Some(para);
        }
        if let Some(para) = note.current_para.as_mut() {
            para.runs.push(node_id);
        }
    } else {
        if current_para.is_none() {
            let mut para = Paragraph::new();
            para.span = Some(SourceSpan::new(source));
            *current_para = Some(para);
        }
        if let Some(para) = current_para.as_mut() {
            para.runs.push(node_id);
        }
    }
}
