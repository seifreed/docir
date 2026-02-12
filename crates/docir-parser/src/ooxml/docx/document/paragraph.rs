use super::*;

pub(super) fn parse_paragraph(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<ParagraphParse, ParseError> {
    let mut para = Paragraph::new();
    let mut field_state = FieldState::new();
    let mut section_ref: Option<SectionRef> = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"w:pPr" => {
                    section_ref = parse_paragraph_properties(reader, &mut para, header_footer_map)?;
                }
                b"w:r" => {
                    let run = parse_run(parser, reader, rels)?;
                    let run_id = run.run_id;
                    para.runs.push(run_id);
                    for emb in &run.embedded {
                        para.runs.push(*emb);
                    }
                    update_field_from_run(&run, run_id, &mut field_state);
                    handle_field_char(
                        parser,
                        &mut para,
                        &mut field_state,
                        run.field_char.as_deref(),
                    );
                }
                b"w:hyperlink" => {
                    let link_id = parse_hyperlink(parser, reader, rels, &e)?;
                    para.runs.push(link_id);
                }
                b"w:sdt" => {
                    let sdt_id = parse_sdt(parser, reader, rels, SdtMode::Inline)?;
                    para.runs.push(sdt_id);
                }
                b"w:fldSimple" => {
                    let instr = attr_value(&e, b"w:instr");
                    let field_id = parse_field(parser, reader, instr)?;
                    para.runs.push(field_id);
                }
                b"w:commentRangeStart" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeStart::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeStart(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:commentRangeEnd" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeEnd::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeEnd(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:commentReference" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentReference::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentReference(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:ins" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::Insert)?;
                    para.runs.push(rev_id);
                }
                b"w:del" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::Delete)?;
                    para.runs.push(rev_id);
                }
                b"w:moveFrom" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::MoveFrom)?;
                    para.runs.push(rev_id);
                }
                b"w:moveTo" => {
                    let rev_id =
                        parse_revision_inline(parser, reader, rels, &e, RevisionType::MoveTo)?;
                    para.runs.push(rev_id);
                }
                b"w:pPrChange" | b"w:rPrChange" => {
                    let rev_id = parse_revision_inline(
                        parser,
                        reader,
                        rels,
                        &e,
                        RevisionType::FormatChange,
                    )?;
                    para.runs.push(rev_id);
                }
                b"w:bookmarkStart" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
                        bm.name = attr_value(&e, b"w:name");
                        bm.col_first = attr_value(&e, b"w:colFirst").and_then(|v| v.parse().ok());
                        bm.col_last = attr_value(&e, b"w:colLast").and_then(|v| v.parse().ok());
                        let bm_id = bm.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::BookmarkStart(bm));
                        para.runs.push(bm_id);
                    }
                }
                b"w:bookmarkEnd" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let bm = docir_core::ir::BookmarkEnd::new(bm_id);
                        let bm_id = bm.id;
                        parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
                        para.runs.push(bm_id);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:commentRangeStart" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeStart::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeStart(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:commentRangeEnd" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentRangeEnd::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentRangeEnd(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:commentReference" => {
                    if let Some(cid) = attr_value(&e, b"w:id") {
                        let mut node = CommentReference::new(cid);
                        node.span = Some(span_from_reader(reader, "word/document.xml"));
                        let node_id = node.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::CommentReference(node));
                        para.runs.push(node_id);
                    }
                }
                b"w:bookmarkStart" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let mut bm = docir_core::ir::BookmarkStart::new(bm_id);
                        bm.name = attr_value(&e, b"w:name");
                        bm.col_first = attr_value(&e, b"w:colFirst").and_then(|v| v.parse().ok());
                        bm.col_last = attr_value(&e, b"w:colLast").and_then(|v| v.parse().ok());
                        let bm_id = bm.id;
                        parser
                            .store
                            .insert(docir_core::ir::IRNode::BookmarkStart(bm));
                        para.runs.push(bm_id);
                    }
                }
                b"w:bookmarkEnd" => {
                    if let Some(bm_id) = attr_value(&e, b"w:id") {
                        let bm = docir_core::ir::BookmarkEnd::new(bm_id);
                        let bm_id = bm.id;
                        parser.store.insert(docir_core::ir::IRNode::BookmarkEnd(bm));
                        para.runs.push(bm_id);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:p" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    para.span = Some(span_from_reader(reader, "word/document.xml"));
    let id = para.id;
    parser.store.insert(docir_core::ir::IRNode::Paragraph(para));
    Ok(ParagraphParse { id, section_ref })
}

pub(super) fn parse_paragraph_simple(
    parser: &mut DocxParser,
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
) -> Result<NodeId, ParseError> {
    Ok(parse_paragraph(parser, reader, rels, None)?.id)
}

struct FieldState {
    active: bool,
    instr_done: bool,
    instr: String,
    runs: Vec<NodeId>,
}

impl FieldState {
    fn new() -> Self {
        Self {
            active: false,
            instr_done: false,
            instr: String::new(),
            runs: Vec::new(),
        }
    }

    fn start(&mut self) {
        self.active = true;
        self.instr_done = false;
        self.instr.clear();
        self.runs.clear();
    }

    fn separate(&mut self) {
        self.instr_done = true;
    }

    fn finish(&mut self) {
        self.active = false;
        self.instr_done = false;
        self.instr.clear();
        self.runs.clear();
    }
}

fn update_field_from_run(run: &RunParse, run_id: NodeId, state: &mut FieldState) {
    if state.active {
        state.runs.push(run_id);
        if run.has_instr && !state.instr_done {
            state.instr.push_str(&run.text);
        }
    }
}

fn handle_field_char(
    parser: &mut DocxParser,
    para: &mut Paragraph,
    state: &mut FieldState,
    char_type: Option<&str>,
) {
    match char_type {
        Some("begin") => state.start(),
        Some("separate") => state.separate(),
        Some("end") => {
            if state.active {
                let instr = if state.instr.trim().is_empty() {
                    None
                } else {
                    Some(state.instr.trim().to_string())
                };
                let mut field = Field::new(instr);
                field.runs = state.runs.clone();
                let field_id = field.id;
                parser.store.insert(docir_core::ir::IRNode::Field(field));
                para.runs.push(field_id);
            }
            state.finish();
        }
        _ => {}
    }
}

fn alignment_from_val(val: &str) -> TextAlignment {
    match val {
        "center" => TextAlignment::Center,
        "right" => TextAlignment::Right,
        "both" => TextAlignment::Justify,
        "distribute" => TextAlignment::Distribute,
        _ => TextAlignment::Left,
    }
}

pub(super) fn parse_paragraph_properties(
    reader: &mut Reader<&[u8]>,
    para: &mut Paragraph,
    header_footer_map: Option<&HashMap<String, NodeId>>,
) -> Result<Option<SectionRef>, ParseError> {
    let mut buf = Vec::new();
    let mut section_ref: Option<SectionRef> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"w:pStyle" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        para.style_id = Some(val);
                    }
                }
                b"w:keepNext" => {
                    para.properties.keep_next = Some(bool_from_val(&e));
                }
                b"w:keepLines" => {
                    para.properties.keep_lines = Some(bool_from_val(&e));
                }
                b"w:pageBreakBefore" => {
                    para.properties.page_break_before = Some(bool_from_val(&e));
                }
                b"w:widowControl" => {
                    para.properties.widow_control = Some(bool_from_val(&e));
                }
                b"w:jc" => {
                    if let Some(val) = attr_value(&e, b"w:val") {
                        para.properties.alignment = Some(alignment_from_val(val.as_str()));
                    }
                }
                b"w:ind" => {
                    let mut indent = para.properties.indentation.clone().unwrap_or_default();
                    if let Some(val) = attr_value(&e, b"w:left").and_then(|v| v.parse().ok()) {
                        indent.left = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:right").and_then(|v| v.parse().ok()) {
                        indent.right = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:firstLine").and_then(|v| v.parse().ok()) {
                        indent.first_line = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:hanging").and_then(|v| v.parse().ok()) {
                        indent.hanging = Some(val);
                    }
                    para.properties.indentation = Some(indent);
                }
                b"w:spacing" => {
                    let mut spacing = para.properties.spacing.clone().unwrap_or_default();
                    if let Some(val) = attr_value(&e, b"w:before").and_then(|v| v.parse().ok()) {
                        spacing.before = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:after").and_then(|v| v.parse().ok()) {
                        spacing.after = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:line").and_then(|v| v.parse().ok()) {
                        spacing.line = Some(val);
                    }
                    if let Some(val) = attr_value(&e, b"w:lineRule") {
                        spacing.line_rule = match val.as_str() {
                            "auto" => Some(LineSpacingRule::Auto),
                            "exact" => Some(LineSpacingRule::Exact),
                            "atLeast" => Some(LineSpacingRule::AtLeast),
                            _ => None,
                        };
                    }
                    para.properties.spacing = Some(spacing);
                }
                b"w:pBdr" => {
                    if let Some(borders) = parse_paragraph_borders(reader)? {
                        para.properties.borders = Some(borders);
                    }
                }
                b"w:outlineLvl" => {
                    if let Some(val) = attr_value(&e, b"w:val").and_then(|v| v.parse().ok()) {
                        para.properties.outline_level = Some(val);
                    }
                }
                b"w:numPr" => {
                    parse_numbering(reader, &mut para.properties)?;
                }
                b"w:sectPr" => {
                    section_ref = Some(apply_section_refs(reader, header_footer_map)?);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:pPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(section_ref)
}

pub(super) fn parse_paragraph_borders(
    reader: &mut Reader<&[u8]>,
) -> Result<Option<docir_core::ir::ParagraphBorders>, ParseError> {
    let mut buf = Vec::new();
    let mut borders = docir_core::ir::ParagraphBorders::default();
    let mut has_any = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let border = parse_border(&e);
                if border.is_none() {
                    continue;
                }
                match e.name().as_ref() {
                    b"w:top" => {
                        borders.top = border;
                        has_any = true;
                    }
                    b"w:bottom" => {
                        borders.bottom = border;
                        has_any = true;
                    }
                    b"w:left" => {
                        borders.left = border;
                        has_any = true;
                    }
                    b"w:right" => {
                        borders.right = border;
                        has_any = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:pBdr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "word/document.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }
    if has_any {
        Ok(Some(borders))
    } else {
        Ok(None)
    }
}
