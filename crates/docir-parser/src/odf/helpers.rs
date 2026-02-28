use super::*;

const ODF_CONTENT_XML: &str = "content.xml";

pub(super) fn parse_notes(reader: &mut OdfReader<'_>) -> Result<Option<String>, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match read_event(reader, &mut buf, ODF_CONTENT_XML)? {
            Event::Start(e) => {
                if e.name().as_ref() == b"text:p" {
                    let para = parse_text_element(reader, b"text:p")?;
                    if !text.is_empty() {
                        text.push('\n');
                    }
                    text.push_str(&para);
                }
            }
            Event::End(e) => {
                if e.name().as_ref() == b"presentation:notes" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if text.is_empty() {
        Ok(None)
    } else {
        Ok(Some(text))
    }
}
#[derive(Debug, Clone)]
pub(super) struct ValidationDef {
    pub(super) validation_type: Option<String>,
    pub(super) operator: Option<String>,
    pub(super) allow_blank: bool,
    pub(super) show_input_message: bool,
    pub(super) show_error_message: bool,
    pub(super) error_title: Option<String>,
    pub(super) error: Option<String>,
    pub(super) prompt_title: Option<String>,
    pub(super) prompt: Option<String>,
    pub(super) formula1: Option<String>,
    pub(super) formula2: Option<String>,
}

pub(super) fn parse_validation_definition(
    start: &BytesStart<'_>,
) -> Option<(String, ValidationDef)> {
    let name = attr_value(start, b"table:name")?;
    let condition = attr_value(start, b"table:condition");
    let allow_blank = attr_value(start, b"table:allow-empty-cell")
        .map(|v| v == "true")
        .unwrap_or(false);
    let show_input_message = attr_value(start, b"table:display-list")
        .map(|v| v == "true")
        .unwrap_or(false);
    let show_error_message = attr_value(start, b"table:display-list")
        .map(|v| v == "true")
        .unwrap_or(false);
    let def = ValidationDef {
        validation_type: condition.clone(),
        operator: None,
        allow_blank,
        show_input_message,
        show_error_message,
        error_title: None,
        error: None,
        prompt_title: None,
        prompt: None,
        formula1: condition,
        formula2: None,
    };
    Some((name, def))
}

pub(super) fn parse_ods_conditional_formatting(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<Option<ConditionalFormat>, ParseError> {
    let mut cf = init_conditional_format(start);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"table:conditional-format" {
                    let rule = build_ods_conditional_rule(&e);
                    cf.rules.push(rule);
                    spreadsheet::skip_element(reader, e.name().as_ref())?;
                }
            }
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"table:conditional-format" {
                    let rule = build_ods_conditional_rule(&e);
                    cf.rules.push(rule);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:conditional-formatting" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(ODF_CONTENT_XML, e)),
            _ => {}
        }
        buf.clear();
    }

    if cf.rules.is_empty() && cf.ranges.is_empty() {
        Ok(None)
    } else {
        Ok(Some(cf))
    }
}

pub(super) fn parse_ods_conditional_formatting_empty(
    start: &BytesStart<'_>,
) -> Result<Option<ConditionalFormat>, ParseError> {
    let cf = init_conditional_format(start);
    if cf.rules.is_empty() && cf.ranges.is_empty() {
        Ok(None)
    } else {
        Ok(Some(cf))
    }
}

fn init_conditional_format(start: &BytesStart<'_>) -> ConditionalFormat {
    let mut cf = ConditionalFormat {
        id: NodeId::new(),
        ranges: Vec::new(),
        rules: Vec::new(),
        span: Some(SourceSpan::new(ODF_CONTENT_XML)),
    };
    if let Some(ranges) = attr_value(start, b"table:target-range-address")
        .or_else(|| attr_value(start, b"table:cell-range-address"))
    {
        cf.ranges = ranges.split_whitespace().map(|s| s.to_string()).collect();
    }
    cf
}

pub(super) fn build_ods_conditional_rule(start: &BytesStart<'_>) -> ConditionalRule {
    let mut rule = ConditionalRule {
        rule_type: "odf-condition".to_string(),
        priority: None,
        operator: None,
        formulae: Vec::new(),
    };
    rule.priority = attr_value(start, b"table:priority").and_then(|v| v.parse::<u32>().ok());
    if let Some(condition) = attr_value(start, b"table:condition") {
        rule.operator = parse_odf_condition_operator(&condition);
        rule.formulae.push(condition);
    }
    if let Some(style_name) = attr_value(start, b"table:apply-style-name") {
        rule.formulae.push(format!("apply-style:{}", style_name));
    }
    rule
}

pub(super) fn parse_odf_condition_operator(condition: &str) -> Option<String> {
    let lower = condition.to_ascii_lowercase();
    if let Some(idx) = lower.find("cell-content-is-") {
        let rest = &lower[idx + "cell-content-is-".len()..];
        let op = rest.split('(').next().unwrap_or(rest);
        return Some(op.to_string());
    }
    if let Some(idx) = lower.find("is-true-formula") {
        let _ = idx;
        return Some("true-formula".to_string());
    }
    if let Some(idx) = lower.find("formula-is") {
        let _ = idx;
        return Some("formula".to_string());
    }
    None
}

#[derive(Debug, Clone)]
pub(super) struct OdsRow {
    pub(super) cells: Vec<OdsCellData>,
}

#[derive(Debug, Clone)]
pub(super) struct OdsCellData {
    pub(super) value: CellValue,
    pub(super) formula: Option<CellFormula>,
    pub(super) style_id: Option<u32>,
    pub(super) col_repeat: u32,
    pub(super) validation_name: Option<String>,
    pub(super) col_span: Option<u32>,
    pub(super) row_span: Option<u32>,
    pub(super) is_covered: bool,
}

impl OdsCellData {
    pub(super) fn should_emit(&self) -> bool {
        !matches!(self.value, CellValue::Empty) || self.formula.is_some()
    }

    pub(super) fn merge_range(&self, row: u32, col: u32) -> Option<MergedCellRange> {
        let col_span = self.col_span.unwrap_or(1);
        let row_span = self.row_span.unwrap_or(1);
        if col_span > 1 || row_span > 1 {
            Some(MergedCellRange {
                start_col: col,
                start_row: row,
                end_col: col + col_span - 1,
                end_row: row + row_span - 1,
            })
        } else {
            None
        }
    }
}

pub(super) fn parse_ods_row(
    reader: &mut OdfReader<'_>,
    _start: &BytesStart<'_>,
    store: &mut IrStore,
    style_map: &mut HashMap<String, u32>,
    next_style_id: &mut u32,
) -> Result<OdsRow, ParseError> {
    let mut buf = Vec::new();
    let mut cells = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    let cell = parse_ods_cell(reader, &e, store, style_map, next_style_id)?;
                    cells.push(cell);
                }
                b"table:covered-table-cell" => {
                    let cell = parse_ods_covered_cell(reader, &e)?;
                    cells.push(cell);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    let cell = parse_ods_cell_empty(&e, style_map, next_style_id)?;
                    cells.push(cell);
                }
                b"table:covered-table-cell" => {
                    let cell = parse_ods_covered_cell_empty(&e)?;
                    cells.push(cell);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table-row" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(ODF_CONTENT_XML, e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(OdsRow { cells })
}

pub(super) fn parse_ods_covered_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
) -> Result<OdsCellData, ParseError> {
    let cell = covered_cell_from_start(start);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::End(e)) if e.name().as_ref() == b"table:covered-table-cell" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(ODF_CONTENT_XML, e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(cell)
}

pub(super) fn parse_ods_covered_cell_empty(
    start: &BytesStart<'_>,
) -> Result<OdsCellData, ParseError> {
    Ok(covered_cell_from_start(start))
}

fn covered_cell_from_start(start: &BytesStart<'_>) -> OdsCellData {
    let col_repeat = attr_value(start, b"table:number-columns-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let col_span =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok());
    let row_span =
        attr_value(start, b"table:number-rows-spanned").and_then(|v| v.parse::<u32>().ok());
    OdsCellData {
        value: CellValue::Empty,
        formula: None,
        style_id: None,
        col_repeat,
        validation_name: None,
        col_span,
        row_span,
        is_covered: true,
    }
}

fn append_text_control(text: &mut String, e: &BytesStart<'_>) {
    match e.name().as_ref() {
        b"text:s" => {
            let count = attr_value(e, b"text:c")
                .and_then(|v| v.parse::<usize>().ok())
                .unwrap_or(1);
            text.extend(std::iter::repeat(' ').take(count));
        }
        b"text:tab" => text.push('\t'),
        b"text:line-break" => text.push('\n'),
        _ => {}
    }
}

pub(super) fn parse_text_element(
    reader: &mut OdfReader<'_>,
    end_name: &[u8],
) -> Result<String, ParseError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => append_text_control(&mut text, &e),
            Ok(Event::Text(e)) => {
                let chunk = e.unescape().unwrap_or_default();
                text.push_str(&chunk);
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_name {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(ODF_CONTENT_XML, e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

pub(super) fn column_index_to_name(mut index: u32) -> String {
    let mut name = String::new();
    index += 1;
    while index > 0 {
        let rem = ((index - 1) % 26) as u8;
        name.push((b'A' + rem) as char);
        index = (index - 1) / 26;
    }
    name.chars().rev().collect()
}

#[derive(Debug, Clone, Copy)]
pub(super) struct ListContext {
    pub(super) num_id: u32,
    pub(super) level: u32,
}

pub(super) fn parse_table(
    reader: &mut OdfReader<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut buf = Vec::new();
    let mut table = Table::new();
    let mut current_row: Option<TableRow> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    current_row = Some(TableRow::new());
                }
                b"table:table-cell" => {
                    let cell_id = parse_table_cell(reader, &e, store, limits)?;
                    if let Some(row) = current_row.as_mut() {
                        row.cells.push(cell_id);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    let row = TableRow::new();
                    let row_id = row.id;
                    store.insert(IRNode::TableRow(row));
                    table.rows.push(row_id);
                }
                b"table:table-cell" => {
                    let mut cell = TableCell::new();
                    if let Some(span) = attr_value(&e, b"table:number-columns-spanned")
                        .and_then(|v| v.parse::<u32>().ok())
                    {
                        let mut props = TableCellProperties::default();
                        props.grid_span = Some(span);
                        cell.properties = props;
                    }
                    let cell_id = cell.id;
                    store.insert(IRNode::TableCell(cell));
                    if let Some(row) = current_row.as_mut() {
                        row.cells.push(cell_id);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"table:table-row" => {
                    if let Some(row) = current_row.take() {
                        let row_id = row.id;
                        store.insert(IRNode::TableRow(row));
                        table.rows.push(row_id);
                    }
                }
                b"table:table" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(ODF_CONTENT_XML, e)),
            _ => {}
        }
        buf.clear();
    }

    let table_id = table.id;
    store.insert(IRNode::Table(table));
    Ok(table_id)
}

pub(super) fn parse_table_cell(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
    limits: &dyn OdfLimitCounter,
) -> Result<NodeId, ParseError> {
    let mut cell = TableCell::new();
    if let Some(span) =
        attr_value(start, b"table:number-columns-spanned").and_then(|v| v.parse::<u32>().ok())
    {
        let mut props = TableCellProperties::default();
        props.grid_span = Some(span);
        cell.properties = props;
    }

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
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
                    cell.content.push(paragraph_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"table:table-cell" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(ODF_CONTENT_XML, e)),
            _ => {}
        }
        buf.clear();
    }

    let cell_id = cell.id;
    store.insert(IRNode::TableCell(cell));
    Ok(cell_id)
}

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

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
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
            Ok(Event::Text(e)) => {
                if let Some(field) = current {
                    let value = e.unescape().unwrap_or_default().to_string();
                    match field {
                        AnnotationField::Creator => comment.author = Some(value),
                        AnnotationField::Date => comment.date = Some(value),
                    }
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"office:annotation" => break,
                b"dc:creator" | b"dc:date" => current = None,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(ODF_CONTENT_XML, e)),
            _ => {}
        }
        buf.clear();
    }

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

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
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
                    content.push(paragraph_id);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"text:note" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(ODF_CONTENT_XML, e)),
            _ => {}
        }
        buf.clear();
    }

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

pub(super) fn parse_draw_frame(
    reader: &mut OdfReader<'_>,
    start: &BytesStart<'_>,
    store: &mut IrStore,
) -> Result<Option<NodeId>, ParseError> {
    let mut shape = Shape::new(ShapeType::Picture);
    shape.transform = parse_frame_transform(start);
    shape.name = attr_value(start, b"draw:name");
    let mut buf = Vec::new();
    let mut has_shape = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"draw:image" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        shape.media_target = Some(href);
                        shape.shape_type = ShapeType::Picture;
                        has_shape = true;
                    }
                }
                b"draw:object" | b"draw:object-ole" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        shape.media_target = Some(href);
                    }
                    shape.shape_type = ShapeType::OleObject;
                    has_shape = true;
                }
                b"draw:plugin" => {
                    if let Some(href) = attr_value(&e, b"xlink:href") {
                        shape.media_target = Some(href.clone());
                        shape.shape_type = classify_media_shape(&href);
                        has_shape = true;
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"draw:frame" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(ODF_CONTENT_XML, e)),
            _ => {}
        }
        buf.clear();
    }

    if has_shape {
        let shape_id = shape.id;
        store.insert(IRNode::Shape(shape));
        Ok(Some(shape_id))
    } else {
        Ok(None)
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

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
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
            Ok(Event::Text(e)) => {
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
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"text:insertion" | b"text:deletion" => {
                    if let Some(rev) = current_revision.take() {
                        let id = rev.id;
                        store.insert(IRNode::Revision(rev));
                        revisions.push(id);
                    }
                }
                b"text:tracked-changes" => break,
                b"dc:creator" | b"dc:date" => current_field = None,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(ODF_CONTENT_XML, e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(revisions)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_validation_definition_reads_contract_fields() {
        let mut start = BytesStart::new("table:content-validation");
        start.push_attribute(("table:name", "val1"));
        start.push_attribute(("table:condition", "cell-content-is-between(1,10)"));
        start.push_attribute(("table:allow-empty-cell", "true"));
        start.push_attribute(("table:display-list", "true"));

        let (name, def) = parse_validation_definition(&start).expect("validation should parse");
        assert_eq!(name, "val1");
        assert_eq!(
            def.validation_type.as_deref(),
            Some("cell-content-is-between(1,10)")
        );
        assert!(def.allow_blank);
        assert!(def.show_input_message);
        assert!(def.show_error_message);
        assert_eq!(
            def.formula1.as_deref(),
            Some("cell-content-is-between(1,10)")
        );
    }

    #[test]
    fn parse_odf_condition_operator_handles_known_forms() {
        assert_eq!(
            parse_odf_condition_operator("cell-content-is-greater-than(5)"),
            Some("greater-than".to_string())
        );
        assert_eq!(
            parse_odf_condition_operator("is-true-formula([.A1]>0)"),
            Some("true-formula".to_string())
        );
        assert_eq!(
            parse_odf_condition_operator("formula-is([.A1]>0)"),
            Some("formula".to_string())
        );
        assert_eq!(parse_odf_condition_operator("unknown()"), None);
    }

    #[test]
    fn parse_text_element_preserves_spacing_controls() {
        let xml = br#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">A<text:s text:c="2"/><text:tab/><text:line-break/>B</text:p>"#;
        let mut reader = Reader::from_reader(std::io::Cursor::new(xml.as_slice()));
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let start = loop {
            match reader.read_event_into(&mut buf).expect("event read") {
                Event::Start(e) if e.name().as_ref() == b"text:p" => break e.into_owned(),
                Event::Eof => panic!("missing text:p start"),
                _ => {}
            }
            buf.clear();
        };

        let parsed = parse_text_element(&mut reader, start.name().as_ref()).expect("parse text");
        assert_eq!(parsed, "A  \t\nB");
    }

    #[test]
    fn parse_ods_covered_cell_empty_sets_repeat_and_span() {
        let mut start = BytesStart::new("table:covered-table-cell");
        start.push_attribute(("table:number-columns-repeated", "3"));
        start.push_attribute(("table:number-columns-spanned", "2"));
        start.push_attribute(("table:number-rows-spanned", "2"));

        let cell = parse_ods_covered_cell_empty(&start).expect("covered cell parse");
        assert!(cell.is_covered);
        assert_eq!(cell.col_repeat, 3);
        let merge = cell.merge_range(4, 5).expect("expected merged range");
        assert_eq!(merge.start_row, 4);
        assert_eq!(merge.start_col, 5);
        assert_eq!(merge.end_row, 5);
        assert_eq!(merge.end_col, 6);
    }
}
