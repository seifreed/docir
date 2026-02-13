use super::*;
use crate::ooxml::part_utils::{read_relationships, read_xml_part_and_rels};
use crate::xml_utils::{reader_from_str, xml_error};
use crate::zip_handler::PackageReader;
use docir_core::ir::{DataValidation, SheetPageMargins};
use quick_xml::events::BytesStart;

struct WorksheetParseAccum {
    columns: HashMap<u32, ColumnDefinition>,
    merged_cells: Vec<MergedCellRange>,
    cells: Vec<NodeId>,
    conditional_formats: Vec<NodeId>,
    data_validations: Vec<NodeId>,
}

impl WorksheetParseAccum {
    fn new() -> Self {
        Self {
            columns: HashMap::new(),
            merged_cells: Vec::new(),
            cells: Vec::new(),
            conditional_formats: Vec::new(),
            data_validations: Vec::new(),
        }
    }
}

impl XlsxParser {
    pub(super) fn parse_worksheet(
        &mut self,
        zip: &mut impl PackageReader,
        xml: &str,
        sheet: &SheetInfo,
        sheet_path: &str,
        relationships: &Relationships,
        kind: SheetKind,
    ) -> Result<NodeId, ParseError> {
        let mut worksheet = Worksheet::new(sheet.name.clone(), sheet.sheet_id);
        worksheet.state = sheet.state;
        worksheet.relationship_id = Some(sheet.rel_id.clone());
        worksheet.kind = kind;
        worksheet.span = Some(SourceSpan::new(sheet_path));

        self.current_sheet_kind = Some(kind);
        self.current_sheet_name = Some(sheet.name.clone());
        self.current_sheet_state = Some(sheet.state);
        self.current_xlm_index = None;

        if kind == SheetKind::MacroSheet {
            self.begin_macro_sheet(sheet);
        }

        let mut accum = self.parse_worksheet_xml(xml, sheet_path, relationships, &mut worksheet)?;

        // Deterministic ordering
        let mut columns_sorted: Vec<ColumnDefinition> = accum.columns.into_values().collect();
        columns_sorted.sort_by_key(|c| c.index);
        accum
            .merged_cells
            .sort_by_key(|r| (r.start_row, r.start_col, r.end_row, r.end_col));

        worksheet.columns = columns_sorted;
        worksheet.merged_cells = accum.merged_cells;
        worksheet.cells = accum.cells;
        worksheet.conditional_formats = accum.conditional_formats;
        worksheet.data_validations = accum.data_validations;

        worksheet.drawings = self.load_worksheet_drawings(zip, sheet_path, relationships)?;
        worksheet.tables = self.load_worksheet_tables(zip, sheet_path, relationships)?;
        worksheet.pivot_tables = self.load_worksheet_pivots(zip, sheet_path, relationships)?;
        worksheet.comments =
            self.load_worksheet_comments(zip, sheet_path, relationships, &sheet.name)?;

        if kind == SheetKind::ChartSheet {
            if let Some(drawing) = self.parse_chartsheet(zip, xml, sheet_path, relationships)? {
                worksheet.drawings.push(drawing);
            }
        }

        self.current_sheet_kind = None;
        self.current_sheet_name = None;
        self.current_sheet_state = None;
        self.current_xlm_index = None;

        let ws_id = worksheet.id;
        self.store.insert(IRNode::Worksheet(worksheet));
        Ok(ws_id)
    }

    fn load_worksheet_drawings(
        &mut self,
        zip: &mut impl PackageReader,
        sheet_path: &str,
        relationships: &Relationships,
    ) -> Result<Vec<NodeId>, ParseError> {
        let mut drawings = Vec::new();
        for rel in relationships.get_by_type(rel_type::DRAWING) {
            let drawing_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&drawing_path) {
                continue;
            }
            let (drawing_xml, drawing_rels) = read_xml_part_and_rels(zip, &drawing_path)?;
            let drawing_id = self.parse_drawing(&drawing_xml, &drawing_path, &drawing_rels, zip)?;
            drawings.push(drawing_id);
        }
        Ok(drawings)
    }

    fn load_worksheet_tables(
        &mut self,
        zip: &mut impl PackageReader,
        sheet_path: &str,
        relationships: &Relationships,
    ) -> Result<Vec<NodeId>, ParseError> {
        self.load_worksheet_parts(
            zip,
            sheet_path,
            relationships,
            rel_type::TABLE,
            |parser, zip, table_path, rel_id| parser.load_worksheet_table(zip, table_path, rel_id),
        )
    }

    fn load_worksheet_pivots(
        &mut self,
        zip: &mut impl PackageReader,
        sheet_path: &str,
        relationships: &Relationships,
    ) -> Result<Vec<NodeId>, ParseError> {
        self.load_worksheet_parts(
            zip,
            sheet_path,
            relationships,
            rel_type::PIVOT_TABLE,
            |parser, zip, pivot_path, rel_id| parser.load_worksheet_pivot(zip, pivot_path, rel_id),
        )
    }

    fn load_worksheet_comments(
        &mut self,
        zip: &mut impl PackageReader,
        sheet_path: &str,
        relationships: &Relationships,
        sheet_name: &str,
    ) -> Result<Vec<NodeId>, ParseError> {
        let mut comments = Vec::new();
        self.load_worksheet_comment_type(
            zip,
            sheet_path,
            relationships,
            sheet_name,
            rel_type::COMMENTS,
            parse_sheet_comments,
            &mut comments,
        )?;
        self.load_worksheet_comment_type(
            zip,
            sheet_path,
            relationships,
            sheet_name,
            rel_type::THREADED_COMMENTS,
            parse_threaded_comments,
            &mut comments,
        )?;
        Ok(comments)
    }

    fn load_worksheet_parts<R, F>(
        &mut self,
        zip: &mut R,
        sheet_path: &str,
        relationships: &Relationships,
        rel_type: &str,
        mut loader: F,
    ) -> Result<Vec<NodeId>, ParseError>
    where
        R: PackageReader,
        F: FnMut(&mut Self, &mut R, &str, &str) -> Result<NodeId, ParseError>,
    {
        let mut ids = Vec::new();
        for rel in relationships.get_by_type(rel_type) {
            let part_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&part_path) {
                continue;
            }
            let id = loader(self, zip, &part_path, &rel.id)?;
            ids.push(id);
        }
        Ok(ids)
    }

    fn load_worksheet_table(
        &mut self,
        zip: &mut impl PackageReader,
        table_path: &str,
        rel_id: &str,
    ) -> Result<NodeId, ParseError> {
        let table_xml = zip.read_file_string(table_path)?;
        let mut table = parse_table_definition(&table_xml, table_path)?;
        table.span = Some(SourceSpan::new(table_path).with_relationship(rel_id.to_string()));
        let id = table.id;
        self.store.insert(IRNode::TableDefinition(table));
        Ok(id)
    }

    fn load_worksheet_pivot(
        &mut self,
        zip: &mut impl PackageReader,
        pivot_path: &str,
        rel_id: &str,
    ) -> Result<NodeId, ParseError> {
        let pivot_xml = zip.read_file_string(pivot_path)?;
        let mut pivot = parse_pivot_table_definition(&pivot_xml, pivot_path)?;
        pivot.span = Some(SourceSpan::new(pivot_path).with_relationship(rel_id.to_string()));
        let id = pivot.id;
        self.store.insert(IRNode::PivotTable(pivot));
        Ok(id)
    }

    fn load_worksheet_comment_type(
        &mut self,
        zip: &mut impl PackageReader,
        sheet_path: &str,
        relationships: &Relationships,
        sheet_name: &str,
        rel_type: &str,
        parse_fn: fn(&str, &str, Option<&str>) -> Result<Vec<SheetComment>, ParseError>,
        out: &mut Vec<NodeId>,
    ) -> Result<(), ParseError> {
        for rel in relationships.get_by_type(rel_type) {
            let comments_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&comments_path) {
                continue;
            }
            let comments_xml = zip.read_file_string(&comments_path)?;
            let parsed = parse_fn(&comments_xml, &comments_path, Some(sheet_name))?;
            self.insert_sheet_comments(parsed, &comments_path, &rel.id, out);
        }
        Ok(())
    }

    fn insert_sheet_comments(
        &mut self,
        parsed: Vec<SheetComment>,
        comments_path: &str,
        rel_id: &str,
        out: &mut Vec<NodeId>,
    ) {
        for mut comment in parsed {
            comment.span =
                Some(SourceSpan::new(comments_path).with_relationship(rel_id.to_string()));
            let id = comment.id;
            self.store.insert(IRNode::SheetComment(comment));
            out.push(id);
        }
    }

    fn parse_worksheet_xml(
        &mut self,
        xml: &str,
        sheet_path: &str,
        relationships: &Relationships,
        worksheet: &mut Worksheet,
    ) -> Result<WorksheetParseAccum, ParseError> {
        let mut accum = WorksheetParseAccum::new();
        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => self.handle_worksheet_start(
                    &e,
                    &mut reader,
                    sheet_path,
                    relationships,
                    worksheet,
                    &mut accum,
                )?,
                Ok(Event::Empty(e)) => self.handle_worksheet_empty(
                    &e,
                    sheet_path,
                    relationships,
                    worksheet,
                    &mut accum,
                )?,
                Ok(Event::Eof) => break,
                Err(e) => return Err(xml_error(sheet_path, e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(accum)
    }

    fn handle_worksheet_start(
        &mut self,
        e: &BytesStart<'_>,
        reader: &mut quick_xml::Reader<&[u8]>,
        sheet_path: &str,
        relationships: &Relationships,
        worksheet: &mut Worksheet,
        accum: &mut WorksheetParseAccum,
    ) -> Result<(), ParseError> {
        if handle_worksheet_common_tag(e, sheet_path, relationships, worksheet, accum, self) {
            return Ok(());
        }
        match e.name().as_ref() {
            b"c" => {
                let cell = self.parse_cell(reader, e, sheet_path)?;
                self.insert_cell(cell, accum);
            }
            b"conditionalFormatting" => {
                let fmt = parse_conditional_formatting(reader, e, sheet_path)?;
                self.insert_conditional_format(fmt, accum);
            }
            b"dataValidations" => {
                let vals = parse_data_validations(reader, sheet_path)?;
                for val in vals {
                    self.insert_data_validation(val, accum);
                }
            }
            _ => {}
        }
        Ok(())
    }

    fn handle_worksheet_empty(
        &mut self,
        e: &BytesStart<'_>,
        sheet_path: &str,
        relationships: &Relationships,
        worksheet: &mut Worksheet,
        accum: &mut WorksheetParseAccum,
    ) -> Result<(), ParseError> {
        if handle_worksheet_common_tag(e, sheet_path, relationships, worksheet, accum, self) {
            return Ok(());
        }
        match e.name().as_ref() {
            b"c" => {
                let cell = self.parse_empty_cell(e, sheet_path)?;
                self.insert_cell(cell, accum);
            }
            b"conditionalFormatting" => {
                let fmt = parse_conditional_formatting_empty(e, sheet_path);
                self.insert_conditional_format(fmt, accum);
            }
            b"dataValidations" => {
                // Empty container, nothing to add
            }
            _ => {}
        }
        Ok(())
    }

    fn insert_cell(&mut self, cell: Cell, accum: &mut WorksheetParseAccum) {
        let cell_id = cell.id;
        self.store.insert(IRNode::Cell(cell));
        accum.cells.push(cell_id);
    }

    fn insert_conditional_format(
        &mut self,
        fmt: ConditionalFormat,
        accum: &mut WorksheetParseAccum,
    ) {
        let id = fmt.id;
        self.store.insert(IRNode::ConditionalFormat(fmt));
        accum.conditional_formats.push(id);
    }

    fn insert_data_validation(
        &mut self,
        validation: DataValidation,
        accum: &mut WorksheetParseAccum,
    ) {
        let id = validation.id;
        self.store.insert(IRNode::DataValidation(validation));
        accum.data_validations.push(id);
    }

    fn parse_chartsheet(
        &mut self,
        zip: &mut impl PackageReader,
        xml: &str,
        sheet_path: &str,
        relationships: &Relationships,
    ) -> Result<Option<NodeId>, ParseError> {
        let mut reader = reader_from_str(xml);
        let mut buf = Vec::new();

        let mut chart_rel: Option<String> = None;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.name().as_ref().ends_with(b"chart") {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r:id" {
                                chart_rel = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: sheet_path.to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        let Some(rel_id) = chart_rel else {
            return Ok(None);
        };
        let Some(rel) = relationships.get(&rel_id) else {
            return Ok(None);
        };
        let chart_path = Relationships::resolve_target(sheet_path, &rel.target);
        if !zip.contains(&chart_path) {
            return Ok(None);
        }
        let chart_xml = zip.read_file_string(&chart_path)?;
        let chart_id = self.parse_chart(&chart_xml, &chart_path);

        let mut drawing = WorksheetDrawing::new();
        drawing.span = Some(SourceSpan::new(sheet_path));
        let mut shape = Shape::new(ShapeType::Chart);
        shape.media_target = Some(chart_path);
        shape.relationship_id = Some(rel_id);
        shape.span = Some(SourceSpan::new(sheet_path));
        let shape_id = shape.id;
        self.store.insert(IRNode::Shape(shape));
        drawing.shapes.push(shape_id);

        if let Some(chart_id) = chart_id {
            self.chart_nodes.push(chart_id);
        }

        let drawing_id = drawing.id;
        self.store.insert(IRNode::WorksheetDrawing(drawing));
        Ok(Some(drawing_id))
    }

    pub(super) fn parse_external_links_and_connections(
        &mut self,
        zip: &mut impl PackageReader,
        workbook_path: &str,
        workbook_rels: &Relationships,
    ) -> Result<(), ParseError> {
        // externalLink parts
        for rel in workbook_rels.get_by_type(rel_type::EXTERNAL_LINK) {
            let external_path = Relationships::resolve_target(workbook_path, &rel.target);
            if !zip.contains(&external_path) {
                continue;
            }
            let rels = read_relationships(zip, &external_path)?;
            if let Ok(xml) = zip.read_file_string(&external_path) {
                if let Ok(mut part) = parse_external_link_part(&xml, &external_path, Some(&rels)) {
                    part.span = Some(SourceSpan::new(&external_path));
                    let part_id = part.id;
                    self.store.insert(IRNode::ExternalLinkPart(part));
                    self.push_shared_part(part_id);
                }
            }
            for ext in rels.by_id.values() {
                let target = &ext.target;
                let ext_ref = ExternalReference::new(ExternalRefType::DataConnection, target);
                let ext_ref = ExternalReference {
                    relationship_id: Some(ext.id.clone()),
                    relationship_type: Some(ext.rel_type.clone()),
                    ..ext_ref
                };
                self.push_external_reference(ext_ref);
            }
        }

        // connections.xml
        if zip.contains("xl/connections.xml") {
            let xml = zip.read_file_string("xl/connections.xml")?;
            if let Ok(mut part) = parse_connections_part(&xml, "xl/connections.xml") {
                part.span = Some(SourceSpan::new("xl/connections.xml"));
                let part_id = part.id;
                let targets = connection_targets(&part);
                self.store.insert(IRNode::ConnectionPart(part));
                self.push_shared_part(part_id);
                for target in targets {
                    let ext_ref = ExternalReference::new(ExternalRefType::DataConnection, target);
                    self.push_external_reference(ext_ref);
                }
            }
        }

        Ok(())
    }

    fn push_shared_part(&mut self, part_id: NodeId) {
        if let Some(IRNode::Document(doc)) = self.store.get_mut(self.root_id) {
            doc.shared_parts.push(part_id);
        }
    }

    fn push_external_reference(&mut self, ext_ref: ExternalReference) {
        let id = ext_ref.id;
        self.store.insert(IRNode::ExternalReference(ext_ref));
        self.security_info.external_refs.push(id);
    }

    pub(super) fn parse_pivot_cache(
        &mut self,
        zip: &mut impl PackageReader,
        xml: &str,
        cache_path: &str,
        cache_id: u32,
    ) -> Result<PivotCache, ParseError> {
        let mut reader = reader_from_str(xml);

        let mut cache = PivotCache::new(cache_id);
        cache.span = Some(SourceSpan::new(cache_path));

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"cacheSource" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"type" {
                                cache.cache_source =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            if attr.key.as_ref() == b"connectionId" {
                                let conn = String::from_utf8_lossy(&attr.value).to_string();
                                let src = format!("connection:{conn}");
                                cache.cache_source = Some(src);
                            }
                        }
                    }
                    b"worksheetSource" => {
                        let mut sheet = None;
                        let mut range = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"sheet" => {
                                    sheet = Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"ref" => {
                                    range = Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                _ => {}
                            }
                        }
                        if let Some(sheet) = sheet {
                            let range = range.unwrap_or_else(|| "-".to_string());
                            cache.cache_source = Some(format!("worksheet:{sheet}!{range}"));
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) if e.name().as_ref() == b"pivotCacheDefinition" => break,
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: cache_path.to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        // Try to read pivot cache records
        let rels = read_relationships(zip, cache_path)?;
        if let Some(rel) = rels.get_first_by_type(rel_type::PIVOT_CACHE_RECORDS) {
            let records_path = Relationships::resolve_target(cache_path, &rel.target);
            if zip.contains(&records_path) {
                let records_xml = zip.read_file_string(&records_path)?;
                let mut records = parse_pivot_cache_records(&records_xml, &records_path)?;
                records.cache_id = Some(cache_id);
                cache.record_count = records.record_count;
                let rec_id = records.id;
                self.store.insert(IRNode::PivotCacheRecords(records));
                cache.records = Some(rec_id);
            }
        }

        Ok(cache)
    }
}

fn handle_worksheet_common_tag(
    e: &BytesStart<'_>,
    sheet_path: &str,
    relationships: &Relationships,
    worksheet: &mut Worksheet,
    accum: &mut WorksheetParseAccum,
    parser: &mut XlsxParser,
) -> bool {
    match e.name().as_ref() {
        b"dimension" => {
            if let Some(val) = attr_value(e, b"ref") {
                worksheet.dimension = Some(val);
            }
            true
        }
        b"tabColor" => {
            worksheet.tab_color = parse_color_attr(e);
            true
        }
        b"pageMargins" => {
            worksheet.page_margins = parse_page_margins(e);
            true
        }
        b"col" => {
            parse_column(e, &mut accum.columns);
            true
        }
        b"mergeCell" => {
            if let Some(range) = parse_merge_cell(e) {
                accum.merged_cells.push(range);
            }
            true
        }
        b"hyperlink" => {
            parser.handle_hyperlink(e, relationships, sheet_path);
            true
        }
        _ => false,
    }
}

fn parse_page_margins(start: &BytesStart) -> Option<SheetPageMargins> {
    let mut margins = SheetPageMargins {
        left: None,
        right: None,
        top: None,
        bottom: None,
        header: None,
        footer: None,
    };
    let mut found = false;
    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"left" => {
                margins.left = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            b"right" => {
                margins.right = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            b"top" => {
                margins.top = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            b"bottom" => {
                margins.bottom = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            b"header" => {
                margins.header = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            b"footer" => {
                margins.footer = String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                found = true;
            }
            _ => {}
        }
    }
    if found {
        Some(margins)
    } else {
        None
    }
}

fn parse_conditional_formatting_empty(start: &BytesStart, sheet_path: &str) -> ConditionalFormat {
    let mut ranges: Vec<String> = Vec::new();
    for attr in start.attributes().flatten() {
        if attr.key.as_ref() == b"sqref" {
            let val = String::from_utf8_lossy(&attr.value).to_string();
            ranges = val.split_whitespace().map(|s| s.to_string()).collect();
        }
    }
    ConditionalFormat {
        id: NodeId::new(),
        ranges,
        rules: Vec::new(),
        span: Some(SourceSpan::new(sheet_path)),
    }
}

fn parse_data_validations(
    reader: &mut Reader<&[u8]>,
    sheet_path: &str,
) -> Result<Vec<DataValidation>, ParseError> {
    let mut validations: Vec<DataValidation> = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"dataValidation" => {
                let val = parse_data_validation(reader, &e, sheet_path)?;
                validations.push(val);
            }
            Ok(Event::Empty(e)) if e.name().as_ref() == b"dataValidation" => {
                let val = parse_data_validation_empty(&e, sheet_path);
                validations.push(val);
            }
            Ok(Event::End(e)) if e.name().as_ref() == b"dataValidations" => break,
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: sheet_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(validations)
}

fn parse_data_validation(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<DataValidation, ParseError> {
    let mut validation = parse_data_validation_empty(start, sheet_path);

    let mut in_formula1 = false;
    let mut in_formula2 = false;
    let mut formula1 = String::new();
    let mut formula2 = String::new();

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"formula1" => {
                    in_formula1 = true;
                    formula1.clear();
                }
                b"formula2" => {
                    in_formula2 = true;
                    formula2.clear();
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let text = e.unescape().unwrap_or_default().to_string();
                if in_formula1 {
                    formula1.push_str(&text);
                } else if in_formula2 {
                    formula2.push_str(&text);
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"formula1" => {
                    in_formula1 = false;
                    if !formula1.is_empty() {
                        validation.formula1 = Some(formula1.clone());
                    }
                }
                b"formula2" => {
                    in_formula2 = false;
                    if !formula2.is_empty() {
                        validation.formula2 = Some(formula2.clone());
                    }
                }
                b"dataValidation" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: sheet_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(validation)
}

fn parse_data_validation_empty(start: &BytesStart, sheet_path: &str) -> DataValidation {
    let mut validation = DataValidation {
        id: NodeId::new(),
        validation_type: None,
        operator: None,
        allow_blank: false,
        show_input_message: false,
        show_error_message: false,
        error_title: None,
        error: None,
        prompt_title: None,
        prompt: None,
        ranges: Vec::new(),
        formula1: None,
        formula2: None,
        span: Some(SourceSpan::new(sheet_path)),
    };

    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"type" => {
                validation.validation_type = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"operator" => {
                validation.operator = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"allowBlank" => {
                let v = String::from_utf8_lossy(&attr.value);
                validation.allow_blank = v == "1" || v.eq_ignore_ascii_case("true");
            }
            b"showInputMessage" => {
                let v = String::from_utf8_lossy(&attr.value);
                validation.show_input_message = v == "1" || v.eq_ignore_ascii_case("true");
            }
            b"showErrorMessage" => {
                let v = String::from_utf8_lossy(&attr.value);
                validation.show_error_message = v == "1" || v.eq_ignore_ascii_case("true");
            }
            b"errorTitle" => {
                validation.error_title = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"error" => {
                validation.error = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"promptTitle" => {
                validation.prompt_title = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"prompt" => {
                validation.prompt = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"sqref" => {
                let val = String::from_utf8_lossy(&attr.value).to_string();
                validation.ranges = val.split_whitespace().map(|s| s.to_string()).collect();
            }
            _ => {}
        }
    }

    validation
}
