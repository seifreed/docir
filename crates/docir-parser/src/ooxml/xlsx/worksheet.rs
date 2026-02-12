use super::*;
use crate::zip_handler::PackageReader;

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

        let mut columns: HashMap<u32, ColumnDefinition> = HashMap::new();
        let mut merged_cells: Vec<MergedCellRange> = Vec::new();
        let mut cells: Vec<NodeId> = Vec::new();
        let mut conditional_formats: Vec<NodeId> = Vec::new();
        let mut data_validations: Vec<NodeId> = Vec::new();

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"dimension" => {
                        if let Some(val) = attr_value(&e, b"ref") {
                            worksheet.dimension = Some(val);
                        }
                    }
                    b"tabColor" => {
                        worksheet.tab_color = parse_color_attr(&e);
                    }
                    b"pageMargins" => {
                        worksheet.page_margins = parse_page_margins(&e);
                    }
                    b"c" => {
                        let cell = self.parse_cell(&mut reader, &e, sheet_path)?;
                        let cell_id = cell.id;
                        self.store.insert(IRNode::Cell(cell));
                        cells.push(cell_id);
                    }
                    b"conditionalFormatting" => {
                        let fmt = parse_conditional_formatting(&mut reader, &e, sheet_path)?;
                        let id = fmt.id;
                        self.store.insert(IRNode::ConditionalFormat(fmt));
                        conditional_formats.push(id);
                    }
                    b"dataValidations" => {
                        let vals = parse_data_validations(&mut reader, sheet_path)?;
                        for val in vals {
                            let id = val.id;
                            self.store.insert(IRNode::DataValidation(val));
                            data_validations.push(id);
                        }
                    }
                    b"col" => {
                        parse_column(&e, &mut columns);
                    }
                    b"mergeCell" => {
                        if let Some(range) = parse_merge_cell(&e) {
                            merged_cells.push(range);
                        }
                    }
                    b"hyperlink" => {
                        self.handle_hyperlink(&e, relationships, sheet_path);
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match e.name().as_ref() {
                    b"dimension" => {
                        if let Some(val) = attr_value(&e, b"ref") {
                            worksheet.dimension = Some(val);
                        }
                    }
                    b"tabColor" => {
                        worksheet.tab_color = parse_color_attr(&e);
                    }
                    b"pageMargins" => {
                        worksheet.page_margins = parse_page_margins(&e);
                    }
                    b"c" => {
                        let cell = self.parse_empty_cell(&e, sheet_path)?;
                        let cell_id = cell.id;
                        self.store.insert(IRNode::Cell(cell));
                        cells.push(cell_id);
                    }
                    b"conditionalFormatting" => {
                        let fmt = parse_conditional_formatting_empty(&e, sheet_path);
                        let id = fmt.id;
                        self.store.insert(IRNode::ConditionalFormat(fmt));
                        conditional_formats.push(id);
                    }
                    b"dataValidations" => {
                        // Empty container, nothing to add
                    }
                    b"col" => {
                        parse_column(&e, &mut columns);
                    }
                    b"mergeCell" => {
                        if let Some(range) = parse_merge_cell(&e) {
                            merged_cells.push(range);
                        }
                    }
                    b"hyperlink" => {
                        self.handle_hyperlink(&e, relationships, sheet_path);
                    }
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

        // Deterministic ordering
        let mut columns_sorted: Vec<ColumnDefinition> = columns.into_values().collect();
        columns_sorted.sort_by_key(|c| c.index);
        merged_cells.sort_by_key(|r| (r.start_row, r.start_col, r.end_row, r.end_col));

        worksheet.columns = columns_sorted;
        worksheet.merged_cells = merged_cells;
        worksheet.cells = cells;
        worksheet.conditional_formats = conditional_formats;
        worksheet.data_validations = data_validations;

        // Drawings (images/charts)
        let mut drawings: Vec<NodeId> = Vec::new();
        for rel in relationships.get_by_type(rel_type::DRAWING) {
            let drawing_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&drawing_path) {
                continue;
            }
            let drawing_xml = zip.read_file_string(&drawing_path)?;
            let rels_path = get_rels_path(&drawing_path);
            let drawing_rels = if zip.contains(&rels_path) {
                let rels_xml = zip.read_file_string(&rels_path)?;
                Relationships::parse(&rels_xml)?
            } else {
                Relationships::default()
            };
            let drawing_id = self.parse_drawing(&drawing_xml, &drawing_path, &drawing_rels, zip)?;
            drawings.push(drawing_id);
        }

        worksheet.drawings = drawings;

        // Table definitions
        let mut tables: Vec<NodeId> = Vec::new();
        for rel in relationships.get_by_type(rel_type::TABLE) {
            let table_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&table_path) {
                continue;
            }
            let table_xml = zip.read_file_string(&table_path)?;
            let mut table = parse_table_definition(&table_xml, &table_path)?;
            table.span = Some(SourceSpan::new(&table_path).with_relationship(rel.id.clone()));
            let id = table.id;
            self.store.insert(IRNode::TableDefinition(table));
            tables.push(id);
        }
        worksheet.tables = tables;

        // Pivot tables
        let mut pivots: Vec<NodeId> = Vec::new();
        for rel in relationships.get_by_type(rel_type::PIVOT_TABLE) {
            let pivot_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&pivot_path) {
                continue;
            }
            let pivot_xml = zip.read_file_string(&pivot_path)?;
            let mut pivot = parse_pivot_table_definition(&pivot_xml, &pivot_path)?;
            pivot.span = Some(SourceSpan::new(&pivot_path).with_relationship(rel.id.clone()));
            let id = pivot.id;
            self.store.insert(IRNode::PivotTable(pivot));
            pivots.push(id);
        }
        worksheet.pivot_tables = pivots;

        // Comments (legacy + threaded)
        let mut comments: Vec<NodeId> = Vec::new();
        for rel in relationships.get_by_type(rel_type::COMMENTS) {
            let comments_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&comments_path) {
                continue;
            }
            let comments_xml = zip.read_file_string(&comments_path)?;
            let parsed = parse_sheet_comments(&comments_xml, &comments_path, Some(&sheet.name))?;
            for mut comment in parsed {
                comment.span =
                    Some(SourceSpan::new(&comments_path).with_relationship(rel.id.clone()));
                let id = comment.id;
                self.store.insert(IRNode::SheetComment(comment));
                comments.push(id);
            }
        }
        for rel in relationships.get_by_type(rel_type::THREADED_COMMENTS) {
            let comments_path = Relationships::resolve_target(sheet_path, &rel.target);
            if !zip.contains(&comments_path) {
                continue;
            }
            let comments_xml = zip.read_file_string(&comments_path)?;
            let parsed = parse_threaded_comments(&comments_xml, &comments_path, Some(&sheet.name))?;
            for mut comment in parsed {
                comment.span =
                    Some(SourceSpan::new(&comments_path).with_relationship(rel.id.clone()));
                let id = comment.id;
                self.store.insert(IRNode::SheetComment(comment));
                comments.push(id);
            }
        }
        worksheet.comments = comments;

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

    fn parse_chartsheet(
        &mut self,
        zip: &mut impl PackageReader,
        xml: &str,
        sheet_path: &str,
        relationships: &Relationships,
    ) -> Result<Option<NodeId>, ParseError> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
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
            let rels_path = get_rels_path(&external_path);
            let rels = if zip.contains(&rels_path) {
                let rels_xml = zip.read_file_string(&rels_path)?;
                Relationships::parse(&rels_xml)?
            } else {
                Relationships::default()
            };
            if let Ok(xml) = zip.read_file_string(&external_path) {
                if let Ok(mut part) = parse_external_link_part(&xml, &external_path, Some(&rels)) {
                    part.span = Some(SourceSpan::new(&external_path));
                    let part_id = part.id;
                    self.store.insert(IRNode::ExternalLinkPart(part));
                    if let Some(IRNode::Document(doc)) = self.store.get_mut(self.root_id) {
                        doc.shared_parts.push(part_id);
                    }
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
                let id = ext_ref.id;
                self.store.insert(IRNode::ExternalReference(ext_ref));
                self.security_info.external_refs.push(id);
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
                if let Some(IRNode::Document(doc)) = self.store.get_mut(self.root_id) {
                    doc.shared_parts.push(part_id);
                }
                for target in targets {
                    let ext_ref = ExternalReference::new(ExternalRefType::DataConnection, target);
                    let id = ext_ref.id;
                    self.store.insert(IRNode::ExternalReference(ext_ref));
                    self.security_info.external_refs.push(id);
                }
            }
        }

        Ok(())
    }

    pub(super) fn parse_pivot_cache(
        &mut self,
        zip: &mut impl PackageReader,
        xml: &str,
        cache_path: &str,
        cache_id: u32,
    ) -> Result<PivotCache, ParseError> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

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
        let rels_path = get_rels_path(cache_path);
        if zip.contains(&rels_path) {
            let rels_xml = zip.read_file_string(&rels_path)?;
            let rels = Relationships::parse(&rels_xml)?;
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
        }

        Ok(cache)
    }
}
