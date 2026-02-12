use super::*;

impl XlsxParser {
    pub(super) fn load_workbook_properties(
        &mut self,
        document: &mut Document,
        workbook_properties: Option<docir_core::ir::WorkbookProperties>,
        workbook_path: &str,
    ) {
        if let Some(mut props) = workbook_properties {
            props.span = Some(SourceSpan::new(workbook_path));
            let props_id = props.id;
            self.store.insert(IRNode::WorkbookProperties(props));
            document.workbook_properties = Some(props_id);
        }
    }

    pub(super) fn load_shared_strings(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
    ) -> Result<(), ParseError> {
        if zip.contains("xl/sharedStrings.xml") {
            let shared_xml = zip.read_file_string("xl/sharedStrings.xml")?;
            let (table, strings) = parse_shared_strings_table(&shared_xml)?;
            self.shared_strings = strings;
            let table_id = table.id;
            self.store.insert(IRNode::SharedStringTable(table));
            document.shared_strings = Some(table_id);
        }
        Ok(())
    }

    pub(super) fn load_styles(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
    ) -> Result<(), ParseError> {
        if let Some(styles) =
            parse_xml_part_with_span(zip, "xl/styles.xml", parse_styles, |styles, path| {
                styles.span = Some(SourceSpan::new(path))
            })?
        {
            let styles_id = styles.id;
            self.store.insert(IRNode::SpreadsheetStyles(styles));
            document.spreadsheet_styles = Some(styles_id);
        }
        Ok(())
    }

    pub(super) fn load_calc_chain(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
    ) -> Result<(), ParseError> {
        if let Some(chain) =
            parse_xml_part_with_span(zip, "xl/calcChain.xml", parse_calc_chain, |chain, path| {
                chain.span = Some(SourceSpan::new(path))
            })?
        {
            let chain_id = chain.id;
            self.store.insert(IRNode::CalcChain(chain));
            document.shared_parts.push(chain_id);
        }
        Ok(())
    }

    pub(super) fn load_people_part(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
    ) -> Result<(), ParseError> {
        if let Some(people) = parse_xml_part_with_span(
            zip,
            "xl/persons/person.xml",
            crate::ooxml::shared::parse_people_part,
            |people, path| people.span = Some(SourceSpan::new(path)),
        )? {
            let id = people.id;
            self.store.insert(IRNode::PeoplePart(people));
            document.shared_parts.push(id);
        }
        Ok(())
    }

    pub(super) fn load_defined_names(
        &mut self,
        document: &mut Document,
        defined_names: Vec<docir_core::ir::DefinedName>,
    ) -> Vec<Option<String>> {
        let mut auto_open_targets: Vec<Option<String>> = Vec::new();
        for defined in defined_names {
            if let Some(target) = auto_open_target_from_defined_name(&defined) {
                auto_open_targets.push(target);
            }
            let id = defined.id;
            self.store.insert(IRNode::DefinedName(defined));
            document.defined_names.push(id);
        }
        auto_open_targets
    }

    pub(super) fn load_sheets(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
        workbook_path: &str,
        workbook_rels: &Relationships,
        sheets: Vec<SheetInfo>,
    ) -> Result<(), ParseError> {
        for sheet in sheets {
            let rel = match workbook_rels.get(&sheet.rel_id) {
                Some(rel) => rel,
                None => {
                    push_warning(
                        &mut self.diagnostics,
                        "MISSING_RELATIONSHIP",
                        format!("Missing relationship for sheet relId {}", sheet.rel_id),
                        Some(workbook_path),
                    );
                    continue;
                }
            };
            let sheet_path = Relationships::resolve_target(workbook_path, &rel.target);

            let sheet_xml = zip.read_file_string(&sheet_path)?;

            let sheet_rels = read_relationships(zip, &sheet_path)?;

            self.process_external_relationships(&sheet_rels, &sheet_path);

            let kind = match rel.rel_type.as_str() {
                rel_type::CHARTSHEET => SheetKind::ChartSheet,
                rel_type::DIALOGSHEET => SheetKind::DialogSheet,
                rel_type::MACROSHEET => SheetKind::MacroSheet,
                _ => SheetKind::Worksheet,
            };
            let sheet_id =
                self.parse_worksheet(zip, &sheet_xml, &sheet, &sheet_path, &sheet_rels, kind)?;
            document.content.push(sheet_id);
        }
        Ok(())
    }

    pub(super) fn load_pivot_caches(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
        workbook_path: &str,
        workbook_rels: &Relationships,
        pivot_cache_refs: Vec<workbook::PivotCacheRef>,
    ) -> Result<(), ParseError> {
        for cache_ref in pivot_cache_refs {
            let Some(rel) = workbook_rels.get(&cache_ref.rel_id) else {
                continue;
            };
            let cache_path = Relationships::resolve_target(workbook_path, &rel.target);
            if !zip.contains(&cache_path) {
                continue;
            }
            let cache_xml = zip.read_file_string(&cache_path)?;
            let cache = self.parse_pivot_cache(zip, &cache_xml, &cache_path, cache_ref.cache_id)?;
            let cache_id = cache.id;
            self.store.insert(IRNode::PivotCache(cache));
            document.pivot_caches.push(cache_id);
        }
        Ok(())
    }

    pub(super) fn load_sheet_metadata(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
    ) -> Result<(), ParseError> {
        if let Some(metadata) = parse_xml_part_with_span(
            zip,
            "xl/metadata.xml",
            parse_sheet_metadata,
            |metadata, path| metadata.span = Some(SourceSpan::new(path)),
        )? {
            let meta_id = metadata.id;
            self.store.insert(IRNode::SheetMetadata(metadata));
            document.sheet_metadata = Some(meta_id);
        }
        Ok(())
    }

    pub(super) fn load_slicer_parts(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
    ) -> Result<(), ParseError> {
        let slicer_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("xl/slicers/") && p.ends_with(".xml"))
            .collect();
        for path in slicer_paths {
            if let Some(slicer) =
                parse_xml_part_with_span(zip, &path, parse_slicer_part, |slicer, path| {
                    slicer.span = Some(SourceSpan::new(path))
                })?
            {
                let id = slicer.id;
                self.store.insert(IRNode::SlicerPart(slicer));
                document.shared_parts.push(id);
            }
        }
        Ok(())
    }

    pub(super) fn load_timeline_parts(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
    ) -> Result<(), ParseError> {
        let timeline_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("xl/timelines/") && p.ends_with(".xml"))
            .collect();
        for path in timeline_paths {
            if let Some(timeline) =
                parse_xml_part_with_span(zip, &path, parse_timeline_part, |timeline, path| {
                    timeline.span = Some(SourceSpan::new(path))
                })?
            {
                let id = timeline.id;
                self.store.insert(IRNode::TimelinePart(timeline));
                document.shared_parts.push(id);
            }
        }
        Ok(())
    }

    pub(super) fn load_query_tables(
        &mut self,
        document: &mut Document,
        zip: &mut impl PackageReader,
    ) -> Result<(), ParseError> {
        let query_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("xl/queryTables/") && p.ends_with(".xml"))
            .collect();
        for path in query_paths {
            if let Some(query) =
                parse_xml_part_with_span(zip, &path, parse_query_table_part, |query, path| {
                    query.span = Some(SourceSpan::new(path))
                })?
            {
                let id = query.id;
                self.store.insert(IRNode::QueryTablePart(query));
                document.shared_parts.push(id);
            }
        }
        Ok(())
    }
}
