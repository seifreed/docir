//! XLSX workbook and worksheet parsing.

use crate::diagnostics::push_warning;
use crate::error::ParseError;
use crate::ooxml::relationships::{rel_type, Relationship, Relationships, TargetMode};
use crate::security_utils::parse_dde_formula;
use crate::zip_handler::SecureZipReader;
use docir_core::ir::{
    parse_cell_reference, BorderDef, BorderSide, CalcChain, CalcChainEntry, Cell, CellAlignment,
    CellError, CellFormat, CellFormula, CellProtection, CellValue, ColumnDefinition,
    ConditionalFormat, ConditionalRule, ConnectionEntry, ConnectionPart, DataValidation,
    Diagnostics, Document, DxfStyle, ExternalLinkPart, ExternalLinkSheet, FillDef, FontDef,
    FormulaType, IRNode, MergedCellRange, NumberFormat, PivotCache, PivotCacheRecords, PivotTable,
    QueryTablePart, Shape, ShapeType, SharedStringItem, SharedStringTable, SheetComment, SheetKind,
    SheetMetadata, SheetMetadataType, SheetPageMargins, SheetState, SlicerPart, SpreadsheetStyles,
    TableColumn, TableDefinition, TableStyleDef, TableStyleInfo, TimelinePart, Worksheet,
    WorksheetDrawing,
};
use docir_core::security::{ExternalRefType, ExternalReference, SecurityInfo};
use docir_core::types::{DocumentFormat, NodeId, SourceSpan};
use docir_core::visitor::IrStore;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

mod workbook;
mod worksheet;
mod xlm;

use workbook::{auto_open_target_from_defined_name, parse_workbook_info, SheetInfo};

/// XLSX parser for workbook.xml and worksheets.
pub struct XlsxParser {
    store: IrStore,
    security_info: SecurityInfo,
    shared_strings: Vec<String>,
    external_rel_ids: HashSet<String>,
    chart_nodes: Vec<NodeId>,
    root_id: NodeId,
    current_sheet_kind: Option<SheetKind>,
    current_sheet_name: Option<String>,
    current_sheet_state: Option<SheetState>,
    current_xlm_index: Option<usize>,
    diagnostics: Diagnostics,
}

impl XlsxParser {
    /// Creates a new XLSX parser.
    pub fn new() -> Self {
        Self {
            store: IrStore::new(),
            security_info: SecurityInfo::default(),
            shared_strings: Vec::new(),
            external_rel_ids: HashSet::new(),
            chart_nodes: Vec::new(),
            root_id: NodeId::new(),
            current_sheet_kind: None,
            current_sheet_name: None,
            current_sheet_state: None,
            current_xlm_index: None,
            diagnostics: Diagnostics::new(),
        }
    }

    /// Parses the workbook and all worksheets.
    pub fn parse_workbook<R: Read + Seek>(
        &mut self,
        zip: &mut SecureZipReader<R>,
        workbook_xml: &str,
        workbook_rels: &Relationships,
        workbook_path: &str,
    ) -> Result<NodeId, ParseError> {
        let mut document = Document::new(DocumentFormat::Spreadsheet);
        self.root_id = document.id;
        document.span = Some(SourceSpan::new(workbook_path));

        self.process_external_relationships(workbook_rels, workbook_path);

        let workbook_info = parse_workbook_info(workbook_xml)?;

        // Workbook properties
        if let Some(mut props) = workbook_info.workbook_properties {
            props.span = Some(SourceSpan::new(workbook_path));
            let props_id = props.id;
            self.store.insert(IRNode::WorkbookProperties(props));
            document.workbook_properties = Some(props_id);
        }

        // Shared strings
        if zip.contains("xl/sharedStrings.xml") {
            let shared_xml = zip.read_file_string("xl/sharedStrings.xml")?;
            let (table, strings) = parse_shared_strings_table(&shared_xml)?;
            self.shared_strings = strings;
            let table_id = table.id;
            self.store.insert(IRNode::SharedStringTable(table));
            document.shared_strings = Some(table_id);
        }

        // Styles
        if zip.contains("xl/styles.xml") {
            let styles_xml = zip.read_file_string("xl/styles.xml")?;
            let mut styles = parse_styles(&styles_xml, "xl/styles.xml")?;
            styles.span = Some(SourceSpan::new("xl/styles.xml"));
            let styles_id = styles.id;
            self.store.insert(IRNode::SpreadsheetStyles(styles));
            document.spreadsheet_styles = Some(styles_id);
        }

        // Calculation chain
        if zip.contains("xl/calcChain.xml") {
            let chain_xml = zip.read_file_string("xl/calcChain.xml")?;
            let mut chain = parse_calc_chain(&chain_xml, "xl/calcChain.xml")?;
            chain.span = Some(SourceSpan::new("xl/calcChain.xml"));
            let chain_id = chain.id;
            self.store.insert(IRNode::CalcChain(chain));
            document.shared_parts.push(chain_id);
        }

        // People part (coauthoring)
        if zip.contains("xl/persons/person.xml") {
            let xml = zip.read_file_string("xl/persons/person.xml")?;
            let mut people =
                crate::ooxml::shared::parse_people_part(&xml, "xl/persons/person.xml")?;
            people.span = Some(SourceSpan::new("xl/persons/person.xml"));
            let id = people.id;
            self.store.insert(IRNode::PeoplePart(people));
            document.shared_parts.push(id);
        }

        // Defined names
        let mut auto_open_targets: Vec<Option<String>> = Vec::new();
        for defined in workbook_info.defined_names {
            if let Some(target) = auto_open_target_from_defined_name(&defined) {
                auto_open_targets.push(target);
            }
            let id = defined.id;
            self.store.insert(IRNode::DefinedName(defined));
            document.defined_names.push(id);
        }

        // Sheets
        let sheets = workbook_info.sheets;
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

            let rels_path = get_rels_path(&sheet_path);
            let sheet_rels = if zip.contains(&rels_path) {
                let rels_xml = zip.read_file_string(&rels_path)?;
                Relationships::parse(&rels_xml)?
            } else {
                Relationships::default()
            };

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

        self.finalize_auto_open_targets(&auto_open_targets, workbook_path);

        // Pivot caches
        for cache_ref in workbook_info.pivot_cache_refs {
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

        self.parse_external_links_and_connections(zip, workbook_path, workbook_rels)?;

        if zip.contains("xl/metadata.xml") {
            let xml = zip.read_file_string("xl/metadata.xml")?;
            let mut metadata = parse_sheet_metadata(&xml, "xl/metadata.xml")?;
            metadata.span = Some(SourceSpan::new("xl/metadata.xml"));
            let meta_id = metadata.id;
            self.store.insert(IRNode::SheetMetadata(metadata));
            document.sheet_metadata = Some(meta_id);
        }

        // slicers
        let slicer_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.starts_with("xl/slicers/") && p.ends_with(".xml"))
            .map(|s| s.to_string())
            .collect();
        for path in slicer_paths {
            let xml = zip.read_file_string(&path)?;
            let mut slicer = parse_slicer_part(&xml, &path)?;
            slicer.span = Some(SourceSpan::new(&path));
            let id = slicer.id;
            self.store.insert(IRNode::SlicerPart(slicer));
            document.shared_parts.push(id);
        }

        // timelines
        let timeline_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.starts_with("xl/timelines/") && p.ends_with(".xml"))
            .map(|s| s.to_string())
            .collect();
        for path in timeline_paths {
            let xml = zip.read_file_string(&path)?;
            let mut timeline = parse_timeline_part(&xml, &path)?;
            timeline.span = Some(SourceSpan::new(&path));
            let id = timeline.id;
            self.store.insert(IRNode::TimelinePart(timeline));
            document.shared_parts.push(id);
        }

        // query tables
        let query_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.starts_with("xl/queryTables/") && p.ends_with(".xml"))
            .map(|s| s.to_string())
            .collect();
        for path in query_paths {
            let xml = zip.read_file_string(&path)?;
            let mut query = parse_query_table_part(&xml, &path)?;
            query.span = Some(SourceSpan::new(&path));
            let id = query.id;
            self.store.insert(IRNode::QueryTablePart(query));
            document.shared_parts.push(id);
        }

        document.shared_parts.extend(self.chart_nodes.drain(..));
        document.security = std::mem::take(&mut self.security_info);
        document.security.recalculate_threat_level();

        let mut diagnostics = std::mem::replace(&mut self.diagnostics, Diagnostics::new());
        if !diagnostics.entries.is_empty() {
            diagnostics.span = Some(SourceSpan::new(workbook_path));
            let diag_id = diagnostics.id;
            self.store.insert(IRNode::Diagnostics(diagnostics));
            document.diagnostics.push(diag_id);
        }

        let doc_id = document.id;
        self.store.insert(IRNode::Document(document));

        Ok(doc_id)
    }

    /// Returns the IR store.
    pub fn into_store(self) -> IrStore {
        self.store
    }

    fn parse_drawing<R: Read + Seek>(
        &mut self,
        xml: &str,
        drawing_path: &str,
        relationships: &Relationships,
        zip: &mut SecureZipReader<R>,
    ) -> Result<NodeId, ParseError> {
        let mut drawing = WorksheetDrawing::new();
        drawing.span = Some(SourceSpan::new(drawing_path));

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut current_shape: Option<Shape> = None;
        let mut current_embed: Option<String> = None;
        let mut current_chart: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"xdr:pic" => {
                        current_shape = Some(Shape::new(ShapeType::Picture));
                    }
                    b"xdr:graphicFrame" => {
                        current_shape = Some(Shape::new(ShapeType::Chart));
                    }
                    b"xdr:cNvPr" => {
                        if let Some(shape) = current_shape.as_mut() {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"name" => {
                                        shape.name =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                    b"descr" => {
                                        shape.alt_text =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    b"a:blip" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r:embed" {
                                current_embed =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ if e.name().as_ref().ends_with(b"chart") => {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            if key == b"r:id" || key.ends_with(b":id") {
                                current_chart =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"xdr:pic" => {
                        if let Some(mut shape) = current_shape.take() {
                            if let Some(rel_id) = current_embed.take() {
                                if let Some(rel) = relationships.get(&rel_id) {
                                    shape.relationship_id = Some(rel_id.clone());
                                    shape.media_target = Some(Relationships::resolve_target(
                                        drawing_path,
                                        &rel.target,
                                    ));
                                    if rel.target_mode == TargetMode::External {
                                        let ext_ref = ExternalReference::new(
                                            ExternalRefType::Image,
                                            &rel.target,
                                        );
                                        let ext_ref = ExternalReference {
                                            relationship_id: Some(rel_id),
                                            ..ext_ref
                                        };
                                        let ext_id = ext_ref.id;
                                        self.store.insert(IRNode::ExternalReference(ext_ref));
                                        self.security_info.external_refs.push(ext_id);
                                    }
                                }
                            }
                            let id = shape.id;
                            self.store.insert(IRNode::Shape(shape));
                            drawing.shapes.push(id);
                        }
                    }
                    b"xdr:graphicFrame" => {
                        if let Some(mut shape) = current_shape.take() {
                            if let Some(rel_id) = current_chart.take() {
                                if let Some(rel) = relationships.get(&rel_id) {
                                    shape.relationship_id = Some(rel_id.clone());
                                    shape.media_target = Some(Relationships::resolve_target(
                                        drawing_path,
                                        &rel.target,
                                    ));
                                    let chart_path =
                                        Relationships::resolve_target(drawing_path, &rel.target);
                                    if zip.contains(&chart_path) {
                                        let chart_xml = zip.read_file_string(&chart_path)?;
                                        if let Some(chart_id) =
                                            self.parse_chart(&chart_xml, &chart_path)
                                        {
                                            self.chart_nodes.push(chart_id);
                                        }
                                    }
                                }
                            }
                            let id = shape.id;
                            self.store.insert(IRNode::Shape(shape));
                            drawing.shapes.push(id);
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: drawing_path.to_string(),
                        message: e.to_string(),
                    });
                }
                _ => {}
            }
            buf.clear();
        }

        let id = drawing.id;
        self.store.insert(IRNode::WorksheetDrawing(drawing));
        Ok(id)
    }

    fn parse_chart(&mut self, xml: &str, chart_path: &str) -> Option<NodeId> {
        crate::ooxml::shared::parse_chart_data(xml, chart_path, &mut self.store)
    }

    fn local_name(name: &[u8]) -> &[u8] {
        match name.iter().rposition(|b| *b == b':') {
            Some(pos) => &name[pos + 1..],
            None => name,
        }
    }

    fn parse_cell(
        &mut self,
        reader: &mut Reader<&[u8]>,
        start: &BytesStart,
        sheet_path: &str,
    ) -> Result<Cell, ParseError> {
        let mut cell_ref: Option<String> = None;
        let mut cell_type: Option<String> = None;
        let mut style_id: Option<u32> = None;

        for attr in start.attributes().flatten() {
            match attr.key.as_ref() {
                b"r" => cell_ref = Some(String::from_utf8_lossy(&attr.value).to_string()),
                b"t" => cell_type = Some(String::from_utf8_lossy(&attr.value).to_string()),
                b"s" => style_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
                _ => {}
            }
        }

        let reference = cell_ref.ok_or_else(|| {
            ParseError::InvalidStructure("Cell missing reference attribute".to_string())
        })?;

        let (col, row) = parse_cell_reference(&reference).ok_or_else(|| {
            ParseError::InvalidStructure(format!("Invalid cell reference: {reference}"))
        })?;

        let mut value_text: Option<String> = None;
        let mut inline_text: Option<String> = None;
        let mut formula: Option<CellFormula> = None;

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"v" => {
                        let text = reader.read_text(e.name()).map_err(|e| ParseError::Xml {
                            file: sheet_path.to_string(),
                            message: e.to_string(),
                        })?;
                        value_text = Some(text.to_string());
                    }
                    b"f" => {
                        let f = parse_formula(reader, &e, sheet_path)?;
                        formula = Some(f);
                    }
                    b"is" => {
                        inline_text = Some(parse_inline_string(reader, sheet_path)?);
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"f" {
                        formula = Some(parse_formula_empty(&e));
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"c" {
                        break;
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

        let mut cell = Cell::new(reference.clone(), col, row);
        cell.style_id = style_id;
        if let Some(f) = &formula {
            self.handle_formula_security(&reference, f, sheet_path);
        }
        cell.formula = formula;
        cell.span = Some(SourceSpan::new(sheet_path));

        cell.value = if let Some(text) = inline_text {
            CellValue::InlineString(text)
        } else if let Some(value) = value_text {
            match cell_type.as_deref() {
                Some("s") => {
                    let idx = value.trim().parse::<u32>().unwrap_or(0);
                    if let Some(s) = self.shared_strings.get(idx as usize) {
                        CellValue::String(s.clone())
                    } else {
                        CellValue::SharedString(idx)
                    }
                }
                Some("b") => {
                    let v = value.trim();
                    CellValue::Boolean(v == "1" || v.eq_ignore_ascii_case("true"))
                }
                Some("str") => CellValue::String(value),
                Some("e") => CellValue::Error(map_cell_error(&value)),
                Some("d") => match value.trim().parse::<f64>() {
                    Ok(v) => CellValue::DateTime(v),
                    Err(_) => CellValue::String(value),
                },
                _ => match value.trim().parse::<f64>() {
                    Ok(v) => CellValue::Number(v),
                    Err(_) => CellValue::String(value),
                },
            }
        } else {
            CellValue::Empty
        };

        Ok(cell)
    }

    fn parse_empty_cell(&self, start: &BytesStart, sheet_path: &str) -> Result<Cell, ParseError> {
        let mut cell_ref: Option<String> = None;
        for attr in start.attributes().flatten() {
            if attr.key.as_ref() == b"r" {
                cell_ref = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
        }

        let reference = cell_ref.ok_or_else(|| {
            ParseError::InvalidStructure("Cell missing reference attribute".to_string())
        })?;
        let (col, row) = parse_cell_reference(&reference).ok_or_else(|| {
            ParseError::InvalidStructure(format!("Invalid cell reference: {reference}"))
        })?;

        let mut cell = Cell::new(reference, col, row);
        cell.span = Some(SourceSpan::new(sheet_path));
        Ok(cell)
    }

    fn handle_hyperlink(
        &mut self,
        element: &BytesStart,
        relationships: &Relationships,
        sheet_path: &str,
    ) {
        let mut rel_id: Option<String> = None;
        for attr in element.attributes().flatten() {
            if attr.key.as_ref() == b"r:id" {
                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
        }

        let Some(rel_id) = rel_id else {
            return;
        };
        let Some(rel) = relationships.get(&rel_id) else {
            return;
        };
        if rel.target_mode != TargetMode::External {
            return;
        }

        let ref_type = classify_relationship(&rel.rel_type);
        self.add_external_reference(rel, ref_type, sheet_path);
    }

    fn process_external_relationships(&mut self, rels: &Relationships, file_path: &str) {
        for rel in rels.external_relationships() {
            let ref_type = classify_relationship(&rel.rel_type);
            self.add_external_reference(rel, ref_type, file_path);
        }
    }

    fn add_external_reference(
        &mut self,
        rel: &Relationship,
        ref_type: ExternalRefType,
        file_path: &str,
    ) {
        let key = format!("{file_path}::{id}", id = rel.id);
        if !self.external_rel_ids.insert(key) {
            return;
        }

        let mut ext_ref = ExternalReference::new(ref_type, &rel.target);
        ext_ref.relationship_id = Some(rel.id.clone());
        ext_ref.relationship_type = Some(rel.rel_type.clone());
        ext_ref.span = Some(SourceSpan::new(file_path).with_relationship(rel.id.clone()));

        let ext_id = ext_ref.id;
        self.store.insert(IRNode::ExternalReference(ext_ref));
        self.security_info.external_refs.push(ext_id);
    }

    fn handle_formula_security(&mut self, cell_ref: &str, formula: &CellFormula, sheet_path: &str) {
        let text = formula.text.trim();
        let upper = text.to_ascii_uppercase();

        // DDE detection in Excel formulas
        if upper.starts_with("DDEAUTO") || upper.starts_with("DDE") {
            if let Some(dde) = parse_dde_formula(
                text,
                SourceSpan::new(sheet_path).with_xml_path(cell_ref),
                true,
            ) {
                self.security_info.dde_fields.push(dde);
            }
        }

        self.record_xlm_formula(cell_ref, text, &upper, sheet_path);
    }
}

impl Default for XlsxParser {
    fn default() -> Self {
        Self::new()
    }
}

fn parse_shared_strings_table(xml: &str) -> Result<(SharedStringTable, Vec<String>), ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut strings: Vec<String> = Vec::new();
    let mut table = SharedStringTable::new();
    table.span = Some(SourceSpan::new("xl/sharedStrings.xml"));

    let mut in_si = false;
    let mut in_t = false;
    let mut in_run = false;
    let mut current = String::new();
    let mut current_run = String::new();
    let mut runs: Vec<String> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"si" => {
                    in_si = true;
                    current.clear();
                    current_run.clear();
                    runs.clear();
                }
                b"r" if in_si => {
                    in_run = true;
                    current_run.clear();
                }
                b"t" if in_si => {
                    in_t = true;
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_si && in_t {
                    let text = e.unescape().map_err(|err| ParseError::Xml {
                        file: "xl/sharedStrings.xml".to_string(),
                        message: err.to_string(),
                    })?;
                    current.push_str(&text);
                    if in_run {
                        current_run.push_str(&text);
                    }
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"t" => in_t = false,
                b"r" => {
                    if in_run {
                        runs.push(current_run.clone());
                        in_run = false;
                        current_run.clear();
                    }
                }
                b"si" => {
                    in_si = false;
                    strings.push(current.clone());
                    table.items.push(SharedStringItem {
                        text: current.clone(),
                        runs: runs.clone(),
                    });
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "xl/sharedStrings.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok((table, strings))
}

fn parse_calc_chain(xml: &str, path: &str) -> Result<CalcChain, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut chain = CalcChain::new();
    chain.span = Some(SourceSpan::new(path));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"c" {
                    let mut cell_ref = None;
                    let mut sheet_id = None;
                    let mut index = None;
                    let mut level = None;
                    let mut new_value = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"r" => {
                                cell_ref = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"i" => {
                                index = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"l" => {
                                level = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"s" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                new_value = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"si" => {
                                sheet_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            _ => {}
                        }
                    }
                    if let Some(cell_ref) = cell_ref {
                        chain.entries.push(CalcChainEntry {
                            cell_ref,
                            sheet_id,
                            index,
                            level,
                            new_value,
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(chain)
}

fn parse_sheet_comments(
    xml: &str,
    path: &str,
    sheet_name: Option<&str>,
) -> Result<Vec<SheetComment>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut authors: Vec<String> = Vec::new();
    let mut in_author = false;
    let mut in_comment = false;
    let mut in_comment_text = false;
    let mut current_ref: Option<String> = None;
    let mut current_author_id: Option<usize> = None;
    let mut current_text = String::new();

    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"author" => in_author = true,
                b"comment" => {
                    in_comment = true;
                    current_ref = attr_value(&e, b"ref");
                    current_author_id =
                        attr_value(&e, b"authorId").and_then(|v| v.parse::<usize>().ok());
                    current_text.clear();
                }
                b"text" | b"t" => {
                    if in_comment {
                        in_comment_text = true;
                    }
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let text = e.unescape().unwrap_or_default().to_string();
                if in_author {
                    authors.push(text);
                } else if in_comment_text {
                    current_text.push_str(&text);
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"author" => in_author = false,
                b"text" | b"t" => in_comment_text = false,
                b"comment" => {
                    if let Some(cell_ref) = current_ref.take() {
                        let mut comment =
                            SheetComment::new(cell_ref, current_text.trim().to_string());
                        comment.sheet_name = sheet_name.map(|s| s.to_string());
                        if let Some(id) = current_author_id.take() {
                            comment.author = authors.get(id).cloned();
                        }
                        out.push(comment);
                    }
                    in_comment = false;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_threaded_comments(
    xml: &str,
    path: &str,
    sheet_name: Option<&str>,
) -> Result<Vec<SheetComment>, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_comment = false;
    let mut in_text = false;
    let mut current_ref: Option<String> = None;
    let mut current_author: Option<String> = None;
    let mut current_text = String::new();

    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"threadedComment" => {
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
            Ok(Event::Text(e)) => {
                if in_text {
                    let text = e.unescape().unwrap_or_default().to_string();
                    current_text.push_str(&text);
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"text" | b"t" => in_text = false,
                b"threadedComment" => {
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
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_styles(xml: &str, styles_path: &str) -> Result<SpreadsheetStyles, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut styles = SpreadsheetStyles::new();
    styles.span = Some(SourceSpan::new(styles_path));

    let mut buf = Vec::new();

    let mut in_num_fmts = false;
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_borders = false;
    let mut in_cell_xfs = false;
    let mut in_cell_style_xfs = false;
    let mut in_dxfs = false;
    let mut in_table_styles = false;

    let mut current_font: Option<FontDef> = None;
    let mut current_fill: Option<FillDef> = None;
    let mut current_border: Option<BorderDef> = None;
    let mut current_border_side: Option<(String, BorderSide)> = None;
    let mut current_xf: Option<CellFormat> = None;
    let mut current_xf_is_style = false;
    let mut current_dxf: Option<DxfStyle> = None;
    let mut current_dxf_font: Option<FontDef> = None;
    let mut current_dxf_fill: Option<FillDef> = None;
    let mut current_dxf_border: Option<BorderDef> = None;
    let mut current_dxf_border_side: Option<(String, BorderSide)> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"numFmts" => in_num_fmts = true,
                b"numFmt" if in_num_fmts => {
                    let mut id = None;
                    let mut code = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"formatCode" => {
                                code = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(code)) = (id, code) {
                        styles.number_formats.push(NumberFormat {
                            id,
                            format_code: code,
                        });
                    }
                }
                b"fonts" => in_fonts = true,
                b"font" if in_fonts => {
                    current_font = Some(FontDef {
                        name: None,
                        size: None,
                        bold: false,
                        italic: false,
                        underline: false,
                        color: None,
                    });
                }
                b"font" if in_dxfs => {
                    current_dxf_font = Some(FontDef {
                        name: None,
                        size: None,
                        bold: false,
                        italic: false,
                        underline: false,
                        color: None,
                    });
                }
                b"name" => {
                    if let Some(font) = current_font.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                font.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                font.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                b"sz" => {
                    if let Some(font) = current_font.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                font.size =
                                    String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                            }
                        }
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                font.size =
                                    String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                            }
                        }
                    }
                }
                b"b" => {
                    if let Some(font) = current_font.as_mut() {
                        font.bold = true;
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        font.bold = true;
                    }
                }
                b"i" => {
                    if let Some(font) = current_font.as_mut() {
                        font.italic = true;
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        font.italic = true;
                    }
                }
                b"u" => {
                    if let Some(font) = current_font.as_mut() {
                        font.underline = true;
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        font.underline = true;
                    }
                }
                b"color" => {
                    if let Some(font) = current_font.as_mut() {
                        font.color = parse_color_attr(&e);
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        font.color = parse_color_attr(&e);
                    } else if let Some((_, side)) = current_border_side.as_mut() {
                        side.color = parse_color_attr(&e);
                    } else if let Some((_, side)) = current_dxf_border_side.as_mut() {
                        side.color = parse_color_attr(&e);
                    } else if let Some(fill) = current_fill.as_mut() {
                        if fill.fg_color.is_none() {
                            fill.fg_color = parse_color_attr(&e);
                        }
                    } else if let Some(fill) = current_dxf_fill.as_mut() {
                        if fill.fg_color.is_none() {
                            fill.fg_color = parse_color_attr(&e);
                        }
                    }
                }
                b"fills" => in_fills = true,
                b"fill" if in_fills => {
                    current_fill = Some(FillDef {
                        pattern_type: None,
                        fg_color: None,
                        bg_color: None,
                    });
                }
                b"fill" if in_dxfs => {
                    current_dxf_fill = Some(FillDef {
                        pattern_type: None,
                        fg_color: None,
                        bg_color: None,
                    });
                }
                b"patternFill" => {
                    if let Some(fill) = current_fill.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"patternType" {
                                fill.pattern_type =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    } else if let Some(fill) = current_dxf_fill.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"patternType" {
                                fill.pattern_type =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                b"fgColor" => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.fg_color = parse_color_attr(&e);
                    } else if let Some(fill) = current_dxf_fill.as_mut() {
                        fill.fg_color = parse_color_attr(&e);
                    }
                }
                b"bgColor" => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.bg_color = parse_color_attr(&e);
                    } else if let Some(fill) = current_dxf_fill.as_mut() {
                        fill.bg_color = parse_color_attr(&e);
                    }
                }
                b"borders" => in_borders = true,
                b"border" if in_borders => {
                    current_border = Some(BorderDef {
                        left: None,
                        right: None,
                        top: None,
                        bottom: None,
                    });
                }
                b"border" if in_dxfs => {
                    current_dxf_border = Some(BorderDef {
                        left: None,
                        right: None,
                        top: None,
                        bottom: None,
                    });
                }
                b"left" | b"right" | b"top" | b"bottom" => {
                    if current_border.is_some() {
                        let mut side = BorderSide {
                            style: None,
                            color: None,
                        };
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"style" {
                                side.style = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        current_border_side =
                            Some((String::from_utf8_lossy(e.name().as_ref()).to_string(), side));
                    } else if current_dxf_border.is_some() {
                        let mut side = BorderSide {
                            style: None,
                            color: None,
                        };
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"style" {
                                side.style = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        current_dxf_border_side =
                            Some((String::from_utf8_lossy(e.name().as_ref()).to_string(), side));
                    }
                }
                b"cellXfs" => in_cell_xfs = true,
                b"cellStyleXfs" => in_cell_style_xfs = true,
                b"dxfs" => in_dxfs = true,
                b"dxf" if in_dxfs => {
                    current_dxf = Some(DxfStyle::new());
                }
                b"xf" if in_cell_xfs => {
                    let mut xf = CellFormat {
                        num_fmt_id: None,
                        font_id: None,
                        fill_id: None,
                        border_id: None,
                        xf_id: None,
                        apply_number_format: false,
                        apply_font: false,
                        apply_fill: false,
                        apply_border: false,
                        apply_alignment: false,
                        apply_protection: false,
                        quote_prefix: false,
                        pivot_button: false,
                        alignment: None,
                        protection: None,
                    };
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                xf.num_fmt_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"fontId" => {
                                xf.font_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"fillId" => {
                                xf.fill_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"borderId" => {
                                xf.border_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"xfId" => {
                                xf.xf_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"applyNumberFormat" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_number_format = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyFont" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_font = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyFill" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_fill = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyBorder" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_border = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyAlignment" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_alignment = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyProtection" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_protection = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"quotePrefix" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.quote_prefix = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"pivotButton" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.pivot_button = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            _ => {}
                        }
                    }
                    current_xf = Some(xf);
                    current_xf_is_style = false;
                }
                b"xf" if in_cell_style_xfs => {
                    let mut xf = CellFormat {
                        num_fmt_id: None,
                        font_id: None,
                        fill_id: None,
                        border_id: None,
                        xf_id: None,
                        apply_number_format: false,
                        apply_font: false,
                        apply_fill: false,
                        apply_border: false,
                        apply_alignment: false,
                        apply_protection: false,
                        quote_prefix: false,
                        pivot_button: false,
                        alignment: None,
                        protection: None,
                    };
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                xf.num_fmt_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"fontId" => {
                                xf.font_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"fillId" => {
                                xf.fill_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"borderId" => {
                                xf.border_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"xfId" => {
                                xf.xf_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"applyNumberFormat" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_number_format = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyFont" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_font = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyFill" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_fill = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyBorder" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_border = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyAlignment" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_alignment = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyProtection" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_protection = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"quotePrefix" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.quote_prefix = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"pivotButton" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.pivot_button = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            _ => {}
                        }
                    }
                    current_xf = Some(xf);
                    current_xf_is_style = true;
                }
                b"numFmt" if in_dxfs => {
                    if let Some(dxf) = current_dxf.as_mut() {
                        let mut id = None;
                        let mut code = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                                }
                                b"formatCode" => {
                                    code = Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                _ => {}
                            }
                        }
                        if let (Some(id), Some(code)) = (id, code) {
                            dxf.num_fmt = Some(NumberFormat {
                                id,
                                format_code: code,
                            });
                        }
                    }
                }
                b"alignment" => {
                    if let Some(xf) = current_xf.as_mut() {
                        let mut alignment = CellAlignment {
                            horizontal: None,
                            vertical: None,
                            wrap_text: false,
                            indent: None,
                            text_rotation: None,
                            shrink_to_fit: false,
                            reading_order: None,
                        };
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    alignment.horizontal =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"vertical" => {
                                    alignment.vertical =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"wrapText" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    alignment.wrap_text =
                                        v == "1" || v.eq_ignore_ascii_case("true");
                                }
                                b"indent" => {
                                    alignment.indent =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                                }
                                b"textRotation" => {
                                    alignment.text_rotation =
                                        String::from_utf8_lossy(&attr.value).parse::<i32>().ok();
                                }
                                b"shrinkToFit" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    alignment.shrink_to_fit =
                                        v == "1" || v.eq_ignore_ascii_case("true");
                                }
                                b"readingOrder" => {
                                    alignment.reading_order =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                                }
                                _ => {}
                            }
                        }
                        xf.alignment = Some(alignment);
                    } else if let Some(dxf) = current_dxf.as_mut() {
                        let mut alignment = CellAlignment {
                            horizontal: None,
                            vertical: None,
                            wrap_text: false,
                            indent: None,
                            text_rotation: None,
                            shrink_to_fit: false,
                            reading_order: None,
                        };
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    alignment.horizontal =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"vertical" => {
                                    alignment.vertical =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"wrapText" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    alignment.wrap_text =
                                        v == "1" || v.eq_ignore_ascii_case("true");
                                }
                                b"indent" => {
                                    alignment.indent =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                                }
                                b"textRotation" => {
                                    alignment.text_rotation =
                                        String::from_utf8_lossy(&attr.value).parse::<i32>().ok();
                                }
                                b"shrinkToFit" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    alignment.shrink_to_fit =
                                        v == "1" || v.eq_ignore_ascii_case("true");
                                }
                                b"readingOrder" => {
                                    alignment.reading_order =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                                }
                                _ => {}
                            }
                        }
                        dxf.alignment = Some(alignment);
                    }
                }
                b"protection" => {
                    let mut protection = CellProtection {
                        locked: None,
                        hidden: None,
                    };
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"locked" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                protection.locked =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"hidden" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                protection.hidden =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            _ => {}
                        }
                    }
                    if let Some(xf) = current_xf.as_mut() {
                        xf.protection = Some(protection);
                    } else if let Some(dxf) = current_dxf.as_mut() {
                        dxf.protection = Some(protection);
                    }
                }
                b"tableStyles" => {
                    let mut info = TableStyleInfo {
                        count: None,
                        default_table_style: None,
                        default_pivot_style: None,
                        styles: Vec::new(),
                    };
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"count" => {
                                info.count =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"defaultTableStyle" => {
                                info.default_table_style =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"defaultPivotStyle" => {
                                info.default_pivot_style =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            _ => {}
                        }
                    }
                    styles.table_styles = Some(info);
                    in_table_styles = true;
                }
                b"tableStyle" if in_table_styles => {
                    if let Some(info) = styles.table_styles.as_mut() {
                        let mut name = None;
                        let mut pivot = None;
                        let mut table = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => {
                                    name = Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"pivot" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    pivot = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                                }
                                b"table" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    table = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                                }
                                _ => {}
                            }
                        }
                        if let Some(name) = name {
                            info.styles.push(TableStyleDef { name, pivot, table });
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"numFmt" if in_num_fmts => {
                    let mut id = None;
                    let mut code = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"formatCode" => {
                                code = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(code)) = (id, code) {
                        styles.number_formats.push(NumberFormat {
                            id,
                            format_code: code,
                        });
                    }
                }
                b"b" => {
                    if let Some(font) = current_font.as_mut() {
                        font.bold = true;
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        font.bold = true;
                    }
                }
                b"i" => {
                    if let Some(font) = current_font.as_mut() {
                        font.italic = true;
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        font.italic = true;
                    }
                }
                b"u" => {
                    if let Some(font) = current_font.as_mut() {
                        font.underline = true;
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        font.underline = true;
                    }
                }
                b"name" => {
                    if let Some(font) = current_font.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                font.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                font.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                b"sz" => {
                    if let Some(font) = current_font.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                font.size =
                                    String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                            }
                        }
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                font.size =
                                    String::from_utf8_lossy(&attr.value).parse::<f64>().ok();
                            }
                        }
                    }
                }
                b"color" => {
                    if let Some(font) = current_font.as_mut() {
                        font.color = parse_color_attr(&e);
                    } else if let Some(font) = current_dxf_font.as_mut() {
                        font.color = parse_color_attr(&e);
                    } else if let Some((_, side)) = current_border_side.as_mut() {
                        side.color = parse_color_attr(&e);
                    } else if let Some((_, side)) = current_dxf_border_side.as_mut() {
                        side.color = parse_color_attr(&e);
                    } else if let Some(fill) = current_fill.as_mut() {
                        if fill.fg_color.is_none() {
                            fill.fg_color = parse_color_attr(&e);
                        }
                    } else if let Some(fill) = current_dxf_fill.as_mut() {
                        if fill.fg_color.is_none() {
                            fill.fg_color = parse_color_attr(&e);
                        }
                    }
                }
                b"patternFill" => {
                    if let Some(fill) = current_fill.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"patternType" {
                                fill.pattern_type =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    } else if let Some(fill) = current_dxf_fill.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"patternType" {
                                fill.pattern_type =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                b"fgColor" => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.fg_color = parse_color_attr(&e);
                    } else if let Some(fill) = current_dxf_fill.as_mut() {
                        fill.fg_color = parse_color_attr(&e);
                    }
                }
                b"bgColor" => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.bg_color = parse_color_attr(&e);
                    } else if let Some(fill) = current_dxf_fill.as_mut() {
                        fill.bg_color = parse_color_attr(&e);
                    }
                }
                b"left" | b"right" | b"top" | b"bottom" => {
                    if let Some(border) = current_border.as_mut() {
                        let mut side = BorderSide {
                            style: None,
                            color: None,
                        };
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"style" {
                                side.style = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        assign_border_side(border, e.name().as_ref(), side);
                    } else if let Some(border) = current_dxf_border.as_mut() {
                        let mut side = BorderSide {
                            style: None,
                            color: None,
                        };
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"style" {
                                side.style = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        assign_border_side(border, e.name().as_ref(), side);
                    }
                }
                b"numFmt" if in_dxfs => {
                    if let Some(dxf) = current_dxf.as_mut() {
                        let mut id = None;
                        let mut code = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                                }
                                b"formatCode" => {
                                    code = Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                _ => {}
                            }
                        }
                        if let (Some(id), Some(code)) = (id, code) {
                            dxf.num_fmt = Some(NumberFormat {
                                id,
                                format_code: code,
                            });
                        }
                    }
                }
                b"alignment" => {
                    if let Some(xf) = current_xf.as_mut() {
                        let mut alignment = CellAlignment {
                            horizontal: None,
                            vertical: None,
                            wrap_text: false,
                            indent: None,
                            text_rotation: None,
                            shrink_to_fit: false,
                            reading_order: None,
                        };
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    alignment.horizontal =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"vertical" => {
                                    alignment.vertical =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"wrapText" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    alignment.wrap_text =
                                        v == "1" || v.eq_ignore_ascii_case("true");
                                }
                                b"indent" => {
                                    alignment.indent =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                                }
                                b"textRotation" => {
                                    alignment.text_rotation =
                                        String::from_utf8_lossy(&attr.value).parse::<i32>().ok();
                                }
                                b"shrinkToFit" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    alignment.shrink_to_fit =
                                        v == "1" || v.eq_ignore_ascii_case("true");
                                }
                                b"readingOrder" => {
                                    alignment.reading_order =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                                }
                                _ => {}
                            }
                        }
                        xf.alignment = Some(alignment);
                    } else if let Some(dxf) = current_dxf.as_mut() {
                        let mut alignment = CellAlignment {
                            horizontal: None,
                            vertical: None,
                            wrap_text: false,
                            indent: None,
                            text_rotation: None,
                            shrink_to_fit: false,
                            reading_order: None,
                        };
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    alignment.horizontal =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"vertical" => {
                                    alignment.vertical =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"wrapText" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    alignment.wrap_text =
                                        v == "1" || v.eq_ignore_ascii_case("true");
                                }
                                b"indent" => {
                                    alignment.indent =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                                }
                                b"textRotation" => {
                                    alignment.text_rotation =
                                        String::from_utf8_lossy(&attr.value).parse::<i32>().ok();
                                }
                                b"shrinkToFit" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    alignment.shrink_to_fit =
                                        v == "1" || v.eq_ignore_ascii_case("true");
                                }
                                b"readingOrder" => {
                                    alignment.reading_order =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                                }
                                _ => {}
                            }
                        }
                        dxf.alignment = Some(alignment);
                    }
                }
                b"protection" => {
                    let mut protection = CellProtection {
                        locked: None,
                        hidden: None,
                    };
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"locked" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                protection.locked =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"hidden" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                protection.hidden =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            _ => {}
                        }
                    }
                    if let Some(xf) = current_xf.as_mut() {
                        xf.protection = Some(protection);
                    } else if let Some(dxf) = current_dxf.as_mut() {
                        dxf.protection = Some(protection);
                    }
                }
                b"xf" if in_cell_style_xfs => {
                    let mut xf = CellFormat {
                        num_fmt_id: None,
                        font_id: None,
                        fill_id: None,
                        border_id: None,
                        xf_id: None,
                        apply_number_format: false,
                        apply_font: false,
                        apply_fill: false,
                        apply_border: false,
                        apply_alignment: false,
                        apply_protection: false,
                        quote_prefix: false,
                        pivot_button: false,
                        alignment: None,
                        protection: None,
                    };
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                xf.num_fmt_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"fontId" => {
                                xf.font_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"fillId" => {
                                xf.fill_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"borderId" => {
                                xf.border_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"xfId" => {
                                xf.xf_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"applyNumberFormat" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_number_format = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyFont" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_font = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyFill" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_fill = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyBorder" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_border = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyAlignment" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_alignment = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"applyProtection" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.apply_protection = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"quotePrefix" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.quote_prefix = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            b"pivotButton" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                xf.pivot_button = v == "1" || v.eq_ignore_ascii_case("true");
                            }
                            _ => {}
                        }
                    }
                    styles.cell_style_xfs.push(xf);
                }
                b"tableStyles" => {
                    let mut info = TableStyleInfo {
                        count: None,
                        default_table_style: None,
                        default_pivot_style: None,
                        styles: Vec::new(),
                    };
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"count" => {
                                info.count =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"defaultTableStyle" => {
                                info.default_table_style =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"defaultPivotStyle" => {
                                info.default_pivot_style =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            _ => {}
                        }
                    }
                    styles.table_styles = Some(info);
                    in_table_styles = true;
                }
                b"tableStyle" if in_table_styles => {
                    if let Some(info) = styles.table_styles.as_mut() {
                        let mut name = None;
                        let mut pivot = None;
                        let mut table = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => {
                                    name = Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"pivot" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    pivot = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                                }
                                b"table" => {
                                    let v = String::from_utf8_lossy(&attr.value);
                                    table = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                                }
                                _ => {}
                            }
                        }
                        if let Some(name) = name {
                            info.styles.push(TableStyleDef { name, pivot, table });
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"numFmts" => in_num_fmts = false,
                b"fonts" => in_fonts = false,
                b"fills" => in_fills = false,
                b"borders" => in_borders = false,
                b"cellXfs" => in_cell_xfs = false,
                b"cellStyleXfs" => in_cell_style_xfs = false,
                b"dxfs" => in_dxfs = false,
                b"tableStyles" => in_table_styles = false,
                b"font" => {
                    if let Some(font) = current_font.take() {
                        styles.fonts.push(font);
                    } else if let Some(font) = current_dxf_font.take() {
                        if let Some(dxf) = current_dxf.as_mut() {
                            dxf.font = Some(font);
                        }
                    }
                }
                b"fill" => {
                    if let Some(fill) = current_fill.take() {
                        styles.fills.push(fill);
                    } else if let Some(fill) = current_dxf_fill.take() {
                        if let Some(dxf) = current_dxf.as_mut() {
                            dxf.fill = Some(fill);
                        }
                    }
                }
                b"border" => {
                    if let Some(border) = current_border.take() {
                        styles.borders.push(border);
                    } else if let Some(border) = current_dxf_border.take() {
                        if let Some(dxf) = current_dxf.as_mut() {
                            dxf.border = Some(border);
                        }
                    }
                }
                b"left" | b"right" | b"top" | b"bottom" => {
                    if let (Some(border), Some((name, side))) =
                        (current_border.as_mut(), current_border_side.take())
                    {
                        assign_border_side(border, name.as_bytes(), side);
                    } else if let (Some(border), Some((name, side))) =
                        (current_dxf_border.as_mut(), current_dxf_border_side.take())
                    {
                        assign_border_side(border, name.as_bytes(), side);
                    }
                }
                b"xf" => {
                    if let Some(xf) = current_xf.take() {
                        if current_xf_is_style {
                            styles.cell_style_xfs.push(xf);
                        } else {
                            styles.cell_xfs.push(xf);
                        }
                        current_xf_is_style = false;
                    }
                }
                b"dxf" => {
                    if let Some(dxf) = current_dxf.take() {
                        styles.dxfs.push(dxf);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: styles_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(styles)
}

fn attr_value(start: &BytesStart, key: &[u8]) -> Option<String> {
    start
        .attributes()
        .flatten()
        .find(|a| a.key.as_ref() == key)
        .map(|a| String::from_utf8_lossy(&a.value).to_string())
}

fn assign_border_side(border: &mut BorderDef, name: &[u8], side: BorderSide) {
    match name {
        b"left" => border.left = Some(side),
        b"right" => border.right = Some(side),
        b"top" => border.top = Some(side),
        b"bottom" => border.bottom = Some(side),
        _ => {}
    }
}

fn parse_color_attr(element: &BytesStart) -> Option<String> {
    let mut rgb = None;
    let mut theme = None;
    let mut indexed = None;
    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"rgb" => rgb = Some(String::from_utf8_lossy(&attr.value).to_string()),
            b"theme" => theme = Some(String::from_utf8_lossy(&attr.value).to_string()),
            b"indexed" => indexed = Some(String::from_utf8_lossy(&attr.value).to_string()),
            _ => {}
        }
    }
    if let Some(rgb) = rgb {
        Some(format!("rgb:{rgb}"))
    } else if let Some(theme) = theme {
        Some(format!("theme:{theme}"))
    } else if let Some(indexed) = indexed {
        Some(format!("indexed:{indexed}"))
    } else {
        None
    }
}

fn parse_conditional_formatting(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<ConditionalFormat, ParseError> {
    let mut ranges: Vec<String> = Vec::new();
    for attr in start.attributes().flatten() {
        if attr.key.as_ref() == b"sqref" {
            let val = String::from_utf8_lossy(&attr.value).to_string();
            ranges = val.split_whitespace().map(|s| s.to_string()).collect();
        }
    }

    let mut rules: Vec<ConditionalRule> = Vec::new();
    let mut current_rule: Option<ConditionalRule> = None;
    let mut in_formula = false;
    let mut formula_text = String::new();

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"cfRule" => {
                    let mut rule_type = "unknown".to_string();
                    let mut priority = None;
                    let mut operator = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"type" => rule_type = String::from_utf8_lossy(&attr.value).to_string(),
                            b"priority" => {
                                priority = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"operator" => {
                                operator = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            _ => {}
                        }
                    }
                    current_rule = Some(ConditionalRule {
                        rule_type,
                        priority,
                        operator,
                        formulae: Vec::new(),
                    });
                }
                b"formula" => {
                    in_formula = true;
                    formula_text.clear();
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_formula {
                    formula_text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"formula" => {
                    in_formula = false;
                    if let Some(rule) = current_rule.as_mut() {
                        if !formula_text.is_empty() {
                            rule.formulae.push(formula_text.clone());
                        }
                    }
                }
                b"cfRule" => {
                    if let Some(rule) = current_rule.take() {
                        rules.push(rule);
                    }
                }
                b"conditionalFormatting" => break,
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

    Ok(ConditionalFormat {
        id: NodeId::new(),
        ranges,
        rules,
        span: Some(SourceSpan::new(sheet_path)),
    })
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

fn parse_table_definition(xml: &str, table_path: &str) -> Result<TableDefinition, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut table = TableDefinition {
        id: NodeId::new(),
        name: None,
        display_name: None,
        ref_range: None,
        header_row_count: None,
        totals_row_count: None,
        columns: Vec::new(),
        span: Some(SourceSpan::new(table_path)),
    };

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                table.name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"displayName" => {
                                table.display_name =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"ref" => {
                                table.ref_range =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"headerRowCount" => {
                                table.header_row_count =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"totalsRowCount" => {
                                table.totals_row_count =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            _ => {}
                        }
                    }
                }
                b"tableColumn" => {
                    let mut id = None;
                    let mut name = None;
                    let mut totals_row_label = None;
                    let mut totals_row_function = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"id" => id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
                            b"name" => {
                                name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"totalsRowLabel" => {
                                totals_row_label =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"totalsRowFunction" => {
                                totals_row_function =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            _ => {}
                        }
                    }
                    if let Some(id) = id {
                        table.columns.push(TableColumn {
                            id,
                            name,
                            totals_row_label,
                            totals_row_function,
                        });
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) if e.name().as_ref() == b"tableColumn" => {
                let mut id = None;
                let mut name = None;
                let mut totals_row_label = None;
                let mut totals_row_function = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"id" => id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
                        b"name" => name = Some(String::from_utf8_lossy(&attr.value).to_string()),
                        b"totalsRowLabel" => {
                            totals_row_label =
                                Some(String::from_utf8_lossy(&attr.value).to_string())
                        }
                        b"totalsRowFunction" => {
                            totals_row_function =
                                Some(String::from_utf8_lossy(&attr.value).to_string())
                        }
                        _ => {}
                    }
                }
                if let Some(id) = id {
                    table.columns.push(TableColumn {
                        id,
                        name,
                        totals_row_label,
                        totals_row_function,
                    });
                }
            }
            Ok(Event::End(e)) if e.name().as_ref() == b"table" => break,
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: table_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(table)
}

fn parse_pivot_table_definition(xml: &str, pivot_path: &str) -> Result<PivotTable, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut pivot = PivotTable {
        id: NodeId::new(),
        name: None,
        cache_id: None,
        ref_range: None,
        span: Some(SourceSpan::new(pivot_path)),
    };

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"pivotTableDefinition" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                pivot.name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"cacheId" => {
                                pivot.cache_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            _ => {}
                        }
                    }
                }
                b"location" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"ref" {
                            pivot.ref_range =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) if e.name().as_ref() == b"location" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"ref" {
                        pivot.ref_range = Some(String::from_utf8_lossy(&attr.value).to_string());
                    }
                }
            }
            Ok(Event::End(e)) if e.name().as_ref() == b"pivotTableDefinition" => break,
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: pivot_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(pivot)
}

fn parse_pivot_cache_records(
    xml: &str,
    records_path: &str,
) -> Result<PivotCacheRecords, ParseError> {
    let mut records = PivotCacheRecords::new();
    records.span = Some(SourceSpan::new(records_path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_record = false;
    let mut current_fields: u32 = 0;
    let mut max_fields: u32 = 0;
    let mut counted_records: u32 = 0;
    let mut has_count_attr = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"pivotCacheRecords" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"count" {
                            records.record_count =
                                String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            has_count_attr = records.record_count.is_some();
                        }
                    }
                } else if e.name().as_ref() == b"r" {
                    in_record = true;
                    current_fields = 0;
                } else if in_record {
                    current_fields = current_fields.saturating_add(1);
                }
            }
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"pivotCacheRecords" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"count" {
                            records.record_count =
                                String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            has_count_attr = records.record_count.is_some();
                        }
                    }
                } else if e.name().as_ref() == b"r" {
                    counted_records = counted_records.saturating_add(1);
                    if max_fields < current_fields {
                        max_fields = current_fields;
                    }
                    in_record = false;
                    current_fields = 0;
                } else if in_record {
                    current_fields = current_fields.saturating_add(1);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"r" {
                    counted_records = counted_records.saturating_add(1);
                    if max_fields < current_fields {
                        max_fields = current_fields;
                    }
                    in_record = false;
                    current_fields = 0;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: records_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    if !has_count_attr {
        records.record_count = Some(counted_records);
    }
    if max_fields > 0 {
        records.field_count = Some(max_fields);
    }

    Ok(records)
}

fn parse_formula(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
    sheet_path: &str,
) -> Result<CellFormula, ParseError> {
    let mut formula_type = FormulaType::Normal;
    let mut shared_index = None;
    let mut shared_ref = None;
    let mut array_ref = None;
    let mut is_array = false;

    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"t" => {
                let v = String::from_utf8_lossy(&attr.value);
                match v.as_ref() {
                    "shared" => formula_type = FormulaType::Shared,
                    "array" => {
                        formula_type = FormulaType::Array;
                        is_array = true;
                    }
                    "dataTable" => formula_type = FormulaType::DataTable,
                    _ => {}
                }
            }
            b"si" => shared_index = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"ref" => {
                let r = String::from_utf8_lossy(&attr.value).to_string();
                if formula_type == FormulaType::Shared {
                    shared_ref = Some(r);
                } else {
                    array_ref = Some(r);
                }
            }
            _ => {}
        }
    }

    let text = reader
        .read_text(start.name())
        .map_err(|e| ParseError::Xml {
            file: sheet_path.to_string(),
            message: e.to_string(),
        })?;

    Ok(CellFormula {
        text: text.to_string(),
        formula_type,
        shared_index,
        shared_ref,
        is_array,
        array_ref,
    })
}

fn extract_formula_function(formula_upper: &str) -> Option<String> {
    let trimmed = formula_upper.trim();
    let trimmed = trimmed.strip_prefix('=').unwrap_or(trimmed);
    let idx = trimmed.find('(')?;
    Some(trimmed[..idx].trim().to_string())
}

fn parse_formula_args_text(formula: &str) -> Option<String> {
    let start = formula.find('(')?;
    let end = formula.rfind(')')?;
    if end > start + 1 {
        Some(formula[start + 1..end].to_string())
    } else {
        None
    }
}

fn parse_formula_empty(start: &BytesStart) -> CellFormula {
    let mut formula_type = FormulaType::Normal;
    let mut shared_index = None;
    let mut shared_ref = None;
    let mut array_ref = None;
    let mut is_array = false;

    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"t" => {
                let v = String::from_utf8_lossy(&attr.value);
                match v.as_ref() {
                    "shared" => formula_type = FormulaType::Shared,
                    "array" => {
                        formula_type = FormulaType::Array;
                        is_array = true;
                    }
                    "dataTable" => formula_type = FormulaType::DataTable,
                    _ => {}
                }
            }
            b"si" => shared_index = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"ref" => {
                let r = String::from_utf8_lossy(&attr.value).to_string();
                if formula_type == FormulaType::Shared {
                    shared_ref = Some(r);
                } else {
                    array_ref = Some(r);
                }
            }
            _ => {}
        }
    }

    CellFormula {
        text: String::new(),
        formula_type,
        shared_index,
        shared_ref,
        is_array,
        array_ref,
    }
}

fn parse_inline_string(reader: &mut Reader<&[u8]>, sheet_path: &str) -> Result<String, ParseError> {
    let mut buf = Vec::new();
    let mut in_t = false;
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"t" {
                    in_t = true;
                }
            }
            Ok(Event::Text(e)) => {
                if in_t {
                    let t = e.unescape().map_err(|err| ParseError::Xml {
                        file: sheet_path.to_string(),
                        message: err.to_string(),
                    })?;
                    text.push_str(&t);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"t" {
                    in_t = false;
                } else if e.name().as_ref() == b"is" {
                    break;
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

    Ok(text)
}

fn parse_column(element: &BytesStart, columns: &mut HashMap<u32, ColumnDefinition>) {
    let mut min = None;
    let mut max = None;
    let mut width = None;
    let mut hidden = false;
    let mut custom_width = false;

    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"min" => min = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"max" => max = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"width" => width = String::from_utf8_lossy(&attr.value).parse::<f64>().ok(),
            b"hidden" => hidden = attr.value.as_ref() == b"1",
            b"customWidth" => custom_width = attr.value.as_ref() == b"1",
            _ => {}
        }
    }

    let (Some(min), Some(max)) = (min, max) else {
        return;
    };
    for idx in min..=max {
        let col_index = idx.saturating_sub(1);
        columns.insert(
            col_index,
            ColumnDefinition {
                index: col_index,
                width,
                hidden,
                custom_width,
            },
        );
    }
}

fn parse_merge_cell(element: &BytesStart) -> Option<MergedCellRange> {
    let mut ref_attr = None;
    for attr in element.attributes().flatten() {
        if attr.key.as_ref() == b"ref" {
            ref_attr = Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }

    let ref_attr = ref_attr?;
    let mut parts = ref_attr.split(':');
    let start = parts.next()?;
    let end = parts.next().unwrap_or(start);

    let (start_col, start_row) = parse_cell_reference(start)?;
    let (end_col, end_row) = parse_cell_reference(end)?;

    Some(MergedCellRange {
        start_col,
        start_row,
        end_col,
        end_row,
    })
}

fn parse_connections_part(xml: &str, path: &str) -> Result<ConnectionPart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut part = ConnectionPart::new();
    part.span = Some(SourceSpan::new(path));
    let mut current: Option<ConnectionEntry> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"connection" => {
                    let mut entry = ConnectionEntry::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"id" => {
                                entry.connection_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"name" => {
                                entry.name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"description" => {
                                entry.description =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"type" => {
                                entry.connection_type =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"refreshedVersion" => {
                                entry.refreshed_version =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"refreshOnLoad" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.refresh_on_load =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"saveData" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.save_data = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"background" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.background = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"sourceFile" => {
                                entry.source_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            b"odcFile" => {
                                entry.connection_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            _ => {}
                        }
                    }
                    current = Some(entry);
                }
                b"dbPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"connection" => {
                                    entry.connection =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"command" => {
                                    entry.command =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"commandType" => {
                                    entry.command_type =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                                }
                                _ => {}
                            }
                        }
                    }
                }
                b"webPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"url" {
                                entry.url = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                b"textPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"sourceFile" || attr.key.as_ref() == b"file" {
                                entry.source_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"connection" => {
                    let mut entry = ConnectionEntry::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"id" => {
                                entry.connection_id =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"name" => {
                                entry.name = Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"description" => {
                                entry.description =
                                    Some(String::from_utf8_lossy(&attr.value).to_string())
                            }
                            b"type" => {
                                entry.connection_type =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"refreshedVersion" => {
                                entry.refreshed_version =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                            }
                            b"refreshOnLoad" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.refresh_on_load =
                                    Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"saveData" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.save_data = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"background" => {
                                let v = String::from_utf8_lossy(&attr.value);
                                entry.background = Some(v == "1" || v.eq_ignore_ascii_case("true"));
                            }
                            b"sourceFile" => {
                                entry.source_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            b"odcFile" => {
                                entry.connection_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            _ => {}
                        }
                    }
                    part.entries.push(entry);
                }
                b"dbPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"connection" => {
                                    entry.connection =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"command" => {
                                    entry.command =
                                        Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                b"commandType" => {
                                    entry.command_type =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                                }
                                _ => {}
                            }
                        }
                    }
                }
                b"webPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"url" {
                                entry.url = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                b"textPr" => {
                    if let Some(entry) = current.as_mut() {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"sourceFile" || attr.key.as_ref() == b"file" {
                                entry.source_file =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"connection" {
                    if let Some(entry) = current.take() {
                        part.entries.push(entry);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(part)
}

fn connection_targets(part: &ConnectionPart) -> Vec<String> {
    let mut targets = Vec::new();
    for entry in &part.entries {
        if let Some(value) = entry.connection.as_ref() {
            targets.push(value.clone());
        }
        if let Some(value) = entry.url.as_ref() {
            targets.push(value.clone());
        }
        if let Some(value) = entry.source_file.as_ref() {
            targets.push(value.clone());
        }
        if let Some(value) = entry.connection_file.as_ref() {
            targets.push(value.clone());
        }
    }
    targets.sort();
    targets.dedup();
    targets
}

fn parse_external_link_part(
    xml: &str,
    path: &str,
    rels: Option<&Relationships>,
) -> Result<ExternalLinkPart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut part = ExternalLinkPart::new();
    part.span = Some(SourceSpan::new(path));

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = XlsxParser::local_name(&name_buf);
                match local {
                    b"externalLink" => {
                        // placeholder for type if present
                        for attr in e.attributes().flatten() {
                            let key = XlsxParser::local_name(attr.key.as_ref());
                            if key == b"linkType" || key == b"type" {
                                part.link_type =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    b"sheetNames" => {}
                    b"sheetName" => {
                        let mut sheet = ExternalLinkSheet {
                            name: None,
                            r_id: None,
                        };
                        for attr in e.attributes().flatten() {
                            let key = XlsxParser::local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"val" || key == b"name" {
                                sheet.name = Some(val);
                            }
                        }
                        if let Some(name) = sheet.name {
                            part.sheets.push(ExternalLinkSheet {
                                name: Some(name),
                                r_id: None,
                            });
                        }
                    }
                    b"externalBook" => {
                        for attr in e.attributes().flatten() {
                            let key = XlsxParser::local_name(attr.key.as_ref());
                            if key == b"id" || key == b"rid" || key == b"rId" {
                                let rel_id = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(rels) = rels {
                                    if let Some(rel) = rels.get(&rel_id) {
                                        part.target = Some(rel.target.clone());
                                        part.link_type = Some(rel.rel_type.clone());
                                    } else {
                                        part.target = Some(rel_id);
                                    }
                                } else {
                                    part.target = Some(rel_id);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(part)
}

fn parse_slicer_part(xml: &str, path: &str) -> Result<SlicerPart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut slicer = SlicerPart::new();
    slicer.span = Some(SourceSpan::new(path));

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = XlsxParser::local_name(&name_buf);
                if local == b"slicer" {
                    for attr in e.attributes().flatten() {
                        let key = XlsxParser::local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"name" => slicer.name = Some(val),
                            b"caption" => slicer.caption = Some(val),
                            b"cache" | b"cacheId" => slicer.cache_id = Some(val),
                            b"ref" | b"pivotRef" => slicer.target_ref = Some(val),
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(slicer)
}

fn parse_timeline_part(xml: &str, path: &str) -> Result<TimelinePart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut timeline = TimelinePart::new();
    timeline.span = Some(SourceSpan::new(path));

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = XlsxParser::local_name(&name_buf);
                if local == b"timeline" {
                    for attr in e.attributes().flatten() {
                        let key = XlsxParser::local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            b"name" => timeline.name = Some(val),
                            b"cache" | b"cacheId" => timeline.cache_id = Some(val),
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(timeline)
}

fn parse_query_table_part(xml: &str, path: &str) -> Result<QueryTablePart, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut query = QueryTablePart::new();
    query.span = Some(SourceSpan::new(path));

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = XlsxParser::local_name(&name_buf);
                match local {
                    b"queryTable" => {
                        for attr in e.attributes().flatten() {
                            let key = XlsxParser::local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key {
                                b"name" => query.name = Some(val),
                                b"connectionId" | b"connection" => query.connection_id = Some(val),
                                _ => {}
                            }
                        }
                    }
                    b"dbPr" => {
                        for attr in e.attributes().flatten() {
                            let key = XlsxParser::local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"command" {
                                query.command = Some(val.clone());
                            }
                            if key == b"connection" {
                                query.connection_id = Some(val);
                            }
                        }
                    }
                    b"webPr" => {
                        for attr in e.attributes().flatten() {
                            let key = XlsxParser::local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == b"url" {
                                query.url = Some(val.clone());
                                query.source = Some(val);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(query)
}

fn parse_sheet_metadata(xml: &str, path: &str) -> Result<SheetMetadata, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut metadata = SheetMetadata::new();
    metadata.span = Some(SourceSpan::new(path));

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_buf = e.name().as_ref().to_vec();
                let local = XlsxParser::local_name(&name_buf);
                match local {
                    b"metadataType" => {
                        let mut mtype = SheetMetadataType::new();
                        for attr in e.attributes().flatten() {
                            let key = XlsxParser::local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key {
                                b"name" => mtype.name = Some(val),
                                b"minSupportedVersion" => mtype.min_supported_version = Some(val),
                                b"copy" => {
                                    mtype.copy =
                                        Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                }
                                b"update" => {
                                    mtype.update =
                                        Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                }
                                _ => {}
                            }
                        }
                        metadata.metadata_types.push(mtype);
                    }
                    b"cellMetadata" => {
                        for attr in e.attributes().flatten() {
                            if XlsxParser::local_name(attr.key.as_ref()) == b"count" {
                                metadata.cell_metadata_count =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            }
                        }
                    }
                    b"valueMetadata" => {
                        for attr in e.attributes().flatten() {
                            if XlsxParser::local_name(attr.key.as_ref()) == b"count" {
                                metadata.value_metadata_count =
                                    String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(metadata)
}

fn map_cell_error(value: &str) -> CellError {
    match value.trim() {
        "#NULL!" => CellError::Null,
        "#DIV/0!" => CellError::DivZero,
        "#VALUE!" => CellError::Value,
        "#REF!" => CellError::Ref,
        "#NAME?" => CellError::Name,
        "#NUM!" => CellError::Num,
        "#N/A" => CellError::NA,
        "#GETTING_DATA" => CellError::GettingData,
        _ => CellError::Value,
    }
}

fn classify_relationship(rel_type_uri: &str) -> ExternalRefType {
    if rel_type_uri.contains("hyperlink") {
        ExternalRefType::Hyperlink
    } else if rel_type_uri.contains("image") {
        ExternalRefType::Image
    } else if rel_type_uri.contains("oleObject") {
        ExternalRefType::OleLink
    } else if rel_type_uri.contains("externalLink") || rel_type_uri.contains("connections") {
        ExternalRefType::DataConnection
    } else if rel_type_uri == rel_type::ATTACHED_TEMPLATE {
        ExternalRefType::AttachedTemplate
    } else {
        ExternalRefType::Other
    }
}

fn get_rels_path(part_path: &str) -> String {
    if let Some(idx) = part_path.rfind('/') {
        let dir = &part_path[..idx + 1];
        let file = &part_path[idx + 1..];
        format!("{}_rels/{}.rels", dir, file)
    } else {
        format!("_rels/{}.rels", part_path)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::IRNode;

    #[test]
    fn test_parse_workbook_info_sheets() {
        let xml = r#"
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <bookViews>
            <workbookView activeTab="1" firstSheet="0" showHorizontalScroll="1"
                          showVerticalScroll="0" showSheetTabs="1" tabRatio="400"
                          windowWidth="12000" windowHeight="8000" xWindow="120" yWindow="240"/>
          </bookViews>
          <sheets>
            <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
            <sheet name="Hidden" sheetId="2" r:id="rId2" state="hidden"/>
            <sheet name="VeryHidden" sheetId="3" r:id="rId3" state="veryHidden"/>
          </sheets>
        </workbook>
        "#;

        let info = parse_workbook_info(xml).expect("parse workbook info");
        assert_eq!(info.sheets.len(), 3);
        assert_eq!(info.sheets[0].name, "Sheet1");
        assert_eq!(info.sheets[1].state, SheetState::Hidden);
        assert_eq!(info.sheets[2].state, SheetState::VeryHidden);
        let props = info.workbook_properties.expect("workbook props");
        assert_eq!(props.active_tab, Some(1));
        assert_eq!(props.first_sheet, Some(0));
        assert_eq!(props.show_horizontal_scroll, Some(true));
        assert_eq!(props.show_vertical_scroll, Some(false));
        assert_eq!(props.show_sheet_tabs, Some(true));
        assert_eq!(props.tab_ratio, Some(400));
        assert_eq!(props.window_width, Some(12000));
        assert_eq!(props.window_height, Some(8000));
        assert_eq!(props.x_window, Some(120));
        assert_eq!(props.y_window, Some(240));
    }

    #[test]
    fn test_parse_shared_strings() {
        let xml = r#"
        <sst>
          <si><t>Hello</t></si>
          <si><r><t>Foo</t></r><r><t>Bar</t></r></si>
        </sst>
        "#;
        let (table, strings) = parse_shared_strings_table(xml).expect("shared strings");
        assert_eq!(strings, vec!["Hello", "FooBar"]);
        assert_eq!(table.items.len(), 2);
        assert_eq!(
            table.items[1].runs,
            vec!["Foo".to_string(), "Bar".to_string()]
        );
    }

    #[test]
    fn test_parse_worksheet_cells() {
        let xml = r#"
        <worksheet>
          <cols>
            <col min="1" max="2" width="10" customWidth="1"/>
          </cols>
          <sheetData>
            <row r="1">
              <c r="A1" t="s"><v>0</v></c>
              <c r="B1"><v>42</v></c>
              <c r="C1" t="b"><v>1</v></c>
              <c r="D1" t="e"><v>#REF!</v></c>
              <c r="E1"><f>SUM(A1:B1)</f><v>3</v></c>
              <c r="F1"><is><t>Inline</t></is></c>
            </row>
          </sheetData>
          <mergeCells>
            <mergeCell ref="A1:B1"/>
          </mergeCells>
        </worksheet>
        "#;

        let mut parser = XlsxParser::new();
        parser.shared_strings = vec!["Hello".to_string()];

        let sheet = SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };

        let mut zip = build_empty_zip();
        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                xml,
                &sheet,
                "xl/worksheets/sheet1.xml",
                &Relationships::default(),
                SheetKind::Worksheet,
            )
            .expect("parse worksheet");
        let store = parser.into_store();

        let worksheet = match store.get(ws_id) {
            Some(IRNode::Worksheet(ws)) => ws,
            _ => panic!("missing worksheet node"),
        };

        assert_eq!(worksheet.columns.len(), 2);
        assert_eq!(worksheet.merged_cells.len(), 1);
        assert_eq!(worksheet.cells.len(), 6);

        let cell_a1 = store
            .get(worksheet.cells[0])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell a1");
        assert!(matches!(cell_a1.value, CellValue::String(ref v) if v == "Hello"));

        let cell_b1 = store
            .get(worksheet.cells[1])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell b1");
        assert!(matches!(cell_b1.value, CellValue::Number(v) if (v - 42.0).abs() < f64::EPSILON));

        let cell_c1 = store
            .get(worksheet.cells[2])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell c1");
        assert!(matches!(cell_c1.value, CellValue::Boolean(true)));

        let cell_d1 = store
            .get(worksheet.cells[3])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell d1");
        assert!(matches!(cell_d1.value, CellValue::Error(CellError::Ref)));

        let cell_e1 = store
            .get(worksheet.cells[4])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell e1");
        assert!(cell_e1.formula.is_some());

        let cell_f1 = store
            .get(worksheet.cells[5])
            .and_then(|n| match n {
                IRNode::Cell(c) => Some(c),
                _ => None,
            })
            .expect("cell f1");
        assert!(matches!(cell_f1.value, CellValue::InlineString(ref v) if v == "Inline"));
    }

    #[test]
    fn test_parse_worksheet_properties() {
        let xml = r#"
        <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <dimension ref="A1:C5"/>
          <sheetPr>
            <tabColor rgb="FF00FF"/>
          </sheetPr>
          <pageMargins left="0.5" right="0.6" top="0.7" bottom="0.8" header="0.3" footer="0.4"/>
          <sheetData/>
        </worksheet>
        "#;

        let mut parser = XlsxParser::new();
        let sheet = SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };
        let mut zip = build_empty_zip();
        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                xml,
                &sheet,
                "xl/worksheets/sheet1.xml",
                &Relationships::default(),
                SheetKind::Worksheet,
            )
            .expect("worksheet");
        let store = parser.into_store();

        let worksheet = match store.get(ws_id) {
            Some(IRNode::Worksheet(ws)) => ws,
            _ => panic!("expected worksheet"),
        };
        assert_eq!(worksheet.dimension.as_deref(), Some("A1:C5"));
        assert_eq!(worksheet.tab_color.as_deref(), Some("rgb:FF00FF"));
        let margins = worksheet.page_margins.as_ref().expect("margins");
        assert_eq!(margins.left, Some(0.5));
        assert_eq!(margins.right, Some(0.6));
        assert_eq!(margins.top, Some(0.7));
        assert_eq!(margins.bottom, Some(0.8));
        assert_eq!(margins.header, Some(0.3));
        assert_eq!(margins.footer, Some(0.4));
    }

    #[test]
    fn test_parse_worksheet_drawing_pic_and_chart() {
        let sheet_xml = r#"
        <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetData/>
        </worksheet>
        "#;

        let drawing_xml = r#"
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                 xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <xdr:twoCellAnchor>
            <xdr:pic>
              <xdr:nvPicPr>
                <xdr:cNvPr id="1" name="Picture 1" descr="Alt text"/>
              </xdr:nvPicPr>
              <xdr:blipFill>
                <a:blip r:embed="rIdImg"/>
              </xdr:blipFill>
            </xdr:pic>
          </xdr:twoCellAnchor>
          <xdr:graphicFrame>
            <xdr:nvGraphicFramePr>
              <xdr:cNvPr id="2" name="Chart 1"/>
            </xdr:nvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rIdChart"/>
              </a:graphicData>
            </a:graphic>
          </xdr:graphicFrame>
        </xdr:wsDr>
        "#;

        let chart_xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:barChart>
              <c:ser><c:tx><c:v>2019</c:v></c:tx></c:ser>
              <c:ser><c:tx><c:v>2020</c:v></c:tx></c:ser>
            </c:barChart>
          </c:chart>
        </c:chartSpace>
        "#;

        let drawing_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image1.png"/>
          <Relationship Id="rIdChart"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="../charts/chart1.xml"/>
        </Relationships>
        "#;

        let sheet_rels = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdDraw"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
            Target="../drawings/drawing1.xml"/>
        </Relationships>
        "#;

        let mut zip = build_zip_with_entries(vec![
            ("xl/drawings/drawing1.xml", drawing_xml),
            ("xl/drawings/_rels/drawing1.xml.rels", drawing_rels),
            ("xl/charts/chart1.xml", chart_xml),
        ]);

        let mut parser = XlsxParser::new();
        let sheet = SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };
        let rels = Relationships::parse(sheet_rels).expect("sheet rels");

        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                sheet_xml,
                &sheet,
                "xl/worksheets/sheet1.xml",
                &rels,
                SheetKind::Worksheet,
            )
            .expect("parse worksheet");
        let store = parser.into_store();
        let ws = match store.get(ws_id) {
            Some(IRNode::Worksheet(w)) => w,
            _ => panic!("missing worksheet"),
        };
        assert_eq!(ws.drawings.len(), 1);
        let drawing = match store.get(ws.drawings[0]) {
            Some(IRNode::WorksheetDrawing(d)) => d,
            _ => panic!("missing drawing"),
        };
        assert_eq!(drawing.shapes.len(), 2);
    }

    #[test]
    fn test_parse_chartsheet_chart() {
        let chartsheet_xml = r#"
        <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <chart r:id="rIdChart"/>
        </chartsheet>
        "#;
        let chart_xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:lineChart/>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdChart"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            Target="../charts/chart1.xml"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels");
        let mut zip = build_zip_with_entries(vec![("xl/charts/chart1.xml", chart_xml)]);
        let mut parser = XlsxParser::new();
        let sheet = SheetInfo {
            name: "Chart1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };
        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                chartsheet_xml,
                &sheet,
                "xl/chartsheets/sheet1.xml",
                &rels,
                SheetKind::ChartSheet,
            )
            .expect("chartsheet");
        let store = parser.into_store();
        let ws = match store.get(ws_id) {
            Some(IRNode::Worksheet(w)) => w,
            _ => panic!("missing worksheet"),
        };
        assert_eq!(ws.drawings.len(), 1);
        let drawing = match store.get(ws.drawings[0]) {
            Some(IRNode::WorksheetDrawing(d)) => d,
            _ => panic!("missing drawing"),
        };
        assert_eq!(drawing.shapes.len(), 1);
        let shape = match store.get(drawing.shapes[0]) {
            Some(IRNode::Shape(s)) => s,
            _ => panic!("missing shape"),
        };
        assert_eq!(shape.shape_type, ShapeType::Chart);
        assert_eq!(shape.media_target.as_deref(), Some("xl/charts/chart1.xml"));
    }

    #[test]
    fn test_parse_dialogsheet_kind() {
        let dialog_xml = r#"
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1" t="str"><v>Hello</v></c>
            </row>
          </sheetData>
        </worksheet>
        "#;
        let mut parser = XlsxParser::new();
        let mut zip = build_empty_zip();
        let sheet = SheetInfo {
            name: "Dialog1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };
        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                dialog_xml,
                &sheet,
                "xl/dialogsheets/sheet1.xml",
                &Relationships::default(),
                SheetKind::DialogSheet,
            )
            .expect("dialogsheet");
        let store = parser.into_store();
        let ws = match store.get(ws_id) {
            Some(IRNode::Worksheet(w)) => w,
            _ => panic!("missing worksheet"),
        };
        assert_eq!(ws.kind, SheetKind::DialogSheet);
        assert_eq!(ws.cells.len(), 1);
    }

    #[test]
    fn test_parse_styles_minimal() {
        let xml = r#"
        <styleSheet>
          <numFmts count="1">
            <numFmt numFmtId="164" formatCode="0.00"/>
          </numFmts>
          <fonts count="1">
            <font>
              <sz val="11"/>
              <name val="Calibri"/>
              <b/>
              <color rgb="FF0000"/>
            </font>
          </fonts>
          <fills count="1">
            <fill>
              <patternFill patternType="solid">
                <fgColor rgb="FFFF00"/>
              </patternFill>
            </fill>
          </fills>
          <borders count="1">
            <border>
              <left style="thin"><color rgb="FF00FF"/></left>
            </border>
          </borders>
          <cellXfs count="1">
            <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"
                applyNumberFormat="1" applyAlignment="1" applyProtection="1" quotePrefix="1">
              <alignment horizontal="center" wrapText="1" indent="2" textRotation="45"
                         shrinkToFit="1" readingOrder="1"/>
              <protection locked="1" hidden="0"/>
            </xf>
          </cellXfs>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>
          </cellStyleXfs>
          <dxfs count="1">
            <dxf>
              <numFmt numFmtId="200" formatCode="0.00"/>
              <font><b/><color rgb="FF0000"/></font>
              <fill><patternFill patternType="solid"><fgColor rgb="00FF00"/></patternFill></fill>
            </dxf>
          </dxfs>
          <tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">
            <tableStyle name="TableStyleMedium2" pivot="0" table="1"/>
          </tableStyles>
        </styleSheet>
        "#;

        let styles = parse_styles(xml, "xl/styles.xml").expect("styles");
        assert_eq!(styles.number_formats.len(), 1);
        assert_eq!(styles.fonts.len(), 1);
        assert_eq!(styles.fills.len(), 1);
        assert_eq!(styles.borders.len(), 1);
        assert_eq!(styles.cell_xfs.len(), 1);
        assert_eq!(styles.cell_style_xfs.len(), 1);
        assert_eq!(styles.dxfs.len(), 1);
        assert!(styles.table_styles.is_some());
        assert_eq!(styles.fonts[0].name.as_deref(), Some("Calibri"));
        assert!(styles.fonts[0].bold);
        assert_eq!(styles.cell_xfs[0].apply_number_format, true);
        assert_eq!(styles.dxfs[0].num_fmt.as_ref().map(|n| n.id), Some(200));
        assert_eq!(
            styles
                .table_styles
                .as_ref()
                .unwrap()
                .default_table_style
                .as_deref(),
            Some("TableStyleMedium2")
        );
        assert_eq!(styles.table_styles.as_ref().unwrap().styles.len(), 1);
        assert_eq!(
            styles.table_styles.as_ref().unwrap().styles[0].name,
            "TableStyleMedium2"
        );
        assert_eq!(
            styles.cell_xfs[0]
                .alignment
                .as_ref()
                .and_then(|a| a.horizontal.as_deref()),
            Some("center")
        );
        assert_eq!(
            styles.cell_xfs[0].alignment.as_ref().and_then(|a| a.indent),
            Some(2)
        );
        assert_eq!(
            styles.cell_xfs[0]
                .alignment
                .as_ref()
                .and_then(|a| a.text_rotation),
            Some(45)
        );
        assert!(styles.cell_xfs[0]
            .alignment
            .as_ref()
            .map(|a| a.shrink_to_fit)
            .unwrap_or(false));
        assert_eq!(
            styles.cell_xfs[0]
                .protection
                .as_ref()
                .and_then(|p| p.locked),
            Some(true)
        );
    }

    #[test]
    fn test_parse_calc_chain() {
        let xml = r#"
        <calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <c r="A1" i="0" l="1" s="1"/>
          <c r="B2" i="2"/>
        </calcChain>
        "#;
        let chain = parse_calc_chain(xml, "xl/calcChain.xml").expect("calc chain");
        assert_eq!(chain.entries.len(), 2);
        assert_eq!(chain.entries[0].cell_ref, "A1");
        assert_eq!(chain.entries[0].level, Some(1));
        assert_eq!(chain.entries[0].new_value, Some(true));
        assert_eq!(chain.entries[1].cell_ref, "B2");
        assert_eq!(chain.entries[1].index, Some(2));
    }

    #[test]
    fn test_parse_sheet_comments() {
        let xml = r#"
        <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <authors>
            <author>Alice</author>
            <author>Bob</author>
          </authors>
          <commentList>
            <comment ref="A1" authorId="0">
              <text><r><t>Hello</t></r></text>
            </comment>
            <comment ref="B2" authorId="1">
              <text><t>World</t></text>
            </comment>
          </commentList>
        </comments>
        "#;
        let comments =
            parse_sheet_comments(xml, "xl/comments1.xml", Some("Sheet1")).expect("comments");
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[0].cell_ref, "A1");
        assert_eq!(comments[0].author.as_deref(), Some("Alice"));
        assert_eq!(comments[0].text, "Hello");
        assert_eq!(comments[1].cell_ref, "B2");
        assert_eq!(comments[1].author.as_deref(), Some("Bob"));
        assert_eq!(comments[1].text, "World");
    }

    #[test]
    fn test_parse_conditional_and_validation() {
        let xml = r#"
        <worksheet>
          <sheetData/>
          <conditionalFormatting sqref="A1:A10">
            <cfRule type="expression" priority="1">
              <formula>SUM(A1)&gt;10</formula>
            </cfRule>
          </conditionalFormatting>
          <dataValidations count="1">
            <dataValidation type="list" allowBlank="1" sqref="B1">
              <formula1>"Yes,No"</formula1>
            </dataValidation>
          </dataValidations>
        </worksheet>
        "#;

        let mut parser = XlsxParser::new();
        let sheet = SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            state: SheetState::Visible,
        };
        let mut zip = build_empty_zip();
        let ws_id = parser
            .parse_worksheet(
                &mut zip,
                xml,
                &sheet,
                "xl/worksheets/sheet1.xml",
                &Relationships::default(),
                SheetKind::Worksheet,
            )
            .expect("worksheet");
        let store = parser.into_store();

        let worksheet = match store.get(ws_id) {
            Some(IRNode::Worksheet(ws)) => ws,
            _ => panic!("expected worksheet"),
        };
        assert_eq!(worksheet.conditional_formats.len(), 1);
        assert_eq!(worksheet.data_validations.len(), 1);
    }

    #[test]
    fn test_parse_table_and_pivot_definition() {
        let table_xml = r#"
        <table name="Table1" displayName="Table1" ref="A1:B3" headerRowCount="1">
          <tableColumns count="2">
            <tableColumn id="1" name="Col1"/>
            <tableColumn id="2" name="Col2" totalsRowFunction="sum"/>
          </tableColumns>
        </table>
        "#;
        let table = parse_table_definition(table_xml, "xl/tables/table1.xml").expect("table");
        assert_eq!(table.columns.len(), 2);
        assert_eq!(table.ref_range.as_deref(), Some("A1:B3"));

        let pivot_xml = r#"
        <pivotTableDefinition name="Pivot1" cacheId="3">
          <location ref="D1:F10"/>
        </pivotTableDefinition>
        "#;
        let pivot = parse_pivot_table_definition(pivot_xml, "xl/pivotTables/pivotTable1.xml")
            .expect("pivot");
        assert_eq!(pivot.cache_id, Some(3));
        assert_eq!(pivot.ref_range.as_deref(), Some("D1:F10"));
    }

    #[test]
    fn test_parse_pivot_cache_records() {
        let records_xml = r#"
        <pivotCacheRecords count="2">
          <r><n v="1"/><s v="0"/><b v="1"/></r>
          <r><n v="2"/><s v="1"/><b v="0"/></r>
        </pivotCacheRecords>
        "#;

        let records =
            parse_pivot_cache_records(records_xml, "xl/pivotCache/pivotCacheRecords1.xml")
                .expect("records");
        assert_eq!(records.record_count, Some(2));
        assert_eq!(records.field_count, Some(3));
    }

    #[test]
    fn test_parse_xlm_macro_sheet() {
        let workbook_xml = r#"
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Macro1" sheetId="1" r:id="rId1"/>
          </sheets>
        </workbook>
        "#;

        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/macrosheet"
            Target="macrosheets/sheet1.xml"/>
        </Relationships>
        "#;

        let sheet_xml = r#"
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1"><f>EXEC(\"calc\")</f></c>
            </row>
          </sheetData>
        </worksheet>
        "#;

        let mut zip = build_zip_with_entries(vec![("xl/macrosheets/sheet1.xml", sheet_xml)]);
        let rels = Relationships::parse(rels_xml).expect("rels");

        let mut parser = XlsxParser::new();
        let root = parser
            .parse_workbook(&mut zip, workbook_xml, &rels, "xl/workbook.xml")
            .expect("workbook");
        let store = parser.into_store();
        let doc = match store.get(root) {
            Some(IRNode::Document(d)) => d,
            _ => panic!("missing document"),
        };
        assert_eq!(doc.security.xlm_macros.len(), 1);
        assert!(doc.security.xlm_macros[0].macro_cells.len() >= 1);
    }

    #[test]
    fn test_parse_xlm_auto_open_defined_name() {
        let workbook_xml = r#"
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <definedNames>
            <definedName name="_xlnm.Auto_Open">Macro1!$A$1</definedName>
          </definedNames>
          <sheets>
            <sheet name="Macro1" sheetId="1" r:id="rId1"/>
          </sheets>
        </workbook>
        "#;

        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/macrosheet"
            Target="macrosheets/sheet1.xml"/>
        </Relationships>
        "#;

        let sheet_xml = r#"
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1"><f>RUN(\"TEST\")</f></c>
            </row>
          </sheetData>
        </worksheet>
        "#;

        let mut zip = build_zip_with_entries(vec![("xl/macrosheets/sheet1.xml", sheet_xml)]);
        let rels = Relationships::parse(rels_xml).expect("rels");

        let mut parser = XlsxParser::new();
        let root = parser
            .parse_workbook(&mut zip, workbook_xml, &rels, "xl/workbook.xml")
            .expect("workbook");
        let store = parser.into_store();
        let doc = match store.get(root) {
            Some(IRNode::Document(d)) => d,
            _ => panic!("missing document"),
        };
        assert_eq!(doc.security.xlm_macros.len(), 1);
        assert!(doc.security.xlm_macros[0].has_auto_open);
    }

    #[test]
    fn test_parse_chart_xml() {
        let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart>
            <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:lineChart>
              <c:ser><c:tx><c:v>Q1</c:v></c:tx></c:ser>
            </c:lineChart>
          </c:chart>
        </c:chartSpace>
        "#;
        let mut parser = XlsxParser::new();
        let id = parser
            .parse_chart(xml, "xl/charts/chart2.xml")
            .expect("chart");
        let store = parser.into_store();
        let chart = match store.get(id) {
            Some(IRNode::ChartData(c)) => c,
            _ => panic!("missing chart"),
        };
        assert!(chart
            .chart_type
            .as_deref()
            .unwrap_or("")
            .contains("lineChart"));
        assert_eq!(chart.title.as_deref(), Some("Revenue"));
        assert_eq!(chart.series.len(), 1);
        assert_eq!(chart.series_data.len(), 1);
    }

    #[test]
    fn test_parse_connections_xml_targets() {
        let xml = r#"
        <connections>
          <connection id="1" name="Conn1" type="1">
            <webPr url="https://example.com/data"/>
          </connection>
          <connection id="2" name="Conn2" type="2">
            <dbPr connection="DatabaseName" command="SELECT * FROM foo" commandType="2"/>
          </connection>
        </connections>
        "#;
        let part = parse_connections_part(xml, "xl/connections.xml").expect("connections");
        assert_eq!(part.entries.len(), 2);
        assert_eq!(part.entries[0].connection_id, Some(1));
        assert_eq!(
            part.entries[0].url.as_deref(),
            Some("https://example.com/data")
        );
        assert_eq!(part.entries[1].connection_id, Some(2));
        assert_eq!(part.entries[1].connection.as_deref(), Some("DatabaseName"));
        assert_eq!(
            part.entries[1].command.as_deref(),
            Some("SELECT * FROM foo")
        );
        assert_eq!(part.entries[1].command_type, Some(2));
        let targets = connection_targets(&part);
        assert!(targets.contains(&"https://example.com/data".to_string()));
        assert!(targets.contains(&"DatabaseName".to_string()));
    }

    #[test]
    fn test_parse_sheet_metadata() {
        let xml = r#"
        <metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <metadataTypes count="1">
            <metadataType name="XLDynamicArray" minSupportedVersion="120000" copy="1" update="0"/>
          </metadataTypes>
          <cellMetadata count="2"/>
          <valueMetadata count="3"/>
        </metadata>
        "#;
        let meta = parse_sheet_metadata(xml, "xl/metadata.xml").expect("metadata");
        assert_eq!(meta.metadata_types.len(), 1);
        assert_eq!(
            meta.metadata_types[0].name.as_deref(),
            Some("XLDynamicArray")
        );
        assert_eq!(meta.cell_metadata_count, Some(2));
        assert_eq!(meta.value_metadata_count, Some(3));
    }

    #[test]
    fn test_parse_people_part() {
        let xml = r#"
        <ppl:people xmlns:ppl="http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes">
          <ppl:person ppl:id="p1" ppl:userId="user1" ppl:displayName="Alice" ppl:initials="A"/>
          <ppl:person ppl:id="p2" ppl:userId="user2" ppl:displayName="Bob"/>
        </ppl:people>
        "#;
        let mut parser = XlsxParser::new();
        let mut zip = build_zip_with_entries(vec![("xl/persons/person.xml", xml)]);
        let workbook_xml =
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
        let doc_id = parser
            .parse_workbook(
                &mut zip,
                workbook_xml,
                &Relationships::default(),
                "xl/workbook.xml",
            )
            .expect("workbook");
        let store = parser.into_store();

        let doc = match store.get(doc_id) {
            Some(IRNode::Document(d)) => d,
            _ => panic!("missing document"),
        };
        assert_eq!(doc.shared_parts.len(), 1);

        let people = match store.get(doc.shared_parts[0]) {
            Some(IRNode::PeoplePart(p)) => p,
            _ => panic!("missing people part"),
        };
        assert_eq!(people.people.len(), 2);
        assert_eq!(people.people[0].display_name.as_deref(), Some("Alice"));
        assert_eq!(people.people[1].display_name.as_deref(), Some("Bob"));
    }

    #[test]
    fn test_parse_external_link_part() {
        let xml = r#"
        <externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <externalBook r:id="rId1">
            <sheetNames>
              <sheetName val="SheetA"/>
              <sheetName val="SheetB"/>
            </sheetNames>
          </externalBook>
        </externalLink>
        "#;
        let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
            Target="file:///C:/data.xlsx"/>
        </Relationships>
        "#;
        let rels = Relationships::parse(rels_xml).expect("rels");
        let part = parse_external_link_part(xml, "xl/externalLinks/externalLink1.xml", Some(&rels))
            .expect("external link");
        assert_eq!(part.target.as_deref(), Some("file:///C:/data.xlsx"));
        assert_eq!(part.sheets.len(), 2);
        assert_eq!(part.sheets[0].name.as_deref(), Some("SheetA"));
    }

    #[test]
    fn test_parse_slicer_part() {
        let xml = r#"
        <slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                name="Slicer1" caption="Region" cache="1" />
        "#;
        let slicer = parse_slicer_part(xml, "xl/slicers/slicer1.xml").expect("slicer");
        assert_eq!(slicer.name.as_deref(), Some("Slicer1"));
        assert_eq!(slicer.caption.as_deref(), Some("Region"));
        assert_eq!(slicer.cache_id.as_deref(), Some("1"));
    }

    #[test]
    fn test_parse_timeline_part() {
        let xml = r#"
        <timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                  name="Timeline1" cache="2" />
        "#;
        let timeline = parse_timeline_part(xml, "xl/timelines/timeline1.xml").expect("timeline");
        assert_eq!(timeline.name.as_deref(), Some("Timeline1"));
        assert_eq!(timeline.cache_id.as_deref(), Some("2"));
    }

    #[test]
    fn test_parse_query_table_part() {
        let xml = r#"
        <queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    name="Query1" connectionId="7">
          <dbPr command="SELECT * FROM tbl"/>
          <webPr url="https://example.com/data"/>
        </queryTable>
        "#;
        let query = parse_query_table_part(xml, "xl/queryTables/queryTable1.xml").expect("query");
        assert_eq!(query.name.as_deref(), Some("Query1"));
        assert_eq!(query.connection_id.as_deref(), Some("7"));
        assert_eq!(query.command.as_deref(), Some("SELECT * FROM tbl"));
        assert_eq!(query.url.as_deref(), Some("https://example.com/data"));
    }

    fn build_empty_zip() -> SecureZipReader<std::io::Cursor<Vec<u8>>> {
        build_zip_with_entries(Vec::new())
    }

    fn build_zip_with_entries(
        entries: Vec<(&str, &str)>,
    ) -> SecureZipReader<std::io::Cursor<Vec<u8>>> {
        let mut data = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(&mut data));
            let options = zip::write::FileOptions::<()>::default();
            for (path, contents) in entries {
                writer.start_file(path, options).expect("start file");
                use std::io::Write;
                writer.write_all(contents.as_bytes()).expect("write file");
            }
            writer.finish().expect("finish zip");
        }
        SecureZipReader::new(std::io::Cursor::new(data), Default::default()).expect("zip")
    }
}
