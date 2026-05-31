//! Workbook parsing helpers for XLSX.

use super::SheetState;
use crate::error::ParseError;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::visit_attributes;
use crate::xml_utils::{
    XmlScanControl, attr_bool_like, dispatch_start_or_empty, local_name, reader_from_str,
    scan_xml_events_with_reader, xml_error,
};
use docir_core::ir::{DefinedName, WorkbookProperties};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::Reader;
use quick_xml::events::BytesStart;

#[derive(Debug, Clone)]
pub(crate) struct SheetInfo {
    pub(crate) name: String,
    pub(crate) sheet_id: u32,
    pub(crate) rel_id: String,
    pub(crate) state: SheetState,
}

#[derive(Debug, Clone)]
pub(crate) struct PivotCacheRef {
    pub(crate) cache_id: u32,
    pub(crate) rel_id: String,
}

#[derive(Debug, Clone)]
pub(crate) struct WorkbookInfo {
    pub(crate) sheets: Vec<SheetInfo>,
    pub(crate) defined_names: Vec<DefinedName>,
    pub(crate) workbook_properties: Option<WorkbookProperties>,
    pub(crate) pivot_cache_refs: Vec<PivotCacheRef>,
}

pub(crate) fn parse_workbook_info(xml: &str) -> Result<WorkbookInfo, ParseError> {
    let mut reader = reader_from_str(xml);

    let mut buf = Vec::new();
    let mut sheets: Vec<SheetInfo> = Vec::new();
    let mut defined_names: Vec<DefinedName> = Vec::new();
    let mut pivot_cache_refs: Vec<PivotCacheRef> = Vec::new();
    let mut workbook_properties: Option<WorkbookProperties> = None;

    scan_xml_events_with_reader(&mut reader, &mut buf, "xl/workbook.xml", |reader, event| {
        let _ = dispatch_start_or_empty(reader, &event, |reader, e, is_start| {
            handle_workbook_event(
                reader,
                e,
                is_start,
                &mut sheets,
                &mut defined_names,
                &mut pivot_cache_refs,
                &mut workbook_properties,
            )
        })?;
        Ok(XmlScanControl::Continue)
    })?;

    Ok(WorkbookInfo {
        sheets,
        defined_names,
        workbook_properties,
        pivot_cache_refs,
    })
}

fn handle_workbook_event(
    reader: &mut Reader<&[u8]>,
    e: &BytesStart<'_>,
    is_start: bool,
    sheets: &mut Vec<SheetInfo>,
    defined_names: &mut Vec<DefinedName>,
    pivot_cache_refs: &mut Vec<PivotCacheRef>,
    workbook_properties: &mut Option<WorkbookProperties>,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"sheet" => parse_sheet_info(e, sheets)?,
        b"definedName" => {
            if is_start {
                if let Some(def) = parse_defined_name(reader, e)? {
                    defined_names.push(def);
                }
            } else if let Some(def) = parse_defined_name_empty(e)? {
                defined_names.push(def);
            }
        }
        b"workbookPr" => parse_workbook_pr(e, workbook_properties)?,
        b"workbookView" => parse_workbook_view(e, workbook_properties)?,
        b"calcPr" => parse_calc_pr(e, workbook_properties)?,
        b"workbookProtection" => {
            let props = workbook_properties.get_or_insert_with(WorkbookProperties::new);
            props.workbook_protected = true;
        }
        b"pivotCache" => parse_pivot_cache_ref(e, pivot_cache_refs)?,
        _ => {}
    }
    Ok(())
}

pub(crate) fn auto_open_target_from_defined_name(name: &DefinedName) -> Option<Option<String>> {
    let upper = name.name.to_ascii_uppercase();
    if upper == "_XLNM.AUTO_OPEN" || upper == "AUTO_OPEN" || upper == "AUTO.OPEN" {
        let val = name.value.trim();
        if let Some((sheet, _)) = val.split_once('!') {
            let cleaned = sheet.trim().trim_matches('\'').to_string();
            if !cleaned.is_empty() {
                return Some(Some(cleaned));
            }
        }
        return Some(None);
    }
    None
}

fn parse_sheet_info(start: &BytesStart, sheets: &mut Vec<SheetInfo>) -> Result<(), ParseError> {
    let mut name = None;
    let mut sheet_id = None;
    let mut rel_id = None;
    let mut state = SheetState::Visible;

    visit_attributes(start, "xl/workbook.xml", |attr| match attr.key.as_ref() {
        b"name" => name = Some(lossy_attr_value(attr).to_string()),
        b"sheetId" => sheet_id = lossy_attr_value(attr).parse::<u32>().ok(),
        key if key.ends_with(b":id") => rel_id = Some(lossy_attr_value(attr).to_string()),
        b"state" => {
            let val = lossy_attr_value(attr);
            state = match val.as_ref() {
                "hidden" => SheetState::Hidden,
                "veryHidden" => SheetState::VeryHidden,
                _ => SheetState::Visible,
            };
        }
        _ => {}
    })?;

    let info = SheetInfo {
        name: name.ok_or_else(|| ParseError::InvalidStructure("Sheet missing name".to_string()))?,
        sheet_id: sheet_id
            .ok_or_else(|| ParseError::InvalidStructure("Sheet missing sheetId".to_string()))?,
        rel_id: rel_id.ok_or_else(|| {
            ParseError::InvalidStructure("Sheet missing relationship id".to_string())
        })?,
        state,
    };

    sheets.push(info);
    Ok(())
}

fn parse_defined_name(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart,
) -> Result<Option<DefinedName>, ParseError> {
    let mut name = None;
    let mut local_sheet_id = None;
    let mut hidden = false;
    let mut comment = None;

    visit_attributes(start, "xl/workbook.xml", |attr| match attr.key.as_ref() {
        b"name" => name = Some(lossy_attr_value(attr).to_string()),
        b"localSheetId" => local_sheet_id = lossy_attr_value(attr).parse::<u32>().ok(),
        b"hidden" => {
            hidden = attr_bool_like(attr.value.as_ref());
        }
        b"comment" => {
            comment = Some(lossy_attr_value(attr).to_string());
        }
        _ => {}
    })?;

    let value = reader
        .read_text(start.name())
        .map_err(|e| xml_error("xl/workbook.xml", e))?;

    Ok(name.map(|name| DefinedName {
        id: NodeId::new(),
        name,
        value: value.to_string(),
        local_sheet_id,
        hidden,
        comment,
        span: Some(SourceSpan::new("xl/workbook.xml")),
    }))
}

fn parse_defined_name_empty(start: &BytesStart) -> Result<Option<DefinedName>, ParseError> {
    let mut name = None;
    let mut local_sheet_id = None;
    let mut hidden = false;
    let mut comment = None;

    visit_attributes(start, "xl/workbook.xml", |attr| match attr.key.as_ref() {
        b"name" => name = Some(lossy_attr_value(attr).to_string()),
        b"localSheetId" => local_sheet_id = lossy_attr_value(attr).parse::<u32>().ok(),
        b"hidden" => {
            hidden = attr_bool_like(attr.value.as_ref());
        }
        b"comment" => {
            comment = Some(lossy_attr_value(attr).to_string());
        }
        _ => {}
    })?;

    Ok(name.map(|name| DefinedName {
        id: NodeId::new(),
        name,
        value: String::new(),
        local_sheet_id,
        hidden,
        comment,
        span: Some(SourceSpan::new("xl/workbook.xml")),
    }))
}

fn parse_workbook_pr(
    start: &BytesStart,
    props: &mut Option<WorkbookProperties>,
) -> Result<(), ParseError> {
    let props = props.get_or_insert_with(WorkbookProperties::new);
    visit_attributes(start, "xl/workbook.xml", |attr| {
        if attr.key.as_ref() == b"date1904" {
            props.date1904 = Some(attr_bool_like(attr.value.as_ref()));
        }
    })
}

fn parse_calc_pr(
    start: &BytesStart,
    props: &mut Option<WorkbookProperties>,
) -> Result<(), ParseError> {
    let props = props.get_or_insert_with(WorkbookProperties::new);
    visit_attributes(start, "xl/workbook.xml", |attr| match attr.key.as_ref() {
        b"calcMode" => {
            props.calc_mode = Some(lossy_attr_value(attr).to_string());
        }
        b"fullCalcOnLoad" => {
            props.calc_full = Some(attr_bool_like(attr.value.as_ref()));
        }
        b"calcOnSave" => {
            props.calc_on_save = Some(attr_bool_like(attr.value.as_ref()));
        }
        _ => {}
    })
}

fn parse_workbook_view(
    start: &BytesStart,
    props: &mut Option<WorkbookProperties>,
) -> Result<(), ParseError> {
    let props = props.get_or_insert_with(WorkbookProperties::new);
    visit_attributes(start, "xl/workbook.xml", |attr| match attr.key.as_ref() {
        b"activeTab" => {
            props.active_tab = lossy_attr_value(attr).parse::<u32>().ok();
        }
        b"firstSheet" => {
            props.first_sheet = lossy_attr_value(attr).parse::<u32>().ok();
        }
        b"showHorizontalScroll" => {
            props.show_horizontal_scroll = Some(attr_bool_like(attr.value.as_ref()));
        }
        b"showVerticalScroll" => {
            props.show_vertical_scroll = Some(attr_bool_like(attr.value.as_ref()));
        }
        b"showSheetTabs" => {
            props.show_sheet_tabs = Some(attr_bool_like(attr.value.as_ref()));
        }
        b"tabRatio" => {
            props.tab_ratio = lossy_attr_value(attr).parse::<u32>().ok();
        }
        b"windowWidth" => {
            props.window_width = lossy_attr_value(attr).parse::<u32>().ok();
        }
        b"windowHeight" => {
            props.window_height = lossy_attr_value(attr).parse::<u32>().ok();
        }
        b"xWindow" => {
            props.x_window = lossy_attr_value(attr).parse::<i32>().ok();
        }
        b"yWindow" => {
            props.y_window = lossy_attr_value(attr).parse::<i32>().ok();
        }
        _ => {}
    })
}

fn parse_pivot_cache_ref(
    start: &BytesStart,
    out: &mut Vec<PivotCacheRef>,
) -> Result<(), ParseError> {
    let mut cache_id = None;
    let mut rel_id = None;
    visit_attributes(start, "xl/workbook.xml", |attr| match attr.key.as_ref() {
        b"cacheId" => {
            cache_id = lossy_attr_value(attr).parse::<u32>().ok();
        }
        key if key.ends_with(b":id") => {
            rel_id = Some(lossy_attr_value(attr).to_string());
        }
        _ => {}
    })?;
    if let (Some(cache_id), Some(rel_id)) = (cache_id, rel_id) {
        out.push(PivotCacheRef { cache_id, rel_id });
    }
    Ok(())
}
