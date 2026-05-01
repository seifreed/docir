//! Workbook parsing helpers for XLSX.

use super::SheetState;
use crate::error::ParseError;
use crate::xml_utils::{
    dispatch_start_or_empty, reader_from_str, scan_xml_events_with_reader, xml_error,
    XmlScanControl,
};
use docir_core::ir::{DefinedName, WorkbookProperties};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::BytesStart;
use quick_xml::Reader;

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
    match e.name().as_ref() {
        b"sheet" => parse_sheet_info(e, sheets)?,
        b"definedName" => {
            if is_start {
                if let Some(def) = parse_defined_name(reader, e)? {
                    defined_names.push(def);
                }
            } else if let Some(def) = parse_defined_name_empty(e) {
                defined_names.push(def);
            }
        }
        b"workbookPr" => parse_workbook_pr(e, workbook_properties),
        b"workbookView" => parse_workbook_view(e, workbook_properties),
        b"calcPr" => parse_calc_pr(e, workbook_properties),
        b"workbookProtection" => {
            let props = workbook_properties.get_or_insert_with(WorkbookProperties::new);
            props.workbook_protected = true;
        }
        b"pivotCache" => parse_pivot_cache_ref(e, pivot_cache_refs),
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

    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"name" => name = Some(String::from_utf8_lossy(&attr.value).to_string()),
            b"sheetId" => sheet_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"r:id" => rel_id = Some(String::from_utf8_lossy(&attr.value).to_string()),
            b"state" => {
                let val = String::from_utf8_lossy(&attr.value);
                state = match val.as_ref() {
                    "hidden" => SheetState::Hidden,
                    "veryHidden" => SheetState::VeryHidden,
                    _ => SheetState::Visible,
                };
            }
            _ => {}
        }
    }

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

    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"name" => name = Some(String::from_utf8_lossy(&attr.value).to_string()),
            b"localSheetId" => {
                local_sheet_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
            }
            b"hidden" => {
                let v = String::from_utf8_lossy(&attr.value);
                hidden = v == "1" || v.eq_ignore_ascii_case("true");
            }
            b"comment" => {
                comment = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            _ => {}
        }
    }

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

fn parse_defined_name_empty(start: &BytesStart) -> Option<DefinedName> {
    let mut name = None;
    let mut local_sheet_id = None;
    let mut hidden = false;
    let mut comment = None;

    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"name" => name = Some(String::from_utf8_lossy(&attr.value).to_string()),
            b"localSheetId" => {
                local_sheet_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
            }
            b"hidden" => {
                let v = String::from_utf8_lossy(&attr.value);
                hidden = v == "1" || v.eq_ignore_ascii_case("true");
            }
            b"comment" => {
                comment = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            _ => {}
        }
    }

    name.map(|name| DefinedName {
        id: NodeId::new(),
        name,
        value: String::new(),
        local_sheet_id,
        hidden,
        comment,
        span: Some(SourceSpan::new("xl/workbook.xml")),
    })
}

fn parse_workbook_pr(start: &BytesStart, props: &mut Option<WorkbookProperties>) {
    let props = props.get_or_insert_with(WorkbookProperties::new);
    for attr in start.attributes().flatten() {
        if attr.key.as_ref() == b"date1904" {
            let v = String::from_utf8_lossy(&attr.value);
            props.date1904 = Some(v == "1" || v.eq_ignore_ascii_case("true"));
        }
    }
}

fn parse_calc_pr(start: &BytesStart, props: &mut Option<WorkbookProperties>) {
    let props = props.get_or_insert_with(WorkbookProperties::new);
    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"calcMode" => {
                props.calc_mode = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"fullCalcOnLoad" => {
                let v = String::from_utf8_lossy(&attr.value);
                props.calc_full = Some(v == "1" || v.eq_ignore_ascii_case("true"));
            }
            b"calcOnSave" => {
                let v = String::from_utf8_lossy(&attr.value);
                props.calc_on_save = Some(v == "1" || v.eq_ignore_ascii_case("true"));
            }
            _ => {}
        }
    }
}

fn parse_workbook_view(start: &BytesStart, props: &mut Option<WorkbookProperties>) {
    let props = props.get_or_insert_with(WorkbookProperties::new);
    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"activeTab" => {
                props.active_tab = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
            }
            b"firstSheet" => {
                props.first_sheet = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
            }
            b"showHorizontalScroll" => {
                let v = String::from_utf8_lossy(&attr.value);
                props.show_horizontal_scroll = Some(v == "1" || v.eq_ignore_ascii_case("true"));
            }
            b"showVerticalScroll" => {
                let v = String::from_utf8_lossy(&attr.value);
                props.show_vertical_scroll = Some(v == "1" || v.eq_ignore_ascii_case("true"));
            }
            b"showSheetTabs" => {
                let v = String::from_utf8_lossy(&attr.value);
                props.show_sheet_tabs = Some(v == "1" || v.eq_ignore_ascii_case("true"));
            }
            b"tabRatio" => {
                props.tab_ratio = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
            }
            b"windowWidth" => {
                props.window_width = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
            }
            b"windowHeight" => {
                props.window_height = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
            }
            b"xWindow" => {
                props.x_window = String::from_utf8_lossy(&attr.value).parse::<i32>().ok();
            }
            b"yWindow" => {
                props.y_window = String::from_utf8_lossy(&attr.value).parse::<i32>().ok();
            }
            _ => {}
        }
    }
}

fn parse_pivot_cache_ref(start: &BytesStart, out: &mut Vec<PivotCacheRef>) {
    let mut cache_id = None;
    let mut rel_id = None;
    for attr in start.attributes().flatten() {
        match attr.key.as_ref() {
            b"cacheId" => {
                cache_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
            }
            b"r:id" => {
                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            _ => {}
        }
    }
    if let (Some(cache_id), Some(rel_id)) = (cache_id, rel_id) {
        out.push(PivotCacheRef { cache_id, rel_id });
    }
}
