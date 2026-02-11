//! Workbook parsing helpers for XLSX.

use super::SheetState;
use crate::error::ParseError;
use docir_core::ir::{DefinedName, WorkbookProperties};
use docir_core::types::{NodeId, SourceSpan};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

#[derive(Debug, Clone)]
pub(super) struct SheetInfo {
    pub(super) name: String,
    pub(super) sheet_id: u32,
    pub(super) rel_id: String,
    pub(super) state: SheetState,
}

#[derive(Debug, Clone)]
pub(super) struct PivotCacheRef {
    pub(super) cache_id: u32,
    pub(super) rel_id: String,
}

#[derive(Debug, Clone)]
pub(super) struct WorkbookInfo {
    pub(super) sheets: Vec<SheetInfo>,
    pub(super) defined_names: Vec<DefinedName>,
    pub(super) workbook_properties: Option<WorkbookProperties>,
    pub(super) pivot_cache_refs: Vec<PivotCacheRef>,
}

pub(super) fn parse_workbook_info(xml: &str) -> Result<WorkbookInfo, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut sheets: Vec<SheetInfo> = Vec::new();
    let mut defined_names: Vec<DefinedName> = Vec::new();
    let mut pivot_cache_refs: Vec<PivotCacheRef> = Vec::new();
    let mut workbook_properties: Option<WorkbookProperties> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"sheet" => parse_sheet_info(&e, &mut sheets)?,
                b"definedName" => {
                    if let Some(def) = parse_defined_name(&mut reader, &e)? {
                        defined_names.push(def);
                    }
                }
                b"workbookPr" => {
                    parse_workbook_pr(&e, &mut workbook_properties);
                }
                b"workbookView" => {
                    parse_workbook_view(&e, &mut workbook_properties);
                }
                b"calcPr" => {
                    parse_calc_pr(&e, &mut workbook_properties);
                }
                b"workbookProtection" => {
                    let props = workbook_properties.get_or_insert_with(WorkbookProperties::new);
                    props.workbook_protected = true;
                }
                b"pivotCache" => {
                    parse_pivot_cache_ref(&e, &mut pivot_cache_refs);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"sheet" => parse_sheet_info(&e, &mut sheets)?,
                b"definedName" => {
                    if let Some(def) = parse_defined_name_empty(&e) {
                        defined_names.push(def);
                    }
                }
                b"workbookPr" => {
                    parse_workbook_pr(&e, &mut workbook_properties);
                }
                b"workbookView" => {
                    parse_workbook_view(&e, &mut workbook_properties);
                }
                b"calcPr" => {
                    parse_calc_pr(&e, &mut workbook_properties);
                }
                b"workbookProtection" => {
                    let props = workbook_properties.get_or_insert_with(WorkbookProperties::new);
                    props.workbook_protected = true;
                }
                b"pivotCache" => {
                    parse_pivot_cache_ref(&e, &mut pivot_cache_refs);
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: "xl/workbook.xml".to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(WorkbookInfo {
        sheets,
        defined_names,
        workbook_properties,
        pivot_cache_refs,
    })
}

pub(super) fn auto_open_target_from_defined_name(name: &DefinedName) -> Option<Option<String>> {
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
        .map_err(|e| ParseError::Xml {
            file: "xl/workbook.xml".to_string(),
            message: e.to_string(),
        })?;

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
