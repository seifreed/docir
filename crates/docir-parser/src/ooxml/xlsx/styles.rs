//! XLSX styles parsing.

use crate::error::ParseError;
use crate::xml_utils::xml_error;
use crate::xml_utils::{attr_f64, attr_value, reader_from_str_with_options};
use docir_core::ir::{
    BorderDef, BorderSide, CellAlignment, CellFormat, CellProtection, DxfStyle, FillDef, FontDef,
    NumberFormat, SpreadsheetStyles, TableStyleDef, TableStyleInfo,
};
use docir_core::types::SourceSpan;
use quick_xml::events::BytesEnd;
use quick_xml::events::{BytesStart, Event};

struct StylesParseState {
    in_num_fmts: bool,
    in_fonts: bool,
    in_fills: bool,
    in_borders: bool,
    in_cell_xfs: bool,
    in_cell_style_xfs: bool,
    in_dxfs: bool,
    in_table_styles: bool,
    current_font: Option<FontDef>,
    current_fill: Option<FillDef>,
    current_border: Option<BorderDef>,
    current_border_side: Option<(String, BorderSide)>,
    current_xf: Option<CellFormat>,
    current_xf_is_style: bool,
    current_dxf: Option<DxfStyle>,
    current_dxf_font: Option<FontDef>,
    current_dxf_fill: Option<FillDef>,
    current_dxf_border: Option<BorderDef>,
    current_dxf_border_side: Option<(String, BorderSide)>,
}

impl StylesParseState {
    fn new() -> Self {
        Self {
            in_num_fmts: false,
            in_fonts: false,
            in_fills: false,
            in_borders: false,
            in_cell_xfs: false,
            in_cell_style_xfs: false,
            in_dxfs: false,
            in_table_styles: false,
            current_font: None,
            current_fill: None,
            current_border: None,
            current_border_side: None,
            current_xf: None,
            current_xf_is_style: false,
            current_dxf: None,
            current_dxf_font: None,
            current_dxf_fill: None,
            current_dxf_border: None,
            current_dxf_border_side: None,
        }
    }
}

pub(crate) fn parse_styles(xml: &str, styles_path: &str) -> Result<SpreadsheetStyles, ParseError> {
    let mut reader = reader_from_str_with_options(xml, true, true);

    let mut styles = SpreadsheetStyles::new();
    styles.span = Some(SourceSpan::new(styles_path));

    let mut buf = Vec::new();
    let mut state = StylesParseState::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => handle_styles_start(&e, &mut state, &mut styles)?,
            Ok(Event::End(e)) => handle_styles_end(&e, &mut state, &mut styles),
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(styles_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(styles)
}

fn handle_styles_start(
    e: &BytesStart<'_>,
    state: &mut StylesParseState,
    styles: &mut SpreadsheetStyles,
) -> Result<(), ParseError> {
    match e.name().as_ref() {
        b"numFmts" => state.in_num_fmts = true,
        b"numFmt" if state.in_num_fmts => {
            if let Some(fmt) = parse_number_format(e) {
                styles.number_formats.push(fmt);
            }
        }
        b"numFmt" if state.in_dxfs => {
            if let Some(fmt) = parse_number_format(e) {
                if let Some(dxf) = state.current_dxf.as_mut() {
                    dxf.num_fmt = Some(fmt);
                }
            }
        }
        b"fonts" => state.in_fonts = true,
        b"font" if state.in_fonts => {
            state.current_font = Some(FontDef {
                name: None,
                size: None,
                bold: false,
                italic: false,
                underline: false,
                color: None,
            });
        }
        b"font" if state.in_dxfs => {
            state.current_dxf_font = Some(FontDef {
                name: None,
                size: None,
                bold: false,
                italic: false,
                underline: false,
                color: None,
            });
        }
        b"name" => {
            if let Some(name) = attr_value(e, b"val") {
                apply_font_attr(
                    &mut state.current_font,
                    &mut state.current_dxf_font,
                    |font| {
                        font.name = Some(name.clone());
                    },
                );
            }
        }
        b"sz" => {
            if let Some(size) = attr_f64(e, b"val") {
                apply_font_attr(
                    &mut state.current_font,
                    &mut state.current_dxf_font,
                    |font| {
                        font.size = Some(size);
                    },
                );
            }
        }
        b"b" => {
            apply_font_attr(
                &mut state.current_font,
                &mut state.current_dxf_font,
                |font| {
                    font.bold = true;
                },
            );
        }
        b"i" => {
            apply_font_attr(
                &mut state.current_font,
                &mut state.current_dxf_font,
                |font| {
                    font.italic = true;
                },
            );
        }
        b"u" => {
            apply_font_attr(
                &mut state.current_font,
                &mut state.current_dxf_font,
                |font| {
                    font.underline = true;
                },
            );
        }
        b"color" => {
            if let Some(font) = state.current_font.as_mut() {
                font.color = parse_color_attr(e);
            } else if let Some(font) = state.current_dxf_font.as_mut() {
                font.color = parse_color_attr(e);
            } else if let Some((_, side)) = state.current_border_side.as_mut() {
                side.color = parse_color_attr(e);
            } else if let Some((_, side)) = state.current_dxf_border_side.as_mut() {
                side.color = parse_color_attr(e);
            } else if let Some(fill) = state.current_fill.as_mut() {
                if fill.fg_color.is_none() {
                    fill.fg_color = parse_color_attr(e);
                }
            } else if let Some(fill) = state.current_dxf_fill.as_mut() {
                if fill.fg_color.is_none() {
                    fill.fg_color = parse_color_attr(e);
                }
            }
        }
        b"fills" => state.in_fills = true,
        b"fill" if state.in_fills => {
            state.current_fill = Some(FillDef {
                pattern_type: None,
                fg_color: None,
                bg_color: None,
            });
        }
        b"fill" if state.in_dxfs => {
            state.current_dxf_fill = Some(FillDef {
                pattern_type: None,
                fg_color: None,
                bg_color: None,
            });
        }
        b"patternFill" => {
            if let Some(pattern_type) = parse_pattern_type(e) {
                if let Some(fill) = state.current_fill.as_mut() {
                    fill.pattern_type = Some(pattern_type.clone());
                } else if let Some(fill) = state.current_dxf_fill.as_mut() {
                    fill.pattern_type = Some(pattern_type);
                }
            }
        }
        b"fgColor" => {
            if let Some(fill) = state.current_fill.as_mut() {
                fill.fg_color = parse_color_attr(e);
            } else if let Some(fill) = state.current_dxf_fill.as_mut() {
                fill.fg_color = parse_color_attr(e);
            }
        }
        b"bgColor" => {
            if let Some(fill) = state.current_fill.as_mut() {
                fill.bg_color = parse_color_attr(e);
            } else if let Some(fill) = state.current_dxf_fill.as_mut() {
                fill.bg_color = parse_color_attr(e);
            }
        }
        b"borders" => state.in_borders = true,
        b"border" if state.in_borders => {
            state.current_border = Some(BorderDef {
                left: None,
                right: None,
                top: None,
                bottom: None,
            });
        }
        b"border" if state.in_dxfs => {
            state.current_dxf_border = Some(BorderDef {
                left: None,
                right: None,
                top: None,
                bottom: None,
            });
        }
        b"left" | b"right" | b"top" | b"bottom" => {
            let side = parse_border_side(e);
            let side_name = String::from_utf8_lossy(e.name().as_ref()).to_string();
            if state.current_border.is_some() {
                state.current_border_side = Some((side_name, side));
            } else if state.current_dxf_border.is_some() {
                state.current_dxf_border_side = Some((side_name, side));
            }
        }
        b"cellXfs" => state.in_cell_xfs = true,
        b"cellStyleXfs" => state.in_cell_style_xfs = true,
        b"dxfs" => state.in_dxfs = true,
        b"dxf" if state.in_dxfs => {
            state.current_dxf = Some(DxfStyle::new());
        }
        b"xf" if state.in_cell_xfs => {
            state.current_xf = Some(parse_xf(e));
            state.current_xf_is_style = false;
        }
        b"xf" if state.in_cell_style_xfs => {
            state.current_xf = Some(parse_xf(e));
            state.current_xf_is_style = true;
        }
        b"alignment" => {
            if let Some(xf) = state.current_xf.as_mut() {
                xf.alignment = Some(parse_alignment(e));
            } else if let Some(dxf) = state.current_dxf.as_mut() {
                dxf.alignment = Some(parse_alignment(e));
            }
        }
        b"protection" => {
            let protection = parse_protection(e);
            if let Some(xf) = state.current_xf.as_mut() {
                xf.protection = Some(protection);
            } else if let Some(dxf) = state.current_dxf.as_mut() {
                dxf.protection = Some(protection);
            }
        }
        b"tableStyles" => {
            styles.table_styles = Some(parse_table_style_info(e));
            state.in_table_styles = true;
        }
        b"tableStyle" if state.in_table_styles => {
            if let Some(info) = styles.table_styles.as_mut() {
                if let Some(style) = parse_table_style_def(e) {
                    info.styles.push(style);
                }
            }
        }
        _ => {}
    }
    Ok(())
}

fn handle_styles_end(
    e: &BytesEnd<'_>,
    state: &mut StylesParseState,
    styles: &mut SpreadsheetStyles,
) {
    match e.name().as_ref() {
        b"numFmts" => state.in_num_fmts = false,
        b"fonts" => state.in_fonts = false,
        b"fills" => state.in_fills = false,
        b"borders" => state.in_borders = false,
        b"cellXfs" => state.in_cell_xfs = false,
        b"cellStyleXfs" => state.in_cell_style_xfs = false,
        b"dxfs" => state.in_dxfs = false,
        b"tableStyles" => state.in_table_styles = false,
        b"font" => {
            if let Some(font) = state.current_font.take() {
                styles.fonts.push(font);
            } else if let Some(font) = state.current_dxf_font.take() {
                if let Some(dxf) = state.current_dxf.as_mut() {
                    dxf.font = Some(font);
                }
            }
        }
        b"fill" => {
            if let Some(fill) = state.current_fill.take() {
                styles.fills.push(fill);
            } else if let Some(fill) = state.current_dxf_fill.take() {
                if let Some(dxf) = state.current_dxf.as_mut() {
                    dxf.fill = Some(fill);
                }
            }
        }
        b"border" => {
            if let Some(border) = state.current_border.take() {
                styles.borders.push(border);
            } else if let Some(border) = state.current_dxf_border.take() {
                if let Some(dxf) = state.current_dxf.as_mut() {
                    dxf.border = Some(border);
                }
            }
        }
        b"left" | b"right" | b"top" | b"bottom" => {
            if let (Some(border), Some((name, side))) = (
                state.current_border.as_mut(),
                state.current_border_side.take(),
            ) {
                assign_border_side(border, name.as_bytes(), side);
            } else if let (Some(border), Some((name, side))) = (
                state.current_dxf_border.as_mut(),
                state.current_dxf_border_side.take(),
            ) {
                assign_border_side(border, name.as_bytes(), side);
            }
        }
        b"xf" => {
            if let Some(xf) = state.current_xf.take() {
                if state.current_xf_is_style {
                    styles.cell_style_xfs.push(xf);
                } else {
                    styles.cell_xfs.push(xf);
                }
                state.current_xf_is_style = false;
            }
        }
        b"dxf" => {
            if let Some(dxf) = state.current_dxf.take() {
                styles.dxfs.push(dxf);
            }
        }
        _ => {}
    }
}

fn apply_font_attr<F>(
    current_font: &mut Option<FontDef>,
    current_dxf_font: &mut Option<FontDef>,
    mut apply: F,
) where
    F: FnMut(&mut FontDef),
{
    if let Some(font) = current_font.as_mut() {
        apply(font);
    } else if let Some(font) = current_dxf_font.as_mut() {
        apply(font);
    }
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

fn parse_number_format(element: &BytesStart) -> Option<NumberFormat> {
    let mut id = None;
    let mut code = None;
    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"numFmtId" => id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"formatCode" => code = Some(String::from_utf8_lossy(&attr.value).to_string()),
            _ => {}
        }
    }
    match (id, code) {
        (Some(id), Some(code)) => Some(NumberFormat {
            id,
            format_code: code,
        }),
        _ => None,
    }
}

fn parse_pattern_type(element: &BytesStart) -> Option<String> {
    for attr in element.attributes().flatten() {
        if attr.key.as_ref() == b"patternType" {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

fn parse_border_side(element: &BytesStart) -> BorderSide {
    let mut side = BorderSide {
        style: None,
        color: None,
    };
    for attr in element.attributes().flatten() {
        if attr.key.as_ref() == b"style" {
            side.style = Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    side
}

fn parse_xf(element: &BytesStart) -> CellFormat {
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
    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"numFmtId" => xf.num_fmt_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"fontId" => xf.font_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"fillId" => xf.fill_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"borderId" => xf.border_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"xfId" => xf.xf_id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
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
    xf
}

fn parse_alignment(element: &BytesStart) -> CellAlignment {
    let mut alignment = CellAlignment {
        horizontal: None,
        vertical: None,
        wrap_text: false,
        indent: None,
        text_rotation: None,
        shrink_to_fit: false,
        reading_order: None,
    };
    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"horizontal" => {
                alignment.horizontal = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"vertical" => {
                alignment.vertical = Some(String::from_utf8_lossy(&attr.value).to_string());
            }
            b"wrapText" => {
                let v = String::from_utf8_lossy(&attr.value);
                alignment.wrap_text = v == "1" || v.eq_ignore_ascii_case("true");
            }
            b"indent" => {
                alignment.indent = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
            }
            b"textRotation" => {
                alignment.text_rotation = String::from_utf8_lossy(&attr.value).parse::<i32>().ok();
            }
            b"shrinkToFit" => {
                let v = String::from_utf8_lossy(&attr.value);
                alignment.shrink_to_fit = v == "1" || v.eq_ignore_ascii_case("true");
            }
            b"readingOrder" => {
                alignment.reading_order = String::from_utf8_lossy(&attr.value).parse::<u32>().ok();
            }
            _ => {}
        }
    }
    alignment
}

fn parse_protection(element: &BytesStart) -> CellProtection {
    let mut protection = CellProtection {
        locked: None,
        hidden: None,
    };
    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"locked" => {
                let v = String::from_utf8_lossy(&attr.value);
                protection.locked = Some(v == "1" || v.eq_ignore_ascii_case("true"));
            }
            b"hidden" => {
                let v = String::from_utf8_lossy(&attr.value);
                protection.hidden = Some(v == "1" || v.eq_ignore_ascii_case("true"));
            }
            _ => {}
        }
    }
    protection
}

fn parse_table_style_info(element: &BytesStart) -> TableStyleInfo {
    let mut info = TableStyleInfo {
        count: None,
        default_table_style: None,
        default_pivot_style: None,
        styles: Vec::new(),
    };
    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"count" => info.count = String::from_utf8_lossy(&attr.value).parse::<u32>().ok(),
            b"defaultTableStyle" => {
                info.default_table_style = Some(String::from_utf8_lossy(&attr.value).to_string())
            }
            b"defaultPivotStyle" => {
                info.default_pivot_style = Some(String::from_utf8_lossy(&attr.value).to_string())
            }
            _ => {}
        }
    }
    info
}

fn parse_table_style_def(element: &BytesStart) -> Option<TableStyleDef> {
    let mut name = None;
    let mut pivot = None;
    let mut table = None;
    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"name" => name = Some(String::from_utf8_lossy(&attr.value).to_string()),
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
    name.map(|name| TableStyleDef { name, pivot, table })
}

pub(crate) fn parse_color_attr(element: &BytesStart) -> Option<String> {
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
