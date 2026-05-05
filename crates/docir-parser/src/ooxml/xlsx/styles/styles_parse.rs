//! XLSX styles parsing.

use crate::error::ParseError;
use crate::xml_utils::{
    attr_f64, attr_u32_from_bytes, attr_value, local_name, reader_from_str_with_options,
};
use crate::xml_utils::{scan_xml_events, XmlScanControl};
use docir_core::ir::{
    BorderDef, BorderSide, CellAlignment, CellFormat, CellProtection, DxfStyle, FillDef, FontDef,
    NumberFormat, SpreadsheetStyles, TableStyleDef, TableStyleInfo,
};
use docir_core::types::SourceSpan;
use quick_xml::events::BytesEnd;
use quick_xml::events::{BytesStart, Event};
#[path = "styles_parse_utils.rs"]
mod styles_parse_utils;
use styles_parse_utils::*;

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

    scan_xml_events(&mut reader, &mut buf, styles_path, |event| {
        match event {
            Event::Start(e) => handle_styles_start(&e, &mut state, &mut styles)?,
            Event::Empty(e) => handle_styles_start(&e, &mut state, &mut styles)?,
            Event::End(e) => handle_styles_end(&e, &mut state, &mut styles),
            _ => {}
        }
        Ok(XmlScanControl::Continue)
    })?;

    if let Some(dxf) = state.current_dxf.take() {
        styles.dxfs.push(dxf);
    }
    Ok(styles)
}

fn handle_styles_start(
    e: &BytesStart<'_>,
    state: &mut StylesParseState,
    styles: &mut SpreadsheetStyles,
) -> Result<(), ParseError> {
    let handled = handle_num_fmt_start(e, state, styles)
        || handle_font_start(e, state)
        || handle_fill_start(e, state)
        || handle_border_start(e, state)
        || handle_xf_start(e, state)
        || handle_table_style_start(e, state, styles);
    #[cfg(test)]
    eprintln!(
        "handle_styles_start event={} handled={handled}",
        String::from_utf8_lossy(e.name().as_ref()),
    );
    if handled {
        return Ok(());
    }
    let raw_name = e.name();
    let name = local_name(raw_name.as_ref());
    if name == b"numFmts" {
        state.in_num_fmts = true;
        #[cfg(test)]
        eprintln!("handle_styles_start entered numFmts");
    } else if name == b"fonts" {
        state.in_fonts = true;
        #[cfg(test)]
        eprintln!("handle_styles_start entered fonts");
    } else if name == b"fills" {
        state.in_fills = true;
        #[cfg(test)]
        eprintln!("handle_styles_start entered fills");
    } else if name == b"borders" {
        state.in_borders = true;
        #[cfg(test)]
        eprintln!("handle_styles_start entered borders");
    } else if name == b"cellXfs" {
        state.in_cell_xfs = true;
        #[cfg(test)]
        eprintln!("handle_styles_start entered cellXfs");
    } else if name == b"cellStyleXfs" {
        state.in_cell_style_xfs = true;
        #[cfg(test)]
        eprintln!("handle_styles_start entered cellStyleXfs");
    } else if name == b"dxfs" {
        state.in_dxfs = true;
        #[cfg(test)]
        eprintln!("handle_styles_start entered dxfs");
    }
    Ok(())
}

fn handle_num_fmt_start(
    e: &BytesStart<'_>,
    state: &mut StylesParseState,
    styles: &mut SpreadsheetStyles,
) -> bool {
    if local_name(e.name().as_ref()) != b"numFmt" {
        return false;
    }
    #[cfg(test)]
    eprintln!(
        "handle_num_fmt_start pre num_fmts={} in_dxfs={}",
        state.in_num_fmts, state.in_dxfs
    );
    if state.in_num_fmts {
        if let Some(fmt) = parse_number_format(e) {
            #[cfg(test)]
            eprintln!("handle_num_fmt_start push num_fmt id={}", fmt.id);
            styles.number_formats.push(fmt);
        }
    } else if state.in_dxfs {
        if let Some(fmt) = parse_number_format(e) {
            if let Some(dxf) = state.current_dxf.as_mut() {
                dxf.num_fmt = Some(fmt);
            }
        }
    }
    true
}

fn handle_font_start(e: &BytesStart<'_>, state: &mut StylesParseState) -> bool {
    let raw_name = e.name();
    let name = local_name(raw_name.as_ref());
    if name == b"font" {
        if state.in_fonts {
            state.current_font = Some(new_font());
            return true;
        }
        if state.in_dxfs {
            state.current_dxf_font = Some(new_font());
            return true;
        }
        return false;
    }

    if !state.in_fonts && !state.in_dxfs {
        return false;
    }

    match name {
        b"name" | b"sz" | b"b" | b"i" | b"u" | b"color" => {
            apply_font_node_attrs(e, state);
            true
        }
        _ => false,
    }
}

fn apply_font_node_attrs(e: &BytesStart<'_>, state: &mut StylesParseState) {
    match local_name(e.name().as_ref()) {
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
        b"b" => apply_font_attr(
            &mut state.current_font,
            &mut state.current_dxf_font,
            |font| font.bold = true,
        ),
        b"i" => apply_font_attr(
            &mut state.current_font,
            &mut state.current_dxf_font,
            |font| font.italic = true,
        ),
        b"u" => apply_font_attr(
            &mut state.current_font,
            &mut state.current_dxf_font,
            |font| font.underline = true,
        ),
        b"color" => apply_color_attr(e, state),
        _ => {}
    }
}

fn apply_color_attr(e: &BytesStart<'_>, state: &mut StylesParseState) {
    if let Some(color) = parse_color_attr(e) {
        if let Some(font) = state.current_font.as_mut() {
            font.color = Some(color.clone());
            return;
        }
        if let Some(font) = state.current_dxf_font.as_mut() {
            font.color = Some(color.clone());
            return;
        }
        if let Some((_, side)) = state.current_border_side.as_mut() {
            side.color = Some(color.clone());
            return;
        }
        if let Some((_, side)) = state.current_dxf_border_side.as_mut() {
            side.color = Some(color.clone());
            return;
        }
        if let Some(fill) = state.current_fill.as_mut() {
            if fill.fg_color.is_none() {
                fill.fg_color = Some(color.clone());
                return;
            }
        }
        if let Some(fill) = state.current_dxf_fill.as_mut() {
            if fill.fg_color.is_none() {
                fill.fg_color = Some(color);
            }
        }
    }
}

fn handle_fill_start(e: &BytesStart<'_>, state: &mut StylesParseState) -> bool {
    match local_name(e.name().as_ref()) {
        b"fill" if state.in_fills => {
            state.current_fill = Some(new_fill());
            true
        }
        b"fill" if state.in_dxfs => {
            state.current_dxf_fill = Some(new_fill());
            true
        }
        b"patternFill" => {
            if let Some(pattern_type) = parse_pattern_type(e) {
                if let Some(fill) = state.current_fill.as_mut() {
                    fill.pattern_type = Some(pattern_type.clone());
                } else if let Some(fill) = state.current_dxf_fill.as_mut() {
                    fill.pattern_type = Some(pattern_type);
                }
            }
            true
        }
        b"fgColor" => {
            if let Some(fill) = state.current_fill.as_mut() {
                fill.fg_color = parse_color_attr(e);
            } else if let Some(fill) = state.current_dxf_fill.as_mut() {
                fill.fg_color = parse_color_attr(e);
            }
            true
        }
        b"bgColor" => {
            if let Some(fill) = state.current_fill.as_mut() {
                fill.bg_color = parse_color_attr(e);
            } else if let Some(fill) = state.current_dxf_fill.as_mut() {
                fill.bg_color = parse_color_attr(e);
            }
            true
        }
        _ => false,
    }
}

fn handle_border_start(e: &BytesStart<'_>, state: &mut StylesParseState) -> bool {
    match local_name(e.name().as_ref()) {
        b"border" if state.in_borders => {
            state.current_border = Some(new_border());
            true
        }
        b"border" if state.in_dxfs => {
            state.current_dxf_border = Some(new_border());
            true
        }
        b"left" | b"right" | b"top" | b"bottom" => {
            let side = parse_border_side(e);
            let side_name = String::from_utf8_lossy(local_name(e.name().as_ref())).to_string();
            if state.current_border.is_some() {
                state.current_border_side = Some((side_name, side));
            } else if state.current_dxf_border.is_some() {
                state.current_dxf_border_side = Some((side_name, side));
            }
            true
        }
        _ => false,
    }
}

fn handle_xf_start(e: &BytesStart<'_>, state: &mut StylesParseState) -> bool {
    match local_name(e.name().as_ref()) {
        b"dxf" if state.in_dxfs => {
            #[cfg(test)]
            eprintln!("handle_xf_start entering dxf");
            state.current_dxf = Some(DxfStyle::new());
            true
        }
        b"xf" if state.in_cell_xfs => {
            state.current_xf = Some(parse_xf(e));
            state.current_xf_is_style = false;
            true
        }
        b"xf" if state.in_cell_style_xfs => {
            state.current_xf = Some(parse_xf(e));
            state.current_xf_is_style = true;
            true
        }
        b"alignment" => {
            if let Some(xf) = state.current_xf.as_mut() {
                xf.alignment = Some(parse_alignment(e));
            } else if let Some(dxf) = state.current_dxf.as_mut() {
                dxf.alignment = Some(parse_alignment(e));
            }
            true
        }
        b"protection" => {
            let protection = parse_protection(e);
            if let Some(xf) = state.current_xf.as_mut() {
                xf.protection = Some(protection);
            } else if let Some(dxf) = state.current_dxf.as_mut() {
                dxf.protection = Some(protection);
            }
            true
        }
        _ => false,
    }
}

fn handle_table_style_start(
    e: &BytesStart<'_>,
    state: &mut StylesParseState,
    styles: &mut SpreadsheetStyles,
) -> bool {
    match local_name(e.name().as_ref()) {
        b"tableStyles" => {
            styles.table_styles = Some(parse_table_style_info(e));
            state.in_table_styles = true;
            true
        }
        b"tableStyle" if state.in_table_styles => {
            if let Some(info) = styles.table_styles.as_mut() {
                if let Some(style) = parse_table_style_def(e) {
                    info.styles.push(style);
                }
            }
            true
        }
        _ => false,
    }
}

fn handle_styles_end(
    e: &BytesEnd<'_>,
    state: &mut StylesParseState,
    styles: &mut SpreadsheetStyles,
) {
    match local_name(e.name().as_ref()) {
        b"numFmts" => state.in_num_fmts = false,
        b"fonts" => state.in_fonts = false,
        b"fills" => state.in_fills = false,
        b"borders" => state.in_borders = false,
        b"cellXfs" => state.in_cell_xfs = false,
        b"cellStyleXfs" => state.in_cell_style_xfs = false,
        // Keep as compatibility with readers that might omit explicit </dxf> events.
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
            #[cfg(test)]
            eprintln!("handle_styles_end dxf");
            if let Some(dxf) = state.current_dxf.take() {
                styles.dxfs.push(dxf);
            }
        }
        b"dxfs" => {
            #[cfg(test)]
            eprintln!(
                "handle_styles_end dxfs before_push in_dxfs={} has_dxf={}",
                state.in_dxfs,
                state.current_dxf.is_some()
            );
            if let Some(dxf) = state.current_dxf.take() {
                styles.dxfs.push(dxf);
            }
            state.in_dxfs = false;
        }
        _ => {}
    }
}

pub(crate) use styles_parse_utils::parse_color_attr;
