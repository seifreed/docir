use super::{
    BorderDef, BorderSide, CellAlignment, CellFormat, CellProtection, FillDef, FontDef,
    NumberFormat, TableStyleDef, TableStyleInfo,
};
use crate::error::ParseError;
use crate::xml_utils::{attr_u32_from_bytes, try_attr_value, xml_error};
use quick_xml::events::BytesStart;
use std::str::FromStr;

pub(super) fn apply_font_attr<F>(
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

pub(super) fn new_font() -> FontDef {
    FontDef {
        name: None,
        size: None,
        bold: false,
        italic: false,
        underline: false,
        color: None,
    }
}

pub(super) fn new_fill() -> FillDef {
    FillDef {
        pattern_type: None,
        fg_color: None,
        bg_color: None,
    }
}

pub(super) fn new_border() -> BorderDef {
    BorderDef {
        left: None,
        right: None,
        top: None,
        bottom: None,
    }
}

pub(super) fn assign_border_side(border: &mut BorderDef, name: &[u8], side: BorderSide) {
    match name {
        b"left" => border.left = Some(side),
        b"right" => border.right = Some(side),
        b"top" => border.top = Some(side),
        b"bottom" => border.bottom = Some(side),
        _ => {}
    }
}

pub(super) fn parse_number_format(
    element: &BytesStart,
    styles_path: &str,
) -> Result<Option<NumberFormat>, ParseError> {
    let id = attr_u32_from_bytes(element, b"numFmtId", styles_path)?;
    let code = try_attr_value(element, b"formatCode", styles_path)?;
    Ok(match (id, code) {
        (Some(id), Some(code)) => Some(NumberFormat {
            id,
            format_code: code,
        }),
        _ => None,
    })
}

pub(super) fn parse_pattern_type(
    element: &BytesStart,
    styles_path: &str,
) -> Result<Option<String>, ParseError> {
    try_attr_value(element, b"patternType", styles_path)
}

pub(super) fn parse_border_side(
    element: &BytesStart,
    styles_path: &str,
) -> Result<BorderSide, ParseError> {
    let mut side = BorderSide {
        style: None,
        color: None,
    };
    side.style = try_attr_value(element, b"style", styles_path)?;
    Ok(side)
}

pub(super) fn parse_xf(element: &BytesStart, styles_path: &str) -> Result<CellFormat, ParseError> {
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
    set_opt_u32_attr(&mut xf.num_fmt_id, element, b"numFmtId", styles_path)?;
    set_opt_u32_attr(&mut xf.font_id, element, b"fontId", styles_path)?;
    set_opt_u32_attr(&mut xf.fill_id, element, b"fillId", styles_path)?;
    set_opt_u32_attr(&mut xf.border_id, element, b"borderId", styles_path)?;
    set_opt_u32_attr(&mut xf.xf_id, element, b"xfId", styles_path)?;

    set_bool_attr(
        &mut xf.apply_number_format,
        element,
        b"applyNumberFormat",
        styles_path,
    )?;
    set_bool_attr(&mut xf.apply_font, element, b"applyFont", styles_path)?;
    set_bool_attr(&mut xf.apply_fill, element, b"applyFill", styles_path)?;
    set_bool_attr(&mut xf.apply_border, element, b"applyBorder", styles_path)?;
    set_bool_attr(
        &mut xf.apply_alignment,
        element,
        b"applyAlignment",
        styles_path,
    )?;
    set_bool_attr(
        &mut xf.apply_protection,
        element,
        b"applyProtection",
        styles_path,
    )?;
    set_bool_attr(&mut xf.quote_prefix, element, b"quotePrefix", styles_path)?;
    set_bool_attr(&mut xf.pivot_button, element, b"pivotButton", styles_path)?;

    Ok(xf)
}

fn set_opt_u32_attr(
    target: &mut Option<u32>,
    element: &BytesStart,
    name: &[u8],
    styles_path: &str,
) -> Result<(), ParseError> {
    if let Some(value) = attr_u32_from_bytes(element, name, styles_path)? {
        *target = Some(value);
    }
    Ok(())
}

fn set_bool_attr(
    target: &mut bool,
    element: &BytesStart,
    name: &[u8],
    styles_path: &str,
) -> Result<(), ParseError> {
    if let Some(value) = parse_bool_attr(element, name, styles_path)? {
        *target = value;
    }
    Ok(())
}

pub(super) fn parse_alignment(
    element: &BytesStart,
    styles_path: &str,
) -> Result<CellAlignment, ParseError> {
    let mut alignment = CellAlignment {
        horizontal: None,
        vertical: None,
        wrap_text: false,
        indent: None,
        text_rotation: None,
        shrink_to_fit: false,
        reading_order: None,
    };
    if let Some(horizontal) = try_attr_value(element, b"horizontal", styles_path)? {
        alignment.horizontal = Some(horizontal);
    }
    if let Some(vertical) = try_attr_value(element, b"vertical", styles_path)? {
        alignment.vertical = Some(vertical);
    }
    alignment.wrap_text = parse_bool_attr(element, b"wrapText", styles_path)?.unwrap_or(false);
    alignment.indent = parse_optional_attr(element, b"indent", styles_path)?;
    alignment.text_rotation = parse_optional_attr(element, b"textRotation", styles_path)?;
    alignment.shrink_to_fit =
        parse_bool_attr(element, b"shrinkToFit", styles_path)?.unwrap_or(false);
    alignment.reading_order = parse_optional_attr(element, b"readingOrder", styles_path)?;
    Ok(alignment)
}

fn parse_optional_attr<T: FromStr>(
    element: &BytesStart,
    name: &[u8],
    styles_path: &str,
) -> Result<Option<T>, ParseError>
where
    T::Err: std::fmt::Display,
{
    try_attr_value(element, name, styles_path)?
        .map(|value| {
            value
                .parse::<T>()
                .map_err(|err| xml_error(styles_path, err))
        })
        .transpose()
}

fn parse_bool_attr(
    element: &BytesStart,
    name: &[u8],
    styles_path: &str,
) -> Result<Option<bool>, ParseError> {
    let Some(value) = try_attr_value(element, name, styles_path)? else {
        return Ok(None);
    };
    match value.as_str() {
        "1" => Ok(Some(true)),
        "0" => Ok(Some(false)),
        value if value.eq_ignore_ascii_case("true") => Ok(Some(true)),
        value if value.eq_ignore_ascii_case("false") => Ok(Some(false)),
        _ => Err(xml_error(
            styles_path,
            format!("Invalid boolean value '{value}'"),
        )),
    }
}

pub(super) fn parse_protection(
    element: &BytesStart,
    styles_path: &str,
) -> Result<CellProtection, ParseError> {
    let mut protection = CellProtection {
        locked: None,
        hidden: None,
    };
    if let Some(locked) = parse_bool_attr(element, b"locked", styles_path)? {
        protection.locked = Some(locked);
    }
    if let Some(hidden) = parse_bool_attr(element, b"hidden", styles_path)? {
        protection.hidden = Some(hidden);
    }
    Ok(protection)
}

pub(super) fn parse_table_style_info(
    element: &BytesStart,
    styles_path: &str,
) -> Result<TableStyleInfo, ParseError> {
    let mut info = TableStyleInfo {
        count: None,
        default_table_style: None,
        default_pivot_style: None,
        styles: Vec::new(),
    };
    info.count = parse_optional_attr(element, b"count", styles_path)?;
    info.default_table_style = try_attr_value(element, b"defaultTableStyle", styles_path)?;
    info.default_pivot_style = try_attr_value(element, b"defaultPivotStyle", styles_path)?;
    Ok(info)
}

pub(super) fn parse_table_style_def(
    element: &BytesStart,
    styles_path: &str,
) -> Result<Option<TableStyleDef>, ParseError> {
    let Some(name) = try_attr_value(element, b"name", styles_path)? else {
        return Ok(None);
    };
    let pivot = parse_bool_attr(element, b"pivot", styles_path)?;
    let table = parse_bool_attr(element, b"table", styles_path)?;
    Ok(Some(TableStyleDef { name, pivot, table }))
}

pub(crate) fn parse_color_attr(
    element: &BytesStart,
    styles_path: &str,
) -> Result<Option<String>, ParseError> {
    let rgb = try_attr_value(element, b"rgb", styles_path)?;
    let theme = try_attr_value(element, b"theme", styles_path)?;
    let indexed = try_attr_value(element, b"indexed", styles_path)?;
    Ok(if let Some(rgb) = rgb {
        Some(format!("rgb:{rgb}"))
    } else if let Some(theme) = theme {
        Some(format!("theme:{theme}"))
    } else {
        indexed.map(|indexed| format!("indexed:{indexed}"))
    })
}
