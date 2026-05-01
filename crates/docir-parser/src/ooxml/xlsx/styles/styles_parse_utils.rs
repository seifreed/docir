use super::{
    attr_u32_from_bytes, attr_value, BorderDef, BorderSide, CellAlignment, CellFormat,
    CellProtection, FillDef, FontDef, NumberFormat, TableStyleDef, TableStyleInfo,
};
use quick_xml::events::BytesStart;

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

pub(super) fn parse_number_format(element: &BytesStart) -> Option<NumberFormat> {
    let id = attr_u32_from_bytes(element, b"numFmtId");
    let code = attr_value(element, b"formatCode");
    match (id, code) {
        (Some(id), Some(code)) => Some(NumberFormat {
            id,
            format_code: code,
        }),
        _ => None,
    }
}

pub(super) fn parse_pattern_type(element: &BytesStart) -> Option<String> {
    attr_value(element, b"patternType")
}

pub(super) fn parse_border_side(element: &BytesStart) -> BorderSide {
    let mut side = BorderSide {
        style: None,
        color: None,
    };
    side.style = attr_value(element, b"style");
    side
}

pub(super) fn parse_xf(element: &BytesStart) -> CellFormat {
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
    set_opt_u32_attr(&mut xf.num_fmt_id, element, b"numFmtId");
    set_opt_u32_attr(&mut xf.font_id, element, b"fontId");
    set_opt_u32_attr(&mut xf.fill_id, element, b"fillId");
    set_opt_u32_attr(&mut xf.border_id, element, b"borderId");
    set_opt_u32_attr(&mut xf.xf_id, element, b"xfId");

    set_bool_attr(&mut xf.apply_number_format, element, b"applyNumberFormat");
    set_bool_attr(&mut xf.apply_font, element, b"applyFont");
    set_bool_attr(&mut xf.apply_fill, element, b"applyFill");
    set_bool_attr(&mut xf.apply_border, element, b"applyBorder");
    set_bool_attr(&mut xf.apply_alignment, element, b"applyAlignment");
    set_bool_attr(&mut xf.apply_protection, element, b"applyProtection");
    set_bool_attr(&mut xf.quote_prefix, element, b"quotePrefix");
    set_bool_attr(&mut xf.pivot_button, element, b"pivotButton");

    xf
}

fn set_opt_u32_attr(target: &mut Option<u32>, element: &BytesStart, name: &[u8]) {
    if let Some(value) = attr_u32_from_bytes(element, name) {
        *target = Some(value);
    }
}

fn set_bool_attr(target: &mut bool, element: &BytesStart, name: &[u8]) {
    if let Some(value) = parse_bool_attr(element, name) {
        *target = value;
    }
}

pub(super) fn parse_alignment(element: &BytesStart) -> CellAlignment {
    let mut alignment = CellAlignment {
        horizontal: None,
        vertical: None,
        wrap_text: false,
        indent: None,
        text_rotation: None,
        shrink_to_fit: false,
        reading_order: None,
    };
    if let Some(horizontal) = attr_value(element, b"horizontal") {
        alignment.horizontal = Some(horizontal);
    }
    if let Some(vertical) = attr_value(element, b"vertical") {
        alignment.vertical = Some(vertical);
    }
    alignment.wrap_text = parse_bool_attr(element, b"wrapText").unwrap_or(false);
    alignment.indent = attr_value(element, b"indent").and_then(|v| v.parse::<u32>().ok());
    alignment.text_rotation =
        attr_value(element, b"textRotation").and_then(|v| v.parse::<i32>().ok());
    alignment.shrink_to_fit = parse_bool_attr(element, b"shrinkToFit").unwrap_or(false);
    alignment.reading_order =
        attr_value(element, b"readingOrder").and_then(|v| v.parse::<u32>().ok());
    alignment
}

fn parse_bool_attr(element: &BytesStart, name: &[u8]) -> Option<bool> {
    attr_value(element, name).map(|value| value == "1" || value.eq_ignore_ascii_case("true"))
}

pub(super) fn parse_protection(element: &BytesStart) -> CellProtection {
    let mut protection = CellProtection {
        locked: None,
        hidden: None,
    };
    if let Some(locked) = parse_bool_attr(element, b"locked") {
        protection.locked = Some(locked);
    }
    if let Some(hidden) = parse_bool_attr(element, b"hidden") {
        protection.hidden = Some(hidden);
    }
    protection
}

pub(super) fn parse_table_style_info(element: &BytesStart) -> TableStyleInfo {
    let mut info = TableStyleInfo {
        count: None,
        default_table_style: None,
        default_pivot_style: None,
        styles: Vec::new(),
    };
    info.count = attr_value(element, b"count").and_then(|v| v.parse::<u32>().ok());
    info.default_table_style = attr_value(element, b"defaultTableStyle");
    info.default_pivot_style = attr_value(element, b"defaultPivotStyle");
    info
}

pub(super) fn parse_table_style_def(element: &BytesStart) -> Option<TableStyleDef> {
    let name = attr_value(element, b"name")?;
    let pivot = parse_bool_attr(element, b"pivot");
    let table = parse_bool_attr(element, b"table");
    Some(TableStyleDef { name, pivot, table })
}

pub(crate) fn parse_color_attr(element: &BytesStart) -> Option<String> {
    let rgb = attr_value(element, b"rgb");
    let theme = attr_value(element, b"theme");
    let indexed = attr_value(element, b"indexed");
    if let Some(rgb) = rgb {
        Some(format!("rgb:{rgb}"))
    } else if let Some(theme) = theme {
        Some(format!("theme:{theme}"))
    } else {
        indexed.map(|indexed| format!("indexed:{indexed}"))
    }
}
