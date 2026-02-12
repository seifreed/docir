//! XLSX styles parsing.

use crate::error::ParseError;
use crate::ooxml::xml_utils::xml_error;
use docir_core::ir::{
    BorderDef, BorderSide, CellAlignment, CellFormat, CellProtection, DxfStyle, FillDef, FontDef,
    NumberFormat, SpreadsheetStyles, TableStyleDef, TableStyleInfo,
};
use docir_core::types::SourceSpan;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

pub(crate) fn parse_styles(xml: &str, styles_path: &str) -> Result<SpreadsheetStyles, ParseError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    reader.config_mut().expand_empty_elements = true;

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
                b"numFmt" if in_dxfs => {
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
                        if let Some(dxf) = current_dxf.as_mut() {
                            dxf.num_fmt = Some(NumberFormat {
                                id,
                                format_code: code,
                            });
                        }
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
                return Err(xml_error(styles_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(styles)
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
