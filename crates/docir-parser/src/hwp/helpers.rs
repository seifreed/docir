use super::part_registry;
use crate::diagnostics::{push_info, push_warning};
use crate::text_utils::parse_text_alignment;
use docir_core::ir::{
    Diagnostics, MediaType, RunProperties, StyleParagraphProperties, StyleRunProperties,
    TableAlignment, TableProperties, TableWidth, TableWidthType,
};
use docir_core::types::{DocumentFormat, SourceSpan};
use quick_xml::events::BytesStart;

pub(super) fn is_hwpx_section(path: &str) -> bool {
    path.starts_with("Contents/section") && path.ends_with(".xml")
}

pub(super) fn is_hwpx_header(path: &str) -> bool {
    path.starts_with("Contents/header") && path.ends_with(".xml")
}

pub(super) fn is_hwpx_footer(path: &str) -> bool {
    path.starts_with("Contents/footer") && path.ends_with(".xml")
}

pub(super) fn is_hwpx_master(path: &str) -> bool {
    path.starts_with("Contents/masterPage") && path.ends_with(".xml")
}

pub(super) fn attr_any(e: &BytesStart, names: &[&[u8]]) -> Option<String> {
    for name in names {
        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == *name {
                if let Ok(value) = attr.unescape_value() {
                    return Some(value.to_string());
                }
            }
        }
    }
    None
}

pub(super) fn run_properties_from_attrs(e: &BytesStart) -> RunProperties {
    let mut props = RunProperties::default();
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let Ok(value) = attr.unescape_value() else {
            continue;
        };
        let value = value.to_string();
        match key {
            b"bold" | b"b" => {
                props.bold = Some(value == "1" || value.eq_ignore_ascii_case("true"));
            }
            b"italic" | b"i" => {
                props.italic = Some(value == "1" || value.eq_ignore_ascii_case("true"));
            }
            b"underline" | b"u" => {
                props.underline = Some(docir_core::ir::UnderlineStyle::Single);
            }
            b"color" => {
                props.color = Some(value.trim_start_matches('#').to_string());
            }
            b"highlight" => {
                props.highlight = Some(value.trim_start_matches('#').to_string());
            }
            b"font" | b"fontName" => {
                props.font_family = Some(value);
            }
            b"size" | b"fontSize" => {
                if let Ok(size) = value.parse::<u32>() {
                    props.font_size = Some(size);
                }
            }
            _ => {}
        }
    }
    props
}

pub(super) fn style_run_props_from_run(run: RunProperties) -> StyleRunProperties {
    StyleRunProperties {
        font_family: run.font_family,
        font_size: run.font_size,
        bold: run.bold,
        italic: run.italic,
        underline: run.underline,
        strike: run.strike,
        color: run.color,
        highlight: run.highlight,
        vertical_align: run.vertical_align,
        all_caps: run.all_caps,
        small_caps: run.small_caps,
    }
}

pub(super) fn parse_hwpx_paragraph_props(e: &BytesStart) -> StyleParagraphProperties {
    let mut props = StyleParagraphProperties::default();
    if let Some(align) = attr_any(e, &[b"align", b"alignment", b"textAlign"]) {
        props.alignment = parse_text_alignment(&align);
    }
    let mut indent = docir_core::ir::Indentation::default();
    let mut has_indent = false;
    if let Some(value) = attr_any(e, &[b"indentLeft", b"indent-left", b"left"]) {
        if let Ok(left) = value.parse::<i32>() {
            indent.left = Some(left);
            has_indent = true;
        }
    }
    if let Some(value) = attr_any(e, &[b"indentRight", b"indent-right", b"right"]) {
        if let Ok(right) = value.parse::<i32>() {
            indent.right = Some(right);
            has_indent = true;
        }
    }
    if let Some(value) = attr_any(e, &[b"firstIndent", b"first-indent", b"first"]) {
        if let Ok(first) = value.parse::<i32>() {
            indent.first_line = Some(first);
            has_indent = true;
        }
    }
    if has_indent {
        props.indentation = Some(indent);
    }
    props
}

pub(super) fn parse_hwpx_table_props(e: &BytesStart) -> Option<TableProperties> {
    let mut props = TableProperties::default();
    let mut has_value = false;
    if let Some(width) = attr_any(e, &[b"width", b"w", b"tableWidth"]) {
        if let Ok(value) = width.parse::<u32>() {
            props.width = Some(TableWidth {
                value,
                width_type: TableWidthType::Dxa,
            });
            has_value = true;
        }
    }
    if let Some(align) = attr_any(e, &[b"align", b"alignment", b"tableAlign"]) {
        let align = align.to_ascii_lowercase();
        props.alignment = match align.as_str() {
            "left" => Some(TableAlignment::Left),
            "center" => Some(TableAlignment::Center),
            "right" => Some(TableAlignment::Right),
            _ => None,
        };
        if props.alignment.is_some() {
            has_value = true;
        }
    }
    if has_value {
        Some(props)
    } else {
        None
    }
}

pub(super) fn media_type_from_path(path: &str) -> MediaType {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".png")
        || lower.ends_with(".jpg")
        || lower.ends_with(".jpeg")
        || lower.ends_with(".gif")
        || lower.ends_with(".bmp")
    {
        MediaType::Image
    } else if lower.ends_with(".mp3") || lower.ends_with(".wav") || lower.ends_with(".aac") {
        MediaType::Audio
    } else if lower.ends_with(".mp4")
        || lower.ends_with(".avi")
        || lower.ends_with(".mov")
        || lower.ends_with(".wmv")
    {
        MediaType::Video
    } else {
        MediaType::Other
    }
}

pub(super) fn build_hwp_diagnostics(format: DocumentFormat, paths: &[String]) -> Diagnostics {
    let registry = part_registry::registry_for(format);
    let mut diagnostics = Diagnostics::new();
    diagnostics.span = Some(SourceSpan::new("package"));

    for path in paths {
        push_info(
            &mut diagnostics,
            "HWP_PART",
            format!("part: {}", path),
            Some(path),
        );
    }

    for spec in registry {
        let mut matched = false;
        for path in paths {
            if part_registry::matches_pattern(path, spec.pattern) {
                matched = true;
                break;
            }
        }
        if !matched {
            push_warning(
                &mut diagnostics,
                "COVERAGE_MISSING",
                format!(
                    "missing part for pattern {} (expected parser={})",
                    spec.pattern, spec.expected_parser
                ),
                Some(spec.pattern),
            );
        }
    }

    diagnostics
}
