use crate::error::ParseError;
use crate::xml_utils::local_name;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{read_event, reader_from_str};
use docir_core::ir::{Theme, ThemeColor, ThemeFontScheme};
use docir_core::types::SourceSpan;
use quick_xml::events::{BytesStart, Event};

#[derive(Default)]
struct ThemeParseState {
    in_clr_scheme: bool,
    current_color_name: Option<String>,
    font_scheme: ThemeFontScheme,
    in_major_font: bool,
    in_minor_font: bool,
}

/// Public API entrypoint: parse_theme.
pub fn parse_theme(xml: &str, path: &str) -> Result<Theme, ParseError> {
    let mut theme = Theme::new();
    theme.span = Some(SourceSpan::new(path));

    let mut reader = reader_from_str(xml);
    let mut buf = Vec::new();
    let mut state = ThemeParseState::default();

    loop {
        match read_event(&mut reader, &mut buf, path)? {
            Event::Start(e) => handle_start_event(&e, &mut theme, &mut state),
            Event::Empty(e) => handle_empty_event(&e, &mut theme, &mut state),
            Event::End(e) => handle_end_event(e.name().as_ref(), &mut state),
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    theme.fonts = state.font_scheme;
    Ok(theme)
}

fn handle_start_event(start: &BytesStart<'_>, theme: &mut Theme, state: &mut ThemeParseState) {
    match local_name(start.name().as_ref()) {
        b"theme" => set_theme_name(start, &mut theme.name),
        b"clrScheme" => state.in_clr_scheme = true,
        b"fontScheme" => {
            if theme.name.is_none() {
                set_theme_name(start, &mut theme.name);
            }
        }
        b"majorFont" => state.in_major_font = true,
        b"minorFont" => state.in_minor_font = true,
        b"latin" => set_latin_typeface(start, state),
        _ if state.in_clr_scheme => {
            let name = String::from_utf8_lossy(start.name().as_ref()).to_string();
            state.current_color_name = Some(name);
        }
        _ => {}
    }
}

fn handle_empty_event(start: &BytesStart<'_>, theme: &mut Theme, state: &mut ThemeParseState) {
    if state.in_clr_scheme {
        let color_value = srgb_value(start);
        if let Some(name) = state.current_color_name.take() {
            if color_value.is_some() {
                theme.colors.push(ThemeColor {
                    name,
                    value: color_value,
                });
            }
        }
    }

    if local_name(start.name().as_ref()) == b"latin" {
        set_latin_typeface(start, state);
    }
}

fn handle_end_event(tag: &[u8], state: &mut ThemeParseState) {
    match local_name(tag) {
        b"clrScheme" => state.in_clr_scheme = false,
        b"majorFont" => state.in_major_font = false,
        b"minorFont" => state.in_minor_font = false,
        _ => state.current_color_name = None,
    }
}

fn set_theme_name(start: &BytesStart<'_>, slot: &mut Option<String>) {
    for attr in start.attributes().flatten() {
        if attr.key.as_ref() == b"name" {
            *slot = Some(lossy_attr_value(&attr).to_string());
            break;
        }
    }
}

fn set_latin_typeface(start: &BytesStart<'_>, state: &mut ThemeParseState) {
    let mut typeface = None;
    for attr in start.attributes().flatten() {
        if attr.key.as_ref() == b"typeface" {
            typeface = Some(lossy_attr_value(&attr).to_string());
            break;
        }
    }
    if let Some(tf) = typeface {
        if state.in_major_font {
            state.font_scheme.major = Some(tf);
        } else if state.in_minor_font {
            state.font_scheme.minor = Some(tf);
        }
    }
}

fn srgb_value(start: &BytesStart<'_>) -> Option<String> {
    if local_name(start.name().as_ref()) != b"srgbClr" {
        return None;
    }
    for attr in start.attributes().flatten() {
        if attr.key.as_ref() == b"val" {
            return Some(lossy_attr_value(&attr).to_string());
        }
    }
    None
}
