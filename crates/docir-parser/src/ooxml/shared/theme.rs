use crate::error::ParseError;
use crate::xml_utils::{read_event, reader_from_str};
use docir_core::ir::{Theme, ThemeColor, ThemeFontScheme};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;

pub fn parse_theme(xml: &str, path: &str) -> Result<Theme, ParseError> {
    let mut theme = Theme::new();
    theme.span = Some(SourceSpan::new(path));

    let mut reader = reader_from_str(xml);

    let mut buf = Vec::new();
    let mut in_clr_scheme = false;
    let mut current_color_name: Option<String> = None;
    let mut font_scheme = ThemeFontScheme::default();
    let mut in_major_font = false;
    let mut in_minor_font = false;

    loop {
        match read_event(&mut reader, &mut buf, path)? {
            Event::Start(e) => match e.name().as_ref() {
                b"a:theme" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"name" {
                            theme.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"a:clrScheme" => {
                    in_clr_scheme = true;
                }
                b"a:fontScheme" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"name" {
                            if theme.name.is_none() {
                                theme.name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                b"a:majorFont" => {
                    in_major_font = true;
                }
                b"a:minorFont" => {
                    in_minor_font = true;
                }
                b"a:latin" => {
                    let mut typeface = None;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"typeface" {
                            typeface = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                    if let Some(tf) = typeface {
                        if in_major_font {
                            font_scheme.major = Some(tf);
                        } else if in_minor_font {
                            font_scheme.minor = Some(tf);
                        }
                    }
                }
                _ => {
                    if in_clr_scheme {
                        let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                        current_color_name = Some(name);
                    }
                }
            },
            Event::Empty(e) => {
                if in_clr_scheme {
                    let mut color_value: Option<String> = None;
                    if e.name().as_ref() == b"a:srgbClr" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                color_value =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    if let Some(name) = current_color_name.take() {
                        if color_value.is_some() {
                            theme.colors.push(ThemeColor {
                                name,
                                value: color_value,
                            });
                        }
                    }
                }

                if e.name().as_ref() == b"a:latin" {
                    let mut typeface = None;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"typeface" {
                            typeface = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                    if let Some(tf) = typeface {
                        if in_major_font {
                            font_scheme.major = Some(tf);
                        } else if in_minor_font {
                            font_scheme.minor = Some(tf);
                        }
                    }
                }
            }
            Event::End(e) => match e.name().as_ref() {
                b"a:clrScheme" => {
                    in_clr_scheme = false;
                }
                b"a:majorFont" => in_major_font = false,
                b"a:minorFont" => in_minor_font = false,
                _ => {
                    current_color_name = None;
                }
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    theme.fonts = font_scheme;
    Ok(theme)
}
