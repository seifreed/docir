use crate::error::ParseError;
use crate::xml_utils::{lossy_attr_value, xml_error};
use docir_core::ir::{ShapeText, ShapeTextParagraph, ShapeTextRun, TextAlignment};
use quick_xml::events::Event;
use quick_xml::Reader;

pub(super) fn parse_text_body(
    reader: &mut Reader<&[u8]>,
    slide_path: &str,
) -> Result<ShapeText, ParseError> {
    parse_text_body_with_end(reader, slide_path, b"p:txBody")
}

pub(super) fn parse_text_body_table(
    reader: &mut Reader<&[u8]>,
    slide_path: &str,
) -> Result<ShapeText, ParseError> {
    parse_text_body_with_end(reader, slide_path, b"a:txBody")
}

fn parse_text_body_with_end(
    reader: &mut Reader<&[u8]>,
    slide_path: &str,
    end_tag: &[u8],
) -> Result<ShapeText, ParseError> {
    let mut paragraphs = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"a:p" {
                    let paragraph = parse_text_paragraph(reader, slide_path)?;
                    paragraphs.push(paragraph);
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == end_tag {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(slide_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(ShapeText { paragraphs })
}

pub(super) fn shape_text_to_plain(text: &ShapeText) -> String {
    let mut out = String::new();
    for (p_idx, para) in text.paragraphs.iter().enumerate() {
        if p_idx > 0 {
            out.push('\n');
        }
        for run in &para.runs {
            out.push_str(&run.text);
        }
    }
    out
}

fn parse_text_paragraph(
    reader: &mut Reader<&[u8]>,
    slide_path: &str,
) -> Result<ShapeTextParagraph, ParseError> {
    let mut runs = Vec::new();
    let mut alignment: Option<TextAlignment> = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"a:pPr" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"algn" {
                            alignment = map_alignment(&lossy_attr_value(&attr));
                        }
                    }
                }
                b"a:r" => {
                    let run = parse_text_run(reader, slide_path)?;
                    runs.push(run);
                }
                b"a:br" => {
                    runs.push(ShapeTextRun {
                        text: "\n".to_string(),
                        bold: None,
                        italic: None,
                        font_size: None,
                        font_family: None,
                    });
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:p" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(slide_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(ShapeTextParagraph { runs, alignment })
}

fn parse_text_run(
    reader: &mut Reader<&[u8]>,
    slide_path: &str,
) -> Result<ShapeTextRun, ParseError> {
    let mut text = String::new();
    let mut bold = None;
    let mut italic = None;
    let mut font_size = None;
    let mut font_family = None;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"a:rPr" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"b" => bold = Some(attr.value.as_ref() == b"1"),
                            b"i" => italic = Some(attr.value.as_ref() == b"1"),
                            b"sz" => font_size = lossy_attr_value(&attr).parse::<u32>().ok(),
                            _ => {}
                        }
                    }
                }
                b"a:t" => {
                    let value = reader
                        .read_text(e.name())
                        .map_err(|e| xml_error(slide_path, e))?;
                    text.push_str(&value);
                }
                b"a:latin" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"typeface" {
                            font_family = Some(lossy_attr_value(&attr).to_string());
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:r" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(slide_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(ShapeTextRun {
        text,
        bold,
        italic,
        font_size,
        font_family,
    })
}

fn map_alignment(value: &str) -> Option<TextAlignment> {
    match value {
        "l" => Some(TextAlignment::Left),
        "r" => Some(TextAlignment::Right),
        "ctr" => Some(TextAlignment::Center),
        "just" => Some(TextAlignment::Justify),
        "dist" => Some(TextAlignment::Distribute),
        _ => None,
    }
}
