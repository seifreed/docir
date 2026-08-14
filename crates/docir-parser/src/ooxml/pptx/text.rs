use crate::error::ParseError;
use crate::xml_utils::{decoded_text, local_name, lossy_attr_value, parse_bool_attr, xml_error};
use docir_core::ir::{ShapeText, ShapeTextParagraph, ShapeTextRun, TextAlignment};
use quick_xml::Reader;
use quick_xml::events::Event;

pub(super) fn parse_text_body(
    reader: &mut Reader<&[u8]>,
    slide_path: &str,
) -> Result<ShapeText, ParseError> {
    parse_text_body_with_end(reader, slide_path, b"txBody")
}

pub(super) fn parse_text_body_table(
    reader: &mut Reader<&[u8]>,
    slide_path: &str,
) -> Result<ShapeText, ParseError> {
    parse_text_body_with_end(reader, slide_path, b"txBody")
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
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"p" => {
                let paragraph = parse_text_paragraph(reader, slide_path)?;
                paragraphs.push(paragraph);
            }
            Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"p" => {
                paragraphs.push(ShapeTextParagraph {
                    runs: Vec::new(),
                    alignment: None,
                });
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == end_tag => {
                break;
            }
            Ok(Event::Eof) => {
                return Err(xml_error(slide_path, "unexpected EOF in text body XML"));
            }
            Err(e) => {
                return Err(xml_error(slide_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(ShapeText { paragraphs })
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
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"pPr" => {
                    for attr in e.attributes() {
                        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
                        if attr.key.as_ref() == b"algn" {
                            alignment = map_alignment(&lossy_attr_value(&attr));
                        }
                    }
                }
                b"r" => {
                    let run = parse_text_run(reader, slide_path)?;
                    runs.push(run);
                }
                b"br" => {
                    runs.push(line_break_run());
                }
                _ => {}
            },
            Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"br" => {
                runs.push(line_break_run());
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"p" => {
                break;
            }
            Ok(Event::Eof) => {
                return Err(xml_error(
                    slide_path,
                    "unexpected EOF in text paragraph XML",
                ));
            }
            Err(e) => {
                return Err(xml_error(slide_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(ShapeTextParagraph { runs, alignment })
}

fn line_break_run() -> ShapeTextRun {
    ShapeTextRun {
        text: "\n".to_string(),
        bold: None,
        italic: None,
        font_size: None,
        font_family: None,
    }
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
            Ok(Event::Start(e)) => {
                if local_name(e.name().as_ref()) == b"t" {
                    let value = reader
                        .read_text(e.name())
                        .map_err(|e| xml_error(slide_path, e))?;
                    text.push_str(&decoded_text(&value).map_err(|err| xml_error(slide_path, err))?);
                } else {
                    apply_run_formatting(
                        &e,
                        slide_path,
                        &mut bold,
                        &mut italic,
                        &mut font_size,
                        &mut font_family,
                    )?;
                }
            }
            Ok(Event::Empty(e)) => apply_run_formatting(
                &e,
                slide_path,
                &mut bold,
                &mut italic,
                &mut font_size,
                &mut font_family,
            )?,
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"r" => {
                break;
            }
            Ok(Event::Eof) => {
                return Err(xml_error(slide_path, "unexpected EOF in text run XML"));
            }
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

fn apply_run_formatting(
    element: &quick_xml::events::BytesStart<'_>,
    slide_path: &str,
    bold: &mut Option<bool>,
    italic: &mut Option<bool>,
    font_size: &mut Option<u32>,
    font_family: &mut Option<String>,
) -> Result<(), ParseError> {
    match local_name(element.name().as_ref()) {
        b"rPr" => {
            for attr in element.attributes() {
                let attr = attr.map_err(|err| xml_error(slide_path, err))?;
                match attr.key.as_ref() {
                    b"b" => *bold = Some(parse_bool_attr(&attr.value, slide_path)?),
                    b"i" => *italic = Some(parse_bool_attr(&attr.value, slide_path)?),
                    b"sz" => {
                        *font_size = Some(
                            lossy_attr_value(&attr)
                                .parse::<u32>()
                                .map_err(|err| xml_error(slide_path, err))?,
                        );
                    }
                    _ => {}
                }
            }
        }
        b"latin" => {
            for attr in element.attributes() {
                let attr = attr.map_err(|err| xml_error(slide_path, err))?;
                if attr.key.as_ref() == b"typeface" {
                    *font_family = Some(lossy_attr_value(&attr).to_string());
                }
            }
        }
        _ => {}
    }
    Ok(())
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

#[cfg(test)]
mod tests {
    use super::{parse_text_body, parse_text_paragraph, parse_text_run};
    use crate::error::ParseError;
    use quick_xml::Reader;
    use quick_xml::events::Event;

    fn reader_after_start<'a>(xml: &'a str, expected_start: &[u8]) -> Reader<&'a [u8]> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        match reader.read_event_into(&mut buf).expect("start event") {
            Event::Start(e) => {
                assert_eq!(e.local_name().as_ref(), expected_start);
            }
            other => panic!("expected start event, got {other:?}"),
        }
        reader
    }

    fn assert_xml_error(result: Result<impl std::fmt::Debug, ParseError>) {
        match result.expect_err("malformed text XML must fail") {
            ParseError::Xml { file, .. } => assert_eq!(file, "ppt/slides/broken-text.xml"),
            other => panic!("expected XML error, got {other:?}"),
        }
    }

    #[test]
    fn parse_text_body_reports_truncated_xml() {
        let xml = r#"
            <p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:p><a:r><a:t>broken</a:t></a:r></a:p>
        "#;
        let mut reader = reader_after_start(xml, b"txBody");

        assert_xml_error(parse_text_body(&mut reader, "ppt/slides/broken-text.xml"));
    }

    #[test]
    fn parse_text_body_preserves_empty_paragraph() {
        let xml = r#"<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:p/></a:txBody>"#;
        let mut reader = reader_after_start(xml, b"txBody");

        let text = parse_text_body(&mut reader, "ppt/slides/empty-text.xml")
            .expect("empty paragraph should parse");

        assert_eq!(text.paragraphs.len(), 1);
        assert!(text.paragraphs[0].runs.is_empty());
    }

    #[test]
    fn parse_text_paragraph_preserves_empty_line_break() {
        let xml = r#"
            <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:r><a:t>before</a:t></a:r>
              <a:br/>
              <a:r><a:t>after</a:t></a:r>
            </a:p>
        "#;
        let mut reader = reader_after_start(xml, b"p");

        let paragraph = parse_text_paragraph(&mut reader, "ppt/slides/line-break.xml")
            .expect("empty line break should parse");

        assert_eq!(paragraph.runs.len(), 3);
        assert_eq!(paragraph.runs[1].text, "\n");
    }

    #[test]
    fn parse_text_paragraph_reports_truncated_xml() {
        let xml = r#"
            <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:r><a:t>broken</a:t></a:r>
        "#;
        let mut reader = reader_after_start(xml, b"p");

        assert_xml_error(parse_text_paragraph(
            &mut reader,
            "ppt/slides/broken-text.xml",
        ));
    }

    #[test]
    fn parse_text_run_reports_truncated_xml() {
        let xml = r#"
            <a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:t>broken</a:t>
        "#;
        let mut reader = reader_after_start(xml, b"r");

        assert_xml_error(parse_text_run(&mut reader, "ppt/slides/broken-text.xml"));
    }

    #[test]
    fn parse_text_paragraph_reports_malformed_paragraph_attrs() {
        let xml = r#"
            <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:pPr algn="ctr" algn="l"></a:pPr>
              <a:r><a:t>text</a:t></a:r>
            </a:p>
        "#;
        let mut reader = reader_after_start(xml, b"p");

        assert_xml_error(parse_text_paragraph(
            &mut reader,
            "ppt/slides/broken-text.xml",
        ));
    }

    #[test]
    fn parse_text_run_reports_malformed_run_attrs() {
        let xml = r#"
            <a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:rPr b="1" b="0"></a:rPr>
              <a:t>text</a:t>
            </a:r>
        "#;
        let mut reader = reader_after_start(xml, b"r");

        assert_xml_error(parse_text_run(&mut reader, "ppt/slides/broken-text.xml"));

        let xml = r#"
            <a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:rPr sz="bad"></a:rPr>
              <a:t>text</a:t>
            </a:r>
        "#;
        let mut reader = reader_after_start(xml, b"r");

        assert_xml_error(parse_text_run(&mut reader, "ppt/slides/broken-text.xml"));
    }

    #[test]
    fn parse_text_run_preserves_empty_formatting_elements() {
        let xml = r#"
            <a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:rPr b="1" i="1" sz="1800"/>
              <a:latin typeface="Calibri"/>
              <a:t>text</a:t>
            </a:r>
        "#;
        let mut reader = reader_after_start(xml, b"r");

        let run = parse_text_run(&mut reader, "ppt/slides/empty-formatting.xml")
            .expect("empty formatting elements should parse");

        assert_eq!(run.text, "text");
        assert_eq!(run.bold, Some(true));
        assert_eq!(run.italic, Some(true));
        assert_eq!(run.font_size, Some(1800));
        assert_eq!(run.font_family.as_deref(), Some("Calibri"));
    }

    #[test]
    fn parse_text_run_accepts_boolean_lexical_values() {
        let xml = r#"
            <a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:rPr b="true" i="false"/>
              <a:t>text</a:t>
            </a:r>
        "#;
        let mut reader = reader_after_start(xml, b"r");

        let run = parse_text_run(&mut reader, "ppt/slides/boolean-formatting.xml")
            .expect("boolean formatting should parse");

        assert_eq!(run.bold, Some(true));
        assert_eq!(run.italic, Some(false));
    }

    #[test]
    fn parse_text_run_reports_malformed_latin_attrs() {
        let xml = r#"
            <a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:rPr><a:latin typeface="Arial" typeface="Calibri"></a:latin></a:rPr>
              <a:t>text</a:t>
            </a:r>
        "#;
        let mut reader = reader_after_start(xml, b"r");

        assert_xml_error(parse_text_run(&mut reader, "ppt/slides/broken-text.xml"));
    }
}
