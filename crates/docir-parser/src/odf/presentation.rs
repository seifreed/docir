//! ODF presentation parsing helpers.

use super::{
    IRNode, IrStore, OdfContentResult, OdfLimitCounter, ParseError, Slide, parse_draw_page,
    parse_odp_transition,
};
use crate::xml_utils::{local_name, try_attr_value_by_suffix, xml_error};
use quick_xml::Reader;
use quick_xml::events::Event;

pub(super) fn parse_content_presentation(
    xml: &[u8],
    store: &mut IrStore,
    _limits: &dyn OdfLimitCounter,
) -> Result<OdfContentResult, ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut in_presentation = false;
    let mut slide_no = 1u32;
    let mut slides = Vec::new();
    let mut root_name: Option<Vec<u8>> = None;
    let mut root_closed = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error("content.xml", err))?;
        if root_closed && matches!(&event, Event::Start(_) | Event::Empty(_)) {
            return Err(xml_error(
                "content.xml",
                "XML document contains multiple roots",
            ));
        }
        match &event {
            Event::Start(e) if root_name.is_none() => {
                root_name = Some(e.name().as_ref().to_vec());
            }
            Event::Empty(e) if root_name.is_none() => {
                root_name = Some(e.name().as_ref().to_vec());
                root_closed = true;
            }
            Event::End(e)
                if root_name
                    .as_deref()
                    .is_some_and(|name| name == e.name().as_ref()) =>
            {
                root_closed = true;
            }
            Event::Eof if root_closed => break,
            Event::Eof => {
                return Err(xml_error(
                    "content.xml",
                    "XML document ends before its root is closed",
                ));
            }
            _ => {}
        }
        match event {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"presentation" => in_presentation = true,
                b"page" if in_presentation => {
                    let slide = parse_draw_page(&mut reader, &e, slide_no, store)?;
                    let slide_id = slide.id;
                    store.insert(IRNode::Slide(slide));
                    slides.push(slide_id);
                    slide_no += 1;
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"page" if in_presentation => {
                    let mut slide = Slide::new(slide_no);
                    slide.name = try_attr_value_by_suffix(&e, &[b":name"], "content.xml")?;
                    slide.transition = parse_odp_transition(&e)?;
                    let slide_id = slide.id;
                    store.insert(IRNode::Slide(slide));
                    slides.push(slide_id);
                    slide_no += 1;
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"presentation" => {
                in_presentation = false;
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(OdfContentResult {
        content: slides,
        ..OdfContentResult::default()
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::odf::{OdfLimits, ShapeType};
    use crate::parser::ParserConfig;
    use docir_core::visitor::IrStore;

    #[test]
    fn parse_content_presentation_accepts_alternate_namespace_prefixes() {
        let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<doc:document-content xmlns:doc="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <o:presentation>
    <d:page d:name="AltSlide">
      <d:frame d:name="Title">
        <d:text-box><t:p>Hello</t:p></d:text-box>
      </d:frame>
    </d:page>
  </o:presentation>
</doc:document-content>
"#;
        let mut store = IrStore::new();
        let limits = OdfLimits::new(&ParserConfig::default(), false);

        let result = parse_content_presentation(xml, &mut store, &limits)
            .expect("presentation should parse");

        assert_eq!(result.content.len(), 1);
        let Some(IRNode::Slide(slide)) = store.get(result.content[0]) else {
            panic!("expected slide");
        };
        assert_eq!(slide.name.as_deref(), Some("AltSlide"));
        assert_eq!(slide.shapes.len(), 1);

        let Some(IRNode::Shape(shape)) = store.get(slide.shapes[0]) else {
            panic!("expected shape");
        };
        assert_eq!(shape.name.as_deref(), Some("Title"));
        assert_eq!(shape.shape_type, ShapeType::TextBox);
    }

    #[test]
    fn parse_content_presentation_reports_malformed_empty_page_attributes() {
        let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<doc:document-content xmlns:doc="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">
  <o:presentation>
    <d:page d:name="Slide 1" d:name="Slide 2"/>
  </o:presentation>
</doc:document-content>
"#;
        let mut store = IrStore::new();
        let limits = OdfLimits::new(&ParserConfig::default(), false);

        let err = parse_content_presentation(xml, &mut store, &limits)
            .expect_err("malformed page attributes must fail");

        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_content_presentation_rejects_truncated_document_root() {
        let xml: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation>"#;
        let mut store = IrStore::new();
        let limits = OdfLimits::new(&ParserConfig::default(), false);

        let err = parse_content_presentation(xml, &mut store, &limits)
            .expect_err("truncated presentation document must fail");
        assert!(matches!(err, ParseError::Xml { file, .. } if file == "content.xml"));
    }

    #[test]
    fn parse_content_presentation_rejects_mismatched_nested_end_tag() {
        let xml: &[u8] =
            br#"<office:document-content><office:body></office:bodyx></office:document-content>"#;
        let mut store = IrStore::new();
        let limits = OdfLimits::new(&ParserConfig::default(), false);

        let err = parse_content_presentation(xml, &mut store, &limits)
            .expect_err("mismatched XML must fail");
        assert!(matches!(err, ParseError::Xml { file, .. } if file == "content.xml"));
    }
}
