//! ODF presentation parsing helpers.

use super::{
    parse_draw_page, parse_odp_transition, IRNode, IrStore, OdfContentResult, OdfLimitCounter,
    ParseError, Slide,
};
use crate::xml_utils::{attr_value_by_suffix, local_name, xml_error};
use quick_xml::events::Event;
use quick_xml::Reader;

pub(super) fn parse_content_presentation(
    xml: &[u8],
    store: &mut IrStore,
    _limits: &dyn OdfLimitCounter,
) -> Result<OdfContentResult, ParseError> {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_presentation = false;
    let mut slide_no = 1u32;
    let mut slides = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
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
            Ok(Event::Empty(e)) => match local_name(e.name().as_ref()) {
                b"page" if in_presentation => {
                    let mut slide = Slide::new(slide_no);
                    slide.name = attr_value_by_suffix(&e, &[b":name"]);
                    slide.transition = parse_odp_transition(&e);
                    let slide_id = slide.id;
                    store.insert(IRNode::Slide(slide));
                    slides.push(slide_id);
                    slide_no += 1;
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"presentation" {
                    in_presentation = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error("content.xml", e)),
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
}
