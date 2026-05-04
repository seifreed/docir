//! ODF presentation parsing helpers.

use super::{
    attr_value, parse_draw_page, parse_odp_transition, IRNode, IrStore, OdfContentResult,
    OdfLimitCounter, ParseError, Slide,
};
use crate::xml_utils::xml_error;
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
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"office:presentation" => in_presentation = true,
                b"draw:page" if in_presentation => {
                    let slide = parse_draw_page(&mut reader, &e, slide_no, store)?;
                    let slide_id = slide.id;
                    store.insert(IRNode::Slide(slide));
                    slides.push(slide_id);
                    slide_no += 1;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"draw:page" if in_presentation => {
                    let mut slide = Slide::new(slide_no);
                    slide.name = attr_value(&e, b"draw:name");
                    slide.transition = parse_odp_transition(&e);
                    let slide_id = slide.id;
                    store.insert(IRNode::Slide(slide));
                    slides.push(slide_id);
                    slide_no += 1;
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"office:presentation" {
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
