use super::{ParseError, Reader, ShapeTransform};
use crate::xml_utils::{local_name, lossy_attr_value, xml_error};
use quick_xml::events::{BytesStart, Event};

pub(super) fn parse_transform(
    reader: &mut Reader<&[u8]>,
    transform: &mut ShapeTransform,
    slide_path: &str,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    let mut depth = 1usize;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                depth += 1;
                apply_transform_event(&e, transform);
            }
            Ok(Event::Empty(e)) => apply_transform_event(&e, transform),
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"xfrm" => {
                if depth == 1 {
                    break;
                }
                depth -= 1;
            }
            Ok(Event::End(_)) => {
                depth = depth.saturating_sub(1);
            }
            Ok(Event::Eof) => {
                return Err(xml_error(slide_path, "unexpected EOF in transform XML"));
            }
            Err(e) => {
                return Err(xml_error(slide_path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}

fn apply_transform_event(e: &BytesStart<'_>, transform: &mut ShapeTransform) {
    match local_name(e.name().as_ref()) {
        b"off" => apply_transform_offset(e, transform),
        b"ext" => apply_transform_extent(e, transform),
        _ => {}
    }
}

fn apply_transform_offset(e: &BytesStart<'_>, transform: &mut ShapeTransform) {
    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"x" => transform.x = lossy_attr_value(&attr).parse::<i64>().unwrap_or(0),
            b"y" => transform.y = lossy_attr_value(&attr).parse::<i64>().unwrap_or(0),
            _ => {}
        }
    }
}

fn apply_transform_extent(e: &BytesStart<'_>, transform: &mut ShapeTransform) {
    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"cx" => transform.width = lossy_attr_value(&attr).parse::<u64>().unwrap_or(0),
            b"cy" => transform.height = lossy_attr_value(&attr).parse::<u64>().unwrap_or(0),
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_transform_fixture(xml: &str, slide_path: &str) -> Result<ShapeTransform, ParseError> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"xfrm" => {
                    let mut transform = ShapeTransform::default();
                    parse_transform(&mut reader, &mut transform, slide_path)?;
                    return Ok(transform);
                }
                Ok(Event::Eof) => return Ok(ShapeTransform::default()),
                Err(e) => return Err(xml_error(slide_path, e)),
                _ => {}
            }
            buf.clear();
        }
    }

    #[test]
    fn parse_transform_reads_start_and_empty_offset_and_extent_nodes() {
        let xml = r#"<a:xfrm>
            <a:off x="120" y="-45"></a:off>
            <a:ext cx="3000" cy="4000"/>
        </a:xfrm>"#;
        let transform = parse_transform_fixture(xml, "slide1.xml").expect("parse transform");
        assert_eq!(transform.x, 120);
        assert_eq!(transform.y, -45);
        assert_eq!(transform.width, 3000);
        assert_eq!(transform.height, 4000);
    }

    #[test]
    fn parse_transform_defaults_to_zero_for_invalid_numbers() {
        let xml = r#"<p:xfrm>
            <a:off x="nan" y="bad"/>
            <a:ext cx="oops" cy="NaN"/>
        </p:xfrm>"#;
        let transform = parse_transform_fixture(xml, "slide2.xml").expect("parse transform");
        assert_eq!(transform.x, 0);
        assert_eq!(transform.y, 0);
        assert_eq!(transform.width, 0);
        assert_eq!(transform.height, 0);
    }

    #[test]
    fn parse_transform_reports_truncated_xml() {
        let xml = "<a:xfrm><a:off x='1' y='2'><a:ext cx='3' cy='4'/>";
        let err =
            parse_transform_fixture(xml, "broken-slide.xml").expect_err("truncated XML must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "broken-slide.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
