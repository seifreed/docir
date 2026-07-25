use super::{ParseError, Reader, ShapeTransform};
use crate::xml_utils::{local_name, lossy_attr_value, xml_error};
use quick_xml::events::attributes::Attribute;
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
                apply_transform_event(&e, transform, slide_path)?;
            }
            Ok(Event::Empty(e)) => apply_transform_event(&e, transform, slide_path)?,
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

fn apply_transform_event(
    e: &BytesStart<'_>,
    transform: &mut ShapeTransform,
    slide_path: &str,
) -> Result<(), ParseError> {
    match local_name(e.name().as_ref()) {
        b"off" => apply_transform_offset(e, transform, slide_path)?,
        b"ext" => apply_transform_extent(e, transform, slide_path)?,
        _ => {}
    }
    Ok(())
}

fn apply_transform_offset(
    e: &BytesStart<'_>,
    transform: &mut ShapeTransform,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in e.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        match attr.key.as_ref() {
            b"x" => transform.x = parse_i64_attr(&attr, slide_path)?,
            b"y" => transform.y = parse_i64_attr(&attr, slide_path)?,
            _ => {}
        }
    }
    Ok(())
}

fn apply_transform_extent(
    e: &BytesStart<'_>,
    transform: &mut ShapeTransform,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in e.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        match attr.key.as_ref() {
            b"cx" => transform.width = parse_u64_attr(&attr, slide_path)?,
            b"cy" => transform.height = parse_u64_attr(&attr, slide_path)?,
            _ => {}
        }
    }
    Ok(())
}

fn parse_i64_attr(attr: &Attribute<'_>, file: &str) -> Result<i64, ParseError> {
    lossy_attr_value(attr)
        .parse()
        .map_err(|err| xml_error(file, err))
}

fn parse_u64_attr(attr: &Attribute<'_>, file: &str) -> Result<u64, ParseError> {
    lossy_attr_value(attr)
        .parse()
        .map_err(|err| xml_error(file, err))
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
    fn parse_transform_rejects_invalid_numbers() {
        let xml = r#"<p:xfrm>
            <a:off x="nan" y="bad"/>
            <a:ext cx="oops" cy="NaN"/>
        </p:xfrm>"#;
        parse_transform_fixture(xml, "slide2.xml").expect_err("invalid transform must fail");
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

    #[test]
    fn parse_transform_reports_malformed_offset_attributes() {
        let xml = r#"<a:xfrm>
            <a:off x="120" x="121" y="-45"/>
            <a:ext cx="3000" cy="4000"/>
        </a:xfrm>"#;
        let err =
            parse_transform_fixture(xml, "broken-slide.xml").expect_err("malformed XML must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "broken-slide.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_transform_reports_malformed_extent_attributes() {
        let xml = r#"<a:xfrm>
            <a:off x="120" y="-45"/>
            <a:ext cx="3000" cx="4000"/>
        </a:xfrm>"#;
        let err =
            parse_transform_fixture(xml, "broken-slide.xml").expect_err("malformed XML must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "broken-slide.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
