#[cfg(test)]
use super::ShapeType;
use super::{ParseError, Reader, Shape, map_shape_type, parse_transform};
use crate::xml_utils::{local_name, lossy_attr_value, xml_error};
use quick_xml::events::{BytesStart, Event};

pub(super) fn parse_shape_properties(
    reader: &mut Reader<&[u8]>,
    shape: &mut Shape,
    slide_path: &str,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()) {
                b"prstGeom" => {
                    apply_preset_geometry(&e, shape, slide_path)?;
                }
                b"xfrm" => {
                    parse_transform(reader, &mut shape.transform, slide_path)?;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) if local_name(e.name().as_ref()) == b"prstGeom" => {
                apply_preset_geometry(&e, shape, slide_path)?;
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"spPr" => {
                break;
            }
            Ok(Event::Eof) => {
                return Err(xml_error(
                    slide_path,
                    "unexpected EOF in shape properties XML",
                ));
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

fn apply_preset_geometry(
    start: &BytesStart<'_>,
    shape: &mut Shape,
    slide_path: &str,
) -> Result<(), ParseError> {
    for attr in start.attributes() {
        let attr = attr.map_err(|err| xml_error(slide_path, err))?;
        if local_name(attr.key.as_ref()) == b"prst" {
            shape.shape_type = map_shape_type(&lossy_attr_value(&attr));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_shape_properties_reads_geometry_and_transform() {
        let xml = r#"<p:spPr>
            <a:xfrm>
                <a:off x="10" y="20"/>
                <a:ext cx="30" cy="40"/>
            </a:xfrm>
            <a:prstGeom prst="ellipse">
                <a:avLst/>
            </a:prstGeom>
        </p:spPr>"#;
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut shape = Shape::new(ShapeType::Unknown);

        let parsed = parse_shape_properties(&mut reader, &mut shape, "slide.xml");
        assert!(parsed.is_ok());
        assert_eq!(shape.shape_type, ShapeType::Ellipse);
        assert_eq!(shape.transform.x, 10);
        assert_eq!(shape.transform.y, 20);
        assert_eq!(shape.transform.width, 30);
        assert_eq!(shape.transform.height, 40);
    }

    #[test]
    fn parse_shape_properties_returns_xml_error_for_malformed_xml() {
        let xml = "<p:spPr><a:xfrm><a:off x='1' y='2'></p:spPr";
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut shape = Shape::new(ShapeType::Unknown);

        let err = parse_shape_properties(&mut reader, &mut shape, "bad-slide.xml")
            .expect_err("expected malformed xml error");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "bad-slide.xml"),
            other => panic!("unexpected error variant: {other:?}"),
        }
    }

    #[test]
    fn parse_shape_properties_reports_truncated_xml() {
        let xml = r#"<p:spPr><a:prstGeom prst="ellipse">"#;
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut shape = Shape::new(ShapeType::Unknown);

        let err = parse_shape_properties(&mut reader, &mut shape, "truncated-slide.xml")
            .expect_err("truncated shape properties must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "truncated-slide.xml"),
            other => panic!("unexpected error variant: {other:?}"),
        }
    }

    #[test]
    fn parse_shape_properties_reports_malformed_geometry_attributes() {
        let xml = r#"<p:spPr>
            <a:prstGeom prst="ellipse" prst="rect">
                <a:avLst/>
            </a:prstGeom>
        </p:spPr>"#;
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut shape = Shape::new(ShapeType::Unknown);

        let err = parse_shape_properties(&mut reader, &mut shape, "bad-slide.xml")
            .expect_err("malformed shape geometry attributes must fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "bad-slide.xml"),
            other => panic!("unexpected error variant: {other:?}"),
        }
    }
}
