use super::*;

pub(super) fn parse_shape_properties(
    reader: &mut Reader<&[u8]>,
    shape: &mut Shape,
    slide_path: &str,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"a:prstGeom" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"prst" {
                            shape.shape_type =
                                map_shape_type(&String::from_utf8_lossy(&attr.value));
                        }
                    }
                }
                b"a:xfrm" => {
                    parse_transform(reader, &mut shape.transform, slide_path)?;
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"p:spPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ParseError::Xml {
                    file: slide_path.to_string(),
                    message: e.to_string(),
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}
