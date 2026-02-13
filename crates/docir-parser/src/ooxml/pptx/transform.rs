use super::*;

pub(super) fn parse_transform(
    reader: &mut Reader<&[u8]>,
    transform: &mut ShapeTransform,
    slide_path: &str,
) -> Result<(), ParseError> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"a:off" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"x" => {
                                transform.x = String::from_utf8_lossy(&attr.value)
                                    .parse::<i64>()
                                    .unwrap_or(0)
                            }
                            b"y" => {
                                transform.y = String::from_utf8_lossy(&attr.value)
                                    .parse::<i64>()
                                    .unwrap_or(0)
                            }
                            _ => {}
                        }
                    }
                }
                b"a:ext" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"cx" => {
                                transform.width = String::from_utf8_lossy(&attr.value)
                                    .parse::<u64>()
                                    .unwrap_or(0)
                            }
                            b"cy" => {
                                transform.height = String::from_utf8_lossy(&attr.value)
                                    .parse::<u64>()
                                    .unwrap_or(0)
                            }
                            _ => {}
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"a:off" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"x" => {
                                transform.x = String::from_utf8_lossy(&attr.value)
                                    .parse::<i64>()
                                    .unwrap_or(0)
                            }
                            b"y" => {
                                transform.y = String::from_utf8_lossy(&attr.value)
                                    .parse::<i64>()
                                    .unwrap_or(0)
                            }
                            _ => {}
                        }
                    }
                }
                b"a:ext" => {
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"cx" => {
                                transform.width = String::from_utf8_lossy(&attr.value)
                                    .parse::<u64>()
                                    .unwrap_or(0)
                            }
                            b"cy" => {
                                transform.height = String::from_utf8_lossy(&attr.value)
                                    .parse::<u64>()
                                    .unwrap_or(0)
                            }
                            _ => {}
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:xfrm" || e.name().as_ref() == b"p:xfrm" {
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
