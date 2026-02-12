use crate::error::ParseError;
use crate::xml_utils::xml_error;
use docir_core::ir::DigitalSignature;
use docir_core::types::SourceSpan;
use quick_xml::events::Event;
use quick_xml::Reader;

pub fn parse_signature(xml: &str, path: &str) -> Result<DigitalSignature, ParseError> {
    let mut sig = DigitalSignature::new();
    sig.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"ds:Signature" | b"Signature" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"Id" {
                            sig.signature_id =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"ds:SignatureMethod" | b"SignatureMethod" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"Algorithm" {
                            sig.signature_method =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"ds:DigestMethod" | b"DigestMethod" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"Algorithm" {
                            sig.digest_methods
                                .push(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"ds:X509SubjectName" | b"X509SubjectName" => {
                    if let Ok(text) = reader.read_text(e.name()) {
                        sig.signer = Some(text.to_string());
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(xml_error(path, e));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(sig)
}
