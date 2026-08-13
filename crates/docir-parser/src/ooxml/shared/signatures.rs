use crate::error::ParseError;
use crate::xml_utils::decoded_text;
use crate::xml_utils::local_name;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::track_xml_document_event;
use crate::xml_utils::visit_attributes;
use crate::xml_utils::xml_error;
use docir_core::ir::DigitalSignature;
use docir_core::types::SourceSpan;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Public API entrypoint: parse_signature.
pub fn parse_signature(xml: &str, path: &str) -> Result<DigitalSignature, ParseError> {
    parse_signature_impl(xml, path)
}

fn parse_signature_impl(xml: &str, path: &str) -> Result<DigitalSignature, ParseError> {
    let mut sig = DigitalSignature::new();
    sig.span = Some(SourceSpan::new(path));

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|err| xml_error(path, err))?;
        if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
            break;
        }
        match event {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"Signature" => {
                    visit_attributes(&e, path, |attr| {
                        if attr.key.as_ref() == b"Id" {
                            sig.signature_id = Some(lossy_attr_value(attr).to_string());
                        }
                    })?;
                }
                b"SignatureMethod" => {
                    visit_attributes(&e, path, |attr| {
                        if attr.key.as_ref() == b"Algorithm" {
                            sig.signature_method = Some(lossy_attr_value(attr).to_string());
                        }
                    })?;
                }
                b"DigestMethod" => {
                    visit_attributes(&e, path, |attr| {
                        if attr.key.as_ref() == b"Algorithm" {
                            sig.digest_methods.push(lossy_attr_value(attr).to_string());
                        }
                    })?;
                }
                b"X509SubjectName" => {
                    let text = reader.read_text(e.name()).map_err(|e| xml_error(path, e))?;
                    depth -= 1;
                    sig.signer = Some(decoded_text(&text).map_err(|err| xml_error(path, err))?);
                }
                _ => {}
            },
            _ => {}
        }
        buf.clear();
    }

    Ok(sig)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_signature_extracts_identity_and_algorithms() {
        let xml = r##"
            <ds:Signature Id="sig-1">
              <ds:SignedInfo>
                <ds:SignatureMethod Algorithm="rsa-sha256"></ds:SignatureMethod>
                <ds:Reference URI="#id">
                  <ds:DigestMethod Algorithm="sha256"></ds:DigestMethod>
                </ds:Reference>
                <ds:Reference URI="#id2">
                  <ds:DigestMethod Algorithm="sha1"></ds:DigestMethod>
                </ds:Reference>
              </ds:SignedInfo>
              <ds:KeyInfo>
                <ds:X509Data>
                  <ds:X509SubjectName>CN=example signer</ds:X509SubjectName>
                </ds:X509Data>
              </ds:KeyInfo>
            </ds:Signature>
        "##;

        let parsed = parse_signature(xml, "ppt/signatures/sig1.xml").expect("signature");
        assert_eq!(parsed.signature_id.as_deref(), Some("sig-1"));
        assert_eq!(parsed.signature_method.as_deref(), Some("rsa-sha256"));
        assert_eq!(parsed.signer.as_deref(), Some("CN=example signer"));
        assert_eq!(parsed.digest_methods, vec!["sha256", "sha1"]);
        assert_eq!(
            parsed.span.as_ref().map(|span| span.file_path.as_str()),
            Some("ppt/signatures/sig1.xml")
        );
    }

    #[test]
    fn parse_signature_handles_plain_tag_names() {
        let xml = r#"
            <Signature Id="sig-plain">
              <SignatureMethod Algorithm="ecdsa-sha256"></SignatureMethod>
              <DigestMethod Algorithm="sha384"></DigestMethod>
              <X509SubjectName>CN=plain signer</X509SubjectName>
            </Signature>
        "#;

        let parsed = parse_signature(xml, "word/signatures/sig.xml").expect("signature");
        assert_eq!(parsed.signature_id.as_deref(), Some("sig-plain"));
        assert_eq!(parsed.signature_method.as_deref(), Some("ecdsa-sha256"));
        assert_eq!(parsed.signer.as_deref(), Some("CN=plain signer"));
        assert_eq!(parsed.digest_methods, vec!["sha384"]);
    }

    #[test]
    fn parse_signature_rejects_incomplete_xml() {
        let xml = "<ds:Signature><ds:SignatureMethod Algorithm=\"sha256\">";
        let err = parse_signature(xml, "bad.xml").expect_err("truncated signature must fail");
        assert!(matches!(err, ParseError::Xml { file, .. } if file == "bad.xml"));
    }

    #[test]
    fn parse_signature_reports_malformed_attributes() {
        let err = parse_signature(
            r#"<ds:Signature Id="sig-1" Id="sig-2"></ds:Signature>"#,
            "word/signatures/sig.xml",
        )
        .expect_err("malformed signature attrs should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "word/signatures/sig.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn parse_signature_reports_malformed_signer_text() {
        let err = parse_signature(
            r#"<Signature><X509SubjectName>CN=broken</Signature>"#,
            "word/signatures/sig.xml",
        )
        .expect_err("malformed signer text should fail");
        match err {
            ParseError::Xml { file, .. } => assert_eq!(file, "word/signatures/sig.xml"),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
