use docir_parser::adapters::{HwpAdapter, HwpxAdapter, OdfAdapter, OoxmlAdapter, RtfAdapter};
use docir_parser::format::FormatParser;
use docir_parser::{ParseError, ParserConfig};
use std::io::Cursor;

#[test]
fn parse_error_variants_and_conversions_are_exercised() {
    let variants = vec![
        ParseError::InvalidZip("bad zip".to_string()),
        ParseError::ResourceLimit("limit hit".to_string()),
        ParseError::Xml {
            file: "x.xml".to_string(),
            message: "broken".to_string(),
        },
        ParseError::MissingPart("word/document.xml".to_string()),
        ParseError::InvalidStructure("invalid structure".to_string()),
        ParseError::InvalidFormat("invalid format".to_string()),
        ParseError::UnsupportedFormat("unknown format".to_string()),
        ParseError::ContentTypeMismatch {
            expected: "application/xml".to_string(),
            actual: "text/plain".to_string(),
        },
        ParseError::RelationshipNotFound("rId1".to_string()),
        ParseError::Encoding("utf8".to_string()),
        ParseError::PathTraversal("../evil".to_string()),
    ];

    for err in variants {
        let rendered = format!("{err}");
        assert!(!rendered.is_empty());
    }

    let bad_zip = zip::ZipArchive::new(Cursor::new(b"not-a-zip".to_vec())).expect_err("zip err");
    let parse_err: ParseError = bad_zip.into();
    assert!(matches!(parse_err, ParseError::InvalidZip(_)));

    let mut xml_reader = quick_xml::Reader::from_str("<root><child/></root>");
    let mut buf = Vec::new();
    let xml_err = xml_reader
        .read_to_end_into(quick_xml::name::QName(b"missing"), &mut buf)
        .expect_err("expected quick-xml mismatch error");
    let parse_xml_err: ParseError = xml_err.into();
    match parse_xml_err {
        ParseError::Xml { file, message } => {
            assert!(file.is_empty());
            assert!(!message.is_empty());
        }
        other => panic!("expected xml parse error, got {other:?}"),
    }
}

#[test]
fn adapters_parse_reader_paths_return_errors_for_invalid_payload() {
    let config = ParserConfig::default();

    let ooxml = OoxmlAdapter::new(config.clone());
    let odf = OdfAdapter::new(config.clone());
    let hwpx = HwpxAdapter::new(config.clone());
    let hwp = HwpAdapter::new(config.clone());
    let rtf = RtfAdapter::new(config);

    for result in [
        ooxml.parse_reader(Cursor::new(b"invalid".to_vec())),
        odf.parse_reader(Cursor::new(b"invalid".to_vec())),
        hwpx.parse_reader(Cursor::new(b"invalid".to_vec())),
        hwp.parse_reader(Cursor::new(b"invalid".to_vec())),
        rtf.parse_reader(Cursor::new(b"invalid".to_vec())),
    ] {
        assert!(result.is_err());
    }
}
