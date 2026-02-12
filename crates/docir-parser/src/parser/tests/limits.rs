use super::super::*;

#[test]
fn test_document_parser_enforces_max_input_size() {
    let mut config = ParserConfig::default();
    config.max_input_size = 32;
    let parser = DocumentParser::with_config(config);
    let data = vec![b'A'; 128];
    let err = parser
        .parse_reader(std::io::Cursor::new(data))
        .expect_err("expected size limit error");
    assert!(matches!(err, ParseError::ResourceLimit(_)));
}
