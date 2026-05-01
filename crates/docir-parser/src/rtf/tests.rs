use super::RtfParser;
use crate::error::ParseError;
use crate::format::FormatParser;
use crate::parser::ParserConfig;
use docir_core::ir::IRNode;
use docir_core::types::DocumentFormat;
use docir_core::types::NodeType;
use std::io::Cursor;

#[test]
fn parse_simple_rtf() {
    let data = b"{\\rtf1\\ansi Hello \\par World}";
    let parser = RtfParser::new();
    let parsed = parser.parse_bytes(data).expect("parse rtf");
    assert_eq!(parsed.format, DocumentFormat::Rtf);
}

#[test]
fn parse_hyperlink_field() {
    let data =
        b"{\\rtf1{\\field{\\fldinst HYPERLINK \\\"https://example.com\\\"}{\\fldrslt Link}}}";
    let parser = RtfParser::new();
    let parsed = parser.parse_bytes(data).expect("parse rtf");
    assert_eq!(parsed.format, DocumentFormat::Rtf);
}

#[test]
fn parse_styles_and_lists() {
    let data = b"{\\rtf1\\ansi{\\stylesheet{\\s1 Heading 1;}{\\cs2 Emphasis;}}\\pard\\ql\\s1\\ls1\\ilvl0 Item}";
    let parser = RtfParser::new();
    let parsed = parser.parse_bytes(data).expect("parse rtf");
    let doc = parsed.document().expect("doc");
    assert!(doc.styles.is_some());
    let has_style_set = parsed
        .store
        .iter_ids_by_type(NodeType::StyleSet)
        .next()
        .is_some();
    assert!(has_style_set);
    let has_numbering = parsed.store.values().any(|node| match node {
        IRNode::Paragraph(p) => p.properties.numbering.is_some(),
        _ => false,
    });
    assert!(has_numbering);
}

#[test]
fn parse_table_borders_and_widths() {
    let data = b"{\\rtf1\\ansi{\\colortbl;\\red255\\green0\\blue0;}\\trowd\\cellx1000\\cellx2000\\clbrdrt\\brdrs\\brdrw10\\clcbpat1\\cell One\\cell Two\\row}";
    let parser = RtfParser::new();
    let parsed = parser.parse_bytes(data).expect("parse rtf");
    let has_cell_props = parsed.store.values().any(|node| match node {
        IRNode::TableCell(cell) => {
            cell.properties.width.is_some()
                || cell.properties.borders.is_some()
                || cell.properties.shading.is_some()
        }
        _ => false,
    });
    assert!(has_cell_props);
}

#[test]
fn parse_paragraph_margins_and_borders() {
    let data = b"{\\rtf1\\ansi\\pard\\li720\\ri360\\fi180\\sb120\\sa240\\sl360\\slmult1\\brdrt\\brdrs\\brdrw15 Paragraph}";
    let parser = RtfParser::new();
    let parsed = parser.parse_bytes(data).expect("parse rtf");
    let has_props = parsed.store.values().any(|node| match node {
        IRNode::Paragraph(p) => {
            p.properties.indentation.is_some()
                || p.properties.spacing.is_some()
                || p.properties.borders.is_some()
        }
        _ => false,
    });
    assert!(has_props);
}

#[test]
fn rtf_group_depth_limit() {
    let data = b"{\\rtf1{{{a}}}}";
    let mut config = ParserConfig::default();
    config.rtf.max_group_depth = 3;
    let parser = RtfParser::with_config(config);
    let err = parser
        .parse_bytes(data)
        .expect_err("should hit depth limit");
    match err {
        ParseError::ResourceLimit(message) => {
            assert!(message.contains("RTF max group depth"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn rtf_object_hex_limit() {
    let data = b"{\\rtf1{\\object{\\objdata 0102030405}}}";
    let mut config = ParserConfig::default();
    config.rtf.max_object_hex_len = 4;
    let parser = RtfParser::with_config(config);
    let err = parser
        .parse_bytes(data)
        .expect_err("should hit objdata limit");
    match err {
        ParseError::ResourceLimit(message) => {
            assert!(message.contains("RTF objdata too large"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn parse_reader_rejects_non_rtf_input() {
    let parser = RtfParser::new();
    let err = parser
        .parse_reader(Cursor::new(b"plain text".as_slice()))
        .expect_err("non-rtf input should fail");
    match err {
        ParseError::UnsupportedFormat(message) => {
            assert!(message.contains("Missing RTF header"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn parse_reader_enforces_max_input_size_before_parse() {
    let config = ParserConfig {
        max_input_size: 8,
        ..ParserConfig::default()
    };
    let parser = RtfParser::with_config(config);
    let data = b"{\\rtf1\\ansi too big}";

    let err = parser
        .parse_reader(Cursor::new(data.as_slice()))
        .expect_err("oversized input should fail");
    match err {
        ParseError::ResourceLimit(message) => {
            assert!(message.contains("Input too large"));
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn format_parser_trait_parse_reader_dispatches_to_rtf_parser() {
    let parser = RtfParser::new();
    let data = b"{\\rtf1\\ansi Trait dispatch}";

    let parsed = <RtfParser as FormatParser>::parse_reader(&parser, Cursor::new(data.as_slice()))
        .expect("trait parse_reader should parse rtf");
    assert_eq!(parsed.format, DocumentFormat::Rtf);
}
