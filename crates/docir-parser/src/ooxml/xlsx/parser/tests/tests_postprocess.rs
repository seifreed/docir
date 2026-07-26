use super::{Relationships, build_zip_with_entries};
use crate::error::ParseError;
use crate::ooxml::xlsx::{
    XlsxParser, extract_formula_function, map_cell_error, parse_formula, parse_formula_args_text,
    parse_formula_empty, parse_inline_string,
};
use docir_core::ir::{CellError, FormulaType, IRNode};
use quick_xml::Reader;
use quick_xml::events::Event;
#[test]
fn test_parse_xlm_macro_sheet() {
    let workbook_xml = r#"
        <workbook xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Macro1" sheetId="1" rel:id="rId1"/>
          </sheets>
        </workbook>
        "#;

    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/macrosheet"
            Target="macrosheets/sheet1.xml"/>
        </Relationships>
        "#;

    let sheet_xml = r#"
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1"><f>EXEC(\"calc\")</f></c>
              <c r="A2"><f>IF(A1;CALL(\"kernel32\",\"WinExec\");0)</f></c>
            </row>
          </sheetData>
        </worksheet>
        "#;

    let mut zip = build_zip_with_entries(vec![("xl/macrosheets/sheet1.xml", sheet_xml)]);
    let rels = Relationships::parse(rels_xml).expect("rels");

    let mut parser = XlsxParser::new();
    let root = parser
        .parse_workbook(&mut zip, workbook_xml, &rels, "xl/workbook.xml")
        .expect("workbook");
    let store = parser.into_store();
    let doc = match store.get(root) {
        Some(IRNode::Document(d)) => d,
        _ => panic!("missing document"),
    };
    assert_eq!(doc.security.xlm_macros.len(), 1);
    assert!(!doc.security.xlm_macros[0].macro_cells.is_empty());
    assert!(
        doc.security.xlm_macros[0]
            .dangerous_functions
            .iter()
            .any(|function| function.name == "EXEC" && function.cell_ref == "A1")
    );
    assert!(
        doc.security.xlm_macros[0]
            .dangerous_functions
            .iter()
            .any(|function| function.name == "CALL" && function.cell_ref == "A2")
    );
}

#[test]
fn test_parse_xlm_auto_open_defined_name() {
    let workbook_xml = r#"
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <definedNames>
            <definedName name="_xlnm.Auto_Open">Macro1!$A$1</definedName>
          </definedNames>
          <sheets>
            <sheet name="Macro1" sheetId="1" r:id="rId1"/>
          </sheets>
        </workbook>
        "#;

    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/macrosheet"
            Target="macrosheets/sheet1.xml"/>
        </Relationships>
        "#;

    let sheet_xml = r#"
        <worksheet>
          <sheetData>
            <row r="1">
              <c r="A1"><f>RUN(\"TEST\")</f></c>
            </row>
          </sheetData>
        </worksheet>
        "#;

    let mut zip = build_zip_with_entries(vec![("xl/macrosheets/sheet1.xml", sheet_xml)]);
    let rels = Relationships::parse(rels_xml).expect("rels");

    let mut parser = XlsxParser::new();
    let root = parser
        .parse_workbook(&mut zip, workbook_xml, &rels, "xl/workbook.xml")
        .expect("workbook");
    let store = parser.into_store();
    let doc = match store.get(root) {
        Some(IRNode::Document(d)) => d,
        _ => panic!("missing document"),
    };
    assert_eq!(doc.security.xlm_macros.len(), 1);
    assert!(doc.security.xlm_macros[0].has_auto_open);
}

#[test]
fn test_parse_empty_cell_and_formula_helpers() {
    let parser = XlsxParser::new();

    let mut reader = Reader::from_str(r#"<c r="C12"/>"#);
    let mut buf = Vec::new();
    let start = match reader.read_event_into(&mut buf).expect("cell start") {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let cell = parser
        .parse_empty_cell(&start, "xl/worksheets/sheet1.xml")
        .expect("empty cell");
    assert_eq!(cell.reference, "C12");
    assert_eq!(cell.column, 2);
    assert_eq!(cell.row, 11);

    let mut missing_reader = Reader::from_str("<c/>");
    buf.clear();
    let missing = match missing_reader
        .read_event_into(&mut buf)
        .expect("cell start")
    {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let err = parser
        .parse_empty_cell(&missing, "xl/worksheets/sheet1.xml")
        .expect_err("missing ref must fail");
    assert!(matches!(err, ParseError::InvalidStructure(_)));

    let mut formula_reader = Reader::from_str(r#"<f t="shared" si="7" ref="A1:A3">SUM(A1:A3)</f>"#);
    formula_reader.config_mut().trim_text(true);
    buf.clear();
    let formula_start = match formula_reader
        .read_event_into(&mut buf)
        .expect("formula start")
    {
        Event::Start(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let formula = parse_formula(
        &mut formula_reader,
        &formula_start,
        "xl/worksheets/sheet1.xml",
    )
    .expect("parse formula");
    assert_eq!(formula.text, "SUM(A1:A3)");
    assert_eq!(formula.formula_type, FormulaType::Shared);
    assert_eq!(formula.shared_index, Some(7));
    assert_eq!(formula.shared_ref.as_deref(), Some("A1:A3"));
    assert!(!formula.is_array);
    assert_eq!(formula.array_ref, None);

    let mut array_reader = Reader::from_str(r#"<f t="array" ref="B1:B9"></f>"#);
    buf.clear();
    let array_start = match array_reader.read_event_into(&mut buf).expect("array start") {
        Event::Start(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let empty_formula =
        parse_formula_empty(&array_start, "xl/worksheets/sheet1.xml").expect("empty formula");
    assert_eq!(empty_formula.formula_type, FormulaType::Array);
    assert!(empty_formula.is_array);
    assert_eq!(empty_formula.array_ref.as_deref(), Some("B1:B9"));
    assert_eq!(empty_formula.shared_ref, None);
}

#[test]
fn test_parse_formula_reports_malformed_attributes() {
    let mut reader = Reader::from_str(r#"<f t="shared" t="array">SUM(A1:A3)</f>"#);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let start = match reader.read_event_into(&mut buf).expect("formula start") {
        Event::Start(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let err = parse_formula(&mut reader, &start, "xl/worksheets/sheet1.xml")
        .expect_err("duplicate formula attrs must fail");
    assert!(matches!(
        err,
        ParseError::Xml {
            ref file,
            ..
        } if file == "xl/worksheets/sheet1.xml"
    ));

    let mut malformed_si_reader = Reader::from_str(r#"<f t="shared" si="bad">SUM(A1:A3)</f>"#);
    malformed_si_reader.config_mut().trim_text(true);
    buf.clear();
    let malformed_si_start = match malformed_si_reader
        .read_event_into(&mut buf)
        .expect("malformed si formula")
    {
        Event::Start(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let err = parse_formula(
        &mut malformed_si_reader,
        &malformed_si_start,
        "xl/worksheets/sheet1.xml",
    )
    .expect_err("malformed formula shared index must fail");
    assert!(matches!(
        err,
        ParseError::Xml {
            ref file,
            ..
        } if file == "xl/worksheets/sheet1.xml"
    ));

    let mut empty_reader = Reader::from_str(r#"<f t="shared" t="array"/>"#);
    buf.clear();
    let empty_start = match empty_reader
        .read_event_into(&mut buf)
        .expect("empty formula")
    {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let err = parse_formula_empty(&empty_start, "xl/worksheets/sheet1.xml")
        .expect_err("duplicate empty formula attrs must fail");
    assert!(matches!(
        err,
        ParseError::Xml {
            ref file,
            ..
        } if file == "xl/worksheets/sheet1.xml"
    ));

    let mut empty_reader = Reader::from_str(r#"<f t="shared" si="bad"/>"#);
    buf.clear();
    let empty_start = match empty_reader
        .read_event_into(&mut buf)
        .expect("empty malformed si formula")
    {
        Event::Empty(e) => e.into_owned(),
        other => panic!("unexpected event: {other:?}"),
    };
    let err = parse_formula_empty(&empty_start, "xl/worksheets/sheet1.xml")
        .expect_err("malformed empty formula shared index must fail");
    assert!(matches!(
        err,
        ParseError::Xml {
            ref file,
            ..
        } if file == "xl/worksheets/sheet1.xml"
    ));
}

#[test]
fn test_parse_formula_text_helpers() {
    assert_eq!(
        extract_formula_function("=SUM(A1:A3)"),
        Some("SUM".to_string())
    );
    assert_eq!(
        extract_formula_function(" COUNTIF(A:A,\">0\") "),
        Some("COUNTIF".to_string())
    );
    assert_eq!(extract_formula_function("A1+1"), None);

    assert_eq!(
        parse_formula_args_text("SUM(A1, B2)"),
        Some("A1, B2".to_string())
    );
    assert_eq!(parse_formula_args_text("NOW()"), None);
    assert_eq!(parse_formula_args_text("BROKEN("), None);
}

#[test]
fn test_parse_inline_string_and_error_mapping() {
    let mut reader = Reader::from_str(
        r#"<is><r><t>Hello</t></r><r><t xml:space="preserve"> World</t></r></is>"#,
    );
    let mut buf = Vec::new();
    match reader.read_event_into(&mut buf).expect("is start") {
        Event::Start(e) => assert_eq!(e.name().as_ref(), b"is"),
        other => panic!("unexpected event: {other:?}"),
    }
    let text = parse_inline_string(&mut reader, "xl/worksheets/sheet1.xml").expect("inline");
    assert_eq!(text, "Hello World");

    assert_eq!(map_cell_error("#NULL!"), CellError::Null);
    assert_eq!(map_cell_error("#DIV/0!"), CellError::DivZero);
    assert_eq!(map_cell_error("#GETTING_DATA"), CellError::GettingData);
    assert_eq!(map_cell_error("unexpected"), CellError::Value);
}
