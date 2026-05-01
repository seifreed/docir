use super::*;
use crate::odf::helpers::ValidationDef;

fn default_limits() -> OdfLimits {
    OdfLimits::new(&ParserConfig::default(), false)
}

fn parse_table_start(xml: &[u8]) -> (Reader<std::io::Cursor<&[u8]>>, BytesStart<'static>) {
    let mut reader = Reader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let table_start = loop {
        match reader.read_event_into(&mut buf).unwrap() {
            Event::Start(e) if e.name().as_ref() == b"table:table" => break e.into_owned(),
            Event::Eof => panic!("missing table start"),
            _ => {}
        }
        buf.clear();
    };
    (reader, table_start)
}

#[test]
fn parse_ods_table_evaluates_formula_and_flushes_validations() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <table:table table:name="SheetCalc">
    <table:table-row>
      <table:table-cell table:cell-value-type="float" table:cell-value="1" />
      <table:table-cell table:cell-value-type="float" table:cell-value="2" />
      <table:table-cell table:formula="of:=SUM([.A1:.B1])" table:content-validation-name="val1" />
    </table:table-row>
  </table:table>
</office:spreadsheet>
"#;

    let (mut reader, table_start) = parse_table_start(xml);
    let mut store = IrStore::new();
    let mut validations = HashMap::new();
    validations.insert(
        "val1".to_string(),
        ValidationDef {
            validation_type: Some("between".to_string()),
            operator: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            formula1: Some("1".to_string()),
            formula2: Some("10".to_string()),
        },
    );

    let worksheet = parse_ods_table(
        &mut reader,
        &table_start,
        1,
        &mut store,
        &validations,
        &default_limits(),
    )
    .unwrap();

    assert_eq!(worksheet.name, "SheetCalc");
    assert_eq!(worksheet.cells.len(), 3);
    assert_eq!(worksheet.data_validations.len(), 1);

    let formula_cell_id = worksheet
        .cells
        .iter()
        .find(|id| {
            matches!(
                store.get(**id),
                Some(IRNode::Cell(Cell {
                    reference,
                    formula: Some(_),
                    ..
                })) if reference == "C1"
            )
        })
        .copied()
        .expect("expected formula cell at C1");
    let Some(IRNode::Cell(cell)) = store.get(formula_cell_id) else {
        panic!("formula cell missing");
    };
    match cell.value {
        CellValue::Number(value) => assert!((value - 3.0).abs() < f64::EPSILON),
        _ => panic!("expected numeric formula result"),
    }
}

#[test]
fn parse_ods_cell_empty_handles_type_and_formula_prefix() {
    let mut cell = BytesStart::new("table:table-cell");
    cell.push_attribute(("table:cell-value-type", "boolean"));
    cell.push_attribute(("table:cell-value", "true"));
    cell.push_attribute(("table:formula", "of:=SUM([.A1];[.A2])"));
    let mut style_map = HashMap::new();
    let mut next_style_id = 1;
    let parsed = parse_ods_cell_empty(&cell, &mut style_map, &mut next_style_id).unwrap();

    match parsed.value {
        CellValue::Boolean(value) => assert!(value),
        _ => panic!("expected boolean value"),
    }
    assert_eq!(
        parsed.formula.as_ref().map(|f| f.text.as_str()),
        Some("SUM([.A1];[.A2])")
    );
}

#[test]
fn parse_ods_table_fast_samples_rows_and_flushes_validation_span() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <table:table table:name="Sampled">
    <table:table-row>
      <table:table-cell table:content-validation-name="rule1"><text:p>kept</text:p></table:table-cell>
      <table:table-cell><text:p>trimmed</text:p></table:table-cell>
    </table:table-row>
    <table:table-row>
      <table:table-cell><text:p>skipped-row</text:p></table:table-cell>
    </table:table-row>
  </table:table>
</office:spreadsheet>
"#;
    let (mut reader, table_start) = parse_table_start(xml);
    let mut store = IrStore::new();
    let mut validations = HashMap::new();
    validations.insert(
        "rule1".to_string(),
        ValidationDef {
            validation_type: Some("list".to_string()),
            operator: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            formula1: Some("\"A;B\"".to_string()),
            formula2: None,
        },
    );
    let mut config = ParserConfig::default();
    config.odf.fast_sample_rows = 1;
    config.odf.fast_sample_cols = 1;
    let limits = OdfLimits::new(&config, true);

    let worksheet = parse_ods_table_fast(
        &mut reader,
        &table_start,
        1,
        &mut store,
        &validations,
        &limits,
    )
    .unwrap();

    assert_eq!(worksheet.name, "Sampled");
    assert_eq!(worksheet.cells.len(), 1);
    assert_eq!(worksheet.data_validations.len(), 1);

    let Some(IRNode::Cell(cell)) = store.get(worksheet.cells[0]) else {
        panic!("expected sampled cell");
    };
    assert_eq!(cell.reference, "A1");

    let Some(IRNode::DataValidation(validation)) = store.get(worksheet.data_validations[0]) else {
        panic!("expected validation node");
    };
    assert_eq!(validation.ranges, vec!["A1"]);
    assert_eq!(
        validation.span.as_ref().map(|span| span.file_path.as_str()),
        Some("content.xml")
    );
}

#[test]
fn parse_ods_table_expands_repeated_rows_cells_and_covered_offsets() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <table:table table:name="Repeated">
    <table:table-row table:number-rows-repeated="2">
      <table:table-cell table:number-columns-repeated="2"><text:p>X</text:p></table:table-cell>
      <table:covered-table-cell table:number-columns-repeated="1"/>
      <table:table-cell table:cell-value-type="boolean" table:cell-value="true"
        table:content-validation-name="rule1" />
    </table:table-row>
  </table:table>
</office:spreadsheet>
"#;
    let (mut reader, table_start) = parse_table_start(xml);
    let mut store = IrStore::new();
    let mut validations = HashMap::new();
    validations.insert(
        "rule1".to_string(),
        ValidationDef {
            validation_type: Some("list".to_string()),
            operator: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            formula1: Some("\"yes;no\"".to_string()),
            formula2: None,
        },
    );

    let worksheet = parse_ods_table(
        &mut reader,
        &table_start,
        1,
        &mut store,
        &validations,
        &default_limits(),
    )
    .expect("table should parse");

    assert_eq!(worksheet.name, "Repeated");
    assert_eq!(worksheet.cells.len(), 6);
    assert_eq!(worksheet.data_validations.len(), 1);

    let refs: Vec<String> = worksheet
        .cells
        .iter()
        .filter_map(|id| match store.get(*id) {
            Some(IRNode::Cell(cell)) => Some(cell.reference.clone()),
            _ => None,
        })
        .collect();
    assert!(refs.contains(&"A1".to_string()));
    assert!(refs.contains(&"B1".to_string()));
    assert!(refs.contains(&"D1".to_string()));
    assert!(refs.contains(&"A2".to_string()));
    assert!(refs.contains(&"B2".to_string()));
    assert!(refs.contains(&"D2".to_string()));

    let Some(IRNode::DataValidation(validation)) = store.get(worksheet.data_validations[0]) else {
        panic!("expected data validation");
    };
    assert_eq!(validation.ranges, vec!["D1", "D2"]);
}

#[test]
fn parse_ods_table_uses_default_sheet_name_and_default_validation_shape() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <table:table>
    <table:table-row>
      <table:table-cell table:content-validation-name="missing"><text:p>v</text:p></table:table-cell>
    </table:table-row>
  </table:table>
</office:spreadsheet>
"#;
    let (mut reader, table_start) = parse_table_start(xml);
    let mut store = IrStore::new();

    let worksheet = parse_ods_table(
        &mut reader,
        &table_start,
        7,
        &mut store,
        &HashMap::new(),
        &default_limits(),
    )
    .expect("table should parse");

    assert_eq!(worksheet.name, "Sheet7");
    assert_eq!(worksheet.data_validations.len(), 1);

    let Some(IRNode::DataValidation(validation)) = store.get(worksheet.data_validations[0]) else {
        panic!("expected validation node");
    };
    assert_eq!(validation.ranges, vec!["A1"]);
    assert!(validation.validation_type.is_none());
    assert!(validation.operator.is_none());
    assert!(!validation.allow_blank);
    assert!(!validation.show_input_message);
    assert!(!validation.show_error_message);
    assert!(validation.span.is_none());
}

#[test]
fn parse_ods_table_fast_sampling_applies_row_and_col_repeat_bounds() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <table:table table:name="FastRepeated">
    <table:table-row table:number-rows-repeated="3">
      <table:table-cell table:number-columns-repeated="3" table:content-validation-name="rule1">
        <text:p>v</text:p>
      </table:table-cell>
      <table:table-cell><text:p>tail</text:p></table:table-cell>
    </table:table-row>
  </table:table>
</office:spreadsheet>
"#;
    let (mut reader, table_start) = parse_table_start(xml);
    let mut store = IrStore::new();
    let mut validations = HashMap::new();
    validations.insert(
        "rule1".to_string(),
        ValidationDef {
            validation_type: Some("list".to_string()),
            operator: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            formula1: Some("\"A;B\"".to_string()),
            formula2: None,
        },
    );
    let mut config = ParserConfig::default();
    config.odf.fast_sample_rows = 2;
    config.odf.fast_sample_cols = 2;
    let limits = OdfLimits::new(&config, true);

    let worksheet = parse_ods_table_fast(
        &mut reader,
        &table_start,
        1,
        &mut store,
        &validations,
        &limits,
    )
    .expect("table should parse in fast mode");

    assert_eq!(worksheet.name, "FastRepeated");
    assert_eq!(worksheet.cells.len(), 4);
    assert_eq!(worksheet.data_validations.len(), 1);

    let refs: Vec<String> = worksheet
        .cells
        .iter()
        .filter_map(|id| match store.get(*id) {
            Some(IRNode::Cell(cell)) => Some(cell.reference.clone()),
            _ => None,
        })
        .collect();
    assert_eq!(refs, vec!["A1", "B1", "A2", "B2"]);

    let Some(IRNode::DataValidation(validation)) = store.get(worksheet.data_validations[0]) else {
        panic!("expected validation node");
    };
    assert_eq!(validation.ranges, vec!["A1", "B1", "A2", "B2"]);
    assert_eq!(
        validation.span.as_ref().map(|span| span.file_path.as_str()),
        Some("content.xml")
    );
}

#[test]
fn parse_ods_table_reports_malformed_xml_as_content_error() {
    let malformed: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <table:table table:name="Broken">
    <table:table-row>
      <table:table-cell table:cell-value-type="string"><text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">oops</text:p>
  </table:table>
</office:spreadsheet>
"#;
    let (mut reader, table_start) = parse_table_start(malformed);
    let mut store = IrStore::new();
    let err = parse_ods_table(
        &mut reader,
        &table_start,
        1,
        &mut store,
        &HashMap::new(),
        &default_limits(),
    )
    .unwrap_err();

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("expected content.xml parse error, got {:?}", other),
    }
}

#[test]
fn parse_ods_table_fast_reports_malformed_xml_as_content_error() {
    let malformed: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <table:table table:name="BrokenFast">
    <table:table-row>
      <table:table-cell table:cell-value-type="string"><text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">oops</text:p>
  </table:table>
</office:spreadsheet>
"#;
    let (mut reader, table_start) = parse_table_start(malformed);
    let mut store = IrStore::new();
    let mut config = ParserConfig::default();
    config.odf.fast_sample_rows = 1;
    config.odf.fast_sample_cols = 1;
    let limits = OdfLimits::new(&config, true);

    let err = parse_ods_table_fast(
        &mut reader,
        &table_start,
        1,
        &mut store,
        &HashMap::new(),
        &limits,
    )
    .unwrap_err();

    match err {
        ParseError::Xml { file, .. } => assert_eq!(file, "content.xml"),
        other => panic!("expected content.xml parse error, got {:?}", other),
    }
}

#[test]
fn parse_ods_cell_empty_reuses_style_id_and_handles_empty_boolean() {
    let mut style_map = HashMap::new();
    let mut next_style_id = 1;

    let mut first = BytesStart::new("table:table-cell");
    first.push_attribute(("table:style-name", "s1"));
    first.push_attribute(("table:cell-value-type", "boolean"));
    let parsed_first = parse_ods_cell_empty(&first, &mut style_map, &mut next_style_id)
        .expect("first cell should parse");
    assert_eq!(parsed_first.style_id, Some(1));
    assert!(matches!(parsed_first.value, CellValue::Empty));

    let mut second = BytesStart::new("table:table-cell");
    second.push_attribute(("table:style-name", "s1"));
    second.push_attribute(("table:formula", "=A1"));
    let parsed_second = parse_ods_cell_empty(&second, &mut style_map, &mut next_style_id)
        .expect("second cell should parse");
    assert_eq!(parsed_second.style_id, Some(1), "style id should be reused");
    assert_eq!(next_style_id, 2, "no extra style ids should be allocated");
    assert_eq!(
        parsed_second.formula.as_ref().map(|f| f.text.as_str()),
        Some("A1")
    );
}

#[test]
fn parse_ods_table_fast_with_disabled_sampling_skips_row_content() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <table:table table:name="FastOff">
    <table:table-row>
      <table:table-cell table:content-validation-name="rule1"><text:p>value</text:p></table:table-cell>
    </table:table-row>
  </table:table>
</office:spreadsheet>
"#;
    let (mut reader, table_start) = parse_table_start(xml);
    let mut store = IrStore::new();
    let mut validations = HashMap::new();
    validations.insert(
        "rule1".to_string(),
        ValidationDef {
            validation_type: Some("list".to_string()),
            operator: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            formula1: Some("\"A;B\"".to_string()),
            formula2: None,
        },
    );

    let mut config = ParserConfig::default();
    config.odf.fast_sample_rows = 0;
    config.odf.fast_sample_cols = 2;
    let limits = OdfLimits::new(&config, true);

    let worksheet = parse_ods_table_fast(
        &mut reader,
        &table_start,
        1,
        &mut store,
        &validations,
        &limits,
    )
    .expect("fast table should parse");

    assert_eq!(worksheet.name, "FastOff");
    assert!(worksheet.cells.is_empty());
    assert!(worksheet.data_validations.is_empty());
}
