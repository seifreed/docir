use docir_core::ir::IRNode;
use docir_parser::DocumentParser;

#[test]
fn parse_minimal_xlsb_cells() {
    let parser = DocumentParser::new();
    let fixture_path = format!(
        "{}/../../fixtures/ooxml/minimal.xlsb",
        env!("CARGO_MANIFEST_DIR")
    );
    let parsed = parser.parse_file(fixture_path).expect("parse xlsb");

    let doc = parsed.document().expect("document");
    assert_eq!(doc.format.display_name(), "Excel Spreadsheet");
    assert!(!doc.content.is_empty(), "expected worksheets");

    let mut cell_count = 0usize;
    for sheet_id in &doc.content {
        if let Some(IRNode::Worksheet(sheet)) = parsed.store.get(*sheet_id) {
            cell_count += sheet.cells.len();
        }
    }
    assert!(cell_count > 0, "expected at least one cell");
}

#[test]
fn parse_values_xlsb_known_cells() {
    let parser = DocumentParser::new();
    let fixture_path = format!(
        "{}/../../fixtures/ooxml/values.xlsb",
        env!("CARGO_MANIFEST_DIR")
    );
    let parsed = parser.parse_file(fixture_path).expect("parse xlsb values");

    let doc = parsed.document().expect("document");
    let mut sheet_ids = Vec::new();
    for sheet_id in &doc.content {
        if let Some(IRNode::Worksheet(sheet)) = parsed.store.get(*sheet_id) {
            sheet_ids.push((sheet.name.clone(), *sheet_id));
        }
    }

    let mut seen_string = false;
    let mut seen_number = false;
    let mut seen_datetime = false;

    for (name, sheet_id) in sheet_ids {
        let sheet = match parsed.store.get(sheet_id) {
            Some(IRNode::Worksheet(sheet)) => sheet,
            _ => continue,
        };
        if name == "Dropdown" {
            for cell_id in &sheet.cells {
                if let Some(IRNode::Cell(cell)) = parsed.store.get(*cell_id) {
                    if cell.reference == "A1" {
                        if let docir_core::ir::CellValue::String(value) = &cell.value {
                            assert_eq!(value, "Deutsch");
                            seen_string = true;
                        }
                    }
                    if cell.reference == "A77" {
                        if let docir_core::ir::CellValue::Number(value) = cell.value {
                            assert!((value - 1.0).abs() < f64::EPSILON);
                            seen_number = true;
                        }
                    }
                }
            }
        }
        if name == "Verpackung_ReSet_Plastic" {
            for cell_id in &sheet.cells {
                if let Some(IRNode::Cell(cell)) = parsed.store.get(*cell_id) {
                    if cell.reference == "L11" {
                        if let docir_core::ir::CellValue::DateTime(value) = cell.value {
                            assert!((value - 45627.0).abs() < 0.0001);
                            seen_datetime = true;
                        }
                    }
                }
            }
        }
    }

    assert!(seen_string, "expected Dropdown!A1 string");
    assert!(seen_number, "expected Dropdown!A77 number");
    assert!(
        seen_datetime,
        "expected Verpackung_ReSet_Plastic!L11 datetime"
    );
}
