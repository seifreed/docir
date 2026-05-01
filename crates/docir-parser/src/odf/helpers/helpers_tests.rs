use super::*;
use docir_core::ir::{CellFormula, CellValue, FormulaType};

#[test]
fn ods_cell_data_should_emit_depends_on_value_or_formula() {
    let empty = OdsCellData {
        value: CellValue::Empty,
        formula: None,
        style_id: None,
        col_repeat: 1,
        validation_name: None,
        col_span: None,
        row_span: None,
        is_covered: false,
    };
    assert!(!empty.should_emit());

    let with_value = OdsCellData {
        value: CellValue::String("x".to_string()),
        ..empty.clone()
    };
    assert!(with_value.should_emit());

    let with_formula = OdsCellData {
        value: CellValue::Empty,
        formula: Some(CellFormula {
            text: "of:=1+1".to_string(),
            formula_type: FormulaType::Normal,
            shared_index: None,
            shared_ref: None,
            is_array: false,
            array_ref: None,
        }),
        ..empty
    };
    assert!(with_formula.should_emit());
}

#[test]
fn ods_cell_data_merge_range_builds_span() {
    let cell = OdsCellData {
        value: CellValue::Empty,
        formula: None,
        style_id: None,
        col_repeat: 1,
        validation_name: None,
        col_span: Some(2),
        row_span: Some(3),
        is_covered: false,
    };

    let merged = cell.merge_range(5, 7).expect("merge range");
    assert_eq!(merged.start_row, 5);
    assert_eq!(merged.start_col, 7);
    assert_eq!(merged.end_row, 7);
    assert_eq!(merged.end_col, 8);
}
