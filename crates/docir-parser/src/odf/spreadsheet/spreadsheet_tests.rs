use super::*;
use crate::odf::{CellValue, IRNode, OdfLimits};
use crate::parser::ParserConfig;
use docir_core::visitor::IrStore;

#[test]
fn parse_content_spreadsheet_accepts_alternate_namespace_prefixes() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<doc:document-content xmlns:doc="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:calc="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
  <calc:spreadsheet>
    <t:table t:name="AltContent">
      <t:table-row>
        <t:table-cell o:value-type="float" o:value="7" />
      </t:table-row>
    </t:table>
  </calc:spreadsheet>
</doc:document-content>
"#;
    let mut store = IrStore::new();
    let limits = OdfLimits::new(&ParserConfig::default(), false);

    let result =
        parse_content_spreadsheet(xml, &mut store, &limits).expect("spreadsheet should parse");

    assert_eq!(result.content.len(), 1);
    let Some(IRNode::Worksheet(sheet)) = store.get(result.content[0]) else {
        panic!("expected worksheet");
    };
    assert_eq!(sheet.name, "AltContent");
    assert_eq!(sheet.cells.len(), 1);

    let Some(IRNode::Cell(cell)) = store.get(sheet.cells[0]) else {
        panic!("expected cell");
    };
    assert_eq!(cell.reference, "A1");
    assert!(matches!(cell.value, CellValue::Number(value) if (value - 7.0).abs() < f64::EPSILON));
}
