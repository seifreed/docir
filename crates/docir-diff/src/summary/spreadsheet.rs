use docir_core::ir::IRNode;
use docir_core::visitor::IrStore;

use super::summary_parse::{abbreviate, opt_str, opt_u32, summarize_cell, summarize_worksheet};

pub(crate) fn summarize(node: &IRNode, _store: &IrStore) -> Option<String> {
    match node {
        IRNode::Worksheet(ws) => Some(summarize_worksheet(ws)),
        IRNode::Cell(cell) => Some(summarize_cell(cell)),
        IRNode::SharedStringTable(table) => Some(format!("items={}", table.items.len())),
        IRNode::ConnectionPart(part) => Some(format!("connections={}", part.entries.len())),
        IRNode::SpreadsheetStyles(styles) => Some(summarize_styles(styles)),
        IRNode::DefinedName(name) => Some(summarize_defined_name(name)),
        IRNode::ConditionalFormat(fmt) => Some(summarize_conditional_format(fmt)),
        IRNode::DataValidation(val) => Some(summarize_data_validation(val)),
        IRNode::TableDefinition(table) => Some(summarize_table_definition(table)),
        IRNode::PivotTable(pivot) => Some(summarize_pivot_table(pivot)),
        IRNode::PivotCache(cache) => Some(summarize_pivot_cache(cache)),
        IRNode::PivotCacheRecords(records) => Some(summarize_pivot_cache_records(records)),
        IRNode::WorkbookProperties(props) => Some(summarize_workbook_properties(props)),
        IRNode::CalcChain(chain) => Some(format!("entries={}", chain.entries.len())),
        IRNode::SheetComment(comment) => Some(summarize_sheet_comment(comment)),
        IRNode::SheetMetadata(meta) => Some(summarize_sheet_metadata(meta)),
        IRNode::WorksheetDrawing(d) => Some(format!("shapes={}", d.shapes.len())),
        IRNode::ChartData(c) => Some(summarize_chart_data(c)),
        IRNode::ExternalLinkPart(part) => Some(summarize_external_link(part)),
        IRNode::SlicerPart(part) => Some(summarize_slicer(part)),
        IRNode::TimelinePart(part) => Some(summarize_timeline(part)),
        IRNode::QueryTablePart(part) => Some(summarize_query_table(part)),
        _ => None,
    }
}

fn summarize_styles(styles: &docir_core::ir::SpreadsheetStyles) -> String {
    format!(
        "numfmts={} fonts={} fills={} borders={} xfs={} style_xfs={} dxfs={} table_styles={}",
        styles.number_formats.len(),
        styles.fonts.len(),
        styles.fills.len(),
        styles.borders.len(),
        styles.cell_xfs.len(),
        styles.cell_style_xfs.len(),
        styles.dxfs.len(),
        styles
            .table_styles
            .as_ref()
            .and_then(|t| t.default_table_style.as_ref())
            .map_or("-", |v| v.as_str())
    )
}

fn summarize_defined_name(name: &docir_core::ir::DefinedName) -> String {
    format!(
        "name={} scope={}",
        name.name,
        name.local_sheet_id
            .map_or("global".to_string(), |id| format!("sheet:{id}"))
    )
}

fn summarize_conditional_format(fmt: &docir_core::ir::ConditionalFormat) -> String {
    format!("ranges={} rules={}", fmt.ranges.len(), fmt.rules.len())
}

fn summarize_data_validation(val: &docir_core::ir::DataValidation) -> String {
    format!(
        "ranges={} type={}",
        val.ranges.len(),
        opt_str(&val.validation_type)
    )
}

fn summarize_table_definition(table: &docir_core::ir::TableDefinition) -> String {
    format!(
        "name={} cols={} ref={}",
        opt_str(&table.display_name),
        table.columns.len(),
        opt_str(&table.ref_range)
    )
}

fn summarize_pivot_table(pivot: &docir_core::ir::PivotTable) -> String {
    format!(
        "name={} cache_id={}",
        opt_str(&pivot.name),
        opt_u32(pivot.cache_id)
    )
}

fn summarize_pivot_cache(cache: &docir_core::ir::PivotCache) -> String {
    format!(
        "cache_id={} source={}",
        cache.cache_id,
        opt_str(&cache.cache_source)
    )
}

fn summarize_pivot_cache_records(records: &docir_core::ir::PivotCacheRecords) -> String {
    format!(
        "records={} fields={}",
        records.record_count.unwrap_or(0),
        records.field_count.unwrap_or(0)
    )
}

fn summarize_workbook_properties(props: &docir_core::ir::WorkbookProperties) -> String {
    format!(
        "calc_mode={} protected={}",
        opt_str(&props.calc_mode),
        props.workbook_protected
    )
}

fn summarize_sheet_comment(comment: &docir_core::ir::SheetComment) -> String {
    format!(
        "cell_ref={} author={} text={}",
        comment.cell_ref,
        opt_str(&comment.author),
        abbreviate(&comment.text, 80)
    )
}

fn summarize_sheet_metadata(meta: &docir_core::ir::SheetMetadata) -> String {
    format!(
        "types={} cell_count={} value_count={}",
        meta.metadata_types.len(),
        opt_u32(meta.cell_metadata_count),
        opt_u32(meta.value_metadata_count)
    )
}

fn summarize_chart_data(c: &docir_core::ir::ChartData) -> String {
    format!(
        "type={} series={} series_data={}",
        opt_str(&c.chart_type),
        c.series.len(),
        c.series_data.len()
    )
}

fn summarize_external_link(part: &docir_core::ir::ExternalLinkPart) -> String {
    format!(
        "type={} target={} sheets={}",
        opt_str(&part.link_type),
        opt_str(&part.target),
        part.sheets.len()
    )
}

fn summarize_slicer(part: &docir_core::ir::SlicerPart) -> String {
    format!(
        "name={} caption={} cache_id={}",
        opt_str(&part.name),
        opt_str(&part.caption),
        opt_str(&part.cache_id)
    )
}

fn summarize_timeline(part: &docir_core::ir::TimelinePart) -> String {
    format!(
        "name={} cache_id={}",
        opt_str(&part.name),
        opt_str(&part.cache_id)
    )
}

fn summarize_query_table(part: &docir_core::ir::QueryTablePart) -> String {
    format!(
        "name={} connection_id={} url={}",
        opt_str(&part.name),
        opt_str(&part.connection_id),
        opt_str(&part.url)
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::{
        CalcChain, CalcChainEntry, Cell, ChartData, ConditionalFormat, DataValidation, DefinedName,
        ExternalLinkPart, PivotCache, PivotCacheRecords, PivotTable, QueryTablePart,
        SharedStringItem, SharedStringTable, SheetComment, SheetMetadata, SheetMetadataType,
        SlicerPart, SpreadsheetStyles, TableColumn, TableDefinition, TimelinePart,
        WorkbookProperties, Worksheet, WorksheetDrawing,
    };
    use docir_core::types::NodeId;
    use docir_core::visitor::IrStore;

    fn assert_summary_contains(node: IRNode, expected: &str) {
        let store = IrStore::new();
        let summary = summarize(&node, &store).unwrap();
        assert!(
            summary.contains(expected),
            "expected summary to contain '{expected}', got '{summary}'"
        );
    }

    fn assert_summary_eq(node: IRNode, expected: &str) {
        let store = IrStore::new();
        let summary = summarize(&node, &store).unwrap();
        assert_eq!(summary, expected);
    }

    #[test]
    fn summarizes_spreadsheet_top_level_nodes() {
        let mut sheet = Worksheet::new("Sheet1", 1);
        sheet.merged_cells.push(docir_core::ir::MergedCellRange {
            start_col: 0,
            start_row: 0,
            end_col: 1,
            end_row: 1,
        });
        assert_summary_contains(IRNode::Worksheet(sheet), "name=Sheet1");

        let mut cell = Cell::new("B2", 1, 1);
        cell.value = docir_core::ir::CellValue::String("value".to_string());
        assert_summary_contains(IRNode::Cell(cell), "ref=B2");

        let mut sst = SharedStringTable::new();
        sst.items.push(SharedStringItem {
            text: "alpha".to_string(),
            runs: vec![],
        });
        assert_summary_eq(IRNode::SharedStringTable(sst), "items=1");
    }

    #[test]
    fn summarizes_spreadsheet_connection_and_style_nodes() {
        let mut conn = docir_core::ir::ConnectionPart::new();
        conn.entries.push(docir_core::ir::ConnectionEntry::new());
        assert_summary_eq(IRNode::ConnectionPart(conn), "connections=1");

        let mut styles = SpreadsheetStyles::new();
        styles.table_styles = Some(docir_core::ir::TableStyleInfo {
            count: Some(1),
            default_table_style: Some("TableStyleLight1".to_string()),
            default_pivot_style: None,
            styles: vec![],
        });
        assert_summary_contains(
            IRNode::SpreadsheetStyles(styles),
            "table_styles=TableStyleLight1",
        );

        let def_name = DefinedName {
            id: NodeId::new(),
            name: "NamedRange".to_string(),
            value: "Sheet1!$A$1".to_string(),
            local_sheet_id: Some(2),
            hidden: false,
            comment: None,
            span: None,
        };
        assert_summary_eq(
            IRNode::DefinedName(def_name),
            "name=NamedRange scope=sheet:2",
        );
    }

    #[test]
    fn summarizes_spreadsheet_validation_and_table_nodes() {
        let cond = ConditionalFormat {
            id: NodeId::new(),
            ranges: vec!["A1:A3".to_string()],
            rules: vec![],
            span: None,
        };
        assert_summary_eq(IRNode::ConditionalFormat(cond), "ranges=1 rules=0");

        let validation = DataValidation {
            id: NodeId::new(),
            validation_type: Some("list".to_string()),
            operator: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            ranges: vec!["C1:C5".to_string()],
            formula1: None,
            formula2: None,
            span: None,
        };
        assert_summary_eq(IRNode::DataValidation(validation), "ranges=1 type=list");

        let table = TableDefinition {
            id: NodeId::new(),
            name: Some("Table1".to_string()),
            display_name: Some("Sales".to_string()),
            ref_range: Some("A1:B10".to_string()),
            header_row_count: Some(1),
            totals_row_count: Some(0),
            columns: vec![TableColumn {
                id: 1,
                name: Some("Amount".to_string()),
                totals_row_label: None,
                totals_row_function: None,
            }],
            span: None,
        };
        assert_summary_eq(
            IRNode::TableDefinition(table),
            "name=Sales cols=1 ref=A1:B10",
        );
    }

    #[test]
    fn summarizes_spreadsheet_pivot_and_chain_nodes() {
        let pivot = PivotTable {
            id: NodeId::new(),
            name: Some("Pivot1".to_string()),
            cache_id: Some(7),
            ref_range: None,
            span: None,
        };
        assert_summary_eq(IRNode::PivotTable(pivot), "name=Pivot1 cache_id=7");

        let mut cache = PivotCache::new(7);
        cache.cache_source = Some("A1:B10".to_string());
        assert_summary_eq(IRNode::PivotCache(cache), "cache_id=7 source=A1:B10");

        let mut records = PivotCacheRecords::new();
        records.record_count = Some(11);
        records.field_count = Some(3);
        assert_summary_eq(IRNode::PivotCacheRecords(records), "records=11 fields=3");

        let mut props = WorkbookProperties::new();
        props.calc_mode = Some("auto".to_string());
        props.workbook_protected = true;
        assert_summary_eq(
            IRNode::WorkbookProperties(props),
            "calc_mode=auto protected=true",
        );

        let mut chain = CalcChain::new();
        chain.entries.push(CalcChainEntry {
            cell_ref: "A1".to_string(),
            sheet_id: Some(1),
            index: None,
            level: None,
            new_value: None,
        });
        assert_summary_eq(IRNode::CalcChain(chain), "entries=1");
    }

    #[test]
    fn summarizes_spreadsheet_metadata_nodes() {
        let mut comment = SheetComment::new("A1", "Very long comment");
        comment.author = Some("Alice".to_string());
        assert_summary_contains(IRNode::SheetComment(comment), "author=Alice");

        let mut metadata = SheetMetadata::new();
        metadata.metadata_types.push(SheetMetadataType::new());
        metadata.cell_metadata_count = Some(2);
        metadata.value_metadata_count = Some(4);
        assert_summary_eq(
            IRNode::SheetMetadata(metadata),
            "types=1 cell_count=2 value_count=4",
        );

        let mut drawing = WorksheetDrawing::new();
        drawing.shapes.push(NodeId::new());
        assert_summary_eq(IRNode::WorksheetDrawing(drawing), "shapes=1");
    }

    #[test]
    fn summarizes_spreadsheet_visual_nodes() {
        let mut chart = ChartData::new();
        chart.chart_type = Some("bar".to_string());
        chart.series.push("Revenue".to_string());
        chart.series_data.push(docir_core::ir::ChartSeries::new());
        assert_summary_eq(IRNode::ChartData(chart), "type=bar series=1 series_data=1");

        let mut ext = ExternalLinkPart::new();
        ext.link_type = Some("externalWorkbook".to_string());
        ext.target = Some("http://example.test".to_string());
        ext.sheets.push(docir_core::ir::ExternalLinkSheet {
            name: Some("Sheet1".to_string()),
            r_id: None,
        });
        assert_summary_eq(
            IRNode::ExternalLinkPart(ext),
            "type=externalWorkbook target=http://example.test sheets=1",
        );

        let mut slicer = SlicerPart::new();
        slicer.name = Some("Region".to_string());
        slicer.caption = Some("Region".to_string());
        slicer.cache_id = Some("Cache1".to_string());
        assert_summary_eq(
            IRNode::SlicerPart(slicer),
            "name=Region caption=Region cache_id=Cache1",
        );

        let mut timeline = TimelinePart::new();
        timeline.name = Some("Date".to_string());
        timeline.cache_id = Some("Cache2".to_string());
        assert_summary_eq(IRNode::TimelinePart(timeline), "name=Date cache_id=Cache2");

        let mut query = QueryTablePart::new();
        query.name = Some("WebQuery".to_string());
        query.connection_id = Some("1".to_string());
        query.url = Some("https://example.test/data.csv".to_string());
        assert_summary_eq(
            IRNode::QueryTablePart(query),
            "name=WebQuery connection_id=1 url=https://example.test/data.csv",
        );
    }
}
