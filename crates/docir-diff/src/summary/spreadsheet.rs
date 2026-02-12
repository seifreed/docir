use docir_core::ir::IRNode;
use docir_core::visitor::IrStore;

use super::{abbreviate, opt_str, opt_u32, summarize_cell, summarize_worksheet};

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
