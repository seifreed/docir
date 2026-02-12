use docir_core::ir::IRNode;
use docir_core::visitor::IrStore;

use super::{abbreviate, opt_str, opt_u32, summarize_cell, summarize_worksheet};

pub(crate) fn summarize(node: &IRNode, _store: &IrStore) -> Option<String> {
    match node {
        IRNode::Worksheet(ws) => Some(summarize_worksheet(ws)),
        IRNode::Cell(cell) => Some(summarize_cell(cell)),
        IRNode::SharedStringTable(table) => Some(format!("items={}", table.items.len())),
        IRNode::ConnectionPart(part) => Some(format!("connections={}", part.entries.len())),
        IRNode::SpreadsheetStyles(styles) => Some(format!(
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
        )),
        IRNode::DefinedName(name) => Some(format!(
            "name={} scope={}",
            name.name,
            name.local_sheet_id
                .map_or("global".to_string(), |id| format!("sheet:{id}"))
        )),
        IRNode::ConditionalFormat(fmt) => Some(format!(
            "ranges={} rules={}",
            fmt.ranges.len(),
            fmt.rules.len()
        )),
        IRNode::DataValidation(val) => Some(format!(
            "ranges={} type={}",
            val.ranges.len(),
            opt_str(&val.validation_type)
        )),
        IRNode::TableDefinition(table) => Some(format!(
            "name={} cols={} ref={}",
            opt_str(&table.display_name),
            table.columns.len(),
            opt_str(&table.ref_range)
        )),
        IRNode::PivotTable(pivot) => Some(format!(
            "name={} cache_id={}",
            opt_str(&pivot.name),
            opt_u32(pivot.cache_id)
        )),
        IRNode::PivotCache(cache) => Some(format!(
            "cache_id={} source={}",
            cache.cache_id,
            opt_str(&cache.cache_source)
        )),
        IRNode::PivotCacheRecords(records) => Some(format!(
            "records={} fields={}",
            records.record_count.unwrap_or(0),
            records.field_count.unwrap_or(0)
        )),
        IRNode::WorkbookProperties(props) => Some(format!(
            "calc_mode={} protected={}",
            opt_str(&props.calc_mode),
            props.workbook_protected
        )),
        IRNode::CalcChain(chain) => Some(format!("entries={}", chain.entries.len())),
        IRNode::SheetComment(comment) => Some(format!(
            "cell_ref={} author={} text={}",
            comment.cell_ref,
            opt_str(&comment.author),
            abbreviate(&comment.text, 80)
        )),
        IRNode::SheetMetadata(meta) => Some(format!(
            "types={} cell_count={} value_count={}",
            meta.metadata_types.len(),
            opt_u32(meta.cell_metadata_count),
            opt_u32(meta.value_metadata_count)
        )),
        IRNode::WorksheetDrawing(d) => Some(format!("shapes={}", d.shapes.len())),
        IRNode::ChartData(c) => Some(format!(
            "type={} series={} series_data={}",
            opt_str(&c.chart_type),
            c.series.len(),
            c.series_data.len()
        )),
        IRNode::ExternalLinkPart(part) => Some(format!(
            "type={} target={} sheets={}",
            opt_str(&part.link_type),
            opt_str(&part.target),
            part.sheets.len()
        )),
        IRNode::SlicerPart(part) => Some(format!(
            "name={} caption={} cache_id={}",
            opt_str(&part.name),
            opt_str(&part.caption),
            opt_str(&part.cache_id)
        )),
        IRNode::TimelinePart(part) => Some(format!(
            "name={} cache_id={}",
            opt_str(&part.name),
            opt_str(&part.cache_id)
        )),
        IRNode::QueryTablePart(part) => Some(format!(
            "name={} connection_id={} url={}",
            opt_str(&part.name),
            opt_str(&part.connection_id),
            opt_str(&part.url)
        )),
        _ => None,
    }
}
