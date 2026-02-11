//! Shared CLI helpers.

use anyhow::{bail, Result};
use docir_core::types::{DocumentFormat, NodeType};
use docir_parser::{DocumentParser, ParserConfig};
use serde::Serialize;
use std::fs::File;
use std::io::{self, Write};
use std::path::PathBuf;

pub fn parse_node_type(input: &str) -> Result<NodeType> {
    let upper = input.trim().to_ascii_uppercase();
    let ty = match upper.as_str() {
        "DOCUMENT" => NodeType::Document,
        "SECTION" => NodeType::Section,
        "PARAGRAPH" => NodeType::Paragraph,
        "RUN" => NodeType::Run,
        "TABLE" => NodeType::Table,
        "TABLEROW" | "TABLE_ROW" => NodeType::TableRow,
        "TABLECELL" | "TABLE_CELL" => NodeType::TableCell,
        "SLIDE" => NodeType::Slide,
        "SHAPE" => NodeType::Shape,
        "WORKSHEET" => NodeType::Worksheet,
        "CELL" => NodeType::Cell,
        "MACROPROJECT" | "MACRO_PROJECT" => NodeType::MacroProject,
        "MACROMODULE" | "MACRO_MODULE" => NodeType::MacroModule,
        "OLEOBJECT" | "OLE_OBJECT" => NodeType::OleObject,
        "EXTERNALREFERENCE" | "EXTERNAL_REFERENCE" => NodeType::ExternalReference,
        "ACTIVEXCONTROL" | "ACTIVEX_CONTROL" => NodeType::ActiveXControl,
        _ => bail!("Unknown node type: {input}"),
    };
    Ok(ty)
}

pub fn parse_doc_format(input: &str) -> Result<DocumentFormat> {
    let upper = input.trim().to_ascii_uppercase();
    let fmt = match upper.as_str() {
        "DOCX" | "WORD" | "WORDPROCESSING" => DocumentFormat::WordProcessing,
        "XLSX" | "EXCEL" | "SPREADSHEET" => DocumentFormat::Spreadsheet,
        "PPTX" | "PPT" | "POWERPOINT" | "PRESENTATION" => DocumentFormat::Presentation,
        "ODT" | "ODF" | "ODFTEXT" => DocumentFormat::OdfText,
        "ODS" | "ODFSPREADSHEET" => DocumentFormat::OdfSpreadsheet,
        "ODP" | "ODFPRESENTATION" => DocumentFormat::OdfPresentation,
        "HWP" => DocumentFormat::Hwp,
        "HWPX" => DocumentFormat::Hwpx,
        "RTF" => DocumentFormat::Rtf,
        _ => bail!("Unknown document format: {input}"),
    };
    Ok(fmt)
}

pub fn build_parser(config: &ParserConfig) -> DocumentParser {
    DocumentParser::with_config(config.clone())
}

pub fn write_json_output<T: Serialize>(
    value: &T,
    pretty: bool,
    output: Option<PathBuf>,
) -> Result<()> {
    let mut writer: Box<dyn Write> = match output {
        Some(path) => Box::new(File::create(path)?),
        None => Box::new(io::stdout()),
    };

    if pretty {
        serde_json::to_writer_pretty(&mut writer, value)?;
    } else {
        serde_json::to_writer(&mut writer, value)?;
    }

    writeln!(writer)?;
    Ok(())
}
