#[path = "builder_hwp.rs"]
mod builder_hwp;
#[path = "builder_hwpx.rs"]
mod builder_hwpx;
use super::{scan_hwpx_security, HwpxParser};
use crate::diagnostics::push_warning;
use crate::error::ParseError;
use crate::input::enforce_input_size;
use crate::input::read_all_with_limit;
use crate::ole::Cfb;
use crate::parse_utils::{finalize_and_normalize, init_store_and_document};
use crate::parser::{
    run_parser_pipeline, NormalizeStage, ParseStage, ParsedDocument, PostprocessStage,
};
use crate::zip_handler::SecureZipReader;
pub(crate) use builder_hwp::HwpHeaderContext;
#[cfg(test)]
use docir_core::ir::IRNode;
#[cfg(test)]
use docir_core::ir::Paragraph;
use docir_core::types::DocumentFormat;
#[cfg(test)]
use docir_core::types::NodeId;
#[cfg(test)]
use docir_core::visitor::IrStore;
use std::io::{Read, Seek};

use super::attach_diagnostics_if_any;
use super::helpers::build_hwp_diagnostics;
use super::io::dump_hwp_streams;
use super::HwpParser;

impl HwpParser {
    /// Public API entrypoint: parse_reader.
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        run_parser_pipeline(self, reader)
    }
}

impl ParseStage for HwpParser {
    fn parse_stage<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        let data = read_all_with_limit(reader, self.config.max_input_size)?;

        let cfb = Cfb::parse(data)?;
        let (mut store, mut doc) = init_store_and_document(DocumentFormat::Hwp);

        let mut stream_names = cfb.list_streams();
        stream_names.sort();
        let mut diagnostics = build_hwp_diagnostics(DocumentFormat::Hwp, &stream_names);

        let header_ctx = self.build_header_context(&cfb, &mut diagnostics)?;

        if self.config.hwp.dump_streams {
            dump_hwp_streams(&cfb, &stream_names, &header_ctx, &mut diagnostics);
        }

        let shared_parts = self.collect_stream_parts(&cfb, &stream_names, &mut store);
        let docinfo_section_count =
            self.read_docinfo_section_count(&cfb, &header_ctx, &mut diagnostics)?;
        let sections = self.parse_sections(
            &cfb,
            &stream_names,
            &header_ctx,
            &mut store,
            &mut diagnostics,
        )?;

        if let Some(expected) = docinfo_section_count {
            if expected as usize != sections.len() {
                push_warning(
                    &mut diagnostics,
                    "HWP_SECTION_MISMATCH",
                    format!(
                        "section count mismatch: docinfo={} parsed={}",
                        expected,
                        sections.len()
                    ),
                    Some("DocInfo"),
                );
            }
        }

        self.parse_default_script(&cfb, &mut store);
        self.scan_external_references(
            &cfb,
            &stream_names,
            &header_ctx,
            &mut store,
            &mut diagnostics,
        );
        self.collect_ole_objects(&cfb, &stream_names, &mut store);

        doc.shared_parts = shared_parts;
        doc.content = sections;

        attach_diagnostics_if_any(&mut store, &mut doc, diagnostics);

        Ok(finalize_and_normalize(DocumentFormat::Hwp, store, doc))
    }
}

impl NormalizeStage for HwpParser {}

impl PostprocessStage for HwpParser {}

impl HwpxParser {
    /// Public API entrypoint: parse_reader.
    pub fn parse_reader<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError> {
        run_parser_pipeline(self, reader)
    }
}

impl ParseStage for HwpxParser {
    fn parse_stage<R: Read + Seek>(&self, mut reader: R) -> Result<ParsedDocument, ParseError> {
        enforce_input_size(&mut reader, self.config.max_input_size)?;
        let mut zip = SecureZipReader::new(reader, self.config.zip_config.clone())?;

        let (mut store, mut doc) = init_store_and_document(DocumentFormat::Hwpx);

        let mut file_names: Vec<String> = zip.file_names().map(|s| s.to_string()).collect();
        file_names.sort();

        let part_data = self.collect_hwpx_parts(&mut zip, &file_names, &mut store);
        doc.shared_parts = part_data.shared_parts;
        doc.shared_parts.extend(part_data.media_assets);

        doc.content = self.parse_hwpx_primary_sections(
            &mut zip,
            &file_names,
            &part_data.media_lookup,
            &mut store,
            &mut doc,
        )?;
        let grouped_sections = self.parse_hwpx_section_groups(
            &mut zip,
            &file_names,
            &part_data.media_lookup,
            &mut store,
            &mut doc,
        )?;
        self.attach_hwpx_headers_and_footers(
            &mut store,
            &doc.content,
            &grouped_sections.header_ids,
            &grouped_sections.footer_ids,
        );
        doc.shared_parts.extend(grouped_sections.master_ids);
        self.parse_hwpx_styles_part(&mut zip, &mut store, &mut doc);

        scan_hwpx_security(&file_names, &mut zip, &mut store, &mut doc);

        let diagnostics = build_hwp_diagnostics(DocumentFormat::Hwpx, &file_names);
        attach_diagnostics_if_any(&mut store, &mut doc, diagnostics);

        Ok(finalize_and_normalize(DocumentFormat::Hwpx, store, doc))
    }
}

impl NormalizeStage for HwpxParser {}

impl PostprocessStage for HwpxParser {}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::{Footer, Header, Section};

    #[test]
    fn attach_hwpx_headers_and_footers_updates_only_section_nodes() {
        let parser = HwpxParser::new();
        let mut store = IrStore::new();

        let mut section_one = Section::new();
        let section_one_id = section_one.id;
        section_one.headers.push(NodeId::new());
        store.insert(IRNode::Section(section_one));

        let section_two = Section::new();
        let section_two_id = section_two.id;
        store.insert(IRNode::Section(section_two));

        let paragraph = Paragraph::new();
        let paragraph_id = paragraph.id;
        store.insert(IRNode::Paragraph(paragraph));

        let header = Header::new();
        let header_id = header.id;
        store.insert(IRNode::Header(header));

        let footer = Footer::new();
        let footer_id = footer.id;
        store.insert(IRNode::Footer(footer));

        parser.attach_hwpx_headers_and_footers(
            &mut store,
            &[section_one_id, paragraph_id, section_two_id],
            &[header_id],
            &[footer_id],
        );

        let Some(IRNode::Section(section_one)) = store.get(section_one_id) else {
            panic!("expected first section");
        };
        assert_eq!(section_one.headers.len(), 2);
        assert_eq!(section_one.headers[1], header_id);
        assert_eq!(section_one.footers, vec![footer_id]);

        let Some(IRNode::Section(section_two)) = store.get(section_two_id) else {
            panic!("expected second section");
        };
        assert_eq!(section_two.headers, vec![header_id]);
        assert_eq!(section_two.footers, vec![footer_id]);
    }
}
