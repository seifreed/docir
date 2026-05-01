//! Shared security scanning interface across formats.

use crate::error::ParseError;
use crate::odf::security::scan_odf_security;
use crate::parser::OoxmlSecurityScanner;
use crate::zip_handler::PackageReader;
use crate::ParserConfig;
use docir_core::ir::{Diagnostics, Document};
use docir_core::visitor::IrStore;

pub struct OdfXmlInputs<'a> {
    pub content_xml: Option<&'a str>,
    pub styles_xml: Option<&'a str>,
    pub settings_xml: Option<&'a str>,
}

pub trait SecurityScanner {
    fn scan_ooxml(
        &self,
        config: &ParserConfig,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
    ) -> Result<(), ParseError>;

    fn scan_odf(
        &self,
        xml: OdfXmlInputs<'_>,
        file_names: &[String],
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        doc: &mut Document,
        diagnostics: &mut Diagnostics,
    );
}

pub struct DefaultSecurityScanner;

impl SecurityScanner for DefaultSecurityScanner {
    fn scan_ooxml(
        &self,
        config: &ParserConfig,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let scanner = OoxmlSecurityScanner::new(config);
        scanner.scan_zip(zip, store)
    }

    fn scan_odf(
        &self,
        xml: OdfXmlInputs<'_>,
        file_names: &[String],
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        doc: &mut Document,
        diagnostics: &mut Diagnostics,
    ) {
        scan_odf_security(xml, file_names, zip, store, doc, diagnostics);
    }
}
