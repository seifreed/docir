//! Shared security scanning interface across formats.

use crate::error::ParseError;
use crate::odf::security::scan_odf_security;
use crate::parser::security::SecurityScanner as OoxmlSecurityScanner;
use crate::zip_handler::SecureZipReader;
use crate::ParserConfig;
use docir_core::ir::{Diagnostics, Document};
use docir_core::visitor::IrStore;
use std::io::{Read, Seek};

pub trait SecurityScanner {
    fn scan_ooxml<R: Read + Seek>(
        &self,
        config: &ParserConfig,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
    ) -> Result<(), ParseError>;

    fn scan_odf<R: Read + Seek>(
        &self,
        content_xml: Option<&str>,
        styles_xml: Option<&str>,
        settings_xml: Option<&str>,
        file_names: &[String],
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
        doc: &mut Document,
        diagnostics: &mut Diagnostics,
    );
}

pub struct DefaultSecurityScanner;

impl SecurityScanner for DefaultSecurityScanner {
    fn scan_ooxml<R: Read + Seek>(
        &self,
        config: &ParserConfig,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
    ) -> Result<(), ParseError> {
        let scanner = OoxmlSecurityScanner::new(config);
        scanner.scan_zip(zip, store)
    }

    fn scan_odf<R: Read + Seek>(
        &self,
        content_xml: Option<&str>,
        styles_xml: Option<&str>,
        settings_xml: Option<&str>,
        file_names: &[String],
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
        doc: &mut Document,
        diagnostics: &mut Diagnostics,
    ) {
        scan_odf_security(
            content_xml,
            styles_xml,
            settings_xml,
            file_names,
            zip,
            store,
            doc,
            diagnostics,
        );
    }
}
