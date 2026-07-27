use std::io::{Read, Seek};

use super::{
    Document, DocumentFormat, IRNode, IrStore, OdfManifestEntry, OdfParser, ParseError,
    SecureZipReader, parse_manifest, parse_styles,
};

type StylesSettingsSignatures = (Option<String>, Option<String>, Option<String>);

impl OdfParser {
    pub(super) fn load_mimetype_and_manifest<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
    ) -> Result<(DocumentFormat, Vec<OdfManifestEntry>), ParseError> {
        let mimetype = zip
            .read_file_string("mimetype")
            .map(|s| s.trim().to_string())
            .map_err(|_| ParseError::UnsupportedFormat("Missing ODF mimetype".to_string()))?;

        let format = detect_odf_format(&mimetype).ok_or_else(|| {
            ParseError::UnsupportedFormat(format!("Unsupported ODF mimetype: {mimetype}"))
        })?;

        let manifest_entries = if zip.contains("META-INF/manifest.xml") {
            let manifest_xml = zip.read_file_string("META-INF/manifest.xml")?;
            parse_manifest(&manifest_xml)?
        } else {
            Vec::new()
        };

        Ok((format, manifest_entries))
    }

    pub(super) fn load_styles_settings_signatures<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        store: &mut IrStore,
        doc: &mut Document,
    ) -> Result<StylesSettingsSignatures, ParseError> {
        let mut styles_xml: Option<String> = None;
        if zip.contains("styles.xml") {
            let xml = zip.read_file_string("styles.xml")?;
            if let Some(styles) = parse_styles(&xml, "styles.xml")? {
                let style_id = styles.id;
                store.insert(IRNode::StyleSet(styles));
                doc.styles = Some(style_id);
            }
            styles_xml = Some(xml);
        }

        let settings_xml = if zip.contains("settings.xml") {
            Some(zip.read_file_string("settings.xml")?)
        } else {
            None
        };

        let signatures_xml = if zip.contains("META-INF/documentsignatures.xml") {
            Some(zip.read_file_string("META-INF/documentsignatures.xml")?)
        } else {
            None
        };

        Ok((styles_xml, settings_xml, signatures_xml))
    }
}

pub(super) fn detect_odf_format(mimetype: &str) -> Option<DocumentFormat> {
    let lower = mimetype.to_ascii_lowercase();
    if lower.contains("opendocument.text") || lower.contains("vnd.sun.xml.writer") {
        Some(DocumentFormat::OdfText)
    } else if lower.contains("opendocument.spreadsheet") || lower.contains("vnd.sun.xml.calc") {
        Some(DocumentFormat::OdfSpreadsheet)
    } else if lower.contains("opendocument.presentation") || lower.contains("vnd.sun.xml.impress") {
        Some(DocumentFormat::OdfPresentation)
    } else {
        None
    }
}
