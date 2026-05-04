use crate::io_support::with_file_bytes_and_config;
use crate::{AppResult, ParserConfig};
use docir_parser::hwp::is_hwpx_mimetype;
use docir_parser::legacy_office::probe_legacy_office_format;
use docir_parser::ole::{is_ole_container, Cfb};
use docir_parser::ooxml::content_types::ContentTypes;
use docir_parser::zip_handler::SecureZipReader;
use serde::Serialize;
use std::io::Cursor;
use std::path::Path;

/// Analyst-facing result for lightweight format triage.
#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
pub struct FormatProbe {
    pub format: String,
    pub container: String,
    pub family: String,
    pub suggested_extension: String,
    pub confidence: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub signals: Vec<String>,
}

impl FormatProbe {
    fn new(
        format: &str,
        container: &str,
        family: &str,
        extension: &str,
        confidence: &str,
        signals: Vec<String>,
    ) -> Self {
        Self {
            format: format.to_string(),
            container: container.to_string(),
            family: family.to_string(),
            suggested_extension: extension.to_string(),
            confidence: confidence.to_string(),
            signals,
        }
    }
}

/// Probes an on-disk file without running the full parser pipeline.
pub fn probe_format_path<P: AsRef<Path>>(path: P, config: &ParserConfig) -> AppResult<FormatProbe> {
    with_file_bytes_and_config(path, config, |bytes, cfg| {
        Ok(probe_format_bytes(bytes, cfg))
    })
}

/// Probes raw bytes and returns a lightweight format classification.
pub fn probe_format_bytes(data: &[u8], config: &ParserConfig) -> FormatProbe {
    if is_rtf_signature(data) {
        return FormatProbe::new(
            "rtf",
            "rtf",
            "rich-text-document",
            "rtf",
            "high",
            vec!["rtf-signature".into()],
        );
    }

    if is_ole_container(data) {
        return probe_cfb(data);
    }

    if is_zip_signature(data) {
        return probe_zip(data, config);
    }

    if data.starts_with(b"%PDF-") {
        return FormatProbe::new(
            "pdf",
            "raw-binary",
            "pdf-document",
            "pdf",
            "high",
            vec!["pdf-signature".into()],
        );
    }

    if has_pe_signature(data) {
        return FormatProbe::new(
            "pe",
            "raw-binary",
            "portable-executable",
            "exe",
            "high",
            vec!["mz-signature".into()],
        );
    }

    if is_png(data) {
        return FormatProbe::new(
            "png",
            "raw-binary",
            "image",
            "png",
            "high",
            vec!["png-signature".into()],
        );
    }

    if is_jpeg(data) {
        return FormatProbe::new(
            "jpeg",
            "raw-binary",
            "image",
            "jpg",
            "high",
            vec!["jpeg-signature".into()],
        );
    }

    if is_gif(data) {
        return FormatProbe::new(
            "gif",
            "raw-binary",
            "image",
            "gif",
            "high",
            vec!["gif-signature".into()],
        );
    }

    if is_swf(data) {
        return FormatProbe::new(
            "swf",
            "raw-binary",
            "flash-object",
            "swf",
            "high",
            vec!["swf-signature".into()],
        );
    }

    FormatProbe::new(
        "unknown",
        "raw-binary",
        "unknown",
        "bin",
        "low",
        vec!["no-known-signature".into()],
    )
}

fn probe_cfb(data: &[u8]) -> FormatProbe {
    let mut signals = vec!["cfb-signature".to_string()];
    let Ok(cfb) = Cfb::parse(data.to_vec()) else {
        signals.push("cfb-open-failed".to_string());
        return FormatProbe::new("ole", "cfb-ole", "compound-file", "ole", "medium", signals);
    };

    if cfb.has_stream("FileHeader") {
        signals.push("stream:FileHeader".to_string());
        return FormatProbe::new(
            "hwp",
            "cfb-ole",
            "hangul-word-processor",
            "hwp",
            "high",
            signals,
        );
    }

    if cfb.has_stream("WordDocument") {
        signals.push("stream:WordDocument".to_string());
        return FormatProbe::new("doc", "cfb-ole", "word-processing", "doc", "high", signals);
    }

    if cfb.has_stream("Workbook") || cfb.has_stream("Book") {
        signals.push("stream:Workbook".to_string());
        return FormatProbe::new("xls", "cfb-ole", "spreadsheet", "xls", "high", signals);
    }

    if cfb.has_stream("PowerPoint Document") {
        signals.push("stream:PowerPoint Document".to_string());
        return FormatProbe::new("ppt", "cfb-ole", "presentation", "ppt", "high", signals);
    }

    if probe_legacy_office_format(&cfb).is_some() {
        signals.push("legacy-office-layout".to_string());
        return FormatProbe::new(
            "office-legacy",
            "cfb-ole",
            "office-legacy",
            "ole",
            "medium",
            signals,
        );
    }

    if cfb.has_stream("!_StringData") || cfb.has_stream("!_StringPool") {
        signals.push("stream:!_StringData".to_string());
        return FormatProbe::new(
            "msi",
            "cfb-ole",
            "installer-package",
            "msi",
            "medium",
            signals,
        );
    }

    FormatProbe::new("ole", "cfb-ole", "compound-file", "ole", "medium", signals)
}

fn probe_zip(data: &[u8], config: &ParserConfig) -> FormatProbe {
    let mut signals = vec!["zip-signature".to_string()];
    let Ok(mut zip) = SecureZipReader::new(Cursor::new(data), config.zip_config.clone()) else {
        signals.push("zip-open-failed".to_string());
        return FormatProbe::new("zip", "zip", "archive", "zip", "medium", signals);
    };

    if zip.contains("[Content_Types].xml") {
        signals.push("zip:[Content_Types].xml".to_string());
        if let Ok(xml) = zip.read_file_string("[Content_Types].xml") {
            if let Ok(content_types) = ContentTypes::parse(&xml) {
                return probe_ooxml_content_types(&content_types, &mut signals);
            }
            signals.push("content-types-parse-failed".to_string());
        }
        return FormatProbe::new(
            "ooxml",
            "zip-ooxml",
            "office-openxml",
            "zip",
            "medium",
            signals,
        );
    }

    if zip.contains("mimetype") {
        signals.push("zip:mimetype".to_string());
        if let Ok(mimetype) = zip.read_file_string("mimetype") {
            let lower = mimetype.trim().to_ascii_lowercase();
            if let Some(probe) = probe_odf_mimetype(&lower, &mut signals) {
                return probe;
            }
        }
    }

    FormatProbe::new("zip", "zip", "archive", "zip", "medium", signals)
}

fn probe_odf_mimetype(lower: &str, signals: &mut Vec<String>) -> Option<FormatProbe> {
    if is_hwpx_mimetype(lower) {
        signals.push(format!("mimetype:{lower}"));
        return Some(FormatProbe::new(
            "hwpx",
            "zip-hwpx",
            "hangul-word-processor",
            "hwpx",
            "high",
            signals.clone(),
        ));
    }
    if lower.contains("opendocument.text") {
        signals.push(format!("mimetype:{lower}"));
        return Some(FormatProbe::new(
            "odt",
            "zip-odf",
            "odf-text",
            "odt",
            "high",
            signals.clone(),
        ));
    }
    if lower.contains("opendocument.spreadsheet") {
        signals.push(format!("mimetype:{lower}"));
        return Some(FormatProbe::new(
            "ods",
            "zip-odf",
            "odf-spreadsheet",
            "ods",
            "high",
            signals.clone(),
        ));
    }
    if lower.contains("opendocument.presentation") {
        signals.push(format!("mimetype:{lower}"));
        return Some(FormatProbe::new(
            "odp",
            "zip-odf",
            "odf-presentation",
            "odp",
            "high",
            signals.clone(),
        ));
    }
    None
}

fn probe_ooxml_content_types(
    content_types: &ContentTypes,
    signals: &mut Vec<String>,
) -> FormatProbe {
    let macro_enabled = content_types.is_macro_enabled();
    let binary_workbook = content_types
        .overrides
        .values()
        .any(|ct| ct.contains("sheet.binary"));
    if macro_enabled {
        signals.push("ooxml:macro-enabled".to_string());
    }
    if binary_workbook {
        signals.push("ooxml:binary-workbook".to_string());
    }

    match content_types.detect_format() {
        Some(docir_core::DocumentFormat::WordProcessing) => {
            let ext = if macro_enabled { "docm" } else { "docx" };
            FormatProbe::new(
                ext,
                "zip-ooxml",
                "word-processing",
                ext,
                "high",
                signals.clone(),
            )
        }
        Some(docir_core::DocumentFormat::Spreadsheet) => {
            let ext = if binary_workbook {
                "xlsb"
            } else if macro_enabled {
                "xlsm"
            } else {
                "xlsx"
            };
            FormatProbe::new(
                ext,
                "zip-ooxml",
                "spreadsheet",
                ext,
                "high",
                signals.clone(),
            )
        }
        Some(docir_core::DocumentFormat::Presentation) => {
            let ext = if macro_enabled { "pptm" } else { "pptx" };
            FormatProbe::new(
                ext,
                "zip-ooxml",
                "presentation",
                ext,
                "high",
                signals.clone(),
            )
        }
        _ => FormatProbe::new(
            "ooxml",
            "zip-ooxml",
            "office-openxml",
            "zip",
            "medium",
            signals.clone(),
        ),
    }
}

fn is_zip_signature(data: &[u8]) -> bool {
    data.starts_with(b"PK\x03\x04")
        || data.starts_with(b"PK\x05\x06")
        || data.starts_with(b"PK\x07\x08")
}

fn is_rtf_signature(data: &[u8]) -> bool {
    data.starts_with(b"{\\rtf")
}

fn has_pe_signature(data: &[u8]) -> bool {
    data.starts_with(b"MZ")
}

fn is_png(data: &[u8]) -> bool {
    data.starts_with(&[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A])
}

fn is_jpeg(data: &[u8]) -> bool {
    data.starts_with(&[0xFF, 0xD8, 0xFF])
}

fn is_gif(data: &[u8]) -> bool {
    data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a")
}

fn is_swf(data: &[u8]) -> bool {
    data.starts_with(b"FWS") || data.starts_with(b"CWS") || data.starts_with(b"ZWS")
}

#[cfg(test)]
mod tests {
    use super::{probe_format_bytes, FormatProbe};
    use crate::{test_support::build_test_cfb, ParserConfig};
    use std::io::Write;
    use zip::write::FileOptions;

    fn write_docx() -> Vec<u8> {
        let cursor = std::io::Cursor::new(Vec::<u8>::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = FileOptions::<()>::default();
        let content_types = r#"
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Override PartName="/word/document.xml"
                ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>"#;
        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.trim().as_bytes()).unwrap();
        zip.add_directory("word/", options).unwrap();
        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(b"<w:document/>").unwrap();
        zip.finish().unwrap().into_inner()
    }

    fn write_odt() -> Vec<u8> {
        let cursor = std::io::Cursor::new(Vec::<u8>::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let stored =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(b"application/vnd.oasis.opendocument.text")
            .unwrap();
        zip.finish().unwrap().into_inner()
    }

    fn write_generic_zip() -> Vec<u8> {
        let cursor = std::io::Cursor::new(Vec::<u8>::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = FileOptions::<()>::default();
        zip.start_file("notes.txt", options).unwrap();
        zip.write_all(b"hello").unwrap();
        zip.finish().unwrap().into_inner()
    }

    fn assert_probe(
        probe: FormatProbe,
        expected_format: &str,
        expected_container: &str,
        expected_extension: &str,
    ) {
        assert_eq!(probe.format, expected_format);
        assert_eq!(probe.container, expected_container);
        assert_eq!(probe.suggested_extension, expected_extension);
    }

    #[test]
    fn probe_format_identifies_docx() {
        let probe = probe_format_bytes(&write_docx(), &ParserConfig::default());
        assert_probe(probe.clone(), "docx", "zip-ooxml", "docx");
        assert_eq!(probe.family, "word-processing");
        assert!(probe.signals.iter().any(|s| s == "zip:[Content_Types].xml"));
    }

    #[test]
    fn probe_format_identifies_odt() {
        let probe = probe_format_bytes(&write_odt(), &ParserConfig::default());
        assert_probe(probe, "odt", "zip-odf", "odt");
    }

    #[test]
    fn probe_format_identifies_legacy_doc() {
        let probe = probe_format_bytes(
            &build_test_cfb(&[("WordDocument", b"doc")]),
            &ParserConfig::default(),
        );
        assert_probe(probe, "doc", "cfb-ole", "doc");
    }

    #[test]
    fn probe_format_identifies_pdf() {
        let probe = probe_format_bytes(b"%PDF-1.7\n", &ParserConfig::default());
        assert_probe(probe, "pdf", "raw-binary", "pdf");
    }

    #[test]
    fn probe_format_identifies_generic_zip() {
        let probe = probe_format_bytes(&write_generic_zip(), &ParserConfig::default());
        assert_probe(probe, "zip", "zip", "zip");
    }
}
