//! Emit the canonical Phase 0 artifact manifest.

use anyhow::Result;
use docir_app::{ExportDocumentRef, ParserConfig, Phase0ArtifactManifestExport};
use std::fs;
use std::path::PathBuf;

use crate::commands::util::{build_app, source_format_label, write_json_output};

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    pretty: bool,
    output: Option<PathBuf>,
    parser_config: &ParserConfig,
) -> Result<()> {
    let app = build_app(parser_config);
    let parsed = app.parse_file(&input)?;
    let bytes = fs::read(&input)?;
    let inventory = app.build_inventory_with_bytes(&parsed, &bytes);
    let export = Phase0ArtifactManifestExport::from_inventory(
        &inventory,
        ExportDocumentRef::new(
            input.display().to_string(),
            source_format_label(&input, parsed.format().extension()),
            Some(input.display().to_string()),
        ),
    );
    write_json_output(&export, pretty, output)
}

#[cfg(test)]
mod tests {
    use super::run;
    use docir_app::{test_support::build_test_cfb, ParserConfig};
    use std::fs;
    use std::io::Write;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};
    use zip::write::FileOptions;

    fn fixture(name: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../fixtures/ooxml")
            .join(name)
    }

    fn temp_file(name: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_manifest_{name}_{nanos}.json"))
    }

    fn temp_input(name: &str, ext: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_manifest_{name}_{nanos}.{ext}"))
    }

    fn write_docx_with_media(path: &PathBuf) {
        let file = fs::File::create(path).expect("create docx");
        let mut zip = zip::ZipWriter::new(file);
        let options = FileOptions::<()>::default();

        let content_types = r#"
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="png" ContentType="image/png"/>
              <Override PartName="/word/document.xml"
                ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>"#;
        let rels = r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
            </Relationships>"#;
        let document = r#"
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
              </w:body>
            </w:document>"#;

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.trim().as_bytes()).unwrap();
        zip.add_directory("_rels/", options).unwrap();
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(rels.trim().as_bytes()).unwrap();
        zip.add_directory("word/", options).unwrap();
        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(document.trim().as_bytes()).unwrap();
        zip.add_directory("word/media/", options).unwrap();
        zip.start_file("word/media/image1.png", options).unwrap();
        zip.write_all(b"PNG").unwrap();
        zip.finish().unwrap();
    }

    #[test]
    fn manifest_run_writes_schema_like_json() {
        let output = temp_file("json");
        run(
            fixture("minimal.docx"),
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("manifest json");
        let text = fs::read_to_string(&output).expect("manifest output");
        assert!(text.contains("\"schema_version\""));
        assert!(text.contains("\"artifacts\""));
        let _ = fs::remove_file(output);
    }

    #[test]
    fn manifest_run_writes_vba_inventory_details_for_legacy_doc() {
        let input = temp_input("legacy_vba", "doc");
        let output = temp_file("legacy_vba");
        let bytes = build_test_cfb(&[
            ("WordDocument", b"doc"),
            (
                "VBA/PROJECT",
                br#"Name="LegacyProject"
Module=Module1/Module1
DPB="AAAA"
Reference=*\G{000204EF-0000-0000-C000-000000000046}#2.0#0#..\stdole2.tlb#OLE Automation
"#,
            ),
            ("VBA/Module1", b"Sub AutoOpen()\nEnd Sub\n"),
        ]);
        fs::write(&input, bytes).expect("fixture");

        run(
            input.clone(),
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("manifest legacy vba");

        let text = fs::read_to_string(&output).expect("manifest output");
        assert!(text.contains("\"vba-project\""));
        assert!(text.contains("\"vba-module-source\""));
        assert!(text.contains(&input.display().to_string()));
        assert!(text.contains("name=LegacyProject"));
        assert!(text.contains("protected=true"));
        assert!(text.contains("autoexec=AutoOpen"));
        assert!(text.contains("type=standard"));
        assert!(text.contains("procedures=AutoOpen"));
        assert!(text.contains(
            "\"sha256\": \"79f3ab2c7767923d61d5feff3eb9348bec5268a6469a95cac521dd858d3540f1\""
        ));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn manifest_run_preserves_ole_sha256_for_legacy_doc() {
        let mut ole10 = Vec::new();
        ole10.extend_from_slice(&64u32.to_le_bytes());
        ole10.extend_from_slice(&2u16.to_le_bytes());
        ole10.extend_from_slice(b"dropper.exe\0");
        ole10.extend_from_slice(b"C:\\src\\dropper.exe\0");
        ole10.extend_from_slice(&0u32.to_le_bytes());
        ole10.extend_from_slice(&0u32.to_le_bytes());
        ole10.extend_from_slice(b"C:\\temp\\dropper.exe\0");
        ole10.extend_from_slice(&4u32.to_le_bytes());
        ole10.extend_from_slice(b"MZ!!");

        let input = temp_input("legacy_ole_hash", "doc");
        let output = temp_file("legacy_ole_hash");
        fs::write(
            &input,
            build_test_cfb(&[("WordDocument", b"doc"), ("Ole10Native", &ole10)]),
        )
        .expect("fixture");

        run(
            input.clone(),
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("manifest legacy ole");

        let text = fs::read_to_string(&output).expect("manifest output");
        assert!(text.contains("\"ole-object\""));
        assert!(text.contains(
            "\"sha256\": \"929f8a8a8dc67c682a47beaddad8b616a2913f90d6e5c51c32c25f10e1242eec\""
        ));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn manifest_run_preserves_media_asset_type_and_sha256_for_docx() {
        let input = temp_input("docx_media", "docx");
        let output = temp_file("docx_media");
        write_docx_with_media(&input);

        run(
            input.clone(),
            true,
            Some(output.clone()),
            &ParserConfig::default(),
        )
        .expect("manifest docx media");

        let text = fs::read_to_string(&output).expect("manifest output");
        assert!(text.contains("\"embedded-file\""));
        assert!(text.contains("\"path\": \"word/media/image1.png\""));
        assert!(text.contains("\"media_type\": \"image/png\""));
        assert!(text.contains(
            "\"sha256\": \"796120837694d3f3f29259cfeb25091698c2a0aa87873658d840b4993ee889b3\""
        ));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn manifest_run_omits_sha256_when_hashes_are_disabled() {
        let input = temp_input("docx_media_no_hashes", "docx");
        let output = temp_file("docx_media_no_hashes");
        write_docx_with_media(&input);

        let config = ParserConfig {
            compute_hashes: false,
            ..ParserConfig::default()
        };
        run(input.clone(), true, Some(output.clone()), &config).expect("manifest docx no hashes");

        let text = fs::read_to_string(&output).expect("manifest output");
        assert!(text.contains("\"embedded-file\""));
        assert!(text.contains("\"media_type\": \"image/png\""));
        assert!(!text.contains("\"sha256\":"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }
}
