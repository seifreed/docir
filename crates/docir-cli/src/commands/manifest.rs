//! Emit the canonical Phase 0 artifact manifest.

use anyhow::Result;
use docir_app::{ExportDocumentRef, ParserConfig, Phase0ArtifactManifestExport};
use std::fs;
use std::path::PathBuf;

use crate::cli::PrettyOutputOpts;
use crate::commands::util::{build_app, source_format_label, write_json_output};

/// Public API entrypoint: run.
pub fn run(input: PathBuf, opts: PrettyOutputOpts, parser_config: &ParserConfig) -> Result<()> {
    let PrettyOutputOpts { pretty, output } = opts;
    let app = build_app(parser_config);
    let source_bytes = fs::read(&input)?;
    let parsed = app.parse_bytes(&source_bytes)?;
    let inventory = app.build_inventory_with_bytes(&parsed, &source_bytes);
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
    use crate::cli::PrettyOutputOpts;
    use crate::test_support;
    use docir_app::{test_support::build_test_cfb, ParserConfig};
    use std::fs;

    #[test]
    fn manifest_run_writes_schema_like_json() {
        let output = test_support::temp_file("json", "json");
        run(
            test_support::fixture("minimal.docx"),
            PrettyOutputOpts {
                pretty: true,
                output: Some(output.clone()),
            },
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
        let input = test_support::temp_file("legacy_vba", "doc");
        let output = test_support::temp_file("legacy_vba", "json");
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
            PrettyOutputOpts {
                pretty: true,
                output: Some(output.clone()),
            },
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

        let input = test_support::temp_file("legacy_ole_hash", "doc");
        let output = test_support::temp_file("legacy_ole_hash", "json");
        fs::write(
            &input,
            build_test_cfb(&[("WordDocument", b"doc"), ("Ole10Native", &ole10)]),
        )
        .expect("fixture");

        run(
            input.clone(),
            PrettyOutputOpts {
                pretty: true,
                output: Some(output.clone()),
            },
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
        let input = test_support::temp_file("docx_media", "docx");
        let output = test_support::temp_file("docx_media", "json");
        test_support::write_docx_with_media(&input);

        run(
            input.clone(),
            PrettyOutputOpts {
                pretty: true,
                output: Some(output.clone()),
            },
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
        let input = test_support::temp_file("docx_media_no_hashes", "docx");
        let output = test_support::temp_file("docx_media_no_hashes", "json");
        test_support::write_docx_with_media(&input);

        let config = ParserConfig {
            compute_hashes: false,
            ..ParserConfig::default()
        };
        run(
            input.clone(),
            PrettyOutputOpts {
                pretty: true,
                output: Some(output.clone()),
            },
            &config,
        )
        .expect("manifest docx no hashes");

        let text = fs::read_to_string(&output).expect("manifest output");
        assert!(text.contains("\"embedded-file\""));
        assert!(text.contains("\"media_type\": \"image/png\""));
        assert!(!text.contains("\"sha256\":"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_file(output);
    }
}
