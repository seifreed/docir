//! Extract VBA modules and write a manifest to disk.

use anyhow::{bail, Context, Result};
use docir_app::{
    ExportDocumentRef, ParserConfig, Phase0ArtifactManifestExport, VbaRecognitionReport,
};
use docir_core::{ExtractedArtifact, ExtractedArtifactKind, ExtractionManifest, ExtractionWarning};
use std::fs;
use std::path::{Path, PathBuf};

use crate::commands::util::build_app_and_parse;
use crate::commands::util::source_format_label;

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    out_dir: PathBuf,
    overwrite: bool,
    best_effort: bool,
    parser_config: &ParserConfig,
) -> Result<()> {
    let mut parse_config = parser_config.clone();
    parse_config.extract_macro_source = true;
    let (app, parsed) = build_app_and_parse(&input, &parse_config)?;
    let report = app.build_vba_recognition(&parsed, true);

    prepare_output_dir(&out_dir, overwrite)?;
    let manifest = write_vba_bundle(&report, &input, &out_dir, best_effort)?;
    let manifest_path = out_dir.join("manifest.json");
    let export = Phase0ArtifactManifestExport::from_manifest(
        &manifest,
        ExportDocumentRef::new(
            input.display().to_string(),
            source_format_label(&input, parsed.format().extension()),
            Some(input.display().to_string()),
        ),
    );
    let manifest_json = serde_json::to_string_pretty(&export)?;
    fs::write(&manifest_path, manifest_json)
        .with_context(|| format!("Failed to write {}", manifest_path.display()))?;
    Ok(())
}

fn prepare_output_dir(out_dir: &Path, overwrite: bool) -> Result<()> {
    if out_dir.exists() {
        if !overwrite {
            bail!(
                "Output directory {} already exists; pass --overwrite to reuse it",
                out_dir.display()
            );
        }
    } else {
        fs::create_dir_all(out_dir)
            .with_context(|| format!("Failed to create {}", out_dir.display()))?;
    }
    Ok(())
}

fn write_vba_bundle(
    report: &VbaRecognitionReport,
    input: &Path,
    out_dir: &Path,
    best_effort: bool,
) -> Result<ExtractionManifest> {
    let mut manifest = ExtractionManifest::new();
    manifest.source_document = Some(input.display().to_string());

    for project in &report.projects {
        let mut project_artifact =
            ExtractedArtifact::new(project.node_id.clone(), ExtractedArtifactKind::MacroProject);
        project_artifact.node_id = None;
        project_artifact.source_path = project.container_path.clone();
        project_artifact.suggested_name = project.project_name.clone();
        project_artifact.size_bytes = Some(project.modules.len() as u64);
        manifest.artifacts.push(project_artifact);

        for module in &project.modules {
            let Some(source_text) = module.source_text.as_ref() else {
                if best_effort {
                    let mut artifact = ExtractedArtifact::new(
                        module.node_id.clone(),
                        ExtractedArtifactKind::MacroModule,
                    );
                    artifact.source_path = module.stream_path.clone();
                    artifact.suggested_name = Some(module.name.clone());
                    artifact.errors = module.extraction_errors.clone();
                    manifest.artifacts.push(artifact);
                    continue;
                }
                bail!(
                    "Module {} has no extracted source; rerun with --best-effort to keep partial bundles",
                    module.name
                );
            };

            let ext = module_extension(&module.kind);
            let file_name = format!("{}_{}.{}", sanitize_name(&module.name), module.node_id, ext);
            let output_path = out_dir.join(&file_name);
            fs::write(&output_path, source_text)
                .with_context(|| format!("Failed to write {}", output_path.display()))?;

            let mut artifact =
                ExtractedArtifact::new(module.node_id.clone(), ExtractedArtifactKind::MacroModule);
            artifact.source_path = module.stream_path.clone();
            artifact.suggested_name = Some(file_name.clone());
            artifact.encoding = module.source_encoding.clone();
            artifact.sha256 = module.source_hash.clone();
            artifact.size_bytes = module.decompressed_size;
            artifact.output_path = Some(file_name);
            artifact.errors = module.extraction_errors.clone();
            manifest.artifacts.push(artifact);
        }
    }

    if report.projects.is_empty() {
        manifest.warnings.push(ExtractionWarning::new(
            "NO_VBA",
            "No VBA project recognized",
        ));
    }

    Ok(manifest)
}

fn module_extension(kind: &str) -> &'static str {
    match kind {
        "class" => "cls",
        "document" => "cls",
        "form" => "frm",
        _ => "bas",
    }
}

fn sanitize_name(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    for ch in input.chars() {
        if ch.is_ascii_alphanumeric() || ch == '-' || ch == '_' {
            out.push(ch);
        } else {
            out.push('_');
        }
    }
    if out.is_empty() {
        "module".to_string()
    } else {
        out
    }
}

#[cfg(test)]
mod tests {
    use super::run;
    use docir_app::{test_support::build_test_cfb, ParserConfig};
    use std::fs;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn fixture(name: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../fixtures/ooxml")
            .join(name)
    }

    fn temp_dir(name: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_extract_vba_{name}_{nanos}"))
    }

    #[test]
    fn extract_vba_without_macros_still_writes_manifest_in_best_effort_mode() {
        let out = temp_dir("manifest");
        run(
            fixture("minimal.docx"),
            out.clone(),
            false,
            true,
            &ParserConfig::default(),
        )
        .expect("best effort manifest");
        assert!(out.join("manifest.json").exists());
        let _ = std::fs::remove_dir_all(out);
    }

    #[test]
    fn extract_vba_legacy_doc_writes_manifest_and_module_files() {
        let input = temp_file("legacy_vba", "doc");
        fs::write(
            &input,
            build_test_cfb(&[
                ("WordDocument", b"doc"),
                (
                    "VBA/PROJECT",
                    br#"Name="LegacyProject"
Module=Module1/Module1
Document=ThisDocument/&H00000000
"#,
                ),
                ("VBA/Module1", b"Sub AutoOpen()\nEnd Sub\n"),
                ("VBA/ThisDocument", b"Sub Document_Open()\nEnd Sub\n"),
            ]),
        )
        .expect("write legacy vba doc");
        let out = temp_dir("legacy_vba");

        run(
            input.clone(),
            out.clone(),
            false,
            false,
            &ParserConfig::default(),
        )
        .expect("extract legacy vba");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"vba-project\""));
        assert!(manifest_text.contains("\"vba-module-source\""));

        let mut module_files = 0usize;
        for entry in fs::read_dir(&out).expect("read out dir") {
            let path = entry.expect("entry").path();
            if path.extension().and_then(|ext| ext.to_str()) == Some("bas")
                || path.extension().and_then(|ext| ext.to_str()) == Some("cls")
            {
                module_files += 1;
            }
        }
        assert!(module_files >= 2);

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_vba_legacy_xls_writes_manifest_and_module_files() {
        let input = temp_file("legacy_vba_xls", "xls");
        fs::write(
            &input,
            build_test_cfb(&[
                ("Workbook", b"wb"),
                (
                    "VBA/PROJECT",
                    br#"Name="WorkbookMacros"
Module=Module1/Module1
Document=ThisWorkbook/&H00000000
"#,
                ),
                ("VBA/Module1", b"Sub Auto_Open()\nEnd Sub\n"),
                (
                    "VBA/ThisWorkbook",
                    b"Private Sub Workbook_Open()\nEnd Sub\n",
                ),
            ]),
        )
        .expect("write legacy vba xls");
        let out = temp_dir("legacy_vba_xls");

        run(
            input.clone(),
            out.clone(),
            false,
            false,
            &ParserConfig::default(),
        )
        .expect("extract legacy xls vba");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"format\": \"xls\""));
        assert!(manifest_text.contains("\"vba-project\""));
        assert!(manifest_text.contains("\"vba-module-source\""));

        let mut module_files = 0usize;
        for entry in fs::read_dir(&out).expect("read out dir") {
            let path = entry.expect("entry").path();
            if path.extension().and_then(|ext| ext.to_str()) == Some("bas")
                || path.extension().and_then(|ext| ext.to_str()) == Some("cls")
            {
                module_files += 1;
            }
        }
        assert!(module_files >= 2);

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_vba_legacy_ppt_writes_manifest_and_module_files() {
        let input = temp_file("legacy_vba_ppt", "ppt");
        fs::write(
            &input,
            build_test_cfb(&[
                ("PowerPoint Document", b"ppt"),
                ("Current User", b"user"),
                (
                    "VBA/PROJECT",
                    br#"Name="PresentationMacros"
Module=Module1/Module1
Document=ThisPresentation/&H00000000
"#,
                ),
                ("VBA/Module1", b"Sub Auto_Open()\nEnd Sub\n"),
                (
                    "VBA/ThisPresentation",
                    b"Private Sub Presentation_Open()\nEnd Sub\n",
                ),
            ]),
        )
        .expect("write legacy vba ppt");
        let out = temp_dir("legacy_vba_ppt");

        run(
            input.clone(),
            out.clone(),
            false,
            false,
            &ParserConfig::default(),
        )
        .expect("extract legacy ppt vba");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"format\": \"ppt\""));
        assert!(manifest_text.contains("\"vba-project\""));
        assert!(manifest_text.contains("\"vba-module-source\""));

        let mut module_files = 0usize;
        for entry in fs::read_dir(&out).expect("read out dir") {
            let path = entry.expect("entry").path();
            if path.extension().and_then(|ext| ext.to_str()) == Some("bas")
                || path.extension().and_then(|ext| ext.to_str()) == Some("cls")
            {
                module_files += 1;
            }
        }
        assert!(module_files >= 2);

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_vba_legacy_partial_requires_best_effort() {
        let input = temp_file("legacy_vba_partial_fail", "doc");
        fs::write(
            &input,
            build_test_cfb(&[
                ("WordDocument", b"doc"),
                (
                    "VBA/PROJECT",
                    br#"Name="PartialProject"
Module=Module1/Module1
Module=MissingMod/MissingMod
"#,
                ),
                ("VBA/Module1", b"Sub AutoOpen()\nEnd Sub\n"),
            ]),
        )
        .expect("write partial legacy vba doc");
        let out = temp_dir("legacy_vba_partial_fail");

        let err = run(
            input.clone(),
            out.clone(),
            false,
            false,
            &ParserConfig::default(),
        )
        .expect_err("extract without best effort should fail");
        assert!(err
            .to_string()
            .contains("rerun with --best-effort to keep partial bundles"));

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_vba_legacy_partial_best_effort_writes_manifest_and_partial_bundle() {
        let input = temp_file("legacy_vba_partial_ok", "doc");
        fs::write(
            &input,
            build_test_cfb(&[
                ("WordDocument", b"doc"),
                (
                    "VBA/PROJECT",
                    br#"Name="PartialProject"
Module=Module1/Module1
Module=MissingMod/MissingMod
DPB="AAAA"
"#,
                ),
                ("VBA/Module1", b"Sub AutoOpen()\nEnd Sub\n"),
            ]),
        )
        .expect("write partial legacy vba doc");
        let out = temp_dir("legacy_vba_partial_ok");

        run(
            input.clone(),
            out.clone(),
            false,
            true,
            &ParserConfig::default(),
        )
        .expect("extract partial legacy vba with best effort");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"format\": \"doc\""));
        assert!(manifest_text.contains("\"vba-project\""));
        assert!(manifest_text.contains("\"vba-module-source\""));
        assert!(manifest_text.contains("MissingMod"));

        let mut bas_files = 0usize;
        for entry in fs::read_dir(&out).expect("read out dir") {
            let path = entry.expect("entry").path();
            if path.extension().and_then(|ext| ext.to_str()) == Some("bas") {
                bas_files += 1;
            }
        }
        assert_eq!(bas_files, 1);

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_vba_legacy_mixed_module_kinds_writes_bas_and_cls_files() {
        let input = temp_file("legacy_vba_mixed", "doc");
        fs::write(
            &input,
            build_test_cfb(&[
                ("WordDocument", b"doc"),
                (
                    "VBA/PROJECT",
                    br#"Name="MixedKinds"
Module=Core/Core
Class=Helper/0
Document=ThisDocument/&H00000000
"#,
                ),
                ("VBA/Core", b"Sub AutoOpen()\nEnd Sub\n"),
                ("VBA/Helper", b"Public Sub Run()\nEnd Sub\n"),
                (
                    "VBA/ThisDocument",
                    b"Private Sub Document_Open()\nEnd Sub\n",
                ),
            ]),
        )
        .expect("write mixed legacy vba doc");
        let out = temp_dir("legacy_vba_mixed");

        run(
            input.clone(),
            out.clone(),
            false,
            false,
            &ParserConfig::default(),
        )
        .expect("extract mixed legacy vba");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"format\": \"doc\""));
        assert!(manifest_text.contains("\"vba-project\""));
        assert!(manifest_text.contains("\"vba-module-source\""));

        let mut bas_files = Vec::new();
        let mut cls_files = Vec::new();
        for entry in fs::read_dir(&out).expect("read out dir") {
            let path = entry.expect("entry").path();
            match path.extension().and_then(|ext| ext.to_str()) {
                Some("bas") => bas_files.push(path),
                Some("cls") => cls_files.push(path),
                _ => {}
            }
        }

        assert_eq!(bas_files.len(), 1);
        assert_eq!(cls_files.len(), 2);
        assert!(bas_files.iter().any(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .unwrap_or_default()
                .contains("Core_")
        }));
        assert!(cls_files.iter().any(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .unwrap_or_default()
                .contains("Helper_")
        }));
        assert!(cls_files.iter().any(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .unwrap_or_default()
                .contains("ThisDocument_")
        }));

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    fn temp_file(name: &str, extension: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_extract_vba_{name}_{nanos}.{extension}"))
    }
}
