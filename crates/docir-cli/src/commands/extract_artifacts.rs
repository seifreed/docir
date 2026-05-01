//! Extract embedded artifacts and write a canonical manifest to disk.

use anyhow::{bail, Context, Result};
use docir_app::{
    extract_artifacts_from_bytes, ArtifactExtractionOptions, ExportDocumentRef, ParserConfig,
    Phase0ArtifactManifestExport,
};
use std::fs;
use std::path::{Path, PathBuf};

use crate::commands::util::{build_app, source_format_label};

#[derive(Debug, Clone, Copy)]
pub struct ExtractArtifactsOptions {
    pub overwrite: bool,
    pub with_raw: bool,
    pub no_media: bool,
    pub only_ole: bool,
    pub only_rtf_objects: bool,
}

/// Public API entrypoint: run.
pub fn run(
    input: PathBuf,
    out_dir: PathBuf,
    options: ExtractArtifactsOptions,
    parser_config: &ParserConfig,
) -> Result<()> {
    prepare_output_dir(&out_dir, options.overwrite)?;

    let app = build_app(parser_config);
    let (parsed, bytes) = app
        .parse_file_with_bytes(&input)
        .with_context(|| format!("Failed to parse {}", input.display()))?;
    let bundle = extract_artifacts_from_bytes(
        &parsed,
        &bytes,
        Some(input.display().to_string()),
        &parser_config.zip_config,
        &ArtifactExtractionOptions {
            compute_hashes: parser_config.compute_hashes,
            with_raw: options.with_raw,
            no_media: options.no_media,
            only_ole: options.only_ole,
            only_rtf_objects: options.only_rtf_objects,
        },
    );

    let mut manifest = bundle.manifest;
    for payload in bundle.payloads {
        let output_path = out_dir.join(&payload.relative_path);
        if let Some(parent) = output_path.parent() {
            fs::create_dir_all(parent)
                .with_context(|| format!("Failed to create {}", parent.display()))?;
        }
        fs::write(&output_path, &payload.data)
            .with_context(|| format!("Failed to write {}", output_path.display()))?;
        if let Some(artifact) = manifest
            .artifacts
            .iter_mut()
            .find(|artifact| artifact.id == payload.artifact_id)
        {
            artifact.output_path = Some(payload.relative_path);
        }
    }

    let export = Phase0ArtifactManifestExport::from_manifest(
        &manifest,
        ExportDocumentRef::new(
            input.display().to_string(),
            source_format_label(&input, parsed.format().extension()),
            Some(input.display().to_string()),
        ),
    );
    let manifest_path = out_dir.join("manifest.json");
    let json = serde_json::to_string_pretty(&export)?;
    fs::write(&manifest_path, json)
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

#[cfg(test)]
mod tests {
    use super::{run, ExtractArtifactsOptions};
    use docir_app::{test_support::build_test_cfb, ParserConfig};
    use std::fs;
    use std::io::Write;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};
    use zip::write::FileOptions;

    fn temp_dir(name: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!("docir_cli_extract_artifacts_{name}_{nanos}"))
    }

    fn temp_file(name: &str, extension: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "docir_cli_extract_artifacts_{name}_{nanos}.{extension}"
        ))
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
    fn extract_artifacts_rtf_writes_manifest_and_blob() {
        let input = temp_file("rtf_objdata", "rtf");
        fs::write(&input, br"{\rtf1{\object{\objdata 4d5a9000}}}").expect("write rtf");
        let out = temp_dir("rtf");

        run(
            input.clone(),
            out.clone(),
            ExtractArtifactsOptions {
                overwrite: false,
                with_raw: false,
                no_media: false,
                only_ole: false,
                only_rtf_objects: false,
            },
            &ParserConfig::default(),
        )
        .expect("extract artifacts");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"embedded-file\""));
        let extracted = fs::read(out.join("rtf/artifact_1.exe")).expect("rtf blob");
        assert_eq!(extracted, vec![0x4d, 0x5a, 0x90, 0x00]);

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_artifacts_docx_media_writes_manifest_with_media_type_and_sha256() {
        let input = temp_file("docx_media", "docx");
        write_docx_with_media(&input);
        let out = temp_dir("docx_media");

        run(
            input.clone(),
            out.clone(),
            ExtractArtifactsOptions {
                overwrite: false,
                with_raw: false,
                no_media: false,
                only_ole: false,
                only_rtf_objects: false,
            },
            &ParserConfig::default(),
        )
        .expect("extract docx media");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"embedded-file\""));
        assert!(manifest_text.contains("\"path\": \"word/media/image1.png\""));
        assert!(manifest_text.contains("\"media_type\": \"image/png\""));
        assert!(manifest_text.contains(
            "\"sha256\": \"796120837694d3f3f29259cfeb25091698c2a0aa87873658d840b4993ee889b3\""
        ));

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_artifacts_docx_media_respects_no_media_flag() {
        let input = temp_file("docx_media_skip", "docx");
        write_docx_with_media(&input);
        let out = temp_dir("docx_media_skip");

        run(
            input.clone(),
            out.clone(),
            ExtractArtifactsOptions {
                overwrite: false,
                with_raw: false,
                no_media: true,
                only_ole: false,
                only_rtf_objects: false,
            },
            &ParserConfig::default(),
        )
        .expect("extract docx media with no-media");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(!manifest_text.contains("\"embedded-file\""));
        assert!(manifest_text.contains("\"NO_ARTIFACTS\""));
        assert!(!out.join("payloads/image1.png").exists());

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_artifacts_docx_media_omits_sha256_when_hashes_are_disabled() {
        let input = temp_file("docx_media_no_hashes", "docx");
        write_docx_with_media(&input);
        let out = temp_dir("docx_media_no_hashes");
        let config = ParserConfig {
            compute_hashes: false,
            ..ParserConfig::default()
        };

        run(
            input.clone(),
            out.clone(),
            ExtractArtifactsOptions {
                overwrite: false,
                with_raw: false,
                no_media: false,
                only_ole: false,
                only_rtf_objects: false,
            },
            &config,
        )
        .expect("extract docx media without hashes");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"embedded-file\""));
        assert!(manifest_text.contains("\"media_type\": \"image/png\""));
        assert!(!manifest_text.contains("\"sha256\":"));
        assert_eq!(
            fs::read(out.join("payloads/image1.png")).expect("payload"),
            b"PNG"
        );

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_artifacts_legacy_doc_writes_manifest_with_cfb_metadata() {
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

        let input = temp_file("legacy_doc", "doc");
        fs::write(
            &input,
            build_test_cfb(&[("WordDocument", b"doc"), ("Ole10Native", &ole10)]),
        )
        .expect("write legacy doc");
        let out = temp_dir("legacy_doc");

        run(
            input.clone(),
            out.clone(),
            ExtractArtifactsOptions {
                overwrite: false,
                with_raw: false,
                no_media: false,
                only_ole: false,
                only_rtf_objects: false,
            },
            &ParserConfig::default(),
        )
        .expect("extract legacy artifacts");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"offset\": 1"));
        assert!(
            manifest_text.contains("\"CFB_CREATED_FILETIME\"")
                || manifest_text.contains("\"CFB_MODIFIED_FILETIME\"")
                || manifest_text.contains("\"ole-object\"")
        );
        assert!(manifest_text.contains("\"ole-object\""));
        assert!(manifest_text.contains("\"embedded-file\""));

        let extracted = fs::read(out.join("payloads/dropper.exe")).expect("legacy payload");
        assert_eq!(extracted, b"MZ!!");

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_artifacts_legacy_xls_package_writes_pdf_payload() {
        let input = temp_file("legacy_xls_package", "xls");
        fs::write(
            &input,
            build_test_cfb(&[("Workbook", b"wb"), ("Package", b"%PDF-1.7")]),
        )
        .expect("write legacy xls");
        let out = temp_dir("legacy_xls_package");

        run(
            input.clone(),
            out.clone(),
            ExtractArtifactsOptions {
                overwrite: false,
                with_raw: false,
                no_media: false,
                only_ole: false,
                only_rtf_objects: false,
            },
            &ParserConfig::default(),
        )
        .expect("extract legacy package");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"format\": \"xls\""));
        assert!(manifest_text.contains("\"application/pdf\""));
        let extracted = fs::read(out.join("payloads/artifact_2.pdf")).expect("pdf payload");
        assert_eq!(extracted, b"%PDF-1.7");

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_artifacts_legacy_ppt_package_writes_pdf_payload() {
        let input = temp_file("legacy_ppt_package", "ppt");
        fs::write(
            &input,
            build_test_cfb(&[
                ("PowerPoint Document", b"ppt"),
                ("Current User", b"user"),
                ("Package", b"%PDF-1.7"),
            ]),
        )
        .expect("write legacy ppt");
        let out = temp_dir("legacy_ppt_package");

        run(
            input.clone(),
            out.clone(),
            ExtractArtifactsOptions {
                overwrite: false,
                with_raw: false,
                no_media: false,
                only_ole: false,
                only_rtf_objects: false,
            },
            &ParserConfig::default(),
        )
        .expect("extract legacy ppt package");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("\"format\": \"ppt\""));
        assert!(manifest_text.contains("\"application/pdf\""));
        let extracted = fs::read(out.join("payloads/artifact_2.pdf")).expect("pdf payload");
        assert_eq!(extracted, b"%PDF-1.7");

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_artifacts_objectpool_package_writes_pdf_payload() {
        let input = temp_file("objectpool_package", "doc");
        fs::write(
            &input,
            build_test_cfb(&[
                ("WordDocument", b"doc"),
                ("ObjectPool/1/Package", b"%PDF-1.7"),
            ]),
        )
        .expect("write objectpool package doc");
        let out = temp_dir("objectpool_package");

        run(
            input.clone(),
            out.clone(),
            ExtractArtifactsOptions {
                overwrite: false,
                with_raw: false,
                no_media: false,
                only_ole: false,
                only_rtf_objects: false,
            },
            &ParserConfig::default(),
        )
        .expect("extract objectpool package");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("ObjectPool/1/Package"));
        assert!(manifest_text.contains("\"application/pdf\""));
        let extracted = fs::read(out.join("payloads/artifact_2.pdf")).expect("pdf payload");
        assert_eq!(extracted, b"%PDF-1.7");

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }

    #[test]
    fn extract_artifacts_objectpool_ole10native_writes_named_payload() {
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

        let input = temp_file("objectpool_ole10", "doc");
        fs::write(
            &input,
            build_test_cfb(&[
                ("WordDocument", b"doc"),
                ("ObjectPool/1/Ole10Native", &ole10),
            ]),
        )
        .expect("write objectpool ole10 doc");
        let out = temp_dir("objectpool_ole10");

        run(
            input.clone(),
            out.clone(),
            ExtractArtifactsOptions {
                overwrite: false,
                with_raw: false,
                no_media: false,
                only_ole: false,
                only_rtf_objects: false,
            },
            &ParserConfig::default(),
        )
        .expect("extract objectpool ole10");

        let manifest_text = fs::read_to_string(out.join("manifest.json")).expect("manifest");
        assert!(manifest_text.contains("ObjectPool/1/Ole10Native"));
        let extracted = fs::read(out.join("payloads/dropper.exe")).expect("ole10 payload");
        assert_eq!(extracted, b"MZ!!");

        let _ = fs::remove_file(input);
        let _ = fs::remove_dir_all(out);
    }
}
