use super::*;
use crate::ooxml::shared as ooxml_shared;
use crate::zip_handler::PackageReader;

impl OoxmlParser {
    /// Parse shared/package parts (themes, media, custom XML, relationships, signatures, legacy).
    pub(crate) fn parse_shared_parts(
        &self,
        zip: &mut impl PackageReader,
        content_types: &ContentTypes,
        store: &mut IrStore,
        root_id: NodeId,
    ) -> Result<(), ParseError> {
        use std::collections::HashSet;

        let mut shared_ids: Vec<NodeId> = Vec::new();
        let mut seen_parts: HashSet<String> = HashSet::new();
        let doc_format = match store.get(root_id) {
            Some(IRNode::Document(doc)) => doc.format,
            _ => DocumentFormat::WordProcessing,
        };

        self.parse_relationship_graphs(zip, store, &mut shared_ids, &mut seen_parts)?;
        self.parse_themes(zip, store, &mut shared_ids, &mut seen_parts)?;
        self.parse_media_assets(zip, content_types, store, &mut shared_ids, &mut seen_parts)?;
        self.parse_thumbnails(zip, content_types, store, &mut shared_ids, &mut seen_parts)?;
        self.parse_chart_parts(zip, store, doc_format, &mut shared_ids, &mut seen_parts)?;
        self.parse_smartart_parts(zip, store, doc_format, &mut shared_ids, &mut seen_parts)?;
        self.parse_people_part(zip, store, doc_format, &mut shared_ids, &mut seen_parts)?;
        self.parse_web_extensions(zip, store, doc_format, &mut shared_ids, &mut seen_parts)?;
        self.parse_custom_xml_parts(zip, store, &mut shared_ids, &mut seen_parts)?;
        self.parse_signatures(zip, store, &mut shared_ids, &mut seen_parts)?;
        self.parse_extension_parts(zip, content_types, store, &mut shared_ids, &mut seen_parts)?;

        if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
            doc.shared_parts.extend(shared_ids);
        }

        Ok(())
    }

    fn parse_relationship_graphs(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        let rel_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.ends_with(".rels"))
            .collect();
        for rel_path in rel_paths {
            let rels_xml = zip.read_file_string(&rel_path)?;
            let rels = Relationships::parse(&rels_xml)?;
            let source = rels_source_from_path(&rel_path);
            let graph = ooxml_shared::build_relationship_graph(&source, &rel_path, &rels);
            let graph_id = graph.id;
            store.insert(IRNode::RelationshipGraph(graph));
            shared_ids.push(graph_id);
            seen_parts.insert(rel_path);
        }
        Ok(())
    }

    fn parse_themes(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        let theme_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.contains("/theme/") && p.ends_with(".xml"))
            .collect();
        for path in theme_paths {
            let theme_xml = zip.read_file_string(&path)?;
            let mut theme = ooxml_shared::parse_theme(&theme_xml, &path)?;
            if theme.name.is_none() {
                theme.name = Some(path.clone());
            }
            let theme_id = theme.id;
            store.insert(IRNode::Theme(theme));
            shared_ids.push(theme_id);
            seen_parts.insert(path);
        }
        Ok(())
    }

    fn parse_media_assets(
        &self,
        zip: &mut impl PackageReader,
        content_types: &ContentTypes,
        store: &mut IrStore,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        let media_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.contains("/media/"))
            .collect();
        self.parse_media_paths(
            zip,
            content_types,
            store,
            shared_ids,
            seen_parts,
            media_paths,
        )
    }

    fn parse_thumbnails(
        &self,
        zip: &mut impl PackageReader,
        content_types: &ContentTypes,
        store: &mut IrStore,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        let thumbnail_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("docProps/thumbnail."))
            .collect();
        self.parse_media_paths(
            zip,
            content_types,
            store,
            shared_ids,
            seen_parts,
            thumbnail_paths,
        )
    }

    fn parse_media_paths(
        &self,
        zip: &mut impl PackageReader,
        content_types: &ContentTypes,
        store: &mut IrStore,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
        paths: Vec<String>,
    ) -> Result<(), ParseError> {
        use sha2::{Digest, Sha256};

        for path in paths {
            let data = zip.read_file(&path)?;
            let media_type = ooxml_shared::classify_media_type(&path);
            let mut media = MediaAsset::new(path.clone(), media_type, data.len() as u64);
            media.content_type = content_types.get_content_type(&path).map(|s| s.to_string());
            if self.config.compute_hashes {
                let mut hasher = Sha256::new();
                hasher.update(&data);
                media.hash = Some(hex::encode(hasher.finalize()));
            }
            media.span = Some(SourceSpan::new(&path));
            let media_id = media.id;
            store.insert(IRNode::MediaAsset(media));
            shared_ids.push(media_id);
            seen_parts.insert(path);
        }
        Ok(())
    }

    fn parse_chart_parts(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        doc_format: DocumentFormat,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        let chart_prefix = match doc_format {
            DocumentFormat::WordProcessing => "word/charts/",
            DocumentFormat::Spreadsheet => "xl/charts/",
            DocumentFormat::Presentation => "ppt/charts/",
            _ => return Ok(()),
        };
        let chart_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with(chart_prefix) && p.ends_with(".xml"))
            .collect();
        for path in chart_paths {
            let xml = zip.read_file_string(&path)?;
            let chart_id = parse_chart_data(&xml, &path, store)?;
            shared_ids.push(chart_id);
            seen_parts.insert(path);
        }
        Ok(())
    }

    fn parse_smartart_parts(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        doc_format: DocumentFormat,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        if doc_format != DocumentFormat::WordProcessing {
            return Ok(());
        }
        let diagram_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("word/diagrams/") && p.ends_with(".xml"))
            .collect();
        for path in diagram_paths {
            let xml = zip.read_file_string(&path)?;
            let part = parse_smartart_part(&xml, &path)?;
            let id = part.id;
            store.insert(IRNode::SmartArtPart(part));
            shared_ids.push(id);
            seen_parts.insert(path);
        }
        Ok(())
    }

    fn parse_people_part(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        doc_format: DocumentFormat,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        if doc_format != DocumentFormat::WordProcessing || !zip.contains("word/people.xml") {
            return Ok(());
        }
        let xml = zip.read_file_string("word/people.xml")?;
        let people = ooxml_shared::parse_people_part(&xml, "word/people.xml")?;
        let id = people.id;
        store.insert(IRNode::PeoplePart(people));
        shared_ids.push(id);
        seen_parts.insert("word/people.xml".to_string());
        Ok(())
    }

    fn parse_web_extensions(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        doc_format: DocumentFormat,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        if doc_format != DocumentFormat::WordProcessing {
            return Ok(());
        }
        let ext_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("word/webExtensions/") && p.ends_with(".xml"))
            .collect();
        for path in ext_paths {
            let xml = zip.read_file_string(&path)?;
            if path.ends_with("/taskpanes.xml") {
                let panes = ooxml_shared::parse_web_extension_taskpanes(&xml, &path)?;
                for pane in panes {
                    let id = pane.id;
                    store.insert(IRNode::WebExtensionTaskpane(pane));
                    shared_ids.push(id);
                }
            } else {
                let ext = ooxml_shared::parse_web_extension(&xml, &path)?;
                let id = ext.id;
                store.insert(IRNode::WebExtension(ext));
                shared_ids.push(id);
            }
            seen_parts.insert(path);
        }
        Ok(())
    }

    fn parse_custom_xml_parts(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        let custom_xml_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("customXml/") && p.ends_with(".xml"))
            .collect();
        for path in custom_xml_paths {
            let xml = zip.read_file_string(&path)?;
            let part =
                ooxml_shared::parse_custom_xml_part(&xml, &path, xml.as_bytes().len() as u64)?;
            let id = part.id;
            store.insert(IRNode::CustomXmlPart(part));
            shared_ids.push(id);
            seen_parts.insert(path);
        }
        Ok(())
    }

    fn parse_signatures(
        &self,
        zip: &mut impl PackageReader,
        store: &mut IrStore,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        let sig_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| p.starts_with("_xmlsignatures/") && p.ends_with(".xml"))
            .collect();
        for path in sig_paths {
            let xml = zip.read_file_string(&path)?;
            let mut sig = ooxml_shared::parse_signature(&xml, &path)?;
            sig.span = Some(SourceSpan::new(&path));
            let id = sig.id;
            store.insert(IRNode::DigitalSignature(sig));
            shared_ids.push(id);
            seen_parts.insert(path);
        }
        Ok(())
    }

    fn parse_extension_parts(
        &self,
        zip: &mut impl PackageReader,
        content_types: &ContentTypes,
        store: &mut IrStore,
        shared_ids: &mut Vec<NodeId>,
        seen_parts: &mut std::collections::HashSet<String>,
    ) -> Result<(), ParseError> {
        let extension_paths: Vec<String> = zip
            .file_names()
            .into_iter()
            .filter(|p| content_types.is_extension_part(p))
            .collect();
        for path in extension_paths {
            if seen_parts.contains(&path) {
                continue;
            }
            let data = zip.read_file(&path)?;
            let mut part = ooxml_shared::legacy_extension_part(&path, data.len() as u64);
            part.content_type = content_types.get_content_type(&path).map(|s| s.to_string());
            part.span = Some(SourceSpan::new(&path));
            let part_id = part.id;
            store.insert(IRNode::ExtensionPart(part));
            shared_ids.push(part_id);
            seen_parts.insert(path);
        }
        Ok(())
    }
}

fn rels_source_from_path(rels_path: &str) -> String {
    if let Some(pos) = rels_path.rfind("/_rels/") {
        let prefix = &rels_path[..pos];
        if prefix.is_empty() {
            return "/".to_string();
        }
        return prefix.to_string();
    }

    if rels_path == "_rels/.rels" {
        return "/".to_string();
    }

    "".to_string()
}
