use super::*;
use crate::ooxml::shared as ooxml_shared;

impl OoxmlParser {
    /// Parse shared/package parts (themes, media, custom XML, relationships, signatures, legacy).
    pub(crate) fn parse_shared_parts<R: Read + Seek>(
        &self,
        zip: &mut SecureZipReader<R>,
        content_types: &ContentTypes,
        store: &mut IrStore,
        root_id: NodeId,
    ) -> Result<(), ParseError> {
        use sha2::{Digest, Sha256};
        use std::collections::HashSet;

        let mut shared_ids: Vec<NodeId> = Vec::new();
        let mut seen_parts: HashSet<String> = HashSet::new();
        let doc_format = match store.get(root_id) {
            Some(IRNode::Document(doc)) => doc.format,
            _ => DocumentFormat::WordProcessing,
        };

        // Relationship graphs
        let rel_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.ends_with(".rels"))
            .map(|s| s.to_string())
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

        // Themes
        let theme_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.contains("/theme/") && p.ends_with(".xml"))
            .map(|s| s.to_string())
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

        // Media assets
        let media_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.contains("/media/"))
            .map(|s| s.to_string())
            .collect();
        for path in media_paths {
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

        // Document thumbnails
        let thumbnail_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.starts_with("docProps/thumbnail."))
            .map(|s| s.to_string())
            .collect();
        for path in thumbnail_paths {
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

        // Word chart parts
        if doc_format == DocumentFormat::WordProcessing {
            let chart_paths: Vec<String> = zip
                .file_names()
                .filter(|p| p.starts_with("word/charts/") && p.ends_with(".xml"))
                .map(|s| s.to_string())
                .collect();
            for path in chart_paths {
                let xml = zip.read_file_string(&path)?;
                let chart_id = parse_chart_data(&xml, &path, store)?;
                shared_ids.push(chart_id);
                seen_parts.insert(path);
            }
        }

        // Spreadsheet chart parts
        if doc_format == DocumentFormat::Spreadsheet {
            let chart_paths: Vec<String> = zip
                .file_names()
                .filter(|p| p.starts_with("xl/charts/") && p.ends_with(".xml"))
                .map(|s| s.to_string())
                .collect();
            for path in chart_paths {
                let xml = zip.read_file_string(&path)?;
                let chart_id = parse_chart_data(&xml, &path, store)?;
                shared_ids.push(chart_id);
                seen_parts.insert(path);
            }
        }

        // Presentation chart parts
        if doc_format == DocumentFormat::Presentation {
            let chart_paths: Vec<String> = zip
                .file_names()
                .filter(|p| p.starts_with("ppt/charts/") && p.ends_with(".xml"))
                .map(|s| s.to_string())
                .collect();
            for path in chart_paths {
                let xml = zip.read_file_string(&path)?;
                let chart_id = parse_chart_data(&xml, &path, store)?;
                shared_ids.push(chart_id);
                seen_parts.insert(path);
            }
        }

        // Word SmartArt parts
        if doc_format == DocumentFormat::WordProcessing {
            let diagram_paths: Vec<String> = zip
                .file_names()
                .filter(|p| p.starts_with("word/diagrams/") && p.ends_with(".xml"))
                .map(|s| s.to_string())
                .collect();
            for path in diagram_paths {
                let xml = zip.read_file_string(&path)?;
                let part = parse_smartart_part(&xml, &path)?;
                let id = part.id;
                store.insert(IRNode::SmartArtPart(part));
                shared_ids.push(id);
                seen_parts.insert(path);
            }
        }

        // Word people part (coauthoring)
        if doc_format == DocumentFormat::WordProcessing && zip.contains("word/people.xml") {
            let xml = zip.read_file_string("word/people.xml")?;
            let people = ooxml_shared::parse_people_part(&xml, "word/people.xml")?;
            let id = people.id;
            store.insert(IRNode::PeoplePart(people));
            shared_ids.push(id);
            seen_parts.insert("word/people.xml".to_string());
        }

        // Word web extensions
        if doc_format == DocumentFormat::WordProcessing {
            let ext_paths: Vec<String> = zip
                .file_names()
                .filter(|p| p.starts_with("word/webExtensions/") && p.ends_with(".xml"))
                .map(|s| s.to_string())
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
        }

        // Custom XML parts
        let custom_xml_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.starts_with("customXml/") && p.ends_with(".xml"))
            .map(|s| s.to_string())
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

        // Digital signatures
        let sig_paths: Vec<String> = zip
            .file_names()
            .filter(|p| p.starts_with("_xmlsignatures/") && p.ends_with(".xml"))
            .map(|s| s.to_string())
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

        // Extension parts
        let extension_paths: Vec<String> = zip
            .file_names()
            .filter(|p| content_types.is_extension_part(p))
            .map(|s| s.to_string())
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

        if let Some(IRNode::Document(doc)) = store.get_mut(root_id) {
            doc.shared_parts.extend(shared_ids);
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
