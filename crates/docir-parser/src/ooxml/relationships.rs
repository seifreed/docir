//! OOXML relationships parser (.rels files).

use crate::error::ParseError;
use crate::xml_utils::local_name;
use crate::xml_utils::{
    read_event, reader_from_str, track_xml_document_event, try_decoded_attr_value,
    visit_attributes_result,
};
use quick_xml::events::Event;
use std::collections::HashMap;

/// A single relationship entry.
#[derive(Debug, Clone)]
pub struct Relationship {
    /// Relationship ID (e.g., "rId1").
    pub id: String,
    /// Relationship type URI.
    pub rel_type: String,
    /// Target path or URL.
    pub target: String,
    /// Target mode (Internal or External).
    pub target_mode: TargetMode,
}

/// Target mode for relationships.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TargetMode {
    /// Internal part within the package.
    #[default]
    Internal,
    /// External resource (URL or path).
    External,
}

/// Collection of relationships from a .rels file.
#[derive(Debug, Clone, Default)]
pub struct Relationships {
    /// Relationships indexed by ID.
    pub by_id: HashMap<String, Relationship>,
    /// Relationships indexed by type.
    pub by_type: HashMap<String, Vec<String>>,
}

impl Relationships {
    /// Parses a .rels file.
    pub fn parse(xml: &str) -> Result<Self, ParseError> {
        Self::parse_with_path(xml, ".rels")
    }

    pub(crate) fn parse_with_path(xml: &str, path: &str) -> Result<Self, ParseError> {
        let mut reader = reader_from_str(xml);

        let mut rels = Relationships::default();
        let mut buf = Vec::new();
        let mut depth = 0usize;
        let mut root_closed = false;

        loop {
            let event = read_event(&mut reader, &mut buf, path)?;
            if track_xml_document_event(&event, &mut depth, &mut root_closed, path)? {
                break;
            }

            match event {
                Event::Empty(e) | Event::Start(e)
                    if local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    let mut id = None;
                    let mut rel_type = None;
                    let mut target = None;
                    let mut target_mode = TargetMode::Internal;

                    visit_attributes_result(&e, path, |attr| {
                        match attr.key.as_ref() {
                            b"Id" => {
                                id = Some(try_decoded_attr_value(attr, e.decoder(), path)?);
                            }
                            b"Type" => {
                                rel_type = Some(try_decoded_attr_value(attr, e.decoder(), path)?);
                            }
                            b"Target" => {
                                target = Some(try_decoded_attr_value(attr, e.decoder(), path)?);
                            }
                            b"TargetMode" => {
                                let mode = try_decoded_attr_value(attr, e.decoder(), path)?;
                                if mode.eq_ignore_ascii_case("External") {
                                    target_mode = TargetMode::External;
                                }
                            }
                            _ => {}
                        }
                        Ok(())
                    })?;

                    let id = id.ok_or_else(|| {
                        ParseError::InvalidStructure(format!("{path} relationship is missing Id"))
                    })?;
                    let rel_type = rel_type.ok_or_else(|| {
                        ParseError::InvalidStructure(format!("{path} relationship is missing Type"))
                    })?;
                    let target = target.ok_or_else(|| {
                        ParseError::InvalidStructure(format!(
                            "{path} relationship is missing Target"
                        ))
                    })?;
                    let rel = Relationship {
                        id: id.clone(),
                        rel_type: rel_type.clone(),
                        target,
                        target_mode,
                    };

                    if rels.by_id.contains_key(&id) {
                        return Err(ParseError::InvalidStructure(format!(
                            "{path} contains duplicate relationship Id: {id}"
                        )));
                    }
                    rels.by_type.entry(rel_type).or_default().push(id.clone());
                    rels.by_id.insert(id, rel);
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(rels)
    }

    /// Gets a relationship by ID.
    pub fn get(&self, id: &str) -> Option<&Relationship> {
        self.by_id.get(id)
    }

    /// Gets relationships by type.
    pub fn get_by_type(&self, rel_type: &str) -> Vec<&Relationship> {
        self.by_type
            .get(rel_type)
            .map(|ids| ids.iter().filter_map(|id| self.by_id.get(id)).collect())
            .unwrap_or_default()
    }

    /// Gets the first relationship of a given type.
    pub fn get_first_by_type(&self, rel_type: &str) -> Option<&Relationship> {
        self.get_by_type(rel_type).into_iter().next()
    }

    /// Returns all external relationships.
    pub fn external_relationships(&self) -> Vec<&Relationship> {
        self.by_id
            .values()
            .filter(|rel| rel.target_mode == TargetMode::External)
            .collect()
    }

    /// Resolves a relationship target relative to a base path.
    pub fn resolve_target(base_path: &str, target: &str) -> String {
        if looks_like_external_target(target) {
            return target.to_string();
        }

        let normalized_target = target.replace('\\', "/");

        // Handle absolute targets
        if normalized_target.starts_with('/') {
            return normalized_target
                .strip_prefix('/')
                .unwrap_or(&normalized_target)
                .to_string();
        }

        // Get directory of base path
        let base_dir = if let Some(idx) = base_path.rfind('/') {
            &base_path[..idx + 1]
        } else {
            ""
        };

        // Simple path resolution (handles ../ references)
        let mut parts: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();

        for component in normalized_target.split('/') {
            match component {
                ".." => {
                    if parts.pop().is_none() {
                        parts.push("..");
                    }
                }
                "." | "" => {}
                other => {
                    parts.push(other);
                }
            }
        }

        parts.join("/")
    }
}

pub(crate) fn looks_like_external_target(target: &str) -> bool {
    if target.starts_with("//") || target.starts_with("\\\\") {
        return true;
    }

    let Some(colon_idx) = target.find(':') else {
        return false;
    };
    let scheme = &target[..colon_idx];
    let Some(first) = scheme.bytes().next() else {
        return false;
    };

    first.is_ascii_alphabetic()
        && scheme
            .bytes()
            .all(|b| b.is_ascii_alphanumeric() || matches!(b, b'+' | b'-' | b'.'))
}

#[cfg(test)]
mod tests {
    use super::{Relationships, rel_type};
    use crate::error::ParseError;

    #[test]
    fn parse_accepts_prefixed_relationship_elements() {
        let xml = r#"
            <rel:Relationships xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships">
              <rel:Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
            </rel:Relationships>
        "#;

        let rels = Relationships::parse(xml).expect("relationships parse");
        let rel = rels.get("rId1").expect("prefixed relationship");

        assert_eq!(rel.rel_type, rel_type::OFFICE_DOCUMENT);
        assert_eq!(rel.target, "word/document.xml");
    }

    #[test]
    fn parse_unescapes_relationship_attribute_values() {
        let xml = r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="https://example.test/a?x=1&amp;y=2"
                TargetMode="External"/>
            </Relationships>
        "#;

        let rels = Relationships::parse(xml).expect("relationships parse");
        let rel = rels.get("rId1").expect("relationship");

        assert_eq!(rel.target, "https://example.test/a?x=1&y=2");
    }

    #[test]
    fn parse_reports_malformed_relationship_attributes() {
        let xml = r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
            </Relationships>
        "#;

        match Relationships::parse(xml).expect_err("malformed relationship attribute must fail") {
            ParseError::Xml { file, .. } => assert_eq!(file, ".rels"),
            other => panic!("expected XML error, got {other:?}"),
        }
    }

    #[test]
    fn parse_rejects_truncated_relationships_document() {
        let xml = r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
        "#;

        let err = Relationships::parse(xml).expect_err("truncated relationships must fail");
        assert!(
            matches!(err, ParseError::Xml { file, message } if file == ".rels" && message.contains("root is closed"))
        );
    }

    #[test]
    fn parse_reports_invalid_relationship_attribute_entity() {
        let xml = r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/&"/>
            </Relationships>
        "#;

        match Relationships::parse(xml).expect_err("invalid relationship entity must fail") {
            ParseError::Xml { file, .. } => assert_eq!(file, ".rels"),
            other => panic!("expected XML error, got {other:?}"),
        }
    }

    #[test]
    fn parse_rejects_relationships_missing_required_attributes() {
        for xml in [
            r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Type="http://example.test/type" Target="document.xml"/>
            </Relationships>
            "#,
            r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Target="document.xml"/>
            </Relationships>
            "#,
            r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://example.test/type"/>
            </Relationships>
            "#,
        ] {
            assert!(Relationships::parse(xml).is_err());
        }
    }

    #[test]
    fn parse_rejects_duplicate_relationship_ids() {
        let xml = r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://example.test/image" Target="image1.png"/>
              <Relationship Id="rId1" Type="http://example.test/image" Target="image2.png"/>
            </Relationships>
        "#;

        assert!(Relationships::parse(xml).is_err());
    }

    #[test]
    fn resolve_target_preserves_external_uris() {
        assert_eq!(
            Relationships::resolve_target("ppt/slides/slide1.xml", "https://example.com/a.wav"),
            "https://example.com/a.wav"
        );
    }

    #[test]
    fn resolve_target_accepts_backslash_components_for_internal_parts() {
        assert_eq!(
            Relationships::resolve_target("xl/worksheets/sheet1.xml", r"..\drawings\drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            Relationships::resolve_target("word/document.xml", r"media\image1.png"),
            "word/media/image1.png"
        );
        assert_eq!(
            Relationships::resolve_target("word/document.xml", r"\\server\share\doc.docx"),
            r"\\server\share\doc.docx"
        );
    }

    #[test]
    fn resolve_target_does_not_collapse_paths_above_package_root() {
        assert_eq!(
            Relationships::resolve_target("word/document.xml", "../../word/styles.xml"),
            "../word/styles.xml"
        );
    }
}

/// Known relationship types.
pub mod rel_type {
    pub const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    pub const CUSTOM_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

    // Word-specific
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub const WEB_SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
    pub const FONT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    pub const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub const ENDNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const THREADED_COMMENTS: &str =
        "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
    pub const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

    // Common
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const OLE_OBJECT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    pub const PACKAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
    pub const VBA_PROJECT: &str =
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
    pub const ATTACHED_TEMPLATE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate";

    // Excel-specific
    pub const WORKSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    pub const CHARTSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    pub const DIALOGSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
    pub const MACROSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/macrosheet";
    pub const SHARED_STRINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    pub const DRAWING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    pub const EXTERNAL_LINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
    pub const CONNECTIONS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
    pub const TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    pub const PIVOT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
    pub const PIVOT_CACHE_DEF: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
    pub const PIVOT_CACHE_RECORDS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";

    // PowerPoint-specific
    pub const SLIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    pub const SLIDE_LAYOUT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    pub const SLIDE_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    pub const NOTES_SLIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
    pub const NOTES_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
    pub const HANDOUT_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster";
}
