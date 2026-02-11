//! OOXML relationships parser (.rels files).

use crate::error::ParseError;
use quick_xml::events::Event;
use quick_xml::Reader;
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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TargetMode {
    /// Internal part within the package.
    Internal,
    /// External resource (URL or path).
    External,
}

impl Default for TargetMode {
    fn default() -> Self {
        Self::Internal
    }
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
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut rels = Relationships::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"Relationship" {
                        let mut id = None;
                        let mut rel_type = None;
                        let mut target = None;
                        let mut target_mode = TargetMode::Internal;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"Id" => {
                                    id = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"Type" => {
                                    rel_type =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"Target" => {
                                    target = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"TargetMode" => {
                                    let mode = String::from_utf8_lossy(&attr.value);
                                    if mode.eq_ignore_ascii_case("External") {
                                        target_mode = TargetMode::External;
                                    }
                                }
                                _ => {}
                            }
                        }

                        if let (Some(id), Some(rel_type), Some(target)) = (id, rel_type, target) {
                            let rel = Relationship {
                                id: id.clone(),
                                rel_type: rel_type.clone(),
                                target,
                                target_mode,
                            };

                            rels.by_type.entry(rel_type).or_default().push(id.clone());
                            rels.by_id.insert(id, rel);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ParseError::Xml {
                        file: ".rels".to_string(),
                        message: e.to_string(),
                    });
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
        // Handle absolute targets
        if target.starts_with('/') {
            return target[1..].to_string();
        }

        // Get directory of base path
        let base_dir = if let Some(idx) = base_path.rfind('/') {
            &base_path[..idx + 1]
        } else {
            ""
        };

        // Simple path resolution (handles ../ references)
        let mut parts: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();

        for component in target.split('/') {
            match component {
                ".." => {
                    parts.pop();
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
