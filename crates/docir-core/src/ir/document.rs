//! Document root IR node.

use crate::ir::new_node_id;
use crate::security::SecurityInfo;
use crate::types::{DocumentFormat, NodeId, SourceSpan};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Root node of the IR tree representing an Office document.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Document {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Document format (Word, Excel, PowerPoint).
    pub format: DocumentFormat,

    /// Document metadata (title, author, etc.).
    pub metadata: Option<NodeId>,

    /// Style definitions (Word only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub styles: Option<NodeId>,

    /// Numbering definitions (Word only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub numbering: Option<NodeId>,

    /// Comments (Word only).
    pub comments: Vec<NodeId>,

    /// Footnotes (Word only).
    pub footnotes: Vec<NodeId>,

    /// Endnotes (Word only).
    pub endnotes: Vec<NodeId>,

    /// Word settings (Word only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub settings: Option<NodeId>,

    /// Web settings (Word only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub web_settings: Option<NodeId>,

    /// Font table (Word only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_table: Option<NodeId>,

    /// Styles with effects (Word only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub styles_with_effects: Option<NodeId>,

    /// Comments extended metadata (Word only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comments_extended: Option<NodeId>,

    /// Comment ID map (Word only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comment_id_map: Option<NodeId>,

    /// Workbook properties (Excel only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub workbook_properties: Option<NodeId>,

    /// Shared strings table (Excel only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shared_strings: Option<NodeId>,

    /// Spreadsheet styles (Excel only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spreadsheet_styles: Option<NodeId>,

    /// Spreadsheet metadata (Excel only).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet_metadata: Option<NodeId>,

    /// Defined names (Excel only).
    #[serde(default)]
    pub defined_names: Vec<NodeId>,

    /// Pivot caches (Excel only).
    #[serde(default)]
    pub pivot_caches: Vec<NodeId>,

    /// Top-level content sections.
    /// For DOCX: sections
    /// For XLSX: worksheets
    /// For PPTX: slides
    pub content: Vec<NodeId>,

    /// Security information (macros, external refs, etc.).
    pub security: SecurityInfo,

    /// Shared/package parts (themes, relationships, media, custom XML, etc.).
    pub shared_parts: Vec<NodeId>,

    /// Diagnostics entries.
    #[serde(default)]
    pub diagnostics: Vec<NodeId>,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Document {
    /// Creates a new Document with the given format.
    pub fn new(format: DocumentFormat) -> Self {
        Self {
            id: new_node_id(),
            format,
            metadata: None,
            styles: None,
            numbering: None,
            comments: Vec::new(),
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            settings: None,
            web_settings: None,
            font_table: None,
            styles_with_effects: None,
            comments_extended: None,
            comment_id_map: None,
            workbook_properties: None,
            shared_strings: None,
            spreadsheet_styles: None,
            sheet_metadata: None,
            defined_names: Vec::new(),
            pivot_caches: Vec::new(),
            content: Vec::new(),
            security: SecurityInfo::default(),
            shared_parts: Vec::new(),
            diagnostics: Vec::new(),
            span: None,
        }
    }

    /// Registers a shared/package part ID on the document.
    ///
    /// Callers should ensure the ID refers to a node stored in the IR store.
    pub fn add_shared_part(&mut self, id: NodeId) {
        self.shared_parts.push(id);
    }

    /// Registers a diagnostics node ID on the document.
    ///
    /// Callers should ensure the ID refers to a diagnostics node stored in the IR store.
    pub fn add_diagnostic(&mut self, id: NodeId) {
        self.diagnostics.push(id);
    }

    /// Returns all child node IDs.
    pub fn children(&self) -> Vec<NodeId> {
        let mut children = Vec::new();
        if let Some(meta_id) = self.metadata {
            children.push(meta_id);
        }
        if let Some(styles_id) = self.styles {
            children.push(styles_id);
        }
        if let Some(numbering_id) = self.numbering {
            children.push(numbering_id);
        }
        children.extend(&self.comments);
        children.extend(&self.footnotes);
        children.extend(&self.endnotes);
        if let Some(settings_id) = self.settings {
            children.push(settings_id);
        }
        if let Some(web_settings_id) = self.web_settings {
            children.push(web_settings_id);
        }
        if let Some(font_table_id) = self.font_table {
            children.push(font_table_id);
        }
        if let Some(sheet_metadata_id) = self.sheet_metadata {
            children.push(sheet_metadata_id);
        }
        if let Some(styles_with_effects_id) = self.styles_with_effects {
            children.push(styles_with_effects_id);
        }
        if let Some(comments_extended_id) = self.comments_extended {
            children.push(comments_extended_id);
        }
        if let Some(comment_id_map_id) = self.comment_id_map {
            children.push(comment_id_map_id);
        }
        if let Some(workbook_props_id) = self.workbook_properties {
            children.push(workbook_props_id);
        }
        if let Some(shared_strings_id) = self.shared_strings {
            children.push(shared_strings_id);
        }
        if let Some(styles_id) = self.spreadsheet_styles {
            children.push(styles_id);
        }
        children.extend(&self.defined_names);
        children.extend(&self.pivot_caches);
        children.extend(&self.content);
        if let Some(macro_id) = self.security.macro_project {
            children.push(macro_id);
        }
        children.extend(&self.security.ole_objects);
        children.extend(&self.security.external_refs);
        children.extend(&self.security.activex_controls);
        children.extend(&self.shared_parts);
        children.extend(&self.diagnostics);
        children
    }
}

/// A logical section within a Word document.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone)]
pub struct Section {
    /// Unique identifier for this node.
    pub id: NodeId,

    /// Section name (if any).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,

    /// Header node IDs for this section.
    pub headers: Vec<NodeId>,

    /// Footer node IDs for this section.
    pub footers: Vec<NodeId>,

    /// Content blocks (paragraphs, tables, etc.).
    pub content: Vec<NodeId>,

    /// Section properties.
    pub properties: SectionProperties,

    /// Source span information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub span: Option<SourceSpan>,
}

impl Section {
    /// Creates a new empty Section.
    pub fn new() -> Self {
        Self {
            id: new_node_id(),
            name: None,
            headers: Vec::new(),
            footers: Vec::new(),
            content: Vec::new(),
            properties: SectionProperties::default(),
            span: None,
        }
    }

    /// Returns all child node IDs.
    pub fn children(&self) -> Vec<NodeId> {
        let mut children = Vec::new();
        children.extend(&self.headers);
        children.extend(&self.footers);
        children.extend(&self.content);
        children
    }
}

impl Default for Section {
    fn default() -> Self {
        Self::new()
    }
}

/// Section layout properties.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct SectionProperties {
    /// Page width in twips (1/20 of a point).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_width: Option<u32>,

    /// Page height in twips.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_height: Option<u32>,

    /// Page orientation.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub orientation: Option<PageOrientation>,

    /// Page margins.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margins: Option<PageMargins>,

    /// Number of columns.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub columns: Option<u32>,

    /// Column spacing in twips.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_spacing: Option<u32>,

    /// Column separator line enabled.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_separator: Option<bool>,

    /// Section type (next page, continuous, etc).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub section_type: Option<SectionType>,

    /// Title page enabled.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_page: Option<bool>,

    /// Page numbering settings.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_numbering: Option<PageNumbering>,

    /// Line numbering settings.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_numbering: Option<LineNumbering>,

    /// Page borders.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_borders: Option<PageBorders>,

    /// Text direction (raw OOXML value).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<String>,
}

/// Page orientation.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

/// Page margins in twips.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct PageMargins {
    pub top: u32,
    pub bottom: u32,
    pub left: u32,
    pub right: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gutter: Option<u32>,
}

/// Section break type.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionType {
    NextPage,
    Continuous,
    EvenPage,
    OddPage,
}

/// Page numbering settings.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct PageNumbering {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
}

/// Line numbering settings.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct LineNumbering {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub count_by: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub distance: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub restart: Option<LineNumberRestart>,
}

/// Line numbering restart rules.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineNumberRestart {
    Continuous,
    NewPage,
    NewSection,
}

/// Page borders.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Default)]
pub struct PageBorders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<crate::ir::Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<crate::ir::Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<crate::ir::Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<crate::ir::Border>,
}
