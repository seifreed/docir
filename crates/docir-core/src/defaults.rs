macro_rules! impl_default_via_new {
    ($($ty:path),+ $(,)?) => {
        $(
            impl Default for $ty {
                fn default() -> Self {
                    Self::new()
                }
            }
        )+
    };
}

impl_default_via_new!(
    crate::ir::CalcChain,
    crate::ir::ChartData,
    crate::ir::ChartSeries,
    crate::ir::CommentExtensionSet,
    crate::ir::CommentIdMap,
    crate::ir::ConnectionEntry,
    crate::ir::ConnectionPart,
    crate::ir::ContentControl,
    crate::ir::Diagnostics,
    crate::ir::DigitalSignature,
    crate::ir::DxfStyle,
    crate::ir::ExternalLinkPart,
    crate::ir::FontTable,
    crate::ir::Footer,
    crate::ir::GlossaryDocument,
    crate::ir::GlossaryEntry,
    crate::ir::HandoutMaster,
    crate::ir::Header,
    crate::ir::NotesMaster,
    crate::ir::NotesSlide,
    crate::ir::NumberingSet,
    crate::ir::PeoplePart,
    crate::ir::PivotCacheRecords,
    crate::ir::PresentationInfo,
    crate::ir::PresentationProperties,
    crate::ir::QueryTablePart,
    crate::ir::SharedStringTable,
    crate::ir::SheetMetadata,
    crate::ir::SheetMetadataType,
    crate::ir::SlicerPart,
    crate::ir::SlideLayout,
    crate::ir::SlideMaster,
    crate::ir::SpreadsheetStyles,
    crate::ir::StyleSet,
    crate::ir::TableStyleSet,
    crate::ir::Theme,
    crate::ir::TimelinePart,
    crate::ir::ViewProperties,
    crate::ir::VmlShape,
    crate::ir::WebExtension,
    crate::ir::WebExtensionTaskpane,
    crate::ir::WebSettings,
    crate::ir::WordSettings,
    crate::ir::WorkbookProperties,
    crate::ir::WorksheetDrawing,
    crate::security::ActiveXControl,
);

#[cfg(test)]
mod tests {
    use crate::ir::{
        CalcChain, CommentExtensionSet, Diagnostics, SheetMetadata, StyleSet, Theme,
        WorksheetDrawing,
    };
    use crate::security::ActiveXControl;

    #[test]
    fn default_impls_delegate_to_new_for_selected_types() {
        let calc_chain = CalcChain::default();
        assert!(calc_chain.entries.is_empty());

        let style_set = StyleSet::default();
        assert!(style_set.styles.is_empty());

        let diagnostics = Diagnostics::default();
        assert!(diagnostics.entries.is_empty());

        let comment_ext = CommentExtensionSet::default();
        assert!(comment_ext.entries.is_empty());

        let sheet_metadata = SheetMetadata::default();
        assert!(sheet_metadata.metadata_types.is_empty());

        let drawing = WorksheetDrawing::default();
        assert!(drawing.shapes.is_empty());

        let theme = Theme::default();
        assert!(theme.name.is_none());

        let activex = ActiveXControl::default();
        assert!(activex.properties.is_empty());
    }
}
