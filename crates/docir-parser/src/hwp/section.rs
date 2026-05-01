mod section_normalize;
mod section_parse;
mod section_postprocess;
#[cfg(test)]
mod section_tests;

pub(crate) use self::section_parse::parse_hwpx_section;
#[cfg(test)]
pub(crate) use self::section_parse::HwpxNoteKind;
#[cfg(test)]
pub(crate) use self::section_parse::{note_kind_from_local, revision_type_from_local};
