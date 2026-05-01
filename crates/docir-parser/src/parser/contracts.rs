use std::io::{Read, Seek};

use crate::{error::ParseError, parser::ParsedDocument};

/// Parse-stage boundary: raw input -> parsed IR.
pub(crate) trait ParseStage {
    /// Low-level parse stage (bytes -> unnormalized parsed IR).
    fn parse_stage<R: Read + Seek>(&self, reader: R) -> Result<ParsedDocument, ParseError>;
}

/// Normalize-stage boundary: parsed IR -> normalized IR.
pub(crate) trait NormalizeStage {
    /// Optional normalization stage for a parsed document.
    fn normalize_stage(&self, parsed: ParsedDocument) -> Result<ParsedDocument, ParseError> {
        Ok(parsed)
    }
}

/// Postprocess-stage boundary: normalized IR -> finalized parsed IR.
pub(crate) trait PostprocessStage {
    /// Optional post-processing stage for a parsed document.
    fn postprocess_stage(&self, parsed: ParsedDocument) -> Result<ParsedDocument, ParseError> {
        Ok(parsed)
    }
}

/// Parser pipeline contract composed from stage boundaries.
pub(crate) trait ParserPipeline: ParseStage + NormalizeStage + PostprocessStage {}

impl<T> ParserPipeline for T where T: ParseStage + NormalizeStage + PostprocessStage {}

pub(crate) fn run_parser_pipeline<P, R>(parser: &P, reader: R) -> Result<ParsedDocument, ParseError>
where
    P: ParserPipeline,
    R: Read + Seek,
{
    let parsed = parser.parse_stage(reader)?;
    let parsed = parser.normalize_stage(parsed)?;
    parser.postprocess_stage(parsed)
}

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::types::DocumentFormat;
    use std::cell::RefCell;
    use std::io::Cursor;

    struct RecordingPipeline {
        calls: RefCell<Vec<&'static str>>,
    }

    impl Default for RecordingPipeline {
        fn default() -> Self {
            Self {
                calls: RefCell::new(Vec::new()),
            }
        }
    }

    impl ParseStage for RecordingPipeline {
        fn parse_stage<R: Read + Seek>(&self, _reader: R) -> Result<ParsedDocument, ParseError> {
            self.calls.borrow_mut().push("parse");
            Ok(sample_parsed_document())
        }
    }

    impl NormalizeStage for RecordingPipeline {
        fn normalize_stage(&self, parsed: ParsedDocument) -> Result<ParsedDocument, ParseError> {
            self.calls.borrow_mut().push("normalize");
            Ok(parsed)
        }
    }

    impl PostprocessStage for RecordingPipeline {
        fn postprocess_stage(&self, parsed: ParsedDocument) -> Result<ParsedDocument, ParseError> {
            self.calls.borrow_mut().push("postprocess");
            Ok(parsed)
        }
    }

    fn sample_parsed_document() -> ParsedDocument {
        let mut store = docir_core::visitor::IrStore::new();
        let doc = docir_core::ir::Document::new(DocumentFormat::WordProcessing);
        let root_id = doc.id;
        store.insert(docir_core::ir::IRNode::Document(doc));
        ParsedDocument {
            root_id,
            format: DocumentFormat::WordProcessing,
            store,
            metrics: None,
        }
    }

    #[test]
    fn run_parser_pipeline_executes_stages_in_order() {
        let pipeline = RecordingPipeline::default();
        let reader = Cursor::new(Vec::<u8>::new());
        let parsed = run_parser_pipeline(&pipeline, reader).expect("pipeline should execute");
        assert!(parsed.document().is_some());

        let calls = pipeline.calls.borrow();
        assert_eq!(calls.as_slice(), ["parse", "normalize", "postprocess"]);
    }

    #[test]
    fn run_parser_pipeline_skips_downstream_stages_if_parse_fails() {
        struct FailingPipeline {
            calls: RefCell<Vec<&'static str>>,
        }

        impl Default for FailingPipeline {
            fn default() -> Self {
                Self {
                    calls: RefCell::new(Vec::new()),
                }
            }
        }

        impl ParseStage for FailingPipeline {
            fn parse_stage<R: Read + Seek>(
                &self,
                _reader: R,
            ) -> Result<ParsedDocument, ParseError> {
                self.calls.borrow_mut().push("parse");
                Err(ParseError::InvalidFormat("parse stage failed".into()))
            }
        }

        impl NormalizeStage for FailingPipeline {
            fn normalize_stage(
                &self,
                _parsed: ParsedDocument,
            ) -> Result<ParsedDocument, ParseError> {
                self.calls.borrow_mut().push("normalize");
                unreachable!("normalize should not run when parse fails");
            }
        }

        impl PostprocessStage for FailingPipeline {
            fn postprocess_stage(
                &self,
                _parsed: ParsedDocument,
            ) -> Result<ParsedDocument, ParseError> {
                self.calls.borrow_mut().push("postprocess");
                unreachable!("postprocess should not run when parse fails");
            }
        }

        let pipeline = FailingPipeline::default();
        let reader = Cursor::new(Vec::<u8>::new());
        let err = run_parser_pipeline(&pipeline, reader).err();
        assert!(matches!(err, Some(ParseError::InvalidFormat(_))));
        let calls = pipeline.calls.borrow();
        assert_eq!(calls.as_slice(), ["parse"]);
    }
}
