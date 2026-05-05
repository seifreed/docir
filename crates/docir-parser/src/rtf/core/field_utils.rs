#[path = "field_utils_parse.rs"]
mod field_utils_parse;
use docir_core::ir::{FieldInstruction, FieldKind};

pub(crate) fn parse_field_instruction(text: &str) -> Option<FieldInstruction> {
    field_utils_parse::parse_field_instruction(text)
}

pub(super) fn parse_hyperlink_instruction(
    text: &str,
) -> Option<(String, Vec<String>, Vec<String>)> {
    field_utils_parse::parse_hyperlink_instruction(text)
}

#[cfg(test)]
pub(super) fn tokenize_field_instruction(text: &str) -> Vec<String> {
    field_utils_parse::tokenize_field_instruction(text)
}

#[cfg(test)]
mod tests {
    use super::{parse_field_instruction, parse_hyperlink_instruction, tokenize_field_instruction};
    use docir_core::ir::FieldKind;

    #[test]
    fn tokenize_handles_quotes_and_switches() {
        let tokens = tokenize_field_instruction(r#"HYPERLINK "https://example.test" \t "_blank""#);
        assert_eq!(
            tokens,
            vec![
                "HYPERLINK".to_string(),
                "https://example.test".to_string(),
                r#"\t"#.to_string(),
                "_blank".to_string(),
            ]
        );
    }

    #[test]
    fn parse_field_instruction_splits_args_and_switches() {
        let parsed = parse_field_instruction(r#"MERGEFIELD customer_name \* MERGEFORMAT"#)
            .expect("field instruction");
        assert!(matches!(parsed.kind, FieldKind::MergeField));
        assert_eq!(
            parsed.args,
            vec!["customer_name".to_string(), "MERGEFORMAT".to_string()]
        );
        assert_eq!(parsed.switches, vec!["*".to_string()]);
    }

    #[test]
    fn parse_field_instruction_recognizes_extended_field_kinds() {
        let parsed =
            parse_field_instruction(r#"DDEAUTO "cmd" "/c calc""#).expect("field instruction");
        assert!(matches!(parsed.kind, FieldKind::DdeAuto));

        let parsed =
            parse_field_instruction(r#"INCLUDEPICTURE "image.png""#).expect("field instruction");
        assert!(matches!(parsed.kind, FieldKind::IncludePicture));

        let parsed = parse_field_instruction("AUTOTEXT MyEntry").expect("field instruction");
        assert!(matches!(parsed.kind, FieldKind::AutoText));

        let parsed = parse_field_instruction("AUTOCORRECT MyEntry").expect("field instruction");
        assert!(matches!(parsed.kind, FieldKind::AutoCorrect));
    }

    #[test]
    fn parse_hyperlink_instruction_extracts_target_and_rest() {
        let (target, args, switches) =
            parse_hyperlink_instruction(r#"HYPERLINK "https://example.test" \l section1"#)
                .expect("hyperlink instruction");
        assert_eq!(target, "https://example.test");
        assert_eq!(args, vec!["section1".to_string()]);
        assert_eq!(switches, vec!["l".to_string()]);
    }
}
