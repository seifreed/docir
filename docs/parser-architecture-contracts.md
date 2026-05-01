# Parser Internal Contracts And Visibility Policy

## Scope

This policy applies to parser orchestration contracts across parser entrypoints:

- `crates/docir-parser/src/parser/*.rs` (primary orchestrators)
- `crates/docir-parser/src/rtf/parser.rs`
- `crates/docir-parser/src/odf/builder.rs`
- `crates/docir-parser/src/hwp/builder.rs`

## Pipeline Contracts

The parser pipeline is split into explicit internal stage contracts:

- `ParseStage`: input reader to parsed IR.
- `NormalizeStage`: parsed IR to normalized IR.
- `PostprocessStage`: normalized IR to finalized IR.

`run_parser_pipeline(...)` executes stages in this order:

1. `parse_stage`
2. `normalize_stage`
3. `postprocess_stage`

Default behavior for `NormalizeStage` and `PostprocessStage` is pass-through.
Parsers can override only the stage they need without changing the pipeline API.

## Visibility Rules

- Default visibility is `pub(crate)` at crate boundaries and private within parser internals.
- `contracts` is private to `parser` module (`mod contracts;`), not directly imported from siblings.
- Sibling modules must use reexported parser boundary items from `crate::parser`:
  - `run_parser_pipeline`
  - `ParseStage`
  - `NormalizeStage`
  - `PostprocessStage`

## PR Gate Expectations

`scripts/quality_parser_pipeline_contracts.sh` enforces:

- Every parser entrypoint type (`DocumentParser`, `OoxmlParser`, `RtfParser`, `OdfParser`, `HwpParser`, `HwpxParser`) implements `ParseStage`.
- `parse_reader` uses `run_parser_pipeline(self, reader)`.
- No direct `parser::contracts::...` imports from sibling modules.
- `contracts` remains private in `parser.rs`.
