# 04-07 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95 || true
cargo llvm-cov --workspace --all-features --summary-only
```

## Workspace Total (Line Coverage)

- Baseline from 04-06 canonical total: `68.62%`
- Current 04-07 canonical total (`summary-only` run): `69.78%`
- Delta vs 04-06: `+1.16` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL, threshold still unmet)
- Observed total in fail-under run: `69.78%`
- Required threshold: `95.00%`
- Remaining gap to threshold: `25.22` percentage points

## Targeted Module Snapshots (04-07 Run)

- `docir-parser/src/odf/spreadsheet.rs`: `58.63%` lines (`424` lines missed)
- `docir-parser/src/odf/ods.rs`: `79.44%` lines (`236` lines missed)
- `docir-parser/src/odf/presentation_helpers.rs`: `55.68%` lines (`320` lines missed)
- `docir-parser/src/odf/styles_support.rs`: `71.76%` lines (`172` lines missed)

## Targeted Module Comparison vs 04-06 Residual Baseline

Residual baseline from `04-06-COVERAGE.md` tracked missed-line counts for these modules.

- `spreadsheet.rs`: `538` -> `424` missed (`-114` lines)
- `ods.rs`: `407` -> `236` missed (`-171` lines)
- `presentation_helpers.rs`: `324` -> `320` missed (`-4` lines)
- `styles_support.rs`: `261` -> `172` missed (`-89` lines)

## Residual Highest-Impact Candidates for 04-08

Below-threshold status remains; next highest-impact candidates by missed lines are:

1. `docir-parser/src/ooxml/docx/document/inline.rs` (`447` lines missed, `57.75%`)
2. `docir-parser/src/ooxml/xlsx/worksheet.rs` (`323` lines missed, `68.36%`)
3. `docir-parser/src/odf/helpers.rs` (`300` lines missed, `69.57%`)
4. `docir-parser/src/rtf/core.rs` (`282` lines missed, `72.49%`)
5. `docir-parser/src/ooxml/pptx.rs` (`248` lines missed, `66.21%`)
