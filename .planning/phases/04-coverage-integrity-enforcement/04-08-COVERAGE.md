# 04-08 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95 || true
cargo llvm-cov --workspace --all-features --summary-only
```

## Workspace Total (Line Coverage)

- Baseline from 04-07 canonical total: `69.78%`
- Current 04-08 canonical total (`summary-only` run): `70.19%`
- Delta vs 04-07: `+0.41` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL, threshold still unmet)
- Observed total in fail-under run: `70.19%`
- Required threshold: `95.00%`
- Remaining gap to threshold: `24.81` percentage points

## Targeted Module Snapshots (04-08 Run)

- `docir-parser/src/ooxml/docx/document/inline.rs`: `64.06%` lines (`428` lines missed)
- `docir-parser/src/ooxml/xlsx/worksheet.rs`: `74.18%` lines (`292` lines missed)
- `docir-parser/src/odf/helpers.rs`: `72.82%` lines (`274` lines missed)
- `docir-parser/src/rtf/core.rs`: `75.26%` lines (`264` lines missed)
- `docir-parser/src/ooxml/pptx.rs`: `67.17%` lines (`241` lines missed)

## Targeted Module Comparison vs 04-07 Residual Baseline

Residual baseline from `04-07-COVERAGE.md` tracked missed-line counts for these modules.

- `inline.rs`: `447` -> `428` missed (`-19` lines)
- `worksheet.rs`: `323` -> `292` missed (`-31` lines)
- `helpers.rs`: `300` -> `274` missed (`-26` lines)
- `core.rs`: `282` -> `264` missed (`-18` lines)
- `pptx.rs`: `248` -> `241` missed (`-7` lines)

## Residual Highest-Impact Candidates for 04-09

Below-threshold status remains; next highest-impact candidates by missed lines are:

1. `docir-parser/src/ooxml/docx/document/inline.rs` (`428` lines missed, `64.06%`)
2. `docir-parser/src/odf/spreadsheet.rs` (`424` lines missed, `58.63%`)
3. `docir-parser/src/odf/presentation_helpers.rs` (`320` lines missed, `55.68%`)
4. `docir-parser/src/ooxml/xlsx/worksheet.rs` (`292` lines missed, `74.18%`)
5. `docir-parser/src/odf/helpers.rs` (`274` lines missed, `72.82%`)
