# 04-05 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95 || true
cargo llvm-cov --workspace --all-features --summary-only
```

## Workspace Total (Line Coverage)

- Baseline from 04-04 summary: `67.10%`
- Current 04-05 canonical total (fail-under run): `68.10%`
- Delta: `+1.00` percentage points

## 95% Gate Status

- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL, threshold still unmet)
- Observed total in fail-under run: `68.10%`

## Targeted ODF File Snapshots (Current Run)

- `docir-parser/src/odf/spreadsheet.rs`: `48.16%` lines (`352` lines missed)
- `docir-parser/src/odf/ods.rs`: `60.79%` lines (`278` lines missed)
- `docir-parser/src/odf/helpers.rs`: `70.81%` lines (`183` lines missed)
- `docir-parser/src/odf/formula.rs`: `82.29%` lines (`79` lines missed)

## Residual Highest-Impact Untouched Candidates (Next Gap Closure)

Excluding 04-03 and 04-04 improved modules and this plan's target files:

1. `docir-parser/src/ooxml/docx/document/inline.rs` (`234` lines missed, `60.27%`)
2. `docir-parser/src/ooxml/xlsx/worksheet.rs` (`188` lines missed, `65.50%`)
3. `docir-parser/src/rtf/core.rs` (`195` lines missed, `67.66%`)
4. `docir-parser/src/odf/presentation_helpers.rs` (`199` lines missed, `46.51%`)
5. `docir-parser/src/odf/styles_support.rs` (`156` lines missed, `55.93%`)
