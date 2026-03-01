# 04-10 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95
cargo llvm-cov --workspace --all-features --summary-only
```

## Workspace Total (Line Coverage)

- Baseline from 04-09 canonical fail-under total: `70.91%`
- 04-10 canonical fail-under run total: `71.29%`
- 04-10 summary-only run total: `71.29%`
- Delta vs 04-09 baseline (canonical fail-under truth): `+0.38` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL, threshold still unmet)
- Observed total in fail-under run: `71.29%`
- Required threshold: `95.00%`
- Remaining gap to threshold: `23.71` percentage points

## Targeted Module Snapshots (04-10 Run)

- `docir-parser/src/odf/spreadsheet.rs`: `63.25%` lines (`441` lines missed)
- `docir-parser/src/ooxml/docx/document/inline.rs`: `71.40%` lines (`423` lines missed)
- `docir-parser/src/ooxml/xlsx/worksheet.rs`: `79.74%` lines (`296` lines missed)
- `docir-parser/src/odf/presentation_helpers.rs`: `75.72%` lines (`212` lines missed)
- `docir-parser/src/odf/helpers.rs`: `80.59%` lines (`229` lines missed)

## Targeted Module Comparison vs 04-09 Residual Baseline

- `spreadsheet.rs`: `439` -> `441` missed (`+2` lines)
- `inline.rs`: `406` -> `423` missed (`+17` lines)
- `worksheet.rs`: `289` -> `296` missed (`+7` lines)
- `presentation_helpers.rs`: `285` -> `212` missed (`-73` lines)
- `helpers.rs`: `226` -> `229` missed (`+3` lines)

## 04-10 Closure Decision

04-10 increased the workspace canonical total from `70.91%` to `71.29%` and materially reduced misses in `presentation_helpers.rs`, but canonical threshold enforcement still fails (`71.29% < 95.00%`). CC-04 completion remains quantitatively blocked until a canonical fail-under run exits `0`.
