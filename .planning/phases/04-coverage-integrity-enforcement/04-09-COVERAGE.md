# 04-09 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95 || true
cargo llvm-cov --workspace --all-features --summary-only
```

## Workspace Total (Line Coverage)

- Baseline from 04-08 canonical total: `70.19%`
- 04-09 canonical fail-under run total: `70.91%`
- 04-09 summary-only run total: `70.90%`
- Delta vs 04-08 baseline (canonical fail-under truth): `+0.72` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL, threshold still unmet)
- Observed total in fail-under run: `70.91%`
- Required threshold: `95.00%`
- Remaining gap to threshold: `24.09` percentage points

## Targeted Module Snapshots (04-09 Run)

- `docir-parser/src/ooxml/docx/document/inline.rs`: `69.31%` lines (`406` lines missed)
- `docir-parser/src/odf/spreadsheet.rs`: `60.59%` lines (`439` lines missed)
- `docir-parser/src/odf/presentation_helpers.rs`: `65.95%` lines (`285` lines missed)
- `docir-parser/src/ooxml/xlsx/worksheet.rs`: `77.82%` lines (`289` lines missed)
- `docir-parser/src/odf/helpers.rs`: `80.07%` lines (`226` lines missed)

## Targeted Module Comparison vs 04-08 Residual Baseline

- `inline.rs`: `428` -> `406` missed (`-22` lines)
- `spreadsheet.rs`: `424` -> `439` missed (`+15` lines)
- `presentation_helpers.rs`: `320` -> `285` missed (`-35` lines)
- `worksheet.rs`: `292` -> `289` missed (`-3` lines)
- `helpers.rs`: `274` -> `226` missed (`-48` lines)

## 04-09 Closure Decision

04-09 improved workspace coverage and reduced four of five targeted residual modules, but canonical threshold enforcement still fails (`70.91% < 95.00%`). Phase 04 remains quantitatively open until a canonical fail-under run returns exit `0`.
