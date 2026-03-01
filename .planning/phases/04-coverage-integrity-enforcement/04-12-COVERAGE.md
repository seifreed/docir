# 04-12 Coverage Evidence

## Canonical Commands

```bash
bash scripts/tests/quality_gate_coverage_commands.sh
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95
cargo llvm-cov --workspace --all-features --summary-only
```

## Coverage Command Contract

- `bash scripts/tests/quality_gate_coverage_commands.sh` result: PASS
  - `coverage-command-contract: OK`
  - `coverage-threshold-fail: OK`
  - `quality_gate_coverage_commands: OK`

## Workspace Total (Line Coverage)

- Baseline from 04-11 canonical fail-under total: `71.67%`
- 04-12 canonical fail-under run total: `72.86%`
- 04-12 summary-only run total: `72.86%`
- Delta vs 04-11 baseline (canonical fail-under truth): `+1.19` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL, threshold still unmet)
- Observed total in fail-under run: `72.86%`
- Required threshold: `95.00%`
- Remaining gap to threshold: `22.14` percentage points

## Targeted Module Snapshots (04-12 Run)

- `docir-parser/src/odf/spreadsheet.rs`: `90.61%` lines (`134` lines missed)
- `docir-parser/src/ooxml/docx/document/inline.rs`: `79.29%` lines (`352` lines missed)
- `docir-parser/src/ooxml/xlsx/worksheet.rs`: `81.77%` lines (`289` lines missed)
- `docir-parser/src/odf/helpers.rs`: `81.17%` lines (`234` lines missed)

## Targeted Module Comparison vs 04-11 Residual Baseline

- `spreadsheet.rs`: `417` -> `134` missed (`-283` lines)
- `inline.rs`: `361` -> `352` missed (`-9` lines)
- `worksheet.rs`: `289` -> `289` missed (`0` lines)
- `helpers.rs`: `234` -> `234` missed (`0` lines)

## Deterministic Residual Handoff (for 04-13 if needed)

Ranked by current missed lines among 04-12 targeted modules:

1. `inline.rs` - `352` missed
2. `worksheet.rs` - `289` missed
3. `helpers.rs` - `234` missed
4. `spreadsheet.rs` - `134` missed

## 04-12 Closure Decision

04-12 materially improved canonical workspace line coverage from `71.67%` to `72.86%`, primarily by closing parallel/pivot residual branches in `spreadsheet.rs`. Canonical threshold enforcement still fails (`72.86% < 95.00%`), so CC-04 remains quantitatively blocked and phase closure cannot be claimed.
