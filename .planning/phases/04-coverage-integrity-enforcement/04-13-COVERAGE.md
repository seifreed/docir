# 04-13 Coverage Evidence

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

- Baseline from 04-12 canonical fail-under total: `72.86%`
- 04-13 canonical fail-under run total: `73.44%`
- 04-13 summary-only run total: `73.44%`
- Delta vs 04-12 baseline (canonical fail-under truth): `+0.58` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL, threshold still unmet)
- Observed total in fail-under run: `73.44%`
- Required threshold: `95.00%`
- Remaining gap to threshold: `21.56` percentage points

## Targeted Module Snapshots (04-13 Run)

- `docir-parser/src/ooxml/docx/document/inline.rs`: `79.77%` lines (`369` lines missed)
- `docir-parser/src/ooxml/xlsx/worksheet.rs`: `93.25%` lines (`115` lines missed)
- `docir-parser/src/odf/helpers.rs`: `82.51%` lines (`235` lines missed)
- `docir-parser/src/odf/spreadsheet.rs`: `90.82%` lines (`135` lines missed)

## Targeted Module Comparison vs 04-12 Residual Baseline

- `inline.rs`: `352` -> `369` missed (`+17` lines)
- `worksheet.rs`: `289` -> `115` missed (`-174` lines)
- `helpers.rs`: `234` -> `235` missed (`+1` line)
- `spreadsheet.rs`: `134` -> `135` missed (`+1` line)

## Deterministic Residual Handoff (for 04-14 if needed)

Ranked by current missed lines among 04-13 targeted modules:

1. `inline.rs` - `369` missed
2. `helpers.rs` - `235` missed
3. `spreadsheet.rs` - `135` missed
4. `worksheet.rs` - `115` missed

## 04-13 Closure Decision

04-13 improved canonical workspace line coverage from `72.86%` to `73.44%` and materially reduced misses in `worksheet.rs`. Canonical threshold enforcement still fails (`73.44% < 95.00%`), so CC-04 remains quantitatively blocked and phase closure cannot be claimed.
