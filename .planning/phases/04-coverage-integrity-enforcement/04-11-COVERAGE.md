# 04-11 Coverage Evidence

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

- Baseline from 04-10 canonical fail-under total: `71.29%`
- 04-11 canonical fail-under run total: `71.67%`
- 04-11 summary-only run total: `71.67%`
- Delta vs 04-10 baseline (canonical fail-under truth): `+0.38` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL, threshold still unmet)
- Observed total in fail-under run: `71.67%`
- Required threshold: `95.00%`
- Remaining gap to threshold: `23.33` percentage points

## Targeted Module Snapshots (04-11 Run)

- `docir-parser/src/odf/spreadsheet.rs`: `66.90%` lines (`417` lines missed)
- `docir-parser/src/ooxml/docx/document/inline.rs`: `77.11%` lines (`361` lines missed)
- `docir-parser/src/ooxml/xlsx/worksheet.rs`: `81.77%` lines (`289` lines missed)
- `docir-parser/src/odf/helpers.rs`: `81.17%` lines (`234` lines missed)

## Targeted Module Comparison vs 04-10 Residual Baseline

- `spreadsheet.rs`: `441` -> `417` missed (`-24` lines)
- `inline.rs`: `423` -> `361` missed (`-62` lines)
- `worksheet.rs`: `296` -> `289` missed (`-7` lines)
- `helpers.rs`: `229` -> `234` missed (`+5` lines)

## Deterministic Residual Handoff (for 04-12 if needed)

Ranked by current missed lines among 04-11 targeted modules:

1. `spreadsheet.rs` - `417` missed
2. `inline.rs` - `361` missed
3. `worksheet.rs` - `289` missed
4. `helpers.rs` - `234` missed

## 04-11 Closure Decision

04-11 improved canonical workspace line coverage from `71.29%` to `71.67%` and reduced misses in three of four targeted residual modules. Canonical threshold enforcement still fails (`71.67% < 95.00%`), so CC-04 remains quantitatively blocked and phase closure cannot be claimed.
