# Phase 04 Goal-Backward Verification

status: gaps_found
phase: `04-coverage-integrity-enforcement`
date: `2026-02-28`
commit: `ba6d8698169fb8eaa2a67927ff883a0df506cd16`

## Goal-Backward Verdict

Phase 04 goal from [`ROADMAP.md`](../../ROADMAP.md): coverage threshold and test integrity are enforced as non-optional gate requirements.

Verdict: **Partially satisfied**. Non-optional enforcement wiring is present and active in canonical gate + CI, and behavior-oriented integrity evidence exists (including 04-05 ODF behavior tests). The phase remains incomplete because enforced coverage threshold `>=95%` is still unmet in canonical measurement.

## Inputs Reviewed

- `.planning/REQUIREMENTS.md`
- `.planning/ROADMAP.md`
- `.planning/phases/04-coverage-integrity-enforcement/04-05-PLAN.md`
- `.planning/phases/04-coverage-integrity-enforcement/04-05-SUMMARY.md`
- `.planning/phases/04-coverage-integrity-enforcement/04-05-COVERAGE.md`
- `.planning/phases/04-coverage-integrity-enforcement/04-04-COVERAGE.md`
- `scripts/quality_gate.sh`
- `scripts/tests/quality_gate_coverage_commands.sh`
- `.github/workflows/quality-gate.yml`
- `README.md`
- `docs/quality-gate-policy.md`
- `crates/docir-parser/src/odf/spreadsheet.rs`
- `crates/docir-parser/src/odf/ods.rs`
- `crates/docir-parser/src/odf/helpers.rs`
- `crates/docir-parser/src/odf/formula.rs`

## Commands Executed

- `bash scripts/tests/quality_gate_coverage_commands.sh`
  - Result: PASS (`coverage-command-contract: OK`, `coverage-threshold-fail: OK`, `quality_gate_coverage_commands: OK`)
- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
  - Result: FAIL (`EXIT:1`)
  - Observed total: `68.09%` lines
- `cargo llvm-cov --workspace --all-features --summary-only`
  - Result: PASS (`EXIT:0`)
  - Observed total: `68.09%` lines
- `rg -n "CC-04|TEST-01|TEST-02" .planning/REQUIREMENTS.md`
- `rg -n "anti-inflation|behavior|synthetic|coverage integrity|coverage" README.md docs/quality-gate-policy.md`
- `rg -n "#\[test\]|assert!\(|assert_eq!\(|assert_ne!\(" crates/docir-parser/src/odf/spreadsheet.rs crates/docir-parser/src/odf/ods.rs crates/docir-parser/src/odf/helpers.rs crates/docir-parser/src/odf/formula.rs`

## Requirement Validation

### CC-04

Requirement: gate enforces test coverage of at least 95% using `cargo llvm-cov`.

Evidence:
- `scripts/quality_gate.sh` sets `COVERAGE_THRESHOLD=95` and runs `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines "${COVERAGE_THRESHOLD}"` in `coverage_check`.
- 04-05 canonical artifact reports fail-under total `68.10%` with exit `1`.
- Fresh canonical runs on current commit show `68.09%` and fail-under `EXIT:1`.

Verdict: **Fail (enforced but unmet)**.

### TEST-01

Requirement: coverage target (>=95%) is measured in canonical run and cannot be skipped in CI.

Evidence:
- `.github/workflows/quality-gate.yml` installs `cargo-llvm-cov` and runs `./scripts/quality_gate.sh` as the quality gate job.
- `scripts/tests/quality_gate_coverage_commands.sh` verifies canonical command contract, including mandatory coverage invocation and failure propagation.

Verdict: **Pass**.

### TEST-02

Requirement: tests used to satisfy gate requirements validate real behavior and are not trivial assertion-only inflation.

Evidence:
- 04-05 artifact and module tests validate concrete outcomes (for example workbook/sheet counts, validation insertion, formula value resolution, malformed XML error behavior) in:
  - `crates/docir-parser/src/odf/spreadsheet.rs`
  - `crates/docir-parser/src/odf/ods.rs`
  - `crates/docir-parser/src/odf/helpers.rs`
  - `crates/docir-parser/src/odf/formula.rs`
- Policy text in `README.md` and `docs/quality-gate-policy.md` explicitly requires behavior-oriented, non-synthetic coverage evidence.

Verdict: **Pass**.

## Must-Have Validation Summary

- Coverage enforcement is non-optional in canonical gate path and CI routing.
- Test-integrity expectations are policy-backed and represented by behavior-first assertions in 04-05 targets.
- Blocking condition remains the numeric CC-04 threshold shortfall.

## Gap Summary

1. **Blocking quantitative gap (CC-04):** canonical workspace line coverage is `68.09%`, below required `95%`.
2. **Residual hotspot concentration:** large missed-line files remain (as captured in `04-05-COVERAGE.md`, e.g. `ooxml/docx/document/inline.rs`, `ooxml/xlsx/worksheet.rs`, `rtf/core.rs`, `odf/presentation_helpers.rs`, `odf/styles_support.rs`).

## Next Action Path

1. Execute Phase `04-06` targeted gap-closure on residual highest-impact hotspots from `04-05-COVERAGE.md`.
2. Re-run canonical coverage truth commands:
   - `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
   - `cargo llvm-cov --workspace --all-features --summary-only`
3. Re-verify phase only when fail-under command returns `EXIT:0` with total `>=95%`.
