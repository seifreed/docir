---
phase: 04-coverage-integrity-enforcement
plan: "16"
subsystem: testing
tags: [coverage, llvm-cov, parser, hwp]
requires:
  - phase: 04-15
    provides: canonical 77.67% baseline and parser hotspot ranking
provides:
  - behavior-first unit tests for hwp legacy parser helpers
  - canonical 04-16 fail-under evidence with residual ranking
affects: [CC-04, TEST-01, TEST-02, phase-04-closure]
completed: 2026-03-01
---

# Phase 04 Plan 16 Summary

Parser hotspot closure in `hwp/legacy.rs` increased canonical coverage from `77.67%` to `78.12%` (`+0.45` points), with fail-under still `EXIT:1`.

## Delivered

- Added dense unit tests for legacy HWP parse/decode/decompress and script/url extraction helpers.
- Validated behavior assertions through parser-local test run.
- Published canonical 04-16 measurement and refreshed residual ranking.

## Verification Commands

- `cargo test -p docir-parser hwp::legacy::tests -- --nocapture`
- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`

## Outcome

- Canonical total: `78.12%`
- Canonical fail-under status: `EXIT:1`
- Remaining threshold gap: `16.88` points
