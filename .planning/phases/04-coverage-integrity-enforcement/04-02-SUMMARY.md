---
phase: 04-coverage-integrity-enforcement
plan: "02"
subsystem: testing
tags: [coverage, fixtures, cli, policy]
requires:
  - phase: 04-coverage-integrity-enforcement
    provides: canonical coverage stage and CI llvm-cov setup from 04-01
provides:
  - Behavior-oriented parser fixture assertions tied to semantic/security outcomes
  - Content-level CLI coverage export contract assertions for JSON and CSV outputs
  - Explicit anti-inflation policy language for coverage acceptance
affects: [quality-gate, parser-tests, cli-tests, docs]
tech-stack:
  added: []
  patterns: [fixture-driven behavior assertions, typed export contract validation]
key-files:
  created: []
  modified:
    - crates/docir-parser/tests/fixtures.rs
    - crates/docir-cli/tests/coverage_export.rs
    - README.md
    - docs/quality-gate-policy.md
key-decisions:
  - "Coverage integrity assertions differentiate OOXML exports from non-OOXML fixtures to reflect real parser behavior."
  - "Anti-inflation policy is reviewer-visible in both README and gate policy docs."
patterns-established:
  - "Coverage tests must validate semantic outputs and exported content contracts, not just process success."
requirements-completed: [TEST-02]
duration: 12 min
completed: 2026-02-28
---

# Phase 04 Plan 02: Coverage Integrity Evidence Summary

**Coverage-related tests now assert parser semantics and CLI export contracts on real fixtures, with explicit repository policy that rejects synthetic coverage-inflation evidence.**

## Performance

- **Duration:** 12 min
- **Started:** 2026-02-28T18:41:30Z
- **Completed:** 2026-02-28T18:53:38Z
- **Tasks:** 3
- **Files modified:** 4

## Accomplishments
- Strengthened parser fixture tests with semantic content checks, OOXML coverage-diagnostic invariants, and security-relevant assertions for object-field fixtures.
- Strengthened CLI coverage export tests with typed JSON contract parsing and CSV/JSON content-level invariants across real fixtures.
- Added anti-inflation policy wording in README and quality policy docs requiring behavior-oriented evidence.

## Task Commits

Each task was committed atomically:

1. **Task 1: Strengthen parser fixture tests with explicit semantic/security behavior assertions** - `8b7c67b` (test)
2. **Task 2: Strengthen CLI coverage export tests with content-level contract assertions** - `b79eaf1` (test)
3. **Task 3: Document anti-inflation test policy for canonical acceptance** - `655d912` (docs)

## Files Created/Modified
- `crates/docir-parser/tests/fixtures.rs` - Adds semantic content assertions, OOXML diagnostic contracts, and security checks for OLE/hyperlink extraction.
- `crates/docir-cli/tests/coverage_export.rs` - Adds typed JSON/CSV contract checks for full and parts exports across fixture classes.
- `README.md` - Adds behavior-oriented coverage acceptance requirement.
- `docs/quality-gate-policy.md` - Adds explicit coverage integrity and anti-inflation rules.

## Decisions Made
- OOXML and non-OOXML fixtures are validated with distinct expectations to avoid false positives and preserve meaningful assertions.
- Coverage contract checks focus on report semantics (summary/counts/part rows/invariants) rather than binary success only.

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 3 - Blocking] Initial CLI export assertions assumed OOXML-style summaries for all fixture formats**
- **Found during:** Task 2 verification
- **Issue:** RTF/ODF/HWP/HWPX fixtures legitimately emit no OOXML coverage summary/part diagnostics; tests failed on valid behavior.
- **Fix:** Split assertions by fixture class (OOXML vs non-OOXML), keeping strong contracts for each class.
- **Files modified:** `crates/docir-cli/tests/coverage_export.rs`
- **Verification:** `cargo test -p docir-cli --test coverage_export -- --nocapture`
- **Committed in:** `b79eaf1`

---

**Total deviations:** 1 auto-fixed (1 blocking)
**Impact on plan:** Improved assertion quality by aligning contracts with real parser behavior per format class.

## Issues Encountered
- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95` still fails at workspace level (observed ~63% lines), so canonical 95% threshold is enforced but not yet satisfiable by current test coverage breadth.

## User Setup Required
None - no external service configuration required.

## Next Phase Readiness
- TEST-02 anti-inflation behavior checks and policy guardrails are in place.
- Additional coverage expansion work is required before the canonical 95% threshold can pass end-to-end.

## Self-Check: PASSED

