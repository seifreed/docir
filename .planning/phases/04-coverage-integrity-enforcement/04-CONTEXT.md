# Phase 4: Coverage Integrity Enforcement - Context

**Gathered:** 2026-02-28
**Status:** Ready for planning

<domain>
## Phase Boundary

Enforce coverage integrity as a hard canonical gate requirement: `cargo llvm-cov` must run in the canonical gate, coverage threshold must be >=95%, and CI/pre-commit/local acceptance must not bypass this measurement.

</domain>

<decisions>
## Implementation Decisions

### Canonical coverage stage contract
- `./scripts/quality_gate.sh` remains the only accepted entrypoint and must include a dedicated coverage stage using `cargo llvm-cov`.
- Coverage stage executes after baseline clean-code stages (`fmt`, strict `clippy`, `test`) and participates in deterministic fail-fast classification.
- Coverage stage failure maps to canonical quality failure (`CLASS=quality_failure`, non-zero exit).

### Threshold and measurement policy
- Coverage threshold is fixed at **>=95%** for Phase 4 acceptance (`CC-04`, `TEST-01`).
- Threshold evaluation must be machine-checked in the gate (not documentation-only).
- Coverage result must be visible in canonical output with unambiguous pass/fail signaling.

### Anti-inflation integrity posture
- No synthetic/trivial tests added solely to inflate coverage metrics (`TEST-02`).
- No allow-list or filtering strategy that hides meaningful production code just to reach threshold without rationale.
- If exclusions are required for technical reasons, they must be explicit, minimal, documented, and reproducible in both local and CI runs.

### CI routing and non-skip enforcement
- CI required job remains `quality-gate` and must not provide alternate acceptance jobs that skip coverage measurement.
- Any helper scripts for diagnostics remain non-authoritative and must not be accepted as merge criteria.
- Pre-commit continues to call canonical gate; partial local shortcuts are informational only.

### Claude's Discretion
- Exact `cargo llvm-cov` invocation flags and report format, as long as threshold enforcement is deterministic and verifiable.
- Exact location/format of persisted coverage artifacts for evidence.
- Whether to add dedicated coverage contract tests around the gate stage behavior.

</decisions>

<specifics>
## Specific Ideas

- Keep threshold value centralized (single source of truth) to prevent drift between shell logic, tests, and docs.
- Preserve final `QUALITY_GATE_RESULT=...` line contract while adding a coverage summary line (`COVERAGE_PERCENT=...`).
- Add explicit negative test path showing gate failure when measured coverage is below threshold.

</specifics>

<code_context>
## Existing Code Insights

### Reusable Assets
- `scripts/quality_gate.sh`: deterministic stage runner currently enforcing fmt/clippy/test.
- `scripts/lib/quality_gate_lib.sh`: shared helpers for stage execution and classification.
- `scripts/tests/quality_gate_exit_codes.sh`: black-box contract harness for exit classes.
- `scripts/tests/quality_gate_baseline_commands.sh`: deterministic stage-order and command-argument harness.

### Established Patterns
- Canonical-only acceptance is already enforced in README, policy docs, hooks, and CI.
- Stage failures map through `classify_failure` and end with a machine-parseable final result line.
- CI quality workflow is single-job (`quality-gate`) and already merge-required.

### Integration Points
- Extend `scripts/quality_gate.sh` stage list with coverage stage tied to `cargo llvm-cov`.
- Extend shell test harnesses (or add new coverage harness) to validate threshold fail/pass deterministically.
- Update docs (`README.md`, `docs/quality-gate-policy.md`, and coverage runbook) to align with canonical coverage enforcement semantics.

</code_context>

<deferred>
## Deferred Ideas

- Forbidden construct scanning (`unwrap/expect/panic/todo/unimplemented`) belongs to Phase 5.
- Dead code/import/docs/complexity hard gates belong to Phase 6.
- Architecture policy and dependency rules belong to Phases 7 and 8.
- Full iterative fail/pass evidence completion contract belongs to Phase 9.

</deferred>

---
*Phase: 04-coverage-integrity-enforcement*
*Context gathered: 2026-02-28*
