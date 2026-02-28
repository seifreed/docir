# Phase 3: Baseline Clean Code Commands - Context

**Gathered:** 2026-02-28
**Status:** Ready for planning

<domain>
## Phase Boundary

Integrate baseline clean-code command enforcement into the canonical gate so acceptance depends on a single run of formatting, linting, and tests with warning-strict behavior. This phase is limited to `CC-01`, `CC-02`, `CC-03`, and `TEST-03`.

</domain>

<decisions>
## Implementation Decisions

### Canonical stage integration
- `./scripts/quality_gate.sh` remains the sole acceptance entrypoint and must orchestrate all Phase 3 checks.
- Baseline checks are implemented as explicit deterministic stages in the canonical gate, not as separate accepted scripts.
- Stage order is locked for diagnosability: `fmt --check` -> `clippy -D warnings` -> `cargo test`.

### Command strictness policy
- Formatting command is fixed to `cargo fmt --all --check`.
- Lint command is fixed to `cargo clippy --all-targets --all-features -- -D warnings`.
- Test command is fixed to `cargo test` for full workspace behavior validation.
- Any non-zero command result must propagate through canonical exit classification without suppression.

### Warning suppression posture
- No acceptance path may inject warning suppression flags (`-A`, env-based warning bypasses, or config-level suppression just to pass gate).
- Gate diagnostics may be verbose, but failure semantics must stay strict and deterministic.
- Existing fast/manual raw command usage remains diagnostic-only and never acceptance-authoritative.

### CI and hook alignment
- Pre-commit and CI continue invoking only canonical gate; they inherit new baseline checks automatically.
- No extra CI quality job is introduced for Phase 3 acceptance.
- Local docs may include raw command troubleshooting, but canonical gate stays normative.

### Claude's Discretion
- Exact implementation split between `scripts/quality_gate.sh` and `scripts/lib/quality_gate_lib.sh`.
- Exact log formatting for per-stage output and failure context.
- Whether to add helper test scripts for stage-level behavior, as long as they are non-authoritative.

</decisions>

<specifics>
## Specific Ideas

- Keep baseline stage names explicit and stable to simplify CI triage and failed-run diagnosis.
- Preserve the final machine-parseable line (`QUALITY_GATE_RESULT=...`) while adding optional stage-level detail.
- Prefer black-box tests for gate behavior (similar style to existing contract/exit-code scripts).

</specifics>

<code_context>
## Existing Code Insights

### Reusable Assets
- `scripts/quality_gate.sh`: canonical deterministic runner with pass/quality-failure/precondition-failure classes.
- `scripts/lib/quality_gate_lib.sh`: reusable stage helpers (`run_stage`, `classify_failure`, `emit_result`, tooling checks).
- `scripts/tests/quality_gate_contract.sh`: canonical-path integrity contract test.
- `scripts/tests/quality_gate_exit_codes.sh`: black-box exit semantics test scaffold.

### Established Patterns
- Single canonical acceptance command enforced in README/policy/hook/CI.
- Terminal result contract is machine-parseable and stable.
- Pre-commit (`.githooks/pre-commit`) and CI (`.github/workflows/quality-gate.yml`) already route through canonical gate.

### Integration Points
- `scripts/quality_gate.sh` default-stage pipeline must be extended for `fmt/clippy/test` execution.
- `scripts/lib/quality_gate_lib.sh` is the right place for command wrappers and consistent failure handling.
- Existing docs (`README.md`, `docs/quality-gate-policy.md`) must remain consistent with canonical-only acceptance semantics after new checks land.

</code_context>

<deferred>
## Deferred Ideas

- Coverage threshold enforcement (`cargo llvm-cov`, >=95%) belongs to Phase 4.
- Forbidden construct enforcement (`unwrap/expect/panic/todo/unimplemented`) belongs to Phase 5.
- Dead code/import/docs/complexity enforcement belongs to Phase 6.
- Architecture boundary enforcement belongs to Phases 7-8.

</deferred>

---
*Phase: 03-baseline-clean-code-commands*
*Context gathered: 2026-02-28*
