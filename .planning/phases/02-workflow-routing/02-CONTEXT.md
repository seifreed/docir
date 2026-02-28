# Phase 2: Workflow Routing - Context

**Gathered:** 2026-02-28
**Status:** Ready for planning

<domain>
## Phase Boundary

Route all routine quality workflows (local development guidance, pre-commit flow, and CI required checks) through the canonical gate command only: `./scripts/quality_gate.sh`. This phase defines invocation paths and enforcement wiring, not additional quality checks.

</domain>

<decisions>
## Implementation Decisions

### Local workflow routing
- Local quality workflow documentation must use `./scripts/quality_gate.sh` as the sole acceptance command.
- Raw commands (`cargo fmt`, `cargo clippy`, `cargo test`) may be documented only as non-authoritative diagnostics.
- Any existing docs that imply equivalent alternate acceptance paths must be normalized to canonical-only wording.

### Pre-commit routing
- Add a repository-managed pre-commit hook path that invokes `./scripts/quality_gate.sh` directly from repo root.
- Hook setup should be deterministic and tool-minimal (native git hook installation script/documentation first; no extra hook framework dependency unless required).
- Hook behavior should fail commit on non-zero gate result and surface the gate result line unchanged.

### CI routing and required check contract
- Add CI workflow that executes only `./scripts/quality_gate.sh` for quality acceptance.
- CI job naming must be stable and intended for branch-protection required-check configuration.
- CI workflow must avoid parallel alternate quality jobs that could be interpreted as equivalent acceptance gates.

### Claude's Discretion
- Exact script names for hook installation helpers, as long as they do not become alternate accepted gate entrypoints.
- CI matrix breadth (single OS vs expanded matrix) if canonical gate invocation remains singular and deterministic.
- Exact README/docs placement for workflow instructions.

</decisions>

<specifics>
## Specific Ideas

- Keep user-facing command examples copy-paste friendly and always rooted at repository root.
- Include a short “why” note wherever raw commands appear: fast feedback only, not acceptance authority.
- CI workflow should preserve and show the final `QUALITY_GATE_RESULT=...` line to simplify required-check triage.

</specifics>

<code_context>
## Existing Code Insights

### Reusable Assets
- `scripts/quality_gate.sh`: Existing canonical entrypoint with deterministic pass/fail classes.
- `scripts/tests/quality_gate_contract.sh`: Existing canonical-path uniqueness contract check that can be reused in routing validation.
- `scripts/tests/quality_gate_exit_codes.sh`: Existing deterministic exit semantics test for workflow confidence.
- `docs/quality-gate-policy.md`: Existing policy baseline for canonical-only acceptance.

### Established Patterns
- Canonical gate emits final machine-parsable result line (`QUALITY_GATE_RESULT=... CLASS=... EXIT_CODE=...`).
- Helper scripts under `scripts/tests/` are non-executable by default and used as verification artifacts, not entrypoints.
- Documentation already distinguishes canonical acceptance from non-authoritative raw checks.

### Integration Points
- README quality workflow section is the local developer entrypoint for route guidance.
- `docs/quality-gate-policy.md` is the normative policy location for allowed/forbidden routing patterns.
- `.github/workflows/` is currently absent and should be created for CI required-check wiring.
- `.git/hooks` currently contains only sample hooks; repository should provide explicit install/update instructions for active pre-commit routing.

</code_context>

<deferred>
## Deferred Ideas

- Expanding CI into multi-job quality decomposition (e.g., separate lint/test/coverage jobs) while retaining canonical-only acceptance.
- Introducing third-party hook managers (pre-commit, lefthook, husky) if native hook routing proves insufficient.
- Additional workflow metrics/reporting dashboards beyond required-check pass/fail routing.

</deferred>

---
*Phase: 02-workflow-routing*
*Context gathered: 2026-02-28*
