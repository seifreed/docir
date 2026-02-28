# Phase 1 Research: Canonical Gate Surface

## Scope and Intent

Phase 1 must establish a single, deterministic quality-gate surface at `./scripts/quality_gate.sh` and prohibit alternate quality entrypoints. This phase defines the contract and enforcement mechanics; deeper check content is handled in later phases.

Mapped requirements: `GATE-01`, `GATE-02`, `GATE-06`.

## Gate Contract (Normative)

### Canonical Entrypoint

- The only accepted quality gate command is:
  - `./scripts/quality_gate.sh`
- Any local docs, CI jobs, helper targets, or hooks that perform quality validation must call this script, not duplicate quality logic.

### Deterministic Exit Behavior

- Exit `0` only when every step selected by the gate contract passes.
- Exit non-zero when any step fails, including preflight errors.
- Exit code semantics for planning:
  - `0`: full pass
  - `1`: quality check failure
  - `2`: invocation/precondition/environment failure (missing tool, wrong directory, non-executable script, etc.)
- Behavior must be deterministic for identical inputs (same commit, lockfile, toolchain, and env).

### Deterministic Execution Semantics

- Fixed, documented step order. No dynamic reordering by environment.
- Strict shell mode in script (`set -euo pipefail`) to avoid partial-success ambiguity.
- No mutating checks in canonical run (for example, no `--fix` modes).
- No “best effort” continuation after mandatory step failure.

### User/CI Output Contract

- Script prints each stage start, pass/fail result, and failing command.
- Last line is a machine-parsable final status line (for example: `QUALITY_GATE_RESULT=PASS` or `QUALITY_GATE_RESULT=FAIL`).
- Failure output must identify whether failure is check-related (`1`) or invocation-related (`2`).

## Non-Bypass Enforcement Patterns

### Repository-Level Patterns

- Single gate file path exists: `scripts/quality_gate.sh`.
- No additional `scripts/*gate*`, `scripts/*check*`, or parallel “quality” wrappers that can serve as substitute acceptance paths.
- Any helper scripts (if needed later) are internal-only and only callable from canonical gate.

### CI Enforcement Patterns

- Required CI quality job executes only `./scripts/quality_gate.sh`.
- CI workflow must not call raw `cargo fmt/clippy/test/...` directly as merge gate.
- Branch protection should require that single canonical job status.

### Local Workflow Patterns

- README/developer docs state canonical command only for final acceptance.
- If pre-commit hooks are added later, they shell out to canonical gate (or strict subset + mandatory full gate before push, documented explicitly).
- Any optional fast commands are documented as non-authoritative and never equivalent to gate pass.

### Governance/Policy Patterns

- Add a lightweight static check (in later phases) that scans repo workflows/docs for non-canonical quality invocations.
- Treat introduction of alternate quality entrypoints as policy violation.

## Repo-Specific Risks and Edge Cases

### Current Repo Reality

- `scripts/` is currently empty; canonical gate is not yet implemented.
- Workspace is multi-crate Rust with heavy parser/test surface (`docir-parser` dominates LOC), so runtime and failure surfacing must stay explicit.
- Existing docs/reference content already mentions canonical gate intent, but enforcement is not yet wired.

### Key Risks for Phase 1

- Accidental alternate entrypoints via helper scripts, Make targets, CI YAML jobs, or docs snippets.
- Non-deterministic behavior from unpinned/local environment differences.
- Ambiguous exit status if shell script uses permissive error handling.
- CI drift: job uses a different command set than local script.

### Edge Cases to Plan For

- Script invoked from non-repo working directory.
- Missing execute permission on `scripts/quality_gate.sh`.
- Missing required tools (`cargo`, future `cargo-llvm-cov`, etc.).
- Unsupported shell assumptions in CI runners.
- Invocation from hooks/CI that alter environment variables and accidentally bypass strict mode.

## Practical Implementation Guidance for Planning

### Minimal Phase-1 Implementation Shape

- Create canonical script skeleton with:
  - strict shell mode
  - preflight checks
  - deterministic stage wrapper
  - explicit result mapping (`0/1/2`)
- Include placeholder/initial mandatory checks sufficient to validate contract behavior now; deeper checks expand in later phases.

### Determinism Controls to Include Early

- Validate execution from repo root or normalize to repo root in script.
- Print version/toolchain metadata at start for reproducibility diagnostics.
- Keep stage list stable and centralized in one place in script.

### Non-Bypass Controls to Include Early

- CI quality workflow (Phase 2 wiring) should be designed now to call only canonical script.
- Avoid adding temporary parallel scripts during rollout; if unavoidable, enforce explicit removal task with deadline in plan.

## Concrete Deliverables for Planning (Phase 1)

1. `scripts/quality_gate.sh` contract spec
- Define accepted invocation, exit code table, deterministic ordering rules, and output schema.

2. Canonical gate script skeleton task breakdown
- Script creation, strict mode, preflight checks, stage runner, final status line, and executable bit.

3. Non-bypass inventory task
- Repository scan checklist for alternate quality entrypoints across `scripts/`, docs, and workflow files.

4. Deterministic behavior verification plan
- Test matrix for pass/fail/precondition states with expected exit codes and expected final status line.

5. Documentation update tasks
- Add canonical gate usage to README/developer workflow docs with explicit “single accepted gate” language.

6. CI routing preparation notes for next phase
- Define required job name and rule that job executes only canonical script (implemented in Phase 2, but specified now).

7. Acceptance checklist for phase completion
- `GATE-01`: canonical file exists at exact path.
- `GATE-02`: pass/fail/precondition exit behavior verified.
- `GATE-06`: no substitute quality entrypoint remains in repo policy surface.

## Suggested Plan Granularity for 01-PLAN

- Plan A: Gate contract and script bootstrap.
- Plan B: Deterministic exit behavior tests and evidence capture.
- Plan C: Non-bypass surface audit and documentation alignment.

## Definition of Phase-1 Done (Research Interpretation)

- One canonical gate entrypoint exists and is executable.
- Exit behavior is deterministic and contractually documented (`0` pass, non-zero fail, with explicit precondition handling).
- Repo policy surface does not present an alternative accepted quality gate path.
- Evidence exists showing at least one pass and one fail outcome under canonical invocation.
