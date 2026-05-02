# Phase 4: Coverage Integrity Enforcement - Research

**Researched:** 2026-02-28
**Domain:** Canonical coverage gate enforcement (`cargo llvm-cov`) in shell/CI workflows
**Confidence:** HIGH

<user_constraints>
## User Constraints (from CONTEXT.md)

### Implementation Decisions

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

## Deferred Ideas

- Forbidden construct scanning (`unwrap/expect/panic/todo/unimplemented`) belongs to Phase 5.
- Dead code/import/docs/complexity hard gates belong to Phase 6.
- Architecture policy and dependency rules belong to Phases 7 and 8.
- Full iterative fail/pass evidence completion contract belongs to Phase 9.
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| CC-04 | Gate enforces test coverage of at least 95% using `cargo llvm-cov`. | Add deterministic coverage stage to canonical gate with machine-enforced `--fail-under-lines 95`. |
| TEST-01 | Coverage target is measured in canonical run and cannot be skipped in CI. | Keep single required CI job (`quality-gate`) invoking only canonical gate; no alternate required quality path. |
| TEST-02 | Coverage-passing tests validate real behavior, not trivial inflation. | Use existing fixture/behavior tests as the source of additional coverage; add anti-inflation checks and review criteria. |
</phase_requirements>

## Summary

This phase should extend the existing deterministic stage runner rather than introduce new orchestration. The current canonical gate already has stable fail-fast behavior, centralized stage dispatch, and machine-parseable final output. The safest implementation is to append one coverage stage after `fmt`, `clippy`, and `test`, then preserve existing `classify_failure` semantics so any coverage miss remains `quality_failure`.

`cargo llvm-cov` already supports hard threshold enforcement (`--fail-under-lines`) and summary reporting, so no custom coverage parser is needed. Planning should also include tooling preconditions (`cargo llvm-cov` availability and LLVM toolchain components) to avoid false negatives from environment drift between local and CI.

For integrity (`TEST-02`), the project already contains behavior-heavy parser/CLI fixture tests. The phase plan should prioritize expanding/using those behavioral tests where gaps exist, and explicitly reject low-value inflation patterns (tautological assertions, tests that only execute getters/setters without validating security or parsing outcomes).

**Primary recommendation:** Add a canonical `coverage_check` stage using `cargo llvm-cov` with hard `>=95%` line threshold, enforce it in the single required `quality-gate` CI job, and validate integrity via black-box gate contract tests plus non-trivial behavior-test criteria.

## Standard Stack

### Core
| Library/Tool | Version | Purpose | Why Standard |
|--------------|---------|---------|--------------|
| Bash canonical gate (`scripts/quality_gate.sh`) | repo-local | Single acceptance orchestrator | Already mandatory acceptance surface (`GATE-01..06`) and wired in pre-commit/CI. |
| Gate support library (`scripts/lib/quality_gate_lib.sh`) | repo-local | Stage lifecycle + failure classification | Existing deterministic primitive; coverage should reuse it. |
| `cargo llvm-cov` | `0.6.21` (installed locally), `0.8.4` current crates.io | Coverage measurement and threshold enforcement | Native Rust ecosystem standard for LLVM source-based coverage with hard fail-under flags. |
| GitHub Actions `quality-gate` job | repo-local workflow | Required merge check | Already canonical CI route; must remain single authoritative quality check. |

### Supporting
| Library/Tool | Version | Purpose | When to Use |
|--------------|---------|---------|-------------|
| `scripts/tests/quality_gate_baseline_commands.sh` | repo-local | Stage order + command-contract validation | Extend expected call sequence to include coverage stage and fail-fast behavior. |
| `scripts/tests/quality_gate_exit_codes.sh` | repo-local | Exit/status contract verification | Add coverage fail-path and pass-path assertions. |
| `taiki-e/install-action` | current | Deterministic CI install for `cargo-llvm-cov` | Use if CI image does not already provide correct version/toolchain components. |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| `cargo llvm-cov --fail-under-lines 95` in canonical gate | Post-process textual reports with custom shell parsing | Adds brittle parsing logic and ambiguity in pass/fail semantics. |
| Canonical coverage stage inside `quality_gate.sh` | Separate CI-only coverage job | Violates canonical-only acceptance and weakens local/CI parity. |
| Behavioral fixture-driven coverage gains | Synthetic tests targeting easy lines | Increases metric without confidence in real malware-analysis behavior. |

## Architecture Patterns

### Recommended Project Structure
```text
scripts/
├── quality_gate.sh                      # add coverage_check stage after test_workspace
├── lib/quality_gate_lib.sh              # unchanged core stage primitives
└── tests/
    ├── quality_gate_baseline_commands.sh  # include coverage command in expected sequence
    ├── quality_gate_exit_codes.sh         # include coverage fail/pass contract checks
    └── quality_gate_contract.sh           # canonical path uniqueness unchanged

.github/workflows/
└── quality-gate.yml                     # still runs only ./scripts/quality_gate.sh
```

### Pattern 1: Canonical Coverage Stage (Fail-Fast)
**What:** Introduce `stage_coverage_check` in canonical stage list after baseline checks.
**When to use:** Always in acceptance flow (local, pre-commit, CI) because coverage is non-optional.
**Example:**
```bash
stage_coverage_check() {
  gate_run_command cargo llvm-cov --workspace --all-features --fail-under-lines 95 --summary-only
}
```

### Pattern 2: Single Source of Truth for Threshold
**What:** Keep threshold in one variable used by command, logs, and tests.
**When to use:** Any thresholded gate to avoid drift across scripts/docs/tests.
**Example:**
```bash
readonly COVERAGE_THRESHOLD=95
gate_run_command cargo llvm-cov ... --fail-under-lines "${COVERAGE_THRESHOLD}" ...
```

### Pattern 3: Contract-First Shell Testing
**What:** Assert stage order, final result line, and exit-class behavior through black-box scripts.
**When to use:** Any gate-stage addition that changes canonical behavior.
**Example:**
```bash
# expected call order includes llvm-cov after test
fmt -> clippy -> test -> llvm-cov
```

### Anti-Patterns to Avoid
- **Custom report scraping:** parsing human-readable percentages instead of using `--fail-under-lines` exit semantics.
- **CI-only coverage enforcement:** measuring coverage outside canonical gate or in optional/non-required job.
- **Coverage inflation tests:** tests that execute code paths without validating parser/security outcomes.
- **Undocumented exclusions:** excluding files/paths without explicit rationale and reproducibility in both local and CI runs.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Coverage threshold engine | Shell/awk percent parser | `cargo llvm-cov --fail-under-lines` | Built-in deterministic failure semantics are simpler and less error-prone. |
| Parallel acceptance paths | Separate accepted CI jobs or helper scripts | Single canonical `./scripts/quality_gate.sh` | Preserves non-bypass contract and requirement traceability. |
| Integrity theater checks | Superficial test-count targets | Behavior assertions tied to parser/security outputs | Directly aligns with `TEST-02` intent. |

**Key insight:** This phase is mostly a contract extension of existing gate infrastructure; risk comes from toolchain/setup drift and weak integrity criteria, not from shell architecture.

## Common Pitfalls

### Pitfall 1: Toolchain mismatch breaks canonical run
**What goes wrong:** `cargo llvm-cov` missing or incompatible in CI/local; gate fails as precondition noise.
**Why it happens:** Tool not installed/pinned; LLVM component assumptions differ by environment.
**How to avoid:** Explicitly verify/install `cargo llvm-cov` in CI and document local setup path.
**Confidence:** HIGH

### Pitfall 2: Coverage threshold drift across scripts/tests/docs
**What goes wrong:** gate uses one threshold, tests/docs assume another.
**Why it happens:** Hardcoded repeated values.
**How to avoid:** Centralize threshold constant and assert it in shell contract tests.
**Confidence:** HIGH

### Pitfall 3: Stale coverage artifacts produce confusing results
**What goes wrong:** runs include stale profile data and produce non-reproducible percentages.
**Why it happens:** mixed raw `cargo` and `cargo llvm-cov` builds without clean strategy.
**How to avoid:** follow `cargo llvm-cov` recommended clean workflow in canonical stage design/documentation.
**Confidence:** MEDIUM

### Pitfall 4: TEST-02 reduced to metric gaming
**What goes wrong:** coverage passes, but tests do not validate meaningful parser/security behavior.
**Why it happens:** team optimizes for percentage only.
**How to avoid:** define anti-inflation acceptance checks (fixture-driven expectations, non-trivial assertions, changed-behavior proof in PR evidence).
**Confidence:** MEDIUM

## Code Examples

Verified patterns from current repository scripts:

### Existing deterministic stage runner
```bash
run_stage() {
  local stage="$1"
  shift
  gate_log "INFO" "STAGE_START ${stage}"
  "$@"
}
```
Source: `scripts/lib/quality_gate_lib.sh`

### Existing final canonical result contract
```bash
printf 'QUALITY_GATE_RESULT=%s CLASS=%s EXIT_CODE=%s\n' "$status" "$class" "$gate_exit"
```
Source: `scripts/quality_gate.sh`

### `cargo llvm-cov` threshold option (official)
```bash
cargo llvm-cov --fail-under-lines MIN
```
Source: `cargo-llvm-cov` docs/README.

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| Canonical gate enforces fmt/clippy/test only | Canonical gate additionally enforces coverage threshold | Phase 4 target | Completes baseline clean-code gate with measurable test-depth requirement. |
| Coverage discussed as policy intent | Coverage enforced by gate exit semantics | Phase 4 target | Removes ambiguity and prevents CI/local skip paths. |
| Coverage percentage as sole success signal | Coverage + meaningful behavior validation requirement | Phase 4 target (`TEST-02`) | Reduces incentive for trivial inflation. |

## Open Questions

1. **Coverage scope decision:** should Phase 4 enforce `line` coverage only (`--fail-under-lines 95`) or include `region`/`function` thresholds too?
   - Recommendation: lock line coverage for Phase 4 (matches requirement wording), defer multi-metric hardening unless explicitly requested.

2. **Version pinning strategy for CI:** install latest `cargo-llvm-cov` each run vs pin explicit version (for reproducibility).
   - Recommendation: pin version in CI or use stable install action config, then update intentionally.

3. **Evidence artifact format:** where to store canonical coverage evidence (`COVERAGE_PERCENT=...`) for repeatable audits.
   - Recommendation: emit machine-parseable line in gate output and optionally persist summary artifact in CI job.

## Sources

### Primary (HIGH confidence)
- Local repository artifacts checked directly:
  - `.planning/phases/04-coverage-integrity-enforcement/04-CONTEXT.md`
  - `.planning/REQUIREMENTS.md`
  - `.planning/STATE.md`
  - `scripts/quality_gate.sh`
  - `scripts/lib/quality_gate_lib.sh`
  - `scripts/tests/quality_gate_baseline_commands.sh`
  - `scripts/tests/quality_gate_exit_codes.sh`
  - `scripts/tests/quality_gate_contract.sh`
  - `.github/workflows/quality-gate.yml`
  - `docs/quality-gate-policy.md`
  - `docs/pre-commit-quality-workflow.md`
  - `docs/ci-required-quality-check.md`

### Secondary (HIGH-MEDIUM confidence)
- cargo-llvm-cov repository and docs:
  - https://github.com/taiki-e/cargo-llvm-cov
  - https://github.com/taiki-e/cargo-llvm-cov/blob/main/README.md
  - https://docs.rs/crate/cargo-llvm-cov/latest
- cargo-llvm-cov latest crate metadata:
  - https://crates.io/crates/cargo-llvm-cov
- GitHub Action for installing cargo subcommands:
  - https://github.com/taiki-e/install-action

### Tertiary (LOW confidence)
- None.

## Metadata

**Confidence breakdown:**
- Canonical integration pattern: HIGH (directly matches existing script/test architecture).
- Coverage command semantics: HIGH (official cargo-llvm-cov docs and crate metadata).
- Anti-inflation enforcement mechanics: MEDIUM (repository-specific policy/process decisions still needed).

**Research date:** 2026-02-28
**Valid until:** 2026-03-31
