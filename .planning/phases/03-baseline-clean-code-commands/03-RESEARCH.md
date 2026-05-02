# Phase 3: Baseline Clean Code Commands - Research

**Researched:** 2026-02-28
**Domain:** Canonical quality-gate stage expansion for Rust formatting, linting, and test enforcement
**Confidence:** HIGH

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
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

### Deferred Ideas (OUT OF SCOPE)
- Coverage threshold enforcement (`cargo llvm-cov`, >=95%) belongs to Phase 4.
- Forbidden construct enforcement (`unwrap/expect/panic/todo/unimplemented`) belongs to Phase 5.
- Dead code/import/docs/complexity enforcement belongs to Phase 6.
- Architecture boundary enforcement belongs to Phases 7-8.
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| CC-01 | Gate enforces `cargo fmt --all --check`. | Adds explicit deterministic fmt stage in canonical stage list with strict non-zero propagation. |
| CC-02 | Gate enforces `cargo clippy --all-targets --all-features -- -D warnings`. | Adds clippy stage with hard `-D warnings`; no bypass flags accepted in gate path. |
| CC-03 | Gate enforces `cargo test`. | Adds full workspace `cargo test` stage in canonical gate sequence. |
| TEST-03 | Warnings and lint checks remain fully enabled; warning suppression is not used to pass the gate. | Enforces clippy warnings-as-errors and preserves policy that suppression flags/config are not acceptance-authoritative. |
</phase_requirements>

## Summary

Phase 3 is an extension of an already-stable canonical gate shell architecture. The current gate already provides deterministic stage execution (`run_stage`), exit class mapping (`classify_failure`), and machine-parseable terminal output (`QUALITY_GATE_RESULT=... CLASS=... EXIT_CODE=...`). Planning should therefore avoid new orchestration surfaces and focus on replacing the current scaffold stage with three explicit baseline stages in locked order.

Because pre-commit and CI already route only through `./scripts/quality_gate.sh`, adding baseline stages automatically propagates enforcement to local commits and merge checks without workflow topology changes. The implementation risk is concentrated in failure semantics and diagnosability: each command must fail fast, map to quality failure (exit 1), and preserve deterministic logging and final status line behavior.

The strongest planning approach is to keep the stage runner pattern and existing black-box test style, then expand the exit-contract tests to cover the real baseline stages. This maintains non-bypass guarantees while validating the exact acceptance semantics required by `CC-01`, `CC-02`, `CC-03`, and `TEST-03`.

**Primary recommendation:** Implement baseline checks as first-class canonical stages (`fmt_check`, `clippy_strict`, `test_workspace`) executed in locked order via existing `run_stage`/`classify_failure`, and validate behavior with black-box gate contract tests.

## Standard Stack

### Core
| Library/Tool | Version | Purpose | Why Standard |
|--------------|---------|---------|--------------|
| Bash canonical gate (`scripts/quality_gate.sh`) | repo-local | Single acceptance orchestrator | Already required by `GATE-01..06` and integrated into hook/CI paths. |
| Gate support library (`scripts/lib/quality_gate_lib.sh`) | repo-local | Shared stage execution/logging/failure classification | Existing stable primitive for deterministic stage handling. |
| `cargo fmt --all --check` | rustfmt via toolchain | Formatting drift detection without mutation | Canonical Rust formatter check mode; deterministic pass/fail signal. |
| `cargo clippy --all-targets --all-features -- -D warnings` | clippy via toolchain | Lint + warnings-as-errors enforcement | Directly satisfies strict warning posture and `TEST-03`. |
| `cargo test` | cargo test runner | Behavioral validation across workspace crates | Required baseline quality signal for `CC-03`. |

### Supporting
| Library/Tool | Version | Purpose | When to Use |
|--------------|---------|---------|-------------|
| `scripts/tests/quality_gate_exit_codes.sh` | repo-local | Black-box exit/result contract verification | Extend for stage-driven fail/pass semantics after baseline integration. |
| `scripts/tests/quality_gate_contract.sh` | repo-local | Canonical gate-path non-bypass guard | Keep running unchanged to prevent alternate accepted gate surfaces. |
| GitHub Actions job `quality-gate` | current repo workflow | Merge check execution surface | Already invokes only canonical gate; inherits Phase 3 checks automatically. |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Canonical gate stages calling raw cargo commands | Separate script per command (fmt/clippy/test) as accepted paths | Violates canonical-only acceptance and increases bypass surface. |
| `cargo clippy ... -D warnings` in gate | Relaxed clippy without denied warnings | Fails `TEST-03` and allows warning-based drift. |
| Fail-fast ordered stages | Parallel checks | Faster runtime but weaker diagnosability and more complex deterministic logging contracts. |

## Architecture Patterns

### Recommended Project Structure
```text
scripts/
├── quality_gate.sh                 # canonical stage sequence and final status
├── lib/
│   └── quality_gate_lib.sh         # run_stage, logging, failure classification
└── tests/
    ├── quality_gate_contract.sh    # non-bypass surface checks
    └── quality_gate_exit_codes.sh  # canonical exit/result contract checks
```

### Pattern 1: Stage Function + Dispatcher
**What:** Each baseline check is represented by a dedicated stage function and dispatched centrally.
**When to use:** Adding required checks while preserving deterministic orchestration.
**Example:**
```bash
stage_fmt_check() { cargo fmt --all --check; }
stage_clippy_strict() { cargo clippy --all-targets --all-features -- -D warnings; }
stage_test_workspace() { cargo test; }
```

### Pattern 2: Fail-Fast Canonical Loop with Exit Classification
**What:** Run stages in locked order; stop on first failure; map failures through one classifier.
**When to use:** Canonical acceptance semantics where one failing required check must fail the gate.
**Example:**
```bash
for stage in validate_repo_root validate_tooling fmt_check clippy_strict test_workspace; do
  run_stage "$stage" dispatch_stage "$stage" || { gate_exit="$(classify_failure "$?")"; break; }
done
```

### Pattern 3: Black-Box Contract Verification
**What:** Validate gate behavior from process boundary (exit code + final result line), not internal functions.
**When to use:** Ensuring canonical semantics remain stable during stage additions.
**Example:**
```bash
./scripts/quality_gate.sh
# verify final line includes QUALITY_GATE_RESULT=PASS CLASS=pass EXIT_CODE=0
```

### Anti-Patterns to Avoid
- **Suppression in canonical path:** Adding `-A`, `RUSTFLAGS` warning bypasses, or clippy allow overrides to force green gate.
- **Stage logic duplication in hooks/CI:** Re-encoding fmt/clippy/test directly in `.githooks` or workflow YAML.
- **Mutating formatter mode in gate:** Using `cargo fmt` (write mode) instead of `--check` for acceptance.
- **Stage-order drift:** Reordering stages ad hoc and making diagnostics inconsistent across runs.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Stage orchestration framework | New task runner around canonical gate | Existing `run_stage` + dispatcher loop | Existing implementation already provides deterministic stage logs and return handling. |
| Warning-policy bypass control | Custom post-processing to ignore warnings | Native `clippy ... -- -D warnings` failure semantics | Native behavior is strict and less error-prone than output parsing. |
| Acceptance verdict parser | Extra parser script to determine pass/fail | Existing final status contract line + process exit code | Current consumers (tests/hooks/CI) already rely on this stable contract. |

**Key insight:** Phase 3 should extend current canonical primitives, not introduce new execution layers.

## Common Pitfalls

### Pitfall 1: Missing stage-level failure isolation
**What goes wrong:** A failing cargo command does not cleanly map to gate exit class or logs become ambiguous.
**Why it happens:** Calling commands outside `run_stage` wrapper or bypassing `classify_failure`.
**How to avoid:** Keep all required checks inside `run_stage` + centralized failure mapping.
**Warning signs:** Missing `STAGE_FAIL <name>` lines or incorrect final class/exit code.

### Pitfall 2: Implicit warning suppression via environment/config
**What goes wrong:** Gate passes despite warnings due to injected suppression knobs.
**Why it happens:** CI/local environment or config adds warning-allow behavior.
**How to avoid:** Keep canonical clippy command fixed with `-D warnings` and avoid acceptance-time suppression flags/config.
**Warning signs:** Clippy warnings present in output while canonical run still passes.

### Pitfall 3: Over-broad precondition classification
**What goes wrong:** Quality failures (fmt/clippy/test) incorrectly map to precondition failure exit class.
**Why it happens:** Stage return handling conflates command failures with setup failures.
**How to avoid:** Preserve `classify_failure` contract where only explicit precondition signals map to exit class 2.
**Warning signs:** Final line shows `CLASS=precondition_failure` for ordinary lint/test failures.

### Pitfall 4: Testability gaps for real baseline stages
**What goes wrong:** Contract tests only exercise synthetic fail env vars, not baseline command path semantics.
**Why it happens:** Legacy scaffold tests not updated after stage additions.
**How to avoid:** Add/adjust black-box tests to assert baseline stage presence/order and strict failure outcomes.
**Warning signs:** Phase merged without evidence of real fmt/clippy/test failure-case validation.

## Code Examples

Verified patterns from current repository scripts:

### Stage wrapper contract (`run_stage`)
```bash
run_stage() {
  local stage="$1"
  shift

  gate_log "INFO" "STAGE_START ${stage}"
  set +e
  "$@"
  local exit_code=$?
  set -e

  if [ "$exit_code" -eq 0 ]; then
    gate_log "INFO" "STAGE_PASS ${stage}"
  else
    gate_log "ERROR" "STAGE_FAIL ${stage} exit_code=${exit_code}"
  fi

  return "$exit_code"
}
```
Source: `scripts/lib/quality_gate_lib.sh`

### Failure class mapping contract (`classify_failure`)
```bash
classify_failure() {
  local exit_code="$1"
  case "$exit_code" in
    2) echo 2 ;;
    *) echo 1 ;;
  esac
}
```
Source: `scripts/lib/quality_gate_lib.sh`

### Final machine-parseable output contract
```bash
printf 'QUALITY_GATE_RESULT=%s CLASS=%s EXIT_CODE=%s\n' "$status" "$class" "$gate_exit"
```
Source: `scripts/quality_gate.sh`

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| Canonical gate scaffold with synthetic quality fail stage | Canonical gate with real baseline clean-code stages | Phase 3 target (current phase) | Converts routing contract into substantive baseline enforcement. |
| Raw cargo commands used as local validation evidence | Canonical gate run as only acceptance authority | Established in Phases 1-2 (2026-02-28) | Removes bypass ambiguity across local, hook, and CI paths. |
| Optional warning handling by convention | Explicit warnings-as-errors via clippy `-D warnings` | Phase 3 target | Makes warning posture objectively enforceable. |

**Deprecated/outdated for this phase:**
- Treating scaffold-only quality gate pass as evidence of clean-code enforcement.
- Treating manual raw command runs as equivalent acceptance proof.

## Open Questions

1. **How should baseline-stage failure tests be made deterministic without introducing accepted bypass scripts?**
   - What we know: Existing exit-contract tests already use controlled env variables for deterministic fail classes.
   - What's unclear: Whether to add stage-specific fail-injection env knobs (test-only) or use fixture-based intentional failure scenarios.
   - Recommendation: Prefer stage-specific deterministic test hooks inside canonical script (non-authoritative) to avoid mutating repository files for test setup.

2. **Should tooling preconditions check for rustfmt/clippy component availability explicitly?**
   - What we know: `validate_tooling` currently checks only `cargo` availability.
   - What's unclear: Desired UX for missing component errors (precondition failure vs quality failure classification).
   - Recommendation: Decide explicitly in plan and codify expected exit class for missing formatter/linter components.

3. **What evidence artifact will satisfy FLOW-02 later while preserving Phase 3 scope?**
   - What we know: This phase must prove fail/pass behavior for fmt/clippy/test checks; FLOW-02 formal completion is Phase 9.
   - What's unclear: Where fail/pass transcripts should be stored now for future completion proof reuse.
   - Recommendation: Capture reproducible command transcripts in phase summary to prevent rework in final enforcement phase.

## Sources

### Primary (HIGH confidence)
- Local repository artifacts checked directly:
  - `.planning/phases/03-baseline-clean-code-commands/03-CONTEXT.md`
  - `.planning/REQUIREMENTS.md`
  - `.planning/STATE.md`
  - `scripts/quality_gate.sh`
  - `scripts/lib/quality_gate_lib.sh`
  - `scripts/tests/quality_gate_contract.sh`
  - `scripts/tests/quality_gate_exit_codes.sh`
  - `.github/workflows/quality-gate.yml`
  - `.githooks/pre-commit`
  - `docs/quality-gate-policy.md`
  - `README.md`

### Secondary (MEDIUM confidence)
- Rustfmt official documentation: https://github.com/rust-lang/rustfmt
- Rust Clippy official documentation: https://doc.rust-lang.org/clippy/
- Cargo test command reference: https://doc.rust-lang.org/cargo/commands/cargo-test.html

### Tertiary (LOW confidence)
- None.

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - commands and orchestration are explicitly locked by CONTEXT.md and already routed in repository.
- Architecture: HIGH - integration points are concrete and already implemented (canonical gate + shared lib + black-box tests).
- Pitfalls: MEDIUM - primary risks are operational/testing details that depend on final plan choices.

**Research date:** 2026-02-28
**Valid until:** 2026-03-30
