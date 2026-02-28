# Phase 2: Workflow Routing - Research

**Researched:** 2026-02-28
**Domain:** Workflow routing to canonical quality gate across local docs, Git hooks, and GitHub CI
**Confidence:** HIGH

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
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

### Deferred Ideas (OUT OF SCOPE)
- Expanding CI into multi-job quality decomposition (e.g., separate lint/test/coverage jobs) while retaining canonical-only acceptance.
- Introducing third-party hook managers (pre-commit, lefthook, husky) if native hook routing proves insufficient.
- Additional workflow metrics/reporting dashboards beyond required-check pass/fail routing.
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| GATE-03 | Local development quality workflow is documented and routed through the canonical gate only. | Local docs routing pattern, canonical-only wording contract, and anti-bypass documentation checks in this research. |
| GATE-04 | Pre-commit quality workflow is documented and routed through the canonical gate only. | Git native hook path (`core.hooksPath`), deterministic install script pattern, and pre-commit script behavior contract. |
| GATE-05 | CI required checks execute the canonical gate script directly. | GitHub workflow/job design guidance that runs only `./scripts/quality_gate.sh` and no equivalent alternate quality job. |
| FLOW-04 | CI marks canonical quality job as required for merge. | Required-check naming rules, branch protection/ruleset mapping guidance, and operational verification checklist. |
</phase_requirements>

## Summary

Phase 2 is primarily routing and governance work, not new quality logic. The canonical gate already exists at `./scripts/quality_gate.sh`; the phase must make every routine path (README guidance, pre-commit, and CI) invoke that exact command as the sole acceptance authority. The existing policy doc and README already state canonical intent, so implementation risk is concentrated in wiring consistency and merge-protection configuration.

Use native Git hooks with a repository-managed hooks directory (`core.hooksPath`) and a thin pre-commit hook that shells to `./scripts/quality_gate.sh` from repository root. This satisfies deterministic setup and avoids introducing a second toolchain. In CI, create one merge-blocking quality job with a stable job name and one execution command (`./scripts/quality_gate.sh`).

`FLOW-04` cannot be fully enforced by code in this repository alone: required-check selection lives in repository settings/rulesets. Planning must therefore include both (a) workflow file creation and (b) explicit configuration and verification steps proving the canonical job is marked required.

**Primary recommendation:** Implement one canonical routing chain: docs -> pre-commit hook -> GitHub required job, with every acceptance surface executing only `./scripts/quality_gate.sh`.

## Standard Stack

### Core
| Library/Tool | Version | Purpose | Why Standard |
|--------------|---------|---------|--------------|
| Bash (`scripts/quality_gate.sh`) | repo-local | Canonical quality acceptance command | Already implemented and policy-authorized gate surface in this repo. |
| Git hooks (`pre-commit`) | Git built-in | Local commit-time enforcement routing | Native, deterministic, no external dependency manager needed. |
| Git `core.hooksPath` | Git built-in | Repository-managed hook location | Standard way to redirect hooks from `.git/hooks` to tracked hook directory. |
| GitHub Actions workflow/job | GitHub platform | CI quality execution and required-check source | Native merge-blocking check surface tied to branch protection/rulesets. |

### Supporting
| Library/Tool | Version | Purpose | When to Use |
|--------------|---------|---------|-------------|
| `scripts/tests/quality_gate_contract.sh` | repo-local | Guard canonical-path uniqueness expectations | Use in verification plan to ensure no alternate executable gate-like scripts appear. |
| `scripts/tests/quality_gate_exit_codes.sh` | repo-local | Verify deterministic gate exit/result semantics | Use to ensure hook/CI consumers receive unchanged canonical result behavior. |
| `docs/quality-gate-policy.md` | repo-local | Normative non-bypass policy | Update when adding pre-commit and CI routing details. |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Native Git hooks + install script | Third-party hook managers (`pre-commit`, `lefthook`, `husky`) | Adds dependency/tooling overhead and an extra policy surface; deferred unless native path is insufficient. |
| Single canonical CI quality job | Split lint/test/format jobs | Better parallelism but creates potential equivalent acceptance surfaces and required-check ambiguity in this phase. |
| `core.hooksPath` tracked directory | Writing into `.git/hooks` directly | Harder to keep deterministic and versioned; not portable across clones without manual copy each time. |

## Architecture Patterns

### Recommended Project Structure
```text
.githooks/
└── pre-commit                 # calls ./scripts/quality_gate.sh only

scripts/
└── install_hooks.sh           # sets core.hooksPath to .githooks

.github/workflows/
└── quality-gate.yml           # single merge-blocking canonical gate job

docs/
└── quality-gate-policy.md     # policy updated with routing contract

README.md                      # local workflow points to canonical gate only
```

### Pattern 1: Canonical Invocation Wrapper (Hook)
**What:** Pre-commit hook is a thin wrapper that only executes canonical gate, with no duplicated quality commands.
**When to use:** Any local pre-commit enforcement.
**Example:**
```bash
#!/usr/bin/env bash
set -euo pipefail
repo_root="$(git rev-parse --show-toplevel)"
cd "${repo_root}"
exec ./scripts/quality_gate.sh
```

### Pattern 2: Deterministic Hook Installation
**What:** Use a repository script to set `core.hooksPath` to a tracked folder, then verify it.
**When to use:** Developer onboarding and docs for reproducible setup.
**Example:**
```bash
git config core.hooksPath .githooks
git config --get core.hooksPath
```

### Pattern 3: Single Canonical CI Job
**What:** One named job executes only the canonical gate and serves as required status check.
**When to use:** Merge-blocking CI quality enforcement.
**Example:**
```yaml
jobs:
  quality-gate:
    name: quality-gate
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - run: ./scripts/quality_gate.sh
```

### Anti-Patterns to Avoid
- **Parallel acceptance path in CI:** Any additional job that directly runs `cargo fmt/clippy/test` as an equivalent required check.
- **Hook logic duplication:** Copying quality commands into `pre-commit` rather than executing canonical gate.
- **Unstable required check names:** Renaming job names frequently; required checks map to check names and become brittle.
- **Non-versioned local hooks:** Manual `.git/hooks/pre-commit` edits that are not tracked or reproducible.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Hook orchestration framework | Custom hook manager layer | Native `pre-commit` hook + `core.hooksPath` | Built-in Git behavior already provides deterministic routing. |
| Alternate quality wrapper scripts | Extra `gate_fast`/`quality_check` accepted commands | `./scripts/quality_gate.sh` only | Additional entrypoints violate canonical gate policy and create bypass risk. |
| CI gate decomposition in this phase | Multi-job acceptance graph | Single canonical gate job | Reduces required-check ambiguity and enforces one acceptance authority. |

**Key insight:** Phase 2 succeeds by minimizing routing surfaces, not adding orchestration abstractions.

## Common Pitfalls

### Pitfall 1: Hook runs from wrong working directory
**What goes wrong:** Hook calls relative path and fails outside repo root or nested invocation contexts.
**Why it happens:** Assuming hook CWD without normalization.
**How to avoid:** Resolve repo root via `git rev-parse --show-toplevel` and `cd` before `exec ./scripts/quality_gate.sh`.
**Warning signs:** Hook passes locally in one shell path but fails for other contributors.

### Pitfall 2: Required check ambiguity blocks merges
**What goes wrong:** Branch protection requires a name that maps ambiguously across checks/workflows.
**Why it happens:** Non-unique or unstable job naming and parallel quality jobs.
**How to avoid:** Use one stable job name (`quality-gate`) and avoid equivalent quality check names.
**Warning signs:** PR UI shows unexpected required checks or unresolved required-check state.

### Pitfall 3: Documentation drift reintroduces alternate acceptance paths
**What goes wrong:** README/docs imply raw cargo commands are equivalent acceptance.
**Why it happens:** Incremental docs edits without policy alignment.
**How to avoid:** Keep explicit wording: raw commands are diagnostic only; acceptance is canonical gate.
**Warning signs:** Phrases like “equivalent to gate” or “either run gate or run fmt/clippy/test”.

### Pitfall 4: Assuming repo files alone satisfy FLOW-04
**What goes wrong:** CI workflow exists but merge protection is not configured to require canonical job.
**Why it happens:** Required checks are repository settings/ruleset configuration, not just YAML presence.
**How to avoid:** Add explicit setup + verification evidence in phase plan.
**Warning signs:** Workflow passes but PR can merge without canonical quality check required.

## Code Examples

Verified routing patterns from official docs and current repository policy:

### Repository-managed hook path
```bash
# one-time setup
./scripts/install_hooks.sh

# expected state
git config --get core.hooksPath
# => .githooks
```

### Pre-commit hook canonical handoff
```bash
#!/usr/bin/env bash
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"
exec ./scripts/quality_gate.sh
```

### CI merge-blocking canonical job
```yaml
name: quality-gate
on:
  pull_request:
  push:
    branches: [main]

jobs:
  quality-gate:
    name: quality-gate
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - run: ./scripts/quality_gate.sh
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| Multiple equivalent local/CI quality commands | Single canonical gate entrypoint | Phase 1 complete (2026-02-28) | Deterministic acceptance semantics and clearer governance. |
| Ad hoc `.git/hooks` local edits | Repo-tracked hooks via `core.hooksPath` | Standard Git capability (current) | Reproducible setup, reviewable hook logic, less onboarding drift. |
| CI job fan-out as acceptance authority | One canonical required job | Recommended for this phase | Prevents ambiguous required checks and bypass-like interpretations. |

**Deprecated/outdated for this phase:**
- Treating raw cargo commands as equivalent acceptance.
- Adding additional gate-like scripts as accepted entrypoints.

## Open Questions

1. **How will FLOW-04 be enforced operationally (branch protection rule or ruleset), and who owns configuration?**
   - What we know: Required status checks are configured in GitHub settings/rulesets, not in workflow YAML alone.
   - What's unclear: Exact governance path and owner permissions in this repository.
   - Recommendation: Add a plan item for post-YAML branch-protection/ruleset configuration plus evidence capture in phase verification.

2. **Should CI run only on `pull_request` or both `pull_request` and protected branch pushes?**
   - What we know: Either can support required checks; both improve observability.
   - What's unclear: Desired runtime/cost tradeoff for this repo.
   - Recommendation: Default to both (`pull_request` + `push` to `main`) unless project policy limits runner usage.

## Sources

### Primary (HIGH confidence)
- Local repository artifacts checked directly:
  - `scripts/quality_gate.sh`
  - `scripts/tests/quality_gate_contract.sh`
  - `scripts/tests/quality_gate_exit_codes.sh`
  - `docs/quality-gate-policy.md`
  - `README.md`
  - `.planning/phases/02-workflow-routing/02-CONTEXT.md`
  - `.planning/REQUIREMENTS.md`
  - `.planning/STATE.md`
- Git hooks documentation (official): https://git-scm.com/docs/githooks
- GitHub Actions workflow syntax (official): https://docs.github.com/en/actions/reference/workflows-and-actions/workflow-syntax
- GitHub required status check naming behavior (official): https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-rulesets/troubleshooting-rules

### Secondary (MEDIUM confidence)
- GitHub protected branches overview (official): https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-protected-branches
- GitHub troubleshooting required checks (official): https://docs.github.com/repositories/configuring-branches-and-merges-in-your-repository/defining-the-mergeability-of-pull-requests/troubleshooting-required-status-checks

### Tertiary (LOW confidence)
- None.

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - driven by existing repo implementation and official Git/GitHub platform docs.
- Architecture: HIGH - phase scope is narrowly constrained by locked decisions and existing canonical gate policy.
- Pitfalls: HIGH - derived from official required-check behavior and concrete repo state (no current workflow/hooks routing).

**Research date:** 2026-02-28
**Valid until:** 2026-03-30
