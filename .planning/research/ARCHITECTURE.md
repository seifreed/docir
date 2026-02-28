# Architecture Research

**Domain:** Rust workspace quality-gate architecture (strict, deterministic enforcement)
**Researched:** 2026-02-28
**Confidence:** HIGH

## Standard Architecture

### System Overview

```text
┌──────────────────────────────────────────────────────────────────────────────┐
│ Interface Layer                                                              │
├──────────────────────────────────────────────────────────────────────────────┤
│  ┌──────────────────────┐  ┌────────────────────────────┐                   │
│  │ Developer Workflows  │  │ CI Required Job            │                   │
│  │ (local/pre-commit)   │  │ (merge-blocking gate run)  │                   │
│  └──────────┬───────────┘  └──────────────┬─────────────┘                   │
│             │                             │                                 │
├─────────────┴─────────────────────────────┴─────────────────────────────────┤
│ Enforcement Orchestration Layer                                              │
├──────────────────────────────────────────────────────────────────────────────┤
│  ┌────────────────────────────────────────────────────────────────────────┐  │
│  │ scripts/quality_gate.sh (single canonical entrypoint)                 │  │
│  │ - deterministic ordering                                               │  │
│  │ - fail-fast policy                                                     │  │
│  │ - shared local/CI execution contract                                  │  │
│  └────────────────────────────────────────────────────────────────────────┘  │
├──────────────────────────────────────────────────────────────────────────────┤
│ Enforcement Engine Layer                                                     │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  ┌─────────────────┐ │
│  │ Build/Test   │  │ Lint/Format  │  │ Coverage     │  │ Policy Checkers │ │
│  │ cargo check  │  │ fmt+clippy   │  │ llvm-cov     │  │ clean+arch rules│ │
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘  └────────┬────────┘ │
│         │                 │                 │                   │           │
├─────────┴─────────────────┴─────────────────┴───────────────────┴───────────┤
│ Codebase Layer                                                               │
│  ┌────────────┐ ┌────────────┐ ┌────────────┐ ┌────────────┐ ┌────────────┐ │
│  │ docir-core │ │ parser/*   │ │ security/* │ │ rules/*    │ │ app/cli/py │ │
│  └────────────┘ └────────────┘ └────────────┘ └────────────┘ └────────────┘ │
└──────────────────────────────────────────────────────────────────────────────┘
```

### Component Responsibilities

| Component | Responsibility | Typical Implementation |
|-----------|----------------|------------------------|
| `scripts/quality_gate.sh` | Canonical sequencing and result contract | shell orchestrator with strict exit semantics |
| Quality policy checks | Enforce Clean Code invariants | AST/grep-based checks, clippy-deny settings, dead-code/forbidden-pattern checks |
| Architecture policy checks | Enforce layer/dependency boundaries | dependency graph validation for crate-level allowed edges |
| CI required job | Non-bypass merge gate | single job that only runs canonical script |
| Workspace crates (`docir-*`) | Product functionality under enforcement | modular Rust crates with `docir-core` as dependency root |

## Recommended Project Structure

```text
scripts/
├── quality_gate.sh              # canonical local+CI gate entrypoint
├── checks/                      # composable check scripts called by gate
│   ├── check_clean_code.sh      # forbidden patterns/complexity/todo policy
│   ├── check_architecture.sh    # crate boundary and dependency rules
│   └── check_coverage.sh        # threshold enforcement wrapper

ci/
└── quality_gate.yml             # required CI workflow; calls script only

policy/
├── clean_code.toml              # thresholds and forbidden patterns
└── architecture.toml            # allowed crate dependency matrix

crates/
├── docir-core/                  # domain model root
├── docir-parser/                # infrastructure-heavy parsing
├── docir-security/
├── docir-rules/
├── docir-diff/
├── docir-serialization/
├── docir-app/                   # application orchestration + ports
├── docir-cli/                   # interface adapter
└── docir-python/                # interface adapter
```

### Structure Rationale

- **`scripts/quality_gate.sh`:** preserves one mandatory enforcement surface and prevents bypass via alternate scripts.
- **`scripts/checks/`:** keeps checks composable while preserving one user-facing entrypoint.
- **`policy/`:** centralizes thresholds and allowed dependency edges as versioned architecture contracts.
- **`ci/quality_gate.yml`:** makes CI behavior isomorphic with local execution.
- **`crates/`:** existing modular architecture remains unchanged; enforcement wraps it rather than reshaping runtime design.

## Architectural Patterns

### Pattern 1: Canonical Gate Orchestrator

**What:** One script owns the full enforcement pipeline and exit status.
**When to use:** Any repository requiring deterministic, non-bypassable quality status.
**Trade-offs:** Strong consistency; reduced flexibility for ad hoc custom flows.

**Example:**
```bash
./scripts/quality_gate.sh
```

### Pattern 2: Policy-as-Code for Architecture

**What:** Explicit allowed dependency matrix checked automatically.
**When to use:** Multi-crate workspace with risk of architectural drift.
**Trade-offs:** Upfront policy maintenance; high long-term boundary stability.

**Example:**
```toml
[allow]
docir-app = ["docir-core", "docir-parser", "docir-security", "docir-rules", "docir-diff", "docir-serialization"]
```

### Pattern 3: Progressive Strictness with Deterministic End State

**What:** Introduce checks incrementally but require all checks to pass in one final canonical run.
**When to use:** Existing codebase with known complexity hotspots (notably parser-heavy modules).
**Trade-offs:** Temporary migration complexity; avoids big-bang breakage.

## Data Flow

### Request Flow

```text
[Developer Commit / PR Push]
    ↓
[scripts/quality_gate.sh]
    ↓
[fmt] → [clippy] → [check/test] → [coverage] → [policy checks]
    ↓
[aggregated pass/fail exit code]
    ↓
[local block + CI required status]
```

### State Management

```text
[policy/*.toml + Cargo workspace metadata]
    ↓ (read by checks)
[check scripts]
    ↓ (emit deterministic pass/fail signals)
[quality_gate.sh]
    ↓
[terminal + CI status]
```

### Key Data Flows

1. **Code Quality Flow:** source changes traverse formatter/linter/test/coverage in fixed order; first failing gate stops completion.
2. **Architecture Compliance Flow:** crate dependency graph is evaluated against allowed-edge policy; violations fail gate before merge.
3. **Evidence Flow:** gate logs produce reproducible failure evidence for local and CI parity.

## Scaling Considerations

| Scale | Architecture Adjustments |
|-------|--------------------------|
| Current workspace (single team) | Single gate script with strict fail-fast and shared policies is sufficient |
| 2x contributors / crate growth | Split checks into cached sub-steps but preserve single canonical wrapper |
| 5x contributors / CI load | Add parallelized internals and artifact caching; keep one mandatory gate contract |

### Scaling Priorities

1. **First bottleneck:** parser-heavy compile/test time; mitigate through selective test partitions behind canonical wrapper.
2. **Second bottleneck:** policy drift across crates; mitigate with dependency-matrix automation and review-required policy changes.

## Anti-Patterns

### Anti-Pattern 1: Multiple Gate Entry Points

**What people do:** Add parallel scripts (`quick_check.sh`, CI-only checks) that become de facto alternatives.
**Why it's wrong:** Produces inconsistent pass criteria and bypass paths.
**Do this instead:** Keep all flows routed through `scripts/quality_gate.sh` only.

### Anti-Pattern 2: Soft-Fail Architecture Checks

**What people do:** Treat dependency violations as warnings during delivery pressure.
**Why it's wrong:** Architectural debt compounds and boundary recovery cost spikes.
**Do this instead:** Enforce hard-fail boundary checks from canonical gate.

## Integration Points

### External Services

| Service | Integration Pattern | Notes |
|---------|---------------------|-------|
| CI provider (required job) | shell execution of canonical gate | no alternate CI command path |
| Rust toolchain (`cargo`, `clippy`, `llvm-cov`) | CLI tool invocation via gate | versions should be pinned in CI image/toolchain file |

### Internal Boundaries

| Boundary | Communication | Notes |
|----------|---------------|-------|
| `docir-cli`/`docir-python` ↔ `docir-app` | crate API calls | keep interfaces thin; avoid duplicate orchestration paths |
| `docir-app` ↔ processing crates | ports/adapters + crate dependencies | preserve dependency direction toward `docir-core` root |
| gate orchestration ↔ workspace crates | command execution + policy evaluation | enforcement must stay external to runtime behavior |

## Build Order

1. **Establish canonical gate skeleton** (`scripts/quality_gate.sh`) with deterministic step order and strict exit behavior.
2. **Integrate baseline toolchain checks** (`fmt`, `clippy`, `check/test`, `coverage`) under the canonical script.
3. **Add clean-code policy enforcement** (forbidden patterns, TODO policy, dead/unused handling) as hard-fail checks.
4. **Add architecture policy enforcement** (crate dependency matrix validation) as hard-fail checks.
5. **Bind CI required job** to execute only canonical gate and block merges on failure.
6. **Capture deterministic evidence** by validating at least one failing run and one passing run in implementation history.
7. **Finalize policy documentation** in planning/docs to codify non-bypass rule and completion contract.

## Sources

- `.planning/PROJECT.md`
- `.planning/codebase/ARCHITECTURE.md`
- `.planning/codebase/STRUCTURE.md`

---
*Architecture research for: docir strict quality enforcement integration*
*Researched: 2026-02-28*
