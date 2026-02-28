# Pitfalls Research

**Domain:** Deterministic quality and architecture gate rollout for a multi-crate Rust workspace
**Researched:** 2026-02-28
**Confidence:** HIGH

## Critical Pitfalls

### Pitfall 1: Parallel Gate Surfaces Create Bypass Paths

**What goes wrong:**
CI, local scripts, and pre-commit run different checks, so code can pass one path and fail another.

**Why it happens:**
Teams preserve legacy scripts while introducing a new gate, then gradually drift into inconsistent enforcement.

**How to avoid:**
Use only `./scripts/quality_gate.sh` as the canonical gate. Make CI required status depend exclusively on it, and retire alternate quality entrypoints.

**Warning signs:**
`cargo test` passes locally but CI fails for policy checks; multiple scripts in docs for "the gate"; merge requests mention "works if you run X instead".

**Phase to address:**
Phase 1 (Canonical Gate Surface and CI wiring)

---

### Pitfall 2: Silent Parse Failures Hide Under Green Gates

**What goes wrong:**
Parser components swallow errors (`ok()`, `unwrap_or_default`, optional fallbacks), yielding incomplete outputs while checks remain green.

**Why it happens:**
Reliability pressure favors permissive parsing, and gates over-index on compile/test status without semantic diagnostics quality.

**How to avoid:**
Promote parse outcomes to typed states (`Parsed`/`Missing`/`Failed`), require surfaced diagnostics for swallowed parse failures, and add regression tests for corrupted fixtures.

**Warning signs:**
Rising "partial parse" incidents without test failures; frequent optional fallback branches; security detections fluctuate across similar files.

**Phase to address:**
Phase 2 (Failure Semantics and Diagnostic Enforcement)

---

### Pitfall 3: Quality Gates Ignore Memory Amplification Risk

**What goes wrong:**
Large documents trigger repeated full-buffer allocations in parse paths; gate passes but production RSS spikes and parsing destabilizes.

**Why it happens:**
Gate criteria focus on formatting/lints/tests/coverage but omit stress and memory regression checks.

**How to avoid:**
Split streaming vs retain-bytes entrypoints, enforce single-buffer ownership in dispatch paths, and add memory regression tests for 100MB+ inputs.

**Warning signs:**
OOM or slowdowns only on large fixtures; parse path duplicates `Vec<u8>`; concurrency worsens failures non-linearly.

**Phase to address:**
Phase 3 (Scalability and Resource-Use Guardrails)

---

### Pitfall 4: God Files Turn Policy Changes into High-Risk Edits

**What goes wrong:**
Large parser modules (>650-750 LOC) become change hotspots; small policy updates require fragile edits in many dense files.

**Why it happens:**
Fast feature growth accumulates orchestration, traversal, and diagnostics responsibilities in single files.

**How to avoid:**
Refactor by responsibility boundaries (traversal, relationship resolution, assembly, diagnostics), assign module ownership, and enforce size/complexity checks in gate policy.

**Warning signs:**
Frequent merge conflicts in `parser/ooxml.rs`, `ooxml/pptx.rs`, `rtf/core.rs`; PRs touch same modules repeatedly; review latency increases.

**Phase to address:**
Phase 4 (Structural Decomposition and Complexity Limits)

---

### Pitfall 5: Repeated XML Event Loops Cause Behavioral Drift

**What goes wrong:**
Equivalent XML constructs are handled differently across modules because traversal loops are reimplemented with minor variations.

**Why it happens:**
No shared traversal helper/policy layer, so teams copy-paste `read_event_into` loops and diverge over time.

**How to avoid:**
Introduce shared XML traversal helpers, normalize unknown-tag and attribute-decoding policy, and add contract tests across DOCX/XLSX/PPTX equivalents.

**Warning signs:**
Bugfixes must be patched in 3+ places; inconsistent unknown-tag behavior; parser parity regressions between formats.

**Phase to address:**
Phase 5 (Shared Parsing Primitives and Contract Tests)

---

### Pitfall 6: Fragmented Security Extraction Weakens Architecture Gates

**What goes wrong:**
Security reference extraction is split by format-specific paths, producing uneven coverage and fragile maintenance.

**Why it happens:**
Security logic evolved near each parser instead of behind a single reusable contract.

**How to avoid:**
Centralize external-reference classification behind one reusable component and enforce format-agnostic test matrices in the canonical gate.

**Warning signs:**
Word/XLSX/PPTX report different classes for equivalent relationships; new relationship types only covered in one format.

**Phase to address:**
Phase 6 (Security Signal Consolidation)

---

### Pitfall 7: Panic Paths Survive Under Strict Gate Branding

**What goes wrong:**
`unreachable!` fallbacks in command/summary paths trigger runtime crashes when enums evolve, despite "strict quality" claims.

**Why it happens:**
Policy scans may target `unwrap`/`expect` while omitting panic macros in production paths.

**How to avoid:**
Replace panic fallbacks with typed errors, enforce no-panic rules for non-test code, and add exhaustiveness tests for routed commands and summaries.

**Warning signs:**
Runtime panics after adding variants; incidents tied to "should be unreachable" assumptions; clippy exceptions for panic macros.

**Phase to address:**
Phase 7 (No-Panic Runtime Invariants)

## Technical Debt Patterns

Shortcuts that seem reasonable but create long-term problems.

| Shortcut | Immediate Benefit | Long-term Cost | When Acceptable |
|----------|-------------------|----------------|-----------------|
| Keep legacy gate scripts alongside canonical gate | Faster rollout with less migration work | Permanent bypass path and nondeterministic policy outcomes | Only during short migration window with hard removal date |
| Keep permissive parser fallbacks without diagnostics | Fewer visible failures now | Hidden data loss and weak incident triage | Never for security-relevant parse stages |
| Defer decomposition of parser god-files | Avoids near-term refactor effort | High defect density and review bottlenecks | Only for isolated hotfixes with follow-up ticket owner/date |
| Duplicate XML loop logic per format | Quick local feature delivery | Drift, duplicate bugfix cost, inconsistent behavior | Never beyond initial prototype |
| Keep panic fallbacks in runtime paths | Simpler control flow in the short term | User-facing crashes and brittle forward compatibility | Never in production command and summary paths |

## Integration Gotchas

Common mistakes when connecting to external services.

| Integration | Common Mistake | Correct Approach |
|-------------|----------------|------------------|
| CI required checks | Marking multiple jobs as merge blockers with inconsistent command sets | Single required job that executes only `./scripts/quality_gate.sh` |
| Coverage tooling (`cargo llvm-cov`) | Treating coverage as advisory and not gate-failing below threshold | Enforce hard 95% minimum inside canonical gate |
| Pre-commit hooks | Running partial checks that differ from CI | Hook should shell out to canonical gate or clearly run a strict subset with required full gate before push |
| Clippy policy | Allowing warnings or local suppressions that CI tolerates | Run clippy with deny warnings and policy checks under same script |
| Static architecture checks | Ad-hoc scripts not run in local developer flow | Put architecture validation into canonical gate and document failure remediation |

## Performance Traps

Patterns that work at small scale but fail as usage grows.

| Trap | Symptoms | Prevention | When It Breaks |
|------|----------|------------|----------------|
| Multi-buffer parsing | RSS spikes, OOM under concurrent large-file parses | Single-buffer ownership and streaming entrypoints | Usually visible at 100MB+ input with parallel workloads |
| Full gate on every tiny edit without caching strategy | Developer throughput collapse, skipped gates | Keep strict gate but optimize caching, fixture scope, and parallelism | Teams begin bypass attempts once cycle time is perceived as blocking |
| Repeated XML traversal per module | CPU overhead and divergent error handling | Shared traversal helpers with normalized policies | As format support grows and parity maintenance cost compounds |
| Security extraction per-format duplication | Uneven detection and repeated scanning costs | Consolidated extraction core with per-format adapters | When new relationship types are added rapidly |

## Security Mistakes

Domain-specific security issues beyond general web security.

| Mistake | Risk | Prevention |
|---------|------|------------|
| Treating parser partial success as acceptable for threat analysis | Missed malicious signals due to dropped sub-parts | Require diagnostics-backed partial outcomes and fail on unclassified critical parse errors |
| Inconsistent external-reference classification across formats | False negatives in malware triage | Centralize relationship classification and assert parity tests |
| Allowing bypass scripts for gate in CI/local | Vulnerable code merged despite "strict" policy | One canonical gate, required CI status, and no alternate merge path |
| Panic-based runtime handling in security paths | Crash-driven denial of analysis | Typed errors with explicit severity and actionable context |

## UX Pitfalls

Common user experience mistakes in this domain.

| Pitfall | User Impact | Better Approach |
|---------|-------------|-----------------|
| Opaque gate failure output | Engineers cannot quickly fix failures, increasing friction | Categorize failures by policy area with clear remediation hints |
| Non-deterministic gate results across environments | "Works on my machine" disputes and trust erosion | Lock toolchain versions and enforce single command path |
| Excessive false positives in strict checks | Teams start suppressing or bypassing policies | Tune rule set using measured violations and explicit ownership |
| Long gate runtime with no progress cues | Perceived instability and premature cancellation | Provide staged output and duration metrics while preserving strictness |

## "Looks Done But Isn't" Checklist

Things that appear complete but are missing critical pieces.

- [ ] **Canonical Gate:** Only one script exists and all docs, hooks, and CI call it directly.
- [ ] **Architecture Enforcement:** Dependency rule violations fail gate in both local and CI execution.
- [ ] **Parse Reliability:** Swallowed parse failures emit structured diagnostics and are test-covered.
- [ ] **Scalability Safety:** Memory regression tests exist for large fixtures and fail on amplification regressions.
- [ ] **Security Consistency:** External-reference detection parity is verified across DOCX/XLSX/PPTX.
- [ ] **Runtime Safety:** No panic fallbacks remain in non-test command/summary execution paths.

## Recovery Strategies

When pitfalls occur despite prevention, how to recover.

| Pitfall | Recovery Cost | Recovery Steps |
|---------|---------------|----------------|
| Parallel gate surfaces already in use | MEDIUM | Freeze merges, route all checks to canonical script, remove alternates, then re-baseline CI |
| Silent parse-failure drift discovered late | HIGH | Add diagnostics immediately, replay corpus, diff outputs, and patch missing detection logic |
| Memory amplification in production parse flow | HIGH | Ship guarded size limits, switch to single-buffer path, add stress tests before reopening throughput |
| XML behavior drift across formats | MEDIUM | Introduce shared helper, backfill contract tests, and reconcile divergent parsing semantics |
| Panic fallback crash in runtime path | MEDIUM | Replace with typed errors, add exhaustive tests, and publish incident postmortem with policy update |

## Pitfall-to-Phase Mapping

How roadmap phases should address these pitfalls.

| Pitfall | Prevention Phase | Verification |
|---------|------------------|--------------|
| Parallel Gate Surfaces Create Bypass Paths | Phase 1 | CI required check calls only canonical script and no alternate gate docs/scripts remain |
| Silent Parse Failures Hide Under Green Gates | Phase 2 | Corrupted fixture suite produces diagnostics and deterministic failure/signal behavior |
| Quality Gates Ignore Memory Amplification Risk | Phase 3 | Large-input regression tests enforce memory and stability thresholds |
| God Files Turn Policy Changes into High-Risk Edits | Phase 4 | Complexity/file-size checks and reduced hotspot conflict frequency |
| Repeated XML Event Loops Cause Behavioral Drift | Phase 5 | Contract tests pass across equivalent XML constructs in multiple formats |
| Fragmented Security Extraction Weakens Architecture Gates | Phase 6 | Single extraction core used by all OOXML formats with parity test matrix |
| Panic Paths Survive Under Strict Gate Branding | Phase 7 | Non-test panic macro scan is clean and routing exhaustiveness tests pass |

## Sources

- `.planning/PROJECT.md`
- `.planning/codebase/CONCERNS.md`
- Existing parser hotspot inventory and execution-order recommendations documented in current codebase concerns

---
*Pitfalls research for: deterministic quality/architecture enforcement in `docir`*
*Researched: 2026-02-28*
