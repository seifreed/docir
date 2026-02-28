# Stack Research

**Domain:** Deterministic Rust quality-gate enforcement for a multi-crate workspace
**Researched:** 2026-02-28
**Confidence:** HIGH

## Recommended Stack

### Core Technologies

| Technology | Version | Purpose | Why Recommended | Confidence |
|------------|---------|---------|-----------------|------------|
| Rust toolchain via `rustup` + `rust-toolchain.toml` | Pin exact stable release (for example `1.xx.y`) | Deterministic compiler/tool behavior across local and CI | Official Rust toolchain manager; exact version pin removes drift and makes gate results reproducible | High |
| Cargo workspace commands (`cargo check/test/doc`) | Bundled with pinned toolchain | Canonical build, test, and docs verification for all crates | Official workflow for Rust workspaces; supports `--workspace`, `--locked`, `--frozen` for deterministic execution | High |
| `rustfmt` (`cargo fmt --all -- --check`) | Component matching pinned toolchain | Enforce formatting as hard gate | Official formatter shipped as toolchain component; check mode gives deterministic pass/fail behavior | High |
| Clippy (`cargo clippy --workspace --all-targets --all-features -- -D warnings`) | Component matching pinned toolchain | Enforce code-quality and policy lints as hard failures | Official Rust lint tool; CI guidance explicitly recommends `-D warnings` for enforcement | High |
| Rustdoc linting (`RUSTDOCFLAGS="-D warnings" cargo doc --workspace --no-deps`) | Bundled with pinned toolchain | Enforce documentation quality as gate criterion | Official rustdoc lint system supports deny-level enforcement (for missing docs and broken intra-doc links) | High |

### Supporting Libraries

| Library | Version | Purpose | When to Use | Confidence |
|---------|---------|---------|-------------|------------|
| `cargo-llvm-cov` | Pin exact `0.6.x` in tooling docs/CI image | Workspace coverage reporting compatible with Rust LLVM instrumentation | Required here because project constraints set a hard 95% threshold; use only as coverage runner on top of official `-C instrument-coverage` model | Medium |
| `llvm-tools-preview` (rustup component) | Match pinned Rust toolchain | Provides `llvm-cov`/`llvm-profdata` used by coverage pipelines | Install when coverage is part of canonical gate; recommended path in rustc coverage docs | High |
| `clippy.toml` (config file, not crate) | Versioned with repo | Central policy for restriction lints (`unwrap_used`, `expect_used`, etc.) | Use when policy must be explicit and reviewable, not hidden in ad hoc CLI flags | High |

### Development Tools

| Tool | Purpose | Notes | Confidence |
|------|---------|-------|------------|
| `scripts/quality_gate.sh` | Single canonical gate entrypoint for local, pre-commit, and CI | Must run all checks in one deterministic sequence; no alternate scripts allowed as quality surface | High |
| `rustup` profiles + components | Predictable CI bootstrap | Prefer `--profile minimal` in CI, then add required components explicitly (`clippy`, `rustfmt`, `llvm-tools-preview`) | High |
| `cargo` lock enforcement flags | Deterministic dependency resolution | Use `--locked` in gate steps; use `--frozen` in CI after dependency cache is prepared | High |

## Installation

```bash
# Core toolchain (pin exact stable in rust-toolchain.toml)
rustup toolchain install stable
rustup component add rustfmt clippy

# Coverage support (required by this project's 95% gate)
rustup component add llvm-tools-preview
cargo install cargo-llvm-cov --version "0.6.*" --locked

# Determinism checks
cargo fmt --all -- --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features --locked
RUSTDOCFLAGS="-D warnings" cargo doc --workspace --no-deps
cargo llvm-cov --workspace --all-features --fail-under-lines 95
```

## Alternatives Considered

| Recommended | Alternative | When to Use Alternative |
|-------------|-------------|-------------------------|
| Pinned stable toolchain (`rust-toolchain.toml`) | Floating `stable` channel without pin | Only for exploratory local work; not acceptable for deterministic gate enforcement |
| Clippy with strict deny flags in canonical gate | Rustc warnings only (`cargo check` without Clippy) | Only for quick local compile feedback; insufficient for policy enforcement |
| `cargo test --workspace --all-features --locked` | Per-crate ad hoc test commands | Only for targeted debugging; not acceptable as final quality gate |
| `cargo-llvm-cov` + official coverage instrumentation | Custom manual scripts invoking `llvm-profdata`/`llvm-cov` directly | Only if `cargo-llvm-cov` becomes unsupported; otherwise higher maintenance and error risk |

## What NOT to Use

| Avoid | Why | Use Instead | Confidence |
|-------|-----|-------------|------------|
| Nightly-only features as canonical gate dependencies | Nightly introduces toolchain drift and unstable behavior; breaks deterministic enforcement | Pinned stable toolchain + stable checks | High |
| `cargo clippy --fix` in CI gate | Mutates source during validation, making gate non-pure and non-reproducible | Non-mutating lint checks with `-D warnings` | High |
| Multiple independent quality scripts | Creates bypass paths and inconsistent pass criteria | Single `./scripts/quality_gate.sh` contract | High |
| Floating coverage tooling versions (`cargo install` without version pin) | Coverage output/behavior can change unexpectedly across runs | Pin exact `cargo-llvm-cov` release and toolchain version | Medium |
| Ignoring lockfile state in CI | Dependency resolution can drift and produce non-deterministic outcomes | `--locked`/`--frozen` for gate commands | High |

## Stack Patterns by Variant

**If running in local developer mode:**
- Use fast partial checks first (`cargo check`, targeted `cargo test -p ...`) before full gate.
- Because fast feedback improves iteration, while final acceptance still requires canonical gate pass.

**If running in CI required job mode:**
- Use only `./scripts/quality_gate.sh` with pinned toolchain and locked dependencies.
- Because one canonical path is necessary to prevent bypass and keep results deterministic.

**If debugging coverage mismatches:**
- Use official instrumentation assumptions (`-C instrument-coverage`) and matching `llvm-tools-preview`.
- Because coverage artifacts are LLVM-version sensitive and must match the Rust toolchain.

## Version Compatibility

| Package A | Compatible With | Notes |
|-----------|-----------------|-------|
| `rustc` (pinned stable in `rust-toolchain.toml`) | `cargo`, `clippy`, `rustfmt` from same toolchain | Keep all components on the same toolchain version to avoid lint/format drift |
| `cargo-llvm-cov` (pinned `0.6.x`) | `llvm-tools-preview` from pinned toolchain | Coverage tools must be version-aligned with compiler LLVM backend |
| Workspace `Cargo.lock` | `cargo ... --locked` | Lockfile must be committed and unchanged during gate run |

## Sources

- https://rust-lang.github.io/rustup/concepts/toolchains.html — toolchain pinning model and versioned toolchain naming (High)
- https://rust-lang.github.io/rustup/concepts/profiles.html — deterministic component installation strategy for CI (High)
- https://doc.rust-lang.org/cargo/commands/cargo-test.html — `--locked`/`--frozen` behavior for deterministic builds/tests (High)
- https://doc.rust-lang.org/stable/clippy/usage.html — official Clippy invocation and lint-level control (High)
- https://doc.rust-lang.org/clippy/continuous_integration/index.html — CI recommendation to use `-D warnings` (High)
- https://github.com/rust-lang/rustfmt — `cargo fmt` and `--check` pass/fail semantics (High)
- https://doc.rust-lang.org/rustdoc/lints.html — rustdoc lint controls and deny-level enforcement (High)
- https://doc.rust-lang.org/beta/rustc/instrument-coverage.html — official coverage instrumentation and `llvm-tools-preview` guidance (High)
- `.planning/PROJECT.md` — project constraints (single gate script, 95% coverage, no bypass)
- `.planning/codebase/STACK.md` — current workspace technology baseline
- `.planning/codebase/ARCHITECTURE.md` — architecture boundaries and scaling risks informing gate strictness

---
*Stack research for: deterministic Rust quality-gate enforcement in `docir`*
*Researched: 2026-02-28*
