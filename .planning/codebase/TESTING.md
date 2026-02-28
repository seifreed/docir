# Testing Patterns Audit (Quality Focus)

## Scope
- Focus: test organization, styles, assertions, fixture strategy, and risk coverage
- Repository: Rust workspace tests under `crates/*/tests` and `#[cfg(test)]` modules in `src`

## Current Testing Topology

### Test volume and placement
- Total `#[test]` occurrences found: 132
- Unit-style tests in `src` modules: 125
- Integration-style tests under crate `tests/`: 7

### Where tests are concentrated
- Parser-heavy suites:
  - `crates/docir-parser/src/ooxml/docx/document/tests.rs`
  - `crates/docir-parser/src/ooxml/pptx/tests.rs`
  - `crates/docir-parser/src/ooxml/xlsx/parser/tests.rs`
  - `crates/docir-parser/src/odf/tests.rs`
  - `crates/docir-parser/src/rtf/tests.rs`
  - `crates/docir-parser/src/hwp/mod.rs` (`mod tests`)
- Integration-level checks:
  - `crates/docir-parser/tests/fixtures.rs`
  - `crates/docir-parser/tests/xlsb.rs`
  - `crates/docir-cli/tests/diff_rules.rs`
  - `crates/docir-cli/tests/coverage_export.rs`

## Strong Patterns

### 1. Format/feature matrix coverage mindset
- Tests validate multiple document families and features, not just happy-path parse success.
- Examples:
  - Cross-format fixture sweep in `crates/docir-parser/tests/fixtures.rs`
  - Format-specific deep checks in `crates/docir-parser/src/odf/tests.rs` and `crates/docir-parser/src/ooxml/*/tests*.rs`

### 2. Security-focused negative testing exists
- Security/resource-limit behavior gets explicit checks.
- Examples:
  - Zip/path traversal protections in `crates/docir-parser/src/zip_handler.rs` tests
  - Input-size limit test in `crates/docir-parser/src/parser/tests/limits.rs`
  - Threat/security indicator tests in:
    - `crates/docir-parser/src/odf/tests/security_indicators.rs`
    - `crates/docir-parser/src/odf/tests/threat_indicators.rs`
    - `crates/docir-security/src/indicators.rs`

### 3. CLI integration tests verify executable behavior
- Tests execute the built binary (`CARGO_BIN_EXE_docir`) and validate exported artifacts/JSON/CSV contracts.
- Examples:
  - `crates/docir-cli/tests/diff_rules.rs`
  - `crates/docir-cli/tests/coverage_export.rs`

### 4. Assertions include semantic checks, not only existence
- Example patterns:
  - Node-type presence assertions in `crates/docir-parser/tests/fixtures.rs`
  - Structured field/property assertions in `crates/docir-parser/src/ooxml/docx/document/tests/advanced_features.rs`

## Testing Pattern Gaps / Quality Risks

### 1. Repeated test bootstrap logic across integration tests
- Duplicated blocks for:
  - workspace root discovery
  - reading `fixtures/manifest.json`
  - tempdir setup
  - command invocation scaffolding
- Repeated in:
  - `crates/docir-cli/tests/diff_rules.rs`
  - `crates/docir-cli/tests/coverage_export.rs`
  - `crates/docir-parser/tests/fixtures.rs`
- Risk: maintenance overhead and drift in setup semantics.

### 2. Large monolithic test files reduce navigability
- Notable large test files:
  - `crates/docir-parser/src/ooxml/docx/document/tests.rs` (788 LOC)
  - `crates/docir-parser/src/ooxml/xlsx/parser/tests.rs` (751 LOC)
  - `crates/docir-parser/src/ooxml/pptx/tests.rs` (708 LOC)
  - `crates/docir-parser/src/odf/tests.rs` (701 LOC)
- Risk: slower local debugging, harder ownership boundaries.

### 3. Limited advanced test techniques
- No evidence of `proptest`, `quickcheck`, `rstest`, snapshot testing, or async test frameworks.
- Current suite is deterministic and explicit, but mostly example-based.
- Risk: edge-case space (especially parser hardening inputs) may rely on manually curated cases only.

### 4. Heavy reliance on `unwrap`/`expect` in test helpers
- Common in fixture builders and setup paths.
- Examples:
  - `crates/docir-parser/src/parser/tests/helpers.rs`
  - `crates/docir-rules/src/tests.rs`
- Acceptable in tests, but pervasive use can obscure failure context for flaky/environment-dependent setup.

## Effective Testing Conventions Observed
- Convention: `#[cfg(test)] mod tests;` or sibling `tests.rs` per module.
- Convention: integration tests under `crates/<crate>/tests/*.rs` for end-to-end crate behavior.
- Convention: fixture-driven validation through repository-level `fixtures/manifest.json`.
- Convention: command-contract verification for CLI outputs (JSON/CSV headers and schema presence).

## Recommended Testing Pattern Improvements (Quality-Focused)
- Introduce shared integration-test helpers for root path and fixture manifest loading.
- Split oversized parser test files by feature slices (e.g., fields, tables, styles, media).
- Keep current deterministic fixtures, but add selective fuzz/property-style checks for parser boundaries.
- Standardize failure messaging in helper setup functions to improve triage speed.
