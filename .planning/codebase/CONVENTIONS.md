# Coding Conventions Audit (Quality Focus)

## Scope
- Repository: Rust workspace (`crates/*`)
- Focus: naming, structure, error handling, module patterns, consistency
- Evidence sampled across `docir-cli`, `docir-app`, `docir-parser`, `docir-core`, `docir-security`, `docir-rules`, `docir-diff`

## High-Confidence Conventions In Use

### 1. Crate/module naming is consistent and domain-driven
- Workspace crates follow `docir-*` naming and clear boundaries (`crates/docir-core`, `crates/docir-parser`, `crates/docir-cli`, etc.; root `Cargo.toml`).
- Rust module/file naming uses snake_case and decomposed subsystems.
- Concrete examples:
  - `crates/docir-parser/src/ooxml/xlsx/worksheet.rs`
  - `crates/docir-parser/src/ooxml/docx/document/paragraph.rs`
  - `crates/docir-core/src/ir/spreadsheet/cell.rs`

### 2. Module-level docs are broadly adopted
- Many files start with `//!` docs that describe module purpose.
- Concrete examples:
  - `crates/docir-cli/src/commands/parse.rs`
  - `crates/docir-parser/src/zip_handler.rs`
  - `crates/docir-core/src/lib.rs`

### 3. Error modeling follows layered conventions
- Domain/infrastructure crates use typed error enums via `thiserror`:
  - `crates/docir-core/src/error.rs` (`CoreError`)
  - `crates/docir-parser/src/error.rs` (`ParseError`)
  - `crates/docir-app/src/lib.rs` (`AppError`, `AppParseError`)
- CLI layer uses `anyhow::Result` plus `.context(...)` for user-facing command failures:
  - `crates/docir-cli/src/commands/parse.rs`
  - `crates/docir-cli/src/commands/summary.rs`
  - `crates/docir-cli/src/commands/coverage.rs`

### 4. Command handler pattern is standardized in CLI
- Command modules expose `run(...) -> Result<()>` with `PathBuf` inputs and parsed options.
- Dispatch routing is centralized:
  - `crates/docir-cli/src/commands/dispatch.rs`
  - `crates/docir-cli/src/commands/mod.rs`

### 5. Use-case and port style appears intentionally applied in app layer
- Traits define ports (`ParserPort`, `SecurityScannerPort`, etc.), use-case structs orchestrate logic.
- Concrete examples:
  - `crates/docir-app/src/lib.rs`
  - `crates/docir-app/src/use_cases.rs`
  - `crates/docir-app/src/adapters.rs`

## Convention Drift / Quality Risks

### 1. File-size pressure in parser and test-heavy modules
- No files exceed 800 LOC in sampled output, but several are near/over 700 LOC, increasing cognitive load and merge risk.
- Concrete large files (from `wc -l`):
  - `crates/docir-parser/src/ooxml/docx/document/tests.rs` (788)
  - `crates/docir-parser/src/ooxml/docx/document/inline.rs` (754)
  - `crates/docir-parser/src/ooxml/xlsx/parser/tests.rs` (751)
  - `crates/docir-parser/src/rtf/core.rs` (749)
  - `crates/docir-parser/src/odf/spreadsheet.rs` (743)

### 2. Repeated control-flow boilerplate in tests
- Same root-discovery pattern appears repeatedly:
  - `PathBuf::from(env!("CARGO_MANIFEST_DIR")).parent().unwrap().parent().unwrap()`
  - Seen in:
    - `crates/docir-cli/tests/diff_rules.rs`
    - `crates/docir-cli/tests/coverage_export.rs`
    - `crates/docir-parser/tests/fixtures.rs`
- Same manifest-loading sequence repeated (`read_to_string` + `serde_json::from_str`) in multiple test files.

### 3. Repeated ad-hoc ZIP fixture builders
- Multiple helper functions manually build ZIP/OOXML payloads with repeated `start_file` + `write_all` sequences.
- Concentrated in:
  - `crates/docir-parser/src/parser/tests/helpers.rs`
  - `crates/docir-rules/src/tests.rs`
  - `crates/docir-parser/src/odf/tests.rs`
- This is useful but currently copy-prone and inconsistently centralized.

### 4. Inconsistent strictness around `unwrap`/`expect`
- Production code mostly avoids panics and returns typed `Result`.
- Test code heavily uses `unwrap`/`expect` (expected in tests), but helper-heavy modules sometimes rely on many `unwrap`s even for setup paths where failure messages could be more targeted.
- Examples:
  - `crates/docir-parser/src/parser/tests/helpers.rs`
  - `crates/docir-cli/tests/diff_rules.rs`

### 5. Minor lint exception indicates API shape strain
- Explicit lint suppression in CLI dispatch path:
  - `#[allow(clippy::too_many_arguments)]` in `crates/docir-cli/src/commands/dispatch.rs` (`build_query_like_command`)
- Not severe, but it marks an area where option structs are only partially adopted.

## Practical Convention Baseline (Current)
- Strong:
  - Crate/module naming
  - Layered error types
  - Command/use-case structure
  - Module documentation habit
- Weakening:
  - Test scaffolding duplication
  - Very large parser/test files
  - Repeated fixture construction patterns

## Suggested Quality Guardrails (Convention-Level)
- Add an internal guideline for shared test bootstrap utilities (repo root + manifest loading + CLI invocation wrappers).
- Set soft file-size guardrails (e.g., warn at 600 LOC for parser modules, 500 LOC for test files).
- Keep `anyhow` at edge layers only (`docir-cli`), and continue typed errors in core/parser/app.
- Prefer option/config structs over argument-heavy helper constructors in CLI command routing.
