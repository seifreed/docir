# Repository Structure Audit (Focus: arch)

## 1. Top-Level Layout

- `Cargo.toml` (workspace manifest and crate membership)
- `crates/` (all runtime/library crates)
- `fixtures/` (sample documents used by tests/manual validation)
- `docs/`, `PROMPTS/`, `tools/`, `scripts/` (supporting assets/tooling)
- `.planning/codebase/` (analysis outputs)

Workspace crate registry source: `Cargo.toml:3-13`.

## 2. Crate Inventory and Size

Rust source footprint by crate (`*.rs` under each `src/`):

| Crate | Rust Files | Approx LOC | Primary Role |
|---|---:|---:|---|
| `docir-parser` | 149 | 33,726 | Multi-format parsing + ingestion/security scan hooks |
| `docir-core` | 46 | 7,280 | IR model, node taxonomy, visitor/query utilities |
| `docir-cli` | 16 | 1,598 | CLI interface and command handlers |
| `docir-diff` | 5 | 1,321 | Structural/semantic diff |
| `docir-rules` | 8 | 1,231 | Rule engine + built-in rules |
| `docir-security` | 8 | 1,010 | Threat enrichment + analyzer |
| `docir-app` | 5 | 764 | App facade, ports, use cases, adapters |
| `docir-serialization` | 2 | 217 | IR JSON serialization |
| `docir-python` | 1 | 164 | Python binding API |

Total Rust LOC across crates: ~47,793.

## 3. Structural Decomposition by Layer

### 3.1 Interface Layer

- CLI entrypoint and parsing:
  - `crates/docir-cli/src/main.rs:5-21`
  - `crates/docir-cli/src/cli.rs` (command schema)
- Command handlers are segmented per use case:
  - `crates/docir-cli/src/commands/parse.rs`
  - `crates/docir-cli/src/commands/security.rs`
  - `crates/docir-cli/src/commands/query.rs`
  - `crates/docir-cli/src/commands/rules.rs`
  - `crates/docir-cli/src/commands/diff.rs`
  - (plus coverage/extract/grep/summary/dump-node/util/dispatch)
- Python public API in one module:
  - `crates/docir-python/src/lib.rs:18-164`

### 3.2 Application Layer

- `crates/docir-app/src/lib.rs`:
  - API surface + port traits + facade construction.
- `crates/docir-app/src/use_cases.rs`:
  - Use-case execution units (parse/security/rules/diff).
- `crates/docir-app/src/adapters.rs`:
  - Default wiring to parser/security/rules/serialization crates.

### 3.3 Domain/Core Layer

- `crates/docir-core/src/lib.rs:8-22` exports:
  - `ir`, `types`, `security`, `visitor`, `query`, normalization, equivalence.
- IR module subtree:
  - `crates/docir-core/src/ir/mod.rs:6-35` (node family modules)
  - `crates/docir-core/src/ir/mod.rs:95-182` (single `IRNode` enum contract)
- Traversal/store abstractions:
  - `crates/docir-core/src/visitor/mod.rs:13-51`

### 3.4 Processing Service Layer

- Parsing:
  - `crates/docir-parser/src/lib.rs:6-25`
  - major subtrees: `ooxml/`, `odf/`, `hwp/`, `rtf/`, `parser/`
- Security:
  - `crates/docir-security/src/analyzer.rs`
  - `crates/docir-security/src/enrich.rs`
- Rules:
  - `crates/docir-rules/src/engine.rs`
  - `crates/docir-rules/src/rules.rs`
- Diff:
  - `crates/docir-diff/src/lib.rs`
  - `crates/docir-diff/src/index.rs`
- Serialization:
  - `crates/docir-serialization/src/json.rs`

## 4. Parser Subsystem Internal Structure

`docir-parser` is the dominant structure hub.

Directory decomposition:

- `crates/docir-parser/src/ooxml` (DOCX/XLSX/PPTX + shared OOXML components)
- `crates/docir-parser/src/odf` (ODT/ODS/ODP)
- `crates/docir-parser/src/hwp` (HWP/HWPX)
- `crates/docir-parser/src/rtf` (RTF)
- `crates/docir-parser/src/parser` (format detection, orchestration, post-processing, tests)

Key orchestration points:

- Unified format detection: `crates/docir-parser/src/parser/document.rs:34-84`
- Dispatch builder: `crates/docir-parser/src/parser/formats.rs:36-44`
- OOXML parser orchestration: `crates/docir-parser/src/parser/ooxml.rs:44-160`
- ODF parser boundary: `crates/docir-parser/src/odf/mod.rs:76-110`
- HWP/HWPX parser boundary: `crates/docir-parser/src/hwp/mod.rs:52-101`
- RTF module boundary: `crates/docir-parser/src/rtf/mod.rs:3-8`

## 5. Structural Hotspots (File Size Concentration)

Largest Rust files (current snapshot):

1. `crates/docir-parser/src/ooxml/docx/document/tests.rs` (~788 LOC)
2. `crates/docir-parser/src/ooxml/docx/document/inline.rs` (~754 LOC)
3. `crates/docir-parser/src/ooxml/xlsx/parser/tests.rs` (~751 LOC)
4. `crates/docir-parser/src/rtf/core.rs` (~749 LOC)
5. `crates/docir-parser/src/odf/spreadsheet.rs` (~743 LOC)
6. `crates/docir-parser/src/odf/helpers.rs` (~741 LOC)
7. `crates/docir-parser/src/ooxml/xlsx/styles.rs` (~712 LOC)
8. `crates/docir-parser/src/ooxml/pptx/tests.rs` (~708 LOC)
9. `crates/docir-core/src/ir/presentation.rs` (~704 LOC)
10. `crates/docir-parser/src/ooxml/pptx.rs` (~703 LOC)

Structural implication:

- Complexity is concentrated in parser format modules and parser-heavy tests.
- No single file currently exceeds 800 LOC in this snapshot, but multiple parser files are close.

## 6. Dependency Structure (Crate-Level)

Observed crate edges from manifests:

- `docir-core`: foundational model crate, no internal crate dependencies.
- `docir-parser` -> `docir-core`
- `docir-security` -> `docir-core`
- `docir-rules` -> `docir-core`
- `docir-diff` -> `docir-core`
- `docir-serialization` -> `docir-core`
- `docir-app` -> `docir-core`, `docir-parser`, `docir-security`, `docir-rules`, `docir-diff`, `docir-serialization`
- `docir-cli` -> `docir-app`, `docir-core`, `docir-security`
- `docir-python` -> `docir-app`, `docir-core`, `docir-parser`, `docir-rules`, `docir-serialization`

Manifest references:

- `crates/docir-core/Cargo.toml`
- `crates/docir-parser/Cargo.toml`
- `crates/docir-security/Cargo.toml`
- `crates/docir-rules/Cargo.toml`
- `crates/docir-diff/Cargo.toml`
- `crates/docir-serialization/Cargo.toml`
- `crates/docir-app/Cargo.toml`
- `crates/docir-cli/Cargo.toml`
- `crates/docir-python/Cargo.toml`

## 7. Structural Data-Flow Anchors (Concrete)

- CLI command routing switchboard: `crates/docir-cli/src/commands/dispatch.rs:22-179`
- CLI shared app bootstrap: `crates/docir-cli/src/commands/util.rs:36-57`
- App parse use case: `crates/docir-app/src/use_cases.rs:32-67`
- App analysis/rules/diff use cases: `crates/docir-app/src/use_cases.rs:94-132`
- Parser document entrypoint: `crates/docir-parser/src/parser/document.rs:24-114`
- Security enrichment mutation of root `Document.security`: `crates/docir-security/src/enrich.rs:41-49`
- Security analyzer traversal pattern: `crates/docir-security/src/analyzer.rs:26-39`
- Rules execution loop: `crates/docir-rules/src/engine.rs:115-123`
- Diff comparison loop: `crates/docir-diff/src/lib.rs:85-125`
- JSON tree serialization recursion: `crates/docir-serialization/src/json.rs:46-57`

## 8. Structure Summary

Current repository structure is well-modularized at crate level, with one dominant complexity center in `docir-parser` and a stable core contract in `docir-core` used by all processing crates.

