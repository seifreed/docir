# CONCERNS

## Audit Scope
- Focus: technical debt, fragile areas, and scale risks.
- Commit analyzed: `cc212b33d8ab1c8a2245e8f35ae3c3c61e2b145f`
- Code scanned: Rust workspace (`crates/**`), coverage debt docs, CLI wiring.

## Priority Risks

### 1) High: Memory amplification on parse path can spike RSS on large inputs
- Evidence:
  - `crates/docir-parser/src/parser/document.rs:95-113` (`parse_file_with_bytes` / `parse_reader_with_bytes`) reads full payload into `Vec<u8>`.
  - `crates/docir-parser/src/parser/ooxml.rs:45-53` reads full payload again for OOXML parse.
  - `crates/docir-parser/src/config.rs:111-127` default `max_input_size` is `512 * 1024 * 1024`.
- Risk:
  - Large files can create multiple in-memory copies before parse completion.
  - Under concurrent parsing this can become a stability bottleneck before CPU becomes the limit.
- Action:
  - Split entrypoints into streaming parse path vs. explicit "retain raw bytes" path.
  - Enforce a single owned buffer through dispatch for OOXML/HWPX/ODF.
  - Add memory-regression tests around 100MB+ fixtures.

### 2) High: Parse errors are frequently downgraded to silent `None` fallbacks
- Evidence:
  - `crates/docir-parser/src/parser/ooxml.rs:255-257, 288, 300, 314, 360, 372, 402, 416, 492-503, 537-559, 657-667` uses `ok()`, `unwrap_or_default()`, and optional fallbacks for parse-part failures.
  - `crates/docir-parser/src/parser/document.rs:60-65` masks `mimetype` read failures with `unwrap_or(false)` during format detection.
- Risk:
  - Corrupted or drifting documents can look "successfully parsed" while dropping critical sub-parts.
  - Detection confidence and incident triage degrade because failures are not surfaced consistently.
- Action:
  - Convert optional part parsing to typed outcomes: `Missing`, `Parsed`, `Failed(parse_error)`.
  - Emit diagnostics for every swallowed parse error.
  - Keep permissive behavior only for explicitly optional parts with severity-tagged diagnostics.

### 3) High: Several parser modules are effectively god-files (high blast radius)
- Evidence (non-test LOC):
  - `crates/docir-parser/src/ooxml/docx/document/inline.rs` (~754 LOC)
  - `crates/docir-parser/src/rtf/core.rs` (~749 LOC)
  - `crates/docir-parser/src/odf/spreadsheet.rs` (~743 LOC)
  - `crates/docir-parser/src/odf/helpers.rs` (~741 LOC)
  - `crates/docir-parser/src/ooxml/pptx.rs` (~703 LOC)
  - `crates/docir-parser/src/parser/ooxml.rs` (~682 LOC)
- Risk:
  - Feature work and bug fixes have high merge-conflict probability.
  - Reviewability is low; regressions are easier to introduce and harder to isolate.
- Action:
  - Slice by responsibility boundaries (tokenization, relationship resolution, node assembly, diagnostics).
  - Introduce module-level ownership to reduce cross-cutting edits.
  - Prioritize extraction of hot change zones first (`parser/ooxml.rs`, `ooxml/pptx.rs`, `rtf/core.rs`).

### 4) High: Repeated XML event-loop patterns create behavior drift risk
- Evidence:
  - `crates/docir-parser/src/ooxml/pptx.rs:97-139, 155-205`
  - `crates/docir-parser/src/ooxml/pptx/slide.rs:30-31, 172-173, 312-313`
  - `crates/docir-parser/src/ooxml/xlsx/worksheet.rs:100-125, 224-240`
  - Repeated `loop { match reader.read_event_into(...) { ... _ => {} } }` pattern.
- Risk:
  - Small fixes (e.g., namespace handling, unknown-tag behavior, error normalization) must be repeated in many places.
  - Inconsistent handling of equivalent XML constructs across formats increases false negatives.
- Action:
  - Introduce shared traversal helpers with explicit policies for unknown tags and attribute decode errors.
  - Add parser-contract tests that assert equivalent behavior for structurally similar XML constructs.

### 5) Medium-High: Security extraction is split across disparate code paths
- Evidence:
  - Central OOXML scanner: `crates/docir-parser/src/parser/security.rs:23-33`.
  - External rel scan in scanner is Word-only: `crates/docir-parser/src/parser/security.rs:125-155`.
  - Additional external-reference extraction exists separately in XLSX/PPTX paths (`crates/docir-parser/src/ooxml/xlsx/relationships.rs`, `crates/docir-parser/src/ooxml/pptx/relationships.rs`).
- Risk:
  - Coverage is maintenance-sensitive: new relationship types can be added in one format path and missed in another.
  - Security signal consistency across formats is fragile.
- Action:
  - Consolidate external reference classification into one reusable component.
  - Define format-agnostic security extraction contracts and test matrix per relationship type.

### 6) Medium: Panic-based fallbacks remain in production command/summary paths
- Evidence:
  - `crates/docir-diff/src/summary.rs:123` uses `unreachable!("summary missing for node")`.
  - `crates/docir-cli/src/commands/dispatch.rs:138` uses `unreachable!("command should be routed by run_command")`.
- Risk:
  - Forward-compatibility failures become runtime panics instead of typed errors.
  - New enum variants or routing mistakes can crash user workflows.
- Action:
  - Replace with typed error returns and include command/node metadata in error context.
  - Add compile-time exhaustiveness tests around command and IR summary coverage.

### 7) Medium: Coverage debt is explicitly tracked but still unresolved in critical areas
- Evidence:
  - `docs/AST_COVERAGE.md:26,39,50,60,68,79` lists open TODOs for fixture completeness and deterministic security regression sets.
  - `docs/AST_COVERAGE.md:64-68,75-79,84,90,92` marks multiple format areas as PARTIAL.
- Risk:
  - Security and format regressions can pass local validation when edge fixtures are missing.
  - Parser maturity claims can outpace deterministic test evidence.
- Action:
  - Treat DOCIR-201..206 as release-gating backlog for parser/security stability.
  - Add fixture provenance + expected-output hashing to reduce non-deterministic assertions.

### 8) Medium: Node identity generation is process-global and non-deterministic
- Evidence:
  - `crates/docir-core/src/types.rs:9-10,46-48` uses global `AtomicU64` for `NodeId::new()`.
- Risk:
  - Cross-run reproducibility of raw node ids is not guaranteed.
  - Any downstream tooling that assumes stable ids between runs can become brittle.
- Action:
  - Document NodeId stability contract explicitly (intra-document vs cross-run).
  - For reproducible diff/reporting modes, add deterministic ID mapping layer at serialization/export boundaries.

## Fragile Hotspots To Watch
- `crates/docir-parser/src/parser/ooxml.rs` (central OOXML orchestration + optional-part parsing policy).
- `crates/docir-parser/src/ooxml/pptx.rs` and `crates/docir-parser/src/ooxml/xlsx/worksheet.rs` (high-volume XML event handling).
- `crates/docir-parser/src/rtf/core.rs` (stateful parser core, high control-flow density).
- `crates/docir-diff/src/summary.rs` (broad enum handling + panic fallback).

## Recommended Execution Order
1. Remove silent-failure paths by adding structured diagnostics for optional-part parse failures.
2. Eliminate multi-buffer parse amplification in `parse_*_with_bytes` flow.
3. Consolidate XML traversal/error-handling helper APIs.
4. Consolidate security external-reference classification across Word/XLSX/PPTX.
5. Replace panic fallbacks with typed errors and add coverage tests for new variants.
