# Architecture Audit (Focus: arch)

## 1. Workspace Architecture Overview

`docir` is a Rust workspace split into 9 crates with mostly clean directional dependencies from interfaces to application to domain-like core modules.

Primary dependency direction:

1. Interface layer: `docir-cli`, `docir-python`
2. Application orchestration: `docir-app`
3. Core services: `docir-parser`, `docir-security`, `docir-rules`, `docir-diff`, `docir-serialization`
4. Core model: `docir-core`

Evidence:

- Workspace members: `Cargo.toml:3-13`
- App depends on parser/security/rules/diff/serialization/core: `crates/docir-app/Cargo.toml:10-17`
- CLI depends on app/core/security only (not parser directly): `crates/docir-cli/Cargo.toml:14-22`
- Python binding depends on app plus lower-level crates: `crates/docir-python/Cargo.toml:13-22`

## 2. Layer Map

### 2.1 Interface Layer

- CLI entrypoint and command routing:
  - `crates/docir-cli/src/main.rs:5-21`
  - `crates/docir-cli/src/commands/dispatch.rs:22-79`
- Python entrypoints:
  - `crates/docir-python/src/lib.rs:18-43` (`parse_json`, `rules`)
  - `crates/docir-python/src/lib.rs:45-86` (`query`)

Responsibilities:

- Parse user input/flags.
- Call application APIs.
- Shape output (JSON/human text).

### 2.2 Application Layer (`docir-app`)

`DocirApp` is the main orchestration facade and port owner:

- Facade, ports, construction:
  - `crates/docir-app/src/lib.rs:99-156`
  - `crates/docir-app/src/lib.rs:167-234`
- Use-case implementations:
  - `crates/docir-app/src/use_cases.rs:13-68` (parse + security scan/enrichment)
  - `crates/docir-app/src/use_cases.rs:79-132` (security analysis, rules, diff)
- Infrastructure adapters for ports:
  - `crates/docir-app/src/adapters.rs:18-68`
  - `crates/docir-app/src/adapters.rs:81-141`

Architecture property:

- Good inversion point: ports (`ParserPort`, `SecurityAnalyzerPort`, `RulesEnginePort`, `SerializerPort`) in app layer keep CLI/Python isolated from concrete engines.

### 2.3 Core Model Layer (`docir-core`)

Defines IR, node taxonomy, traversal, query, normalization, security schema.

- Module exports: `crates/docir-core/src/lib.rs:8-22`
- IR graph and node enum: `crates/docir-core/src/ir/mod.rs:74-182`
- Traversal/store abstractions: `crates/docir-core/src/visitor/mod.rs:13-51`

Architecture property:

- `docir-core` is dependency root for most crates and appears framework-agnostic (pure data + traversal abstractions).

### 2.4 Processing/Infrastructure Services

- Parser subsystem (`docir-parser`): multi-format ingestion into IR
  - Public boundary: `crates/docir-parser/src/lib.rs:6-33`
  - Format detection/dispatch: `crates/docir-parser/src/parser/document.rs:24-92`
  - Dispatch to concrete parsers: `crates/docir-parser/src/parser/formats.rs:5-44`
- Security analysis/enrichment (`docir-security`)
  - Analyzer visitor: `crates/docir-security/src/analyzer.rs:21-77`, `86-207`
  - Indicator enrichment pipeline: `crates/docir-security/src/enrich.rs:18-50`
- Rules engine (`docir-rules`)
  - Core engine/run path: `crates/docir-rules/src/engine.rs:78-131`
- Serialization (`docir-serialization`)
  - JSON serializer and tree builder: `crates/docir-serialization/src/json.rs:46-63`, `65-111`
- Diff engine (`docir-diff`)
  - Index + diff comparison: `crates/docir-diff/src/index.rs:16-27`, `crates/docir-diff/src/lib.rs:62-133`

## 3. Runtime Data Flows

### 3.1 Parse Flow (CLI)

1. CLI parses args and builds parser config.
   - `crates/docir-cli/src/main.rs:18-20`
2. Command dispatcher routes to parse handler.
   - `crates/docir-cli/src/commands/dispatch.rs:24-30`
3. Helper builds `DocirApp` and parses file.
   - `crates/docir-cli/src/commands/util.rs:45-57`
4. `DocirApp` parse use-case runs parser, optional security scanning, enrichment.
   - `crates/docir-app/src/lib.rs:236-249`
   - `crates/docir-app/src/use_cases.rs:32-55`
5. JSON serialization executes through app serializer port.
   - `crates/docir-app/src/lib.rs:251-254`
   - `crates/docir-serialization/src/json.rs:145-157`

### 3.2 Security Flow

1. CLI command calls `app.analyze_security`.
   - `crates/docir-cli/src/commands/security.rs:10-18`
2. App use-case creates analyzer via factory and runs visitor.
   - `crates/docir-app/src/lib.rs:256-259`
   - `crates/docir-app/src/use_cases.rs:94-97`
3. Analyzer walks IR and emits findings.
   - `crates/docir-security/src/analyzer.rs:26-39`
   - `crates/docir-security/src/analyzer.rs:86-207`

### 3.3 Rules Flow

1. Interface invokes rules with profile.
   - CLI: `crates/docir-cli/src/commands/dispatch.rs:71-77`
   - Python: `crates/docir-python/src/lib.rs:32-43`
2. App delegates to `RunRules` use case.
   - `crates/docir-app/src/lib.rs:261-264`
   - `crates/docir-app/src/use_cases.rs:115-123`
3. Rule engine builds context and evaluates enabled rules.
   - `crates/docir-rules/src/engine.rs:109-131`

### 3.4 Diff Flow

1. CLI dispatches `diff` command.
   - `crates/docir-cli/src/commands/dispatch.rs:65-70`
2. App computes diff from parsed stores.
   - `crates/docir-app/src/lib.rs:266-269`
   - `crates/docir-app/src/use_cases.rs:126-131`
3. Diff engine indexes both IR trees and compares signatures.
   - `crates/docir-diff/src/lib.rs:70-126`

## 4. Dependency and Boundary Assessment

### 4.1 Strong Architectural Decisions

- Clear workspace modularization by concern (`core`, `parser`, `security`, `rules`, `serialization`, `diff`, adapters/interfaces).
- App-layer ports + adapters provide a stable orchestration boundary:
  - Ports in `crates/docir-app/src/lib.rs:99-147`
  - Adapters in `crates/docir-app/src/adapters.rs:18-68`
- Core model is reusable and shared by all processing crates.

### 4.2 Boundary Blurs / Architectural Risks

1. `docir-app` acts as both application facade and adapter composition root.
   - Same crate defines ports/use-cases and concrete defaults (`docir_parser`, `docir_security`, `docir_rules`, `docir_serialization`) in `crates/docir-app/src/adapters.rs:10-14`.
   - This is pragmatic but couples app crate to infrastructure crate churn.

2. `docir-python` partially bypasses app abstraction by directly using `docir_core::query::Query` and `docir_serialization::JsonSerializer`.
   - `crates/docir-python/src/lib.rs:5-12`
   - Not a hard violation, but it creates a second orchestration style besides `DocirApp`.

3. Parser crate is very broad (format parsing + security scanning helpers + diagnostics + zip hardening) and is the primary complexity hotspot.
   - Entrypoint aggregation in `crates/docir-parser/src/lib.rs:6-25`
   - Large subtrees under `ooxml/`, `odf/`, `hwp/`, `rtf/`, and `parser/`.

## 5. Scaling Notes (2x-5x complexity)

What likely breaks first:

1. `docir-parser` maintainability due to breadth and large format-specific modules.
2. Inconsistent orchestration across interfaces (CLI via app facade, Python partly direct-to-core/services).
3. Application-level coupling to concrete adapters in a single crate (`docir-app`).

What scales well already:

1. Core IR + visitor/query abstractions (`docir-core`) as a stable contract.
2. Independent service crates (`rules`, `diff`, `serialization`, `security`) operating on `IrStore` + `NodeId`.
3. Command routing segmentation in CLI (`crates/docir-cli/src/commands/*.rs`).

## 6. Concise Architecture Diagram

```text
[docir-cli] ---------\
                      \             +--> [docir-security]
[docir-python] ----> [docir-app] ---+--> [docir-rules]
                       |             +--> [docir-diff]
                       |             +--> [docir-serialization]
                       +-------------> [docir-parser]
                                      \
                                       +--> [docir-core]  (shared model root)

[docir-parser] -----> [docir-core]
[docir-security] ---> [docir-core]
[docir-rules] ------> [docir-core]
[docir-diff] -------> [docir-core]
[docir-serialization]-> [docir-core]
```

