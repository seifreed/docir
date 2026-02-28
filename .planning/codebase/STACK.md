# Technology Stack Audit

## Scope
- Repository: `docir` Rust workspace
- Focus: implementation stack, tooling stack, and delivery surfaces
- Evidence model: direct references to repository files with line anchors

## 1) Primary Language and Runtime
- Language: Rust 2021 edition
  - Evidence: `Cargo.toml:17`
- Workspace layout: multi-crate monorepo
  - Evidence: `Cargo.toml:1-13`
- Licensing/repo metadata (project-level packaging metadata)
  - Evidence: `Cargo.toml:19-21`

## 2) Workspace Composition (Internal Modules)
The codebase is split into focused crates:

| Crate | Role | Evidence |
|---|---|---|
| `docir-core` | IR model, types, query/visitor primitives | `crates/docir-core/Cargo.toml:1-16`, `crates/docir-core/src/types.rs:127-207` |
| `docir-parser` | Format detection + parsing for OOXML/ODF/HWP/HWPX/RTF | `crates/docir-parser/Cargo.toml:1-25`, `crates/docir-parser/src/parser/document.rs:24-114` |
| `docir-security` | Threat analysis and indicator logic | `crates/docir-security/Cargo.toml:1-13` |
| `docir-rules` | Rule engine and profiles | `crates/docir-rules/Cargo.toml:1-14`, `crates/docir-rules/src/lib.rs:5-10` |
| `docir-diff` | Structural diff and hashing support | `crates/docir-diff/Cargo.toml:1-16` |
| `docir-serialization` | JSON serialization for IR | `crates/docir-serialization/Cargo.toml:1-12` |
| `docir-app` | Application orchestration/use-case facade and ports | `crates/docir-app/src/lib.rs:98-156` |
| `docir-cli` | Command-line entrypoint and command handlers | `crates/docir-cli/src/main.rs:15-20`, `crates/docir-cli/src/cli.rs:6-320` |
| `docir-python` | PyO3-based Python bindings (`cdylib`) | `crates/docir-python/Cargo.toml:9-22`, `crates/docir-python/src/lib.rs:157-164` |

## 3) Dependency Stack by Concern

### 3.1 Core/Data Modeling
- `serde`, `serde_json`, `thiserror`
  - Evidence: `Cargo.toml:25-27`
- `docir-core` optionally gates `serde` feature (enabled by default)
  - Evidence: `crates/docir-core/Cargo.toml:8-15`

### 3.2 Parsing and Document Formats
- Archive/container handling: `zip`
- XML parsing: `quick-xml`
- Text encoding: `encoding_rs`
- Spreadsheet binary parsing: `calamine` (used for XLSB)
  - Evidence: `Cargo.toml:30-33`, `crates/docir-parser/src/parser/parser_xlsx.rs:26-37`

### 3.3 Security/Cryptography
- Hashing: `sha2`, `sha1`
- KDF: `pbkdf2`
- Encoding: `base64`
- Ciphers: `aes`, `cbc`
  - Evidence: `Cargo.toml:36-41`
- Concrete usage:
  - ODF encrypted part decryption: `crates/docir-parser/src/odf/container.rs:231-260`
  - HWP stream decryption: `crates/docir-parser/src/hwp/io.rs:124-146`

### 3.4 CLI/Operational
- CLI parser: `clap` derive
- Error handling: `anyhow`
- Logging: `log` + `env_logger`
  - Evidence: `Cargo.toml:44-49`, `crates/docir-cli/src/main.rs:16`, `crates/docir-cli/src/cli.rs:3`

### 3.5 Python Surface
- `pyo3` with `extension-module` and `abi3-py37`
- Crate type: `cdylib`
  - Evidence: `crates/docir-python/Cargo.toml:11,22`

## 4) Supported Data/Document Standards in Code
- Office OOXML: Word/Excel/PowerPoint
- ODF: text/spreadsheet/presentation
- Hangul: HWP and HWPX
- RTF
  - Evidence: `crates/docir-core/src/types.rs:130-149`
- Parser dispatch based on file signatures/container markers (RTF, OLE/CFB, ZIP)
  - Evidence: `crates/docir-parser/src/parser/document.rs:43-83`

## 5) Exposed Product Surfaces
- CLI binary: `docir`
  - Evidence: `crates/docir-cli/Cargo.toml:8-10`
- Python module: `docir` (functions: `parse_json`, `rules`, `query`, `summary`)
  - Evidence: `crates/docir-python/src/lib.rs:18-129`, `crates/docir-python/src/lib.rs:157-164`
- Library workspace crates for embedding/integration
  - Evidence: `Cargo.toml:3-13`

## 6) Build, Package, and Distribution Signals
- Cargo workspace as canonical build system
  - Evidence: root and per-crate `Cargo.toml` files
- Rust dependency resolution via crates.io index lock entries
  - Evidence: `Cargo.lock` (`registry+https://github.com/rust-lang/crates.io-index` entries)
- Public package/repo references in docs
  - Evidence: `README.md:12-15`, `README.md:231`

## 7) Technology Boundaries (Current State)
- No runtime HTTP client stack detected (`reqwest`, `hyper`, `tokio` absent in manifests/imports)
- No database driver stack detected (`sqlx`, `rusqlite`, `postgres`, etc. absent)
- Architecture is primarily local file parsing + in-process analysis + optional Python FFI
  - Evidence: dependency and import scans across `Cargo.toml` and `crates/**/src`

## 8) Practical Stack Summary
- Core stack: Rust workspace + typed IR (`docir-core`) + parser/security/rules/diff modules
- IO model: local file/reader input, JSON/text output
- Security model: static analysis with cryptographic helpers for archive/encrypted format handling
- Integration surface: CLI and Python extension, no active network backends
